"""Copy compressed ZIP members without inflate/deflate or private zipfile APIs.

Local records (including ZIP64 headers and data descriptors) are copied verbatim.
Only central-directory offsets are relocated. Changed entries are encoded by the
standard library. Multi-disk and prefixed/SFX archives are deliberately rejected.
"""

from __future__ import annotations

import copy
import hashlib
import shutil
import struct
from pathlib import Path
from zipfile import ZipFile

from .errors import SheetJetError, UnsupportedOperation

MAX32 = 0xFFFFFFFF
MAX16 = 0xFFFF


def _exact(stream, length):
    data = stream.read(length)
    if len(data) != length:
        raise SheetJetError("Truncated ZIP structure")
    return data


def directory(path, infos):
    """Return the raw directory entries and local-record spans of a standard ZIP."""
    with Path(path).open("rb") as stream:
        stream.seek(0, 2)
        size = stream.tell()
        start = max(0, size - 65557)
        stream.seek(start)
        tail = stream.read()
        at = len(tail)
        while True:
            at = tail.rfind(b"PK\x05\x06", 0, at)
            if at < 0:
                raise SheetJetError("ZIP end record not found")
            if at + 22 <= len(tail) and at + 22 + struct.unpack_from("<H", tail, at + 20)[0] == len(
                tail
            ):
                break
        _, disk, cd_disk, on_disk, count, cd_size, offset, _ = struct.unpack_from(
            "<4s4H2IH", tail, at
        )
        if disk or cd_disk or on_disk != count:
            raise UnsupportedOperation("Multi-disk ZIP archives are unsupported")
        if count == MAX16 or cd_size == MAX32 or offset == MAX32:
            stream.seek(start + at - 20)
            sig, disk, zoffset, disks = struct.unpack("<4sIQI", _exact(stream, 20))
            if sig != b"PK\x06\x07" or disk or disks != 1:
                raise UnsupportedOperation("Invalid or multi-disk ZIP64 locator")
            stream.seek(zoffset)
            z = struct.unpack("<4sQ2H2I4Q", _exact(stream, 56))
            if z[0] != b"PK\x06\x06" or z[4] or z[5] or z[6] != z[7]:
                raise UnsupportedOperation("Invalid ZIP64 end record")
            count, cd_size, offset = z[7:10]
        if count != len(infos) or offset + cd_size > start + at:
            raise SheetJetError("Inconsistent ZIP directory")
        stream.seek(offset)
        entries = {}
        for info in infos:
            fixed = _exact(stream, 46)
            if fixed[:4] != b"PK\x01\x02":
                raise UnsupportedOperation("Nonstandard/prefixed ZIP directory layout")
            n, e, c = struct.unpack_from("<3H", fixed, 28)
            rest = _exact(stream, n + e + c)
            if (
                rest[:n].decode("utf-8" if info.flag_bits & 0x800 else "cp437")
                != info.orig_filename
            ):
                raise SheetJetError("ZIP directory filename mismatch")
            entries[info.filename] = fixed + rest
        if stream.tell() != offset + cd_size:
            raise UnsupportedOperation("Extra records in ZIP central directory are unsupported")
        ordered = sorted(infos, key=lambda i: i.header_offset)
        if ordered and ordered[0].header_offset != 0:
            raise UnsupportedOperation("Prefixed/SFX ZIP archives are unsupported")
        spans = {}
        for i, info in enumerate(ordered):
            end = ordered[i + 1].header_offset if i + 1 < len(ordered) else offset
            stream.seek(info.header_offset)
            header = _exact(stream, 30)
            if header[:4] != b"PK\x03\x04":
                raise SheetJetError("Invalid ZIP local header")
            n, e = struct.unpack_from("<2H", header, 26)
            if info.header_offset + 30 + n + e + info.compress_size > end:
                raise SheetJetError("Overlapping ZIP member records")
            spans[info.filename] = (info.header_offset, end)
        return entries, spans


def relocate(entry, offset):
    """Rewrite a local-record offset, preserving all other central-directory data."""
    fixed = bytearray(entry[:46])
    n, e, c = struct.unpack_from("<3H", fixed, 28)
    old_offset = struct.unpack_from("<I", fixed, 42)[0]
    if old_offset != MAX32 and offset < MAX32:
        struct.pack_into("<I", fixed, 42, offset)
        return bytes(fixed) + entry[46:]
    filename = entry[46 : 46 + n]
    extra = entry[46 + n : 46 + n + e]
    comment = entry[46 + n + e :]
    comp, uncomp = struct.unpack_from("<2I", fixed, 20)
    pos = (8 if uncomp == MAX32 else 0) + (8 if comp == MAX32 else 0)
    output, found = bytearray(), False
    cursor = 0
    while cursor < len(extra):
        if cursor + 4 > len(extra):
            raise SheetJetError("Malformed ZIP extra field")
        tag, length = struct.unpack_from("<HH", extra, cursor)
        data = extra[cursor + 4 : cursor + 4 + length]
        if len(data) != length:
            raise SheetJetError("Truncated ZIP extra field")
        if tag == 1:
            if found or len(data) < pos + (8 if old_offset == MAX32 else 0):
                raise SheetJetError("Invalid ZIP64 offset field")
            data = (
                data[:pos]
                + struct.pack("<Q", offset)
                + data[pos + (8 if old_offset == MAX32 else 0) :]
            )
            found = True
        output.extend(struct.pack("<HH", tag, len(data)) + data)
        cursor += 4 + length
    if not found:
        if pos or old_offset == MAX32:
            raise SheetJetError("Missing ZIP64 field")
        output.extend(struct.pack("<HHQ", 1, 8, offset))
    if len(output) > MAX16:
        raise SheetJetError("ZIP extra fields exceed format limit")
    struct.pack_into("<I", fixed, 42, MAX32)
    struct.pack_into("<H", fixed, 30, len(output))
    version = struct.unpack_from("<H", fixed, 6)[0]
    struct.pack_into("<H", fixed, 6, max(version, 45))
    return bytes(fixed) + filename + output + comment


def _end(stream, count, offset, size, comment):
    if count >= MAX16 or offset >= MAX32 or size >= MAX32:
        zoffset = stream.tell()
        stream.write(
            struct.pack("<4sQ2H2I4Q", b"PK\x06\x06", 44, 45, 45, 0, 0, count, count, size, offset)
        )
        stream.write(struct.pack("<4sIQI", b"PK\x06\x07", 0, zoffset, 1))
    stream.write(
        struct.pack(
            "<4s4H2IH",
            b"PK\x05\x06",
            0,
            0,
            min(count, MAX16),
            min(count, MAX16),
            min(size, MAX32),
            min(offset, MAX32),
            len(comment),
        )
    )
    stream.write(comment)


def copy_archive(package, output, changes, removed, directory_path):
    """Rebuild a ZIP with byte-identical untouched local records; return copy evidence."""
    encoded = Path(directory_path) / "changed-members.zip"
    with ZipFile(encoded, "w") as dst:
        for name, replacement in changes.items():
            info = copy.copy(package.parts[name])
            # Sizes/offsets are regenerated for changed members. Retaining an old ZIP64
            # field would create duplicate or stale local-header size information.
            extra = bytearray()
            cursor = 0
            while cursor < len(info.extra):
                tag, length = struct.unpack_from("<HH", info.extra, cursor)
                field = info.extra[cursor : cursor + 4 + length]
                if len(field) != 4 + length:
                    raise SheetJetError("Truncated ZIP extra field")
                if tag != 1:
                    extra.extend(field)
                cursor += 4 + length
            info.extra = bytes(extra)
            if isinstance(replacement, bytes):
                dst.writestr(info, replacement)
            else:
                with open(replacement, "rb") as src, dst.open(info, "w", force_zip64=True) as out:
                    shutil.copyfileobj(src, out, 1024 * 1024)
    entries, spans = directory(package.path, package.zip.infolist())
    with ZipFile(encoded) as z:
        changed_entries, changed_spans = directory(encoded, z.infolist())
    central, certificates = [], []
    with (
        package.path.open("rb") as src,
        encoded.open("rb") as updated,
        Path(output).open("wb") as dst,
    ):
        for info in package.zip.infolist():
            name = info.filename
            if name in removed:
                continue
            changed = name in changes
            stream = updated if changed else src
            begin, end = (changed_spans if changed else spans)[name]
            central.append(relocate((changed_entries if changed else entries)[name], dst.tell()))
            stream.seek(begin)
            remaining = end - begin
            digest = hashlib.sha256()
            while remaining:
                chunk = _exact(stream, min(remaining, 1024 * 1024))
                dst.write(chunk)
                if not changed:
                    digest.update(chunk)
                remaining -= len(chunk)
            if not changed:
                certificates.append(
                    {"part": name, "bytes": end - begin, "sha256": digest.hexdigest()}
                )
        start = dst.tell()
        for entry in central:
            dst.write(entry)
        _end(dst, len(central), start, dst.tell() - start, package.zip.comment)
    package.metrics.add("compressed_bytes_copied", sum(c["bytes"] for c in certificates))
    package.metrics.add("members_recompressed", len(changes))
    return certificates
