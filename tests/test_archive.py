import io
import struct
from types import SimpleNamespace
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

import pytest

from sheetjet.archive import MAX32, copy_archive, directory, relocate
from sheetjet.core import Metrics


class NonSeekable(io.BytesIO):
    def seekable(self):
        return False

    def seek(self, *args):
        raise io.UnsupportedOperation("not seekable")


@pytest.mark.parametrize("descriptor", [False, True])
@pytest.mark.parametrize("zip64", [False, True])
def test_compressed_records_roundtrip(tmp_path, monkeypatch, descriptor, zip64):
    import zipfile

    source = tmp_path / "source.zip"
    sink = NonSeekable() if descriptor else io.BytesIO()
    # A low threshold exercises real ZIP64 central records without multi-GB fixtures.
    with monkeypatch.context() as patch:
        if zip64:
            patch.setattr(zipfile, "ZIP64_LIMIT", 100)
        with ZipFile(sink, "w", compression=ZIP_DEFLATED) as z:
            z.comment = b"archive comment retained verbatim"
            info = ZipInfo("opaque/data.bin")
            info.compress_type = ZIP_DEFLATED
            info.comment = b"member comment"
            z.writestr(info, bytes(range(256)) * 100)
            z.writestr("change.xml", b"<a>old</a>")
            z.writestr("remove.xml", b"remove")
            z.writestr("Unicode/\u4e2d\u6587.txt", b"unchanged")
    source.write_bytes(sink.getvalue())
    target = tmp_path / "target.zip"
    with ZipFile(source) as z:
        package = SimpleNamespace(
            path=source, zip=z, parts={i.filename: i for i in z.infolist()}, metrics=Metrics()
        )
        evidence = copy_archive(
            package, target, {"change.xml": b"<a>new</a>"}, {"remove.xml"}, tmp_path
        )
        _, a_spans = directory(source, z.infolist())
        with ZipFile(target) as out:
            assert out.testzip() is None
            assert out.comment == z.comment
            assert out.read("change.xml") == b"<a>new</a>"
            assert "remove.xml" not in out.namelist()
            _, b_spans = directory(target, out.infolist())
            for record in evidence:
                name = record["part"]
                assert out.read(name) == z.read(name)
                assert (
                    source.read_bytes()[slice(*a_spans[name])]
                    == target.read_bytes()[slice(*b_spans[name])]
                )
    assert len(evidence) == 2


def test_relocation_crosses_four_gib_boundary(tmp_path):
    path = tmp_path / "small.zip"
    with ZipFile(path, "w") as z:
        z.writestr("x", b"x")
    with ZipFile(path) as z:
        entries, _ = directory(path, z.infolist())
    rewritten = relocate(entries["x"], MAX32 + 200)
    assert struct.unpack_from("<I", rewritten, 42)[0] == MAX32
    n, e = struct.unpack_from("<HH", rewritten, 28)
    extra = rewritten[46 + n : 46 + n + e]
    assert struct.unpack("<HHQ", extra) == (1, 8, MAX32 + 200)
    # ZIP64 offset can be relocated again to a small archive without losing the field.
    again = relocate(rewritten, 200)
    assert struct.unpack("<HHQ", again[46 + n : 46 + n + e]) == (1, 8, 200)


def test_actual_zip64_entry_count(tmp_path):
    path = tmp_path / "many.zip"
    with ZipFile(path, "w") as z:
        for i in range(65536):
            z.writestr(str(i), b"")
    with ZipFile(path) as z:
        entries, spans = directory(path, z.infolist())
        assert len(entries) == len(spans) == 65536
