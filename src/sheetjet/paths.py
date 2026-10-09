"""Workspace path policy for agent-facing command-line entry points."""

from pathlib import Path

from .errors import SheetJetError

_DEVICES = {"CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$"} | {
    prefix + digit for prefix in ("COM", "LPT") for digit in "123456789¹²³"
}


def _valid_filename(part):
    return (
        ":" not in part
        and part.rstrip(" .") == part
        and part.split(".", 1)[0].upper() not in _DEVICES
    )


def _resolve_path(candidate):
    """Resolve existing ancestors strictly, allowing only absent trailing components."""
    suffix = []
    while True:
        try:
            return candidate.resolve(strict=True).joinpath(*reversed(suffix))
        except FileNotFoundError:
            try:
                candidate.lstat()
            except FileNotFoundError:
                if candidate == candidate.parent:
                    raise
                suffix.append(candidate.name)
                candidate = candidate.parent
            else:
                raise SheetJetError("Cannot resolve an existing path or broken link") from None


def workspace_path(value, workspace):
    """Resolve a path under a caller-selected root, including existing symlinks.

    This validates paths, not concurrent filesystem mutation. Untrusted processes
    must additionally run in an OS sandbox with a fixed workspace and permissions.
    """
    root = Path(workspace).resolve(strict=True)
    if not root.is_dir():
        raise SheetJetError("Workspace must be an existing directory")
    candidate = root / Path(value)
    names = [part for part in candidate.parts if part not in {candidate.anchor, ".", ".."}]
    if not all(_valid_filename(part) for part in names):
        raise SheetJetError("Path contains an unsupported filename")
    candidate = _resolve_path(candidate)
    if not candidate.is_relative_to(root):
        raise SheetJetError("Path is outside the workspace; select the intended --workspace")
    return candidate
