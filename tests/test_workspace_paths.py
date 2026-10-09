import json
from pathlib import Path

import pytest

from benchmarks.text_queries import run
from sheetjet import SheetJetError
from sheetjet.cli import main
from sheetjet.paths import workspace_path


def test_resolves_relative_and_absolute_paths_inside_workspace(tmp_path):
    nested = tmp_path / "data" / "new.xlsx"
    assert workspace_path("data/new.xlsx", tmp_path) == nested
    assert workspace_path(nested, tmp_path) == nested


@pytest.mark.parametrize("value", ["../secret.json", "../workspace-other/secret.json"])
def test_parent_traversal_and_sibling_prefix_are_rejected(tmp_path, value):
    root = tmp_path / "workspace"
    root.mkdir()
    with pytest.raises(SheetJetError, match="outside the workspace"):
        workspace_path(value, root)


def test_absolute_path_outside_workspace_is_rejected(tmp_path):
    root = tmp_path / "workspace"
    root.mkdir()
    outside = tmp_path / "secret.json"
    with pytest.raises(SheetJetError, match="outside the workspace"):
        workspace_path(outside, root)


@pytest.mark.parametrize("name", ["NUL", "con.txt", "data.json:stream", "report.", "LPT1.xlsx"])
def test_non_file_windows_paths_are_rejected(tmp_path, name):
    with pytest.raises(SheetJetError, match="unsupported filename"):
        workspace_path(name, tmp_path)


def test_symlink_escapes_are_rejected(tmp_path):
    root = tmp_path / "workspace"
    outside = tmp_path / "outside"
    root.mkdir()
    outside.mkdir()
    try:
        (root / "link").symlink_to(outside, target_is_directory=True)
    except OSError as error:
        if getattr(error, "winerror", None) == 1314:
            pytest.skip("Windows requires symbolic-link privileges")
        raise
    with pytest.raises(SheetJetError, match="outside the workspace"):
        workspace_path("link/secret.json", root)


def test_workspace_must_be_a_directory(tmp_path):
    file = tmp_path / "file"
    file.write_text("test")
    with pytest.raises(SheetJetError, match="directory"):
        workspace_path("data.xlsx", file)


@pytest.mark.parametrize("argument", ["--spec", "--output", "--cache-dir"])
def test_cli_rejects_escaped_paths_before_opening_workbooks(tmp_path, capsys, argument):
    root = tmp_path / "workspace"
    root.mkdir()
    args = ["--workspace", str(root)]
    if argument == "--cache-dir":
        args += [argument, "../outside"]
    args += ["reconcile", "before.xlsx", "after.xlsx", "--spec", "spec.json"]
    if argument != "--cache-dir":
        args += [argument, "../outside"]
    assert main(args) == 2
    assert "outside the workspace" in json.loads(capsys.readouterr().err)["message"]


def test_default_cli_workspace_is_current_directory(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)
    assert main(["inspect", "../outside.xlsx"]) == 2
    assert "outside the workspace" in json.loads(capsys.readouterr().err)["message"]


def test_spec_cannot_inject_execution_options(tmp_path, capsys):
    (tmp_path / "spec.json").write_text(json.dumps({"overwrite": True}))
    assert (
        main(
            [
                "--workspace",
                str(tmp_path),
                "reconcile",
                "before.xlsx",
                "after.xlsx",
                "--spec",
                "spec.json",
            ]
        )
        == 2
    )
    assert "spec must contain only" in json.loads(capsys.readouterr().err)["message"]


def test_benchmark_rejects_outside_output_and_input_overwrite(tmp_path):
    source = tmp_path / "book.xlsx"
    source.write_bytes(b"protected input")
    outside = Path("../outside.json")
    with pytest.raises(SheetJetError, match="outside the workspace"):
        run(source, outside, workspace=tmp_path)
    with pytest.raises(ValueError, match="replace its input"):
        run(source, source, workspace=tmp_path)
    assert source.read_bytes() == b"protected input"
