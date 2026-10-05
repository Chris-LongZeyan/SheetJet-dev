"""Copy frozen submissions and bind machine-specific executable paths."""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from pathlib import Path


def bind(destination, python, node=None, artifact_python=None):
    if bool(node) != bool(artifact_python):
        raise ValueError("Provide both --node and --artifact-python for the OpenAI submission")
    source = Path(__file__).parent / "submissions"
    peers = ["sheetjet", "anthropic"] + (["openai"] if node else [])
    for peer in peers:
        target = destination.resolve() / peer
        target.mkdir(parents=True, exist_ok=False)
        for file in (source / peer).iterdir():
            if file.suffix in {".py", ".mjs"}:
                shutil.copyfile(file, target / file.name)
        builder = target / ("builder.mjs" if peer == "openai" else "builder.py")
        if peer == "openai":
            original = builder.read_text(encoding="utf-8")
            lines = original.splitlines(keepends=True)
            bindings = [i for i, line in enumerate(lines) if line.startswith("const python=")]
            if len(bindings) != 1:
                raise ValueError("Unexpected frozen runtime binding")
            lines[bindings[0]] = (
                "const python=" + json.dumps(str(artifact_python.resolve())) + ";\n"
            )
            builder.write_text("".join(lines), encoding="utf-8", newline="\n")
        registry = {
            case: {
                "argv": [
                    str((node if peer == "openai" else python).resolve()),
                    str(builder),
                    case,
                    "{input}",
                    "{output}",
                    "{result}",
                ]
            }
            for case in ["edit", "aggregate", "wide", "create"]
        }
        path = target / "commands.json"
        path.write_text(json.dumps(registry, indent=2), encoding="utf-8")
        print(f"--registry {peer}={path}")
    if node:
        print("Link the bundled node_modules into the openai submission directory before running.")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("destination", type=Path)
    parser.add_argument("--python", type=Path, default=Path(sys.executable))
    parser.add_argument("--node", type=Path)
    parser.add_argument("--artifact-python", type=Path)
    args = parser.parse_args()
    bind(args.destination, args.python, args.node, args.artifact_python)


if __name__ == "__main__":
    main()
