# Frozen independent skill submissions

These are original task builders authored by three isolated, same-model agents on
5 October 2026. Each agent received the same four tasks and inputs and its assigned
skill. They could inspect their own outputs and repair them before freezing.
They did not receive competitors' implementations or timing feedback. One SheetJet
cache-directory correction enforced fresh per-trial data caches before measurement.
The evaluation is one builder-generation session per skill, not a distribution of
agent reasoning attempts. Execution repeats do not increase that sample size.

The archived builder files are byte-identical to the measured submissions. Formatting
them would change their evidence hashes, so this directory is excluded from Ruff.
Peer SKILL.md files and proprietary toolkit code are **not** redistributed here.

| Assigned skill | Source/version | SKILL.md SHA-256 |
|---|---|---|
| SheetJet | repository skill before the v0.4 documentation addition | `116134efda4509899edc3b24ac3168ed7f729afe3286787766962aacd8bbd218` |
| Anthropic XLSX | [official skills repository](https://github.com/anthropics/skills/blob/8a1541c4a3ffa5a20a5a91de0dcf3f0bab1d1ef4/skills/xlsx/SKILL.md), commit `8a1541c4a3ffa5a20a5a91de0dcf3f0bab1d1ef4` | `111e06521941d60feea6c160d9f67a0b9b95a383d3db0b4c95e65937eb12fb5a` |
| OpenAI Spreadsheets | official installed bundle `26.921.10847`, `skills/Spreadsheets/SKILL.md` | `d863237ee6ca690f52146fed9de1a8c8f18aca942b61c23d2a4d798a462535b8` |

All agents inherited the parent model without an override. Exact serving revision,
model-token telemetry, and reasoning-time comparisons were unavailable.

## What was actually executed

- **SheetJet:** SheetJet inspection/query/patch APIs for existing workbooks;
  openpyxl fallback for creating the new workbook and chart.
- **Anthropic:** openpyxl streaming reads and original targeted ZIP/XML edits;
  openpyxl for the new workbook. This is a preservation-aware implementation,
  not a naive load/save proxy.
- **OpenAI:** artifact-tool authors the output cells/new workbook, with an original
  Python standard-library adapter to preserve existing packages and avoid formula
  recalculation. See [the agent's implementation notes](openai/SUBMISSION_NOTES.md).
  Rendering was checked separately during authoring and is outside timed execution.
  The installed renderer wrote previews but exited with status 1 without diagnostics;
  this limitation is not counted as an execution failure of the frozen builders.

## Reproduce the execution comparison

Install SheetJet and the evaluator in the Python interpreter used below:

```console
python -m pip install -e ".[test,benchmark]"
python -m benchmarks.skill_trials --prepare benchmark-output/skill-inputs --rows 25000
python -m benchmarks.bind_skill_trials benchmark-output/skill-builders
python -m benchmarks.skill_trials --run benchmark-output/skill-inputs --registry sheetjet=benchmark-output/skill-builders/sheetjet/commands.json --registry anthropic=benchmark-output/skill-builders/anthropic/commands.json --output benchmark-output/skill-results --repeats 3
```

Use exactly 25,000 rows per period with these frozen submissions: the SheetJet agent
bound its selected ranges to that task. The fixture generator's other sizes are for
new agent trials, not silently interchangeable with this submission.

The OpenAI peer additionally requires access to its installed bundled artifact
runtime. It is not installed by pip. Pass `--node ABSOLUTE_NODE_EXECUTABLE` and
`--artifact-python ABSOLUTE_BUNDLED_PYTHON` to the binder, and expose that runtime's
`node_modules` through a directory link at `skill-builders/openai/node_modules`.
For example, use `New-Item -ItemType Junction` on Windows or `ln -s` on Unix with
the actual dependency paths returned by the runtime loader. Then include
`--registry openai=benchmark-output/skill-builders/openai/commands.json`.
The binder changes only the copied OpenAI builder's Python executable binding;
the runner records the resulting local hashes. It never modifies the frozen archive.

Every run needs a new output directory. Timed subprocesses run in isolated output
directories with fresh SheetJet data caches. Inputs are hashed before and after
every trial. Registries are trusted local executable configurations: inspect them
before running. The task prompt is saved as `skill-inputs/tasks.json`.

The JSON report includes every timing, return code, pass/fail check, measured
process-tree RSS peak, input hash, and submission hash. Independent artifact checks
run after the timing ends. Reported wall times include imports and the builder's
own validation. OS file caches are not flushed. Memory sampling every 10 ms may
miss brief peaks and double-count shared pages in multi-process workflows.

See [the results and limits](../../docs/skill-workflows-v0.4.md).
