# Submission implementation and verification

The assigned OpenAI Spreadsheets skill (26.921.10847) and its API quick start, editing/creation workflows, style, chart, clarification, and finance guidance were read. Task instructions override conflicting default recalculation and layout guidance.

`commands.json` is the executable registry. All four entries call `builder.mjs` with task, absolute input path, absolute output XLSX path, and absolute result JSON path. The create command ignores its input argument. The builders have no shared workbook-derived cache. Temporary authored workbooks, inspector output, and runtime temporary files reside beneath each fresh output's parent. A private submission-level initial TEMP/TMP is supplied for runtime startup.

## Dependencies

- Bundled Node: `C:/Users/Chris_longzeyan/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node.exe`
- Bundled Python: `C:/Users/Chris_longzeyan/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe`
- `@oai/artifact-tool`, resolved through the submission's `node_modules` junction to the loader's bundled `dependencies/node/node_modules`.
- Python uses only its standard library. No SheetJet, openpyxl, xlsxwriter, pandas writer, copied toolkit internals, or external downloads are used.

The Windows junction initially required an approved sandbox escalation; it was created successfully.

## Explicit preservation fallback

For edit, aggregate, and wide, the original `package_adapter.py` extracts only task-relevant data with stdlib ZIP/XML parsing. The artifact tool authors each target numeric value into a small temporary workbook, exports it, and the adapter reads those actual exported values. It transplants their numeric XML payloads into the requested existing target cells, leaving source cell formatting and every other package entry unchanged. The adapter verifies that every unrelated ZIP entry's uncompressed content is byte-identical. It preserves original formula text, caches, custom XML, styles, sheet/hidden states, names, and external parts without evaluating or refreshing them.

This fallback addresses capabilities absent from the documented artifact authoring API: surgical package preservation and writing without automatic dependent-formula recalculation. The API quick start explicitly says value writes automatically recalculate dependents; a single bounded `SpreadsheetFile.exportXlsx` help query returned no entries. The original full workbook is therefore not mutated by the artifact tool. This is a preservation adapter, not a replacement spreadsheet authoring engine.

For create, artifact-tool alone creates all cells, formulas, formatting, frozen panes, the native chart, recalculated values, and the final XLSX. No post-export package edits occur.

## Tests and observations

All four commands completed successfully with valid outputs and compact result JSON. Create was retried after isolating a rendering-only runtime issue; no task logic changed during that repair. Source and changed views were imported, inspected, rendered with artifact-tool, and visually reviewed. The existing Summary source layout has narrow columns; it was preserved as explicitly requested. The new Budget view is readable and the native column chart does not cover the table.

`audit_saved.py` independently checks saved XML/packages. It confirmed that only allowed numeric cells changed, every other package entry and non-value target-cell detail was preserved, and create contains the exact requested values/formulas, bold headers, two-decimal formats, frozen header, one column chart (`barDir=col`) bound to Jan–Mar Revenue/Cost, excluding totals. The edit leaves the existing formula cache at 8 and reports recalculation needed.

Verified snapshots: APAC 6,237,500; EMEA 6,250,000; Americas 6,262,500; Other 6,275,000 from Period01 and Period03. Wide APAC Amount is 1,260 from Region A2:A41 and Amount CRD2:CRD41. Edit changes Assumptions!B1 from 0.08 to 0.12.

Rendering limitation: this installed bundled runtime writes valid PNG previews and returns requested inspection results, but a process that calls `workbook.render` terminates with status 1 even after all awaits and successful file writes, without a diagnostic. Rendering was isolated in `verify_views.mjs` and completed during authoring QA; frozen builders do not call render and return status 0. This is recorded as an environment/runtime limitation, not a skill failure. Excel itself was not available as a verification engine; package-level native chart/formula checks and artifact recalculation/rendering were used.

The first operation marker was run successfully exactly once before spreadsheet authoring, with edit operation, expected output count 4, format xlsx.

Builders and registry are frozen after these checks. No benchmark results or comparisons were inspected.
