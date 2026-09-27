Build an extremely high-performance Excel-processing skill intended for large, complex, real-world workbooks.

The primary goals, in priority order, are:

1. Minimize LLM token usage.
2. Minimize workbook processing time.
3. Handle very large and complex Excel files reliably.
4. Preserve workbook formulas, formatting, structure, and semantics whenever possible.
5. Avoid loading or exposing irrelevant workbook content to the language model.
6. Make spreadsheet operations deterministic where deterministic code is sufficient, and use LLM reasoning only where semantic interpretation is actually required.

## Core architectural principle

Do NOT treat an Excel workbook as text to be read by the LLM.

Treat the workbook as a structured data source that should first be inspected, indexed, queried, modified, and validated using code.

Use this pipeline:

Workbook
→ structural inspection
→ compact workbook index
→ task interpretation
→ targeted range discovery
→ minimal data extraction
→ deterministic computation/edit
→ local validation
→ concise result returned to the LLM

The LLM should see only the minimum information necessary to make a decision.

## Phase 1: Fast workbook reconnaissance

When receiving an Excel workbook, first create a compact workbook map without loading unnecessary cell contents.

Collect metadata such as:

- workbook filename/type
- sheet names
- visible/hidden state
- used range of each sheet
- approximate row/column count
- tables
- table ranges
- named ranges
- merged ranges
- freeze panes
- formula count
- non-empty cell count where cheap to obtain
- likely header rows
- likely data regions
- cells containing formulas
- external links
- charts/pivots where detectable
- hidden rows/columns
- workbook calculation settings

Create a compact representation similar to:

{
  "sheets": {
    "Revenue": {
      "range": "A1:AZ52000",
      "headers": ["Region", "Product", "FY25", "FY26"],
      "tables": [...],
      "formula_regions": [...],
      "likely_data_regions": [...]
    }
  }
}

Do NOT send the entire workbook map to the LLM if it is large. Query the index as needed.

## Phase 2: Build a lightweight searchable workbook index

Create a local index allowing fast lookup of:

- sheet names
- table names
- column headers
- row labels
- important text cells
- named ranges
- formulas
- formula dependencies where practical
- keywords and their locations

For example, a request mentioning:

"China revenue growth assumption"

should allow the system to search the workbook index and return something like:

Possible matches:
- Assumptions!B42:G48
- Revenue_Model!D11:M22
- China Forecast table

rather than reading every worksheet.

Use lexical matching first and semantic matching only when useful.

## Phase 3: Query planning before reading cells

Before accessing large ranges, create a small internal execution plan.

Example:

Task:
"Calculate 2026 EBITDA margin for APAC."

Plan:
1. Find APAC revenue.
2. Find APAC EBITDA.
3. Determine whether values already exist or require calculation.
4. Read only required ranges.
5. Calculate result programmatically.
6. Validate units and period.

Never fetch an entire sheet merely because the required information is somewhere on that sheet.

## Phase 4: Progressive disclosure

Use progressive data retrieval.

Level 0:
Workbook metadata only.

Level 1:
Headers and labels.

Level 2:
Small candidate ranges.

Level 3:
Expanded range around the relevant data.

Level 4:
Large-range extraction only when unavoidable.

The agent should always attempt the cheapest level first.

## Phase 5: Large-table processing

For large datasets, never send rows individually to the LLM unless semantic reasoning is required.

Prefer:

- Polars
- DuckDB
- pandas where appropriate
- Arrow
- vectorized NumPy operations

Use filters, aggregations, joins, sorting, grouping, window functions, and calculations programmatically.

For example:

Instead of sending 200,000 transaction rows to the LLM and asking it to identify the top customers:

Run:

GROUP BY customer
SUM(revenue)
ORDER BY revenue DESC
LIMIT 20

Then provide only the resulting 20 rows to the model if interpretation is needed.

## Phase 6: Efficient Excel libraries

Select tools according to the operation.

Possible choices:

openpyxl:
- editing XLSX/XLSM structures
- formulas
- styles
- merged cells
- workbook metadata

Polars / DuckDB:
- high-speed tabular analysis
- large datasets
- filtering
- aggregation
- joins

pandas:
- compatibility or operations not convenient elsewhere

pyxlsb:
- XLSB reading where needed

Use read-only/write-only modes where useful.

Avoid repeatedly reopening the workbook.

Avoid cell-by-cell loops where vectorized or range-based operations are possible.

Batch reads and writes whenever possible.

## Phase 7: Formula awareness

Treat formulas differently from values.

When the task requires formula analysis:

- identify formula cells
- identify repeated formula patterns
- normalize formulas where useful
- locate precedent/dependent ranges when feasible
- detect suspicious deviations in formula patterns

Do NOT expand every formula into verbose natural language.

For repeated formulas, represent them compactly.

Example:

Instead of storing 10,000 formulas individually:

Rows 12–10011:
Column H pattern = F[row] * G[row]

Record the pattern once.

## Phase 8: Formatting awareness

Formatting can contain business meaning.

Preserve:

- number formats
- fonts
- fills
- borders
- alignment
- row heights
- column widths
- conditional formatting
- merged cells

Do not serialize every style into the LLM context.

Represent styles using IDs or signatures.

Example:

style_7 = {
  number_format: "$#,##0.0",
  font: "Arial 9 bold",
  fill: "003C8C"
}

Cells should reference style_7 rather than repeating the complete style specification.

## Phase 9: Editing strategy

For edits, use a patch-based approach.

Represent modifications internally as operations such as:

SET_VALUE
SET_FORMULA
COPY_FORMULA
INSERT_ROWS
DELETE_ROWS
SET_STYLE
COPY_STYLE
CREATE_TABLE
RESIZE_TABLE
RENAME_SHEET

Example:

[
  {
    "operation": "SET_VALUE",
    "sheet": "Assumptions",
    "cell": "F23",
    "value": 0.08
  }
]

Apply patches in batches.

Do not reconstruct the workbook unnecessarily.

## Phase 10: Minimal-diff philosophy

When modifying a workbook:

- change only requested cells/regions
- preserve untouched formulas
- preserve styles
- preserve worksheets
- preserve names
- preserve workbook metadata
- preserve VBA content when possible
- avoid accidental formatting drift

After editing, compare the modified workbook with the original at the structural level and confirm that unexpected regions were not modified.

## Phase 11: Validation

Validation should be proportional to the modification.

Do NOT reread an entire 100 MB workbook after changing one cell.

Instead validate:

- changed cells
- nearby region
- affected formulas
- relevant totals
- workbook integrity

Examples:

If a table is modified:
- verify row count
- verify formulas
- verify totals
- verify styles
- verify table range

If a formula is changed:
- verify the formula
- check neighboring formulas for consistency
- inspect dependent outputs where practical

## Phase 12: Token-budget awareness

Implement explicit token-saving behavior.

Before returning spreadsheet data to the LLM:

1. Remove irrelevant columns.
2. Remove irrelevant rows.
3. Aggregate where possible.
4. Replace repetitive structures with patterns.
5. Compress style metadata.
6. Summarize repeated formulas.
7. Prefer coordinates and references over repeated values.
8. Limit preview rows.
9. Use statistics rather than raw observations when sufficient.

Example:

Bad:

Send 50,000 rows.

Better:

Rows: 50,000
Revenue:
  mean: ...
  median: ...
  min: ...
  max: ...
Top anomalies:
  [10 rows]

Only retrieve raw rows if subsequent reasoning requires them.

## Phase 13: Separate reasoning from execution

The LLM should decide:

- what the user means
- what information is needed
- which candidate region is semantically correct
- how ambiguous labels should be interpreted
- whether output satisfies user intent

Code should handle:

- reading cells
- writing cells
- sorting
- filtering
- calculations
- joins
- aggregation
- copying formulas
- formatting
- workbook inspection
- comparison
- validation

Never use LLM tokens for arithmetic or repetitive spreadsheet manipulation that can be performed deterministically.

## Phase 14: Caching

Cache reusable information during the session:

- workbook index
- sheet metadata
- detected headers
- table locations
- formula maps
- style definitions
- already-read ranges

If the next request concerns the same workbook, do not rebuild or resend unchanged information.

Invalidate only affected cache entries after modifications.

## Phase 15: Range-first APIs

Design internal functions around ranges and tables rather than individual cells.

Examples:

inspect_workbook()
inspect_sheet(sheet)
find_text(query)
find_table(query)
read_range(sheet, range)
read_table(table_name, columns=None, filters=None)
write_range(...)
patch_cells(...)
get_formulas(range)
get_dependencies(range)
get_styles(range)
profile_table(...)
validate_range(...)

Avoid APIs that encourage thousands of calls like:

get_cell(A1)
get_cell(A2)
get_cell(A3)
...

## Phase 16: Adaptive strategy based on workbook size

Small workbook:
Full inspection may be acceptable.

Medium workbook:
Metadata + targeted extraction.

Large workbook:
Lazy loading + indexing + query execution.

Huge workbook:
Streaming/chunking, DuckDB/Polars processing, minimal Excel-object interaction.

The system should automatically choose a strategy based on:

- file size
- worksheet dimensions
- cell count
- formula density
- formatting complexity
- task type

## Phase 17: Performance instrumentation

Record useful performance metrics during development:

- workbook open time
- indexing time
- cells inspected
- cells sent to LLM
- approximate tokens consumed
- number of range reads
- number of write operations
- validation time
- total execution time
- peak memory usage

The objective is not only correctness but efficiency.

## Phase 18: Benchmark suite

Create stress-test workbooks including:

1. 1M-row transaction dataset.
2. Multi-sheet financial model.
3. Formula-heavy workbook.
4. Formatting-heavy workbook.
5. Workbook with hidden sheets and named ranges.
6. Workbook with merged cells.
7. Workbook containing pivot tables/charts.
8. Messy analyst-created spreadsheet with multiple tables on one sheet.
9. Workbook with 30–100 interconnected sheets.
10. XLSM workbook where VBA must remain intact.

Benchmark common tasks:

- locate a value
- summarize a table
- change assumptions
- update formulas
- detect errors
- add columns
- apply formatting
- reconcile two sheets
- perform financial analysis
- create charts/tables

Compare:

- runtime
- token consumption
- memory
- correctness
- workbook fidelity

against a naive implementation that loads workbook contents directly into model context.

## Phase 19: Safety / reliability

Never silently modify ambiguous cells.

If several candidate regions exist, gather enough metadata to distinguish them.

Maintain an internal change log.

Before saving, verify that the output workbook opens successfully and has the expected worksheet structure.

Prefer producing an explicit error over corrupting a workbook.

## Main optimization target

The most important design objective is:

MAXIMIZE:
useful reasoning performed per LLM token

MINIMIZE:
raw Excel content exposed to the LLM

The ideal system should be capable of processing a workbook containing millions of cells while the language model sees only tens or hundreds of relevant cells.

Do not optimize merely by shortening prompts. Optimize the entire spreadsheet-processing architecture so that unnecessary information never enters the model context in the first place.

## Deliverables

Design and implement:

1. The Excel skill architecture.
2. Its SKILL.md/instruction file.
3. Supporting Python modules.
4. Workbook inspection/indexing utilities.
5. Large-table query engine.
6. Patch-based editing engine.
7. Validation utilities.
8. Token-efficient serialization format.
9. Performance benchmark suite.
10. Example tasks and regression tests.

Prioritize correctness, speed, minimal token consumption, and preservation of complex Excel files.

Where multiple technical approaches exist, benchmark them instead of assuming one is faster.

The final implementation should behave more like a database query engine and spreadsheet compiler than an AI that reads Excel spreadsheets line by line.