"""Reproducible fixtures and independent judging for agent-authored skill workflows.

Peer skills are read at their pinned sources, never vendored into this repository.
Agent-generated builders are independent submissions, not hand-written peer proxies.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import random
import statistics
import subprocess
import threading
import time
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree as ET

from .peers import fixture

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
Q = "{" + NS + "}"


def cell_fingerprints(path, changed=None):
    """Independent openpyxl decoding, excluding requested values but not their styles."""
    from openpyxl import load_workbook

    changed = changed or {}
    book = load_workbook(path, read_only=True, data_only=False)
    result = {}
    try:
        for sheet in book:
            values, styles = hashlib.sha256(), hashlib.sha256()
            style_cache = {}
            if sheet.max_row * sheet.max_column > 5_000_000:
                raise ValueError("Output dimensions exceed the evaluator's fixture limit")
            for row in sheet:
                for cell in row:
                    if cell.value is None:
                        continue
                    sid = cell._style_id
                    if sid not in style_cache:
                        style_cache[sid] = hashlib.sha256(
                            json.dumps(
                                [
                                    str(cell.font),
                                    str(cell.fill),
                                    str(cell.border),
                                    str(cell.alignment),
                                    str(cell.protection),
                                    cell.number_format,
                                ],
                                ensure_ascii=False,
                            ).encode()
                        ).digest()
                    styles.update(cell.coordinate.encode() + b"\0" + style_cache[sid])
                    if cell.coordinate not in changed.get(sheet.title, set()):
                        value = cell.value
                        if isinstance(value, float) and value.is_integer():
                            value = int(value)
                        values.update(
                            json.dumps(
                                [cell.coordinate, cell.data_type, value],
                                ensure_ascii=False,
                                default=str,
                            ).encode()
                        )
            result[sheet.title] = {
                "values": values.hexdigest(),
                "styles": styles.hexdigest(),
                "state": sheet.sheet_state,
                "shape": [sheet.max_row, sheet.max_column],
            }
        result["_sheet_order"] = book.sheetnames
        return result
    finally:
        book.close()


def part_hashes(path):
    result = {}
    with ZipFile(path) as archive:
        for part in archive.infolist():
            with archive.open(part) as stream:
                result[part.filename] = hashlib.file_digest(stream, "sha256").hexdigest()
    return result


def expected_changes(case, rows):
    if case == "edit":
        return {"Assumptions": {"B1": 0.12}}
    if case == "wide":
        return {"Summary": {"B2": sum(i * 3 for i in range(1, 41) if i % 2 == 0)}}
    if case == "aggregate":
        totals = {key: 0 for key in ["APAC", "EMEA", "Americas", "Other"]}
        for i in range(1, rows + 1):
            totals[["APAC", "EMEA", "Americas", "Other"][i % 4]] += 2 * (i % 1000 + 1)
        return {"Summary": {f"B{i}": totals[key] for i, key in enumerate(totals, 2)}}
    return {}


def cached_assumption(path):
    """Read the fixture's existing formula cache without evaluating it."""
    from openpyxl import load_workbook

    book = load_workbook(path, read_only=True, data_only=True)
    try:
        return book["Assumptions"]["C1"].value
    finally:
        book.close()


def judge(case, source, output, result_path, rows, reference=None):
    from openpyxl import load_workbook

    checks = {}
    report = {"checks": checks}
    try:
        raw = result_path.read_bytes()
        json.loads(raw)
        checks["bounded_json"] = len(raw) <= 6000
        report["result_bytes"] = len(raw)
        if case == "create":
            book = load_workbook(output)
            try:
                checks["sheet_names"] = book.sheetnames == ["Budget"]
                sheet = book["Budget"]
                expected = {
                    "A1": "Month",
                    "B1": "Revenue",
                    "C1": "Cost",
                    "D1": "Profit",
                    "A5": "Total",
                }
                for rn, row in enumerate([["Jan", 100, 60], ["Feb", 120, 70], ["Mar", 90, 50]], 2):
                    expected.update(
                        {f"{col}{rn}": value for col, value in zip("ABC", row, strict=True)}
                    )
                    expected[f"D{rn}"] = f"=B{rn}-C{rn}"
                expected.update({f"{col}5": f"=SUM({col}2:{col}4)" for col in "BCD"})
                checks["values_and_formulas"] = all(
                    sheet[cell].value == value for cell, value in expected.items()
                )
                checks["bold_headers"] = all(sheet[f"{col}1"].font.bold for col in "ABCD")
                checks["number_formats"] = all(
                    ".00" in sheet[f"{col}{rn}"].number_format
                    for col in "BCD"
                    for rn in range(2, 6)
                )
                checks["freeze_panes"] = sheet.freeze_panes == "A2"
                checks["column_widths"] = all(
                    sheet.column_dimensions[col].width >= 10 for col in "ABCD"
                )
                from openpyxl.chart import BarChart

                charts = [
                    chart
                    for chart in sheet._charts
                    if isinstance(chart, BarChart) and chart.type == "col"
                ]
                checks["column_chart"] = bool(charts)
                references = [
                    series.val.numRef.f.replace("$", "")
                    for chart in charts
                    for series in chart.series
                    if series.val and series.val.numRef
                ]
                checks["chart_series"] = len(references) == 2 and all(
                    any(ref.endswith(f"!{col}2:{col}4") for ref in references) for col in "BC"
                )
            finally:
                book.close()
        else:
            changed = expected_changes(case, rows)
            original = reference or cell_fingerprints(source, changed)
            actual = cell_fingerprints(output, changed)
            checks["unchanged_values_formulas_structure_styles"] = actual == original
            if not checks["unchanged_values_formulas_structure_styles"]:
                report["changed_fingerprint_fields"] = {
                    name: [
                        k
                        for k in original[name]
                        if original[name][k] != actual.get(name, {}).get(k)
                    ]
                    for name in original
                    if isinstance(original[name], dict) and original[name] != actual.get(name)
                }
            book = load_workbook(output, read_only=True, data_only=False)
            try:
                checks["requested_values"] = all(
                    book[sheet][cell].value == value
                    for sheet, cells in changed.items()
                    for cell, value in cells.items()
                )
            finally:
                book.close()
            if case in {"edit", "aggregate"}:
                checks["existing_formula_cache_preserved"] = cached_assumption(
                    source
                ) == cached_assumption(output)
            before, after = part_hashes(source), part_hashes(output)
            report["byte_identical_source_parts"] = sum(
                after.get(name) == digest for name, digest in before.items()
            )
            report["source_part_count"] = len(before)
            report["changed_or_missing_parts"] = [
                name for name, digest in before.items() if after.get(name) != digest
            ]
            if "customXml/item1.xml" in before:
                checks["custom_xml_payload"] = (
                    after.get("customXml/item1.xml") == before["customXml/item1.xml"]
                )
                with ZipFile(output) as archive:
                    links = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
                    checks["custom_xml_relationship"] = any(
                        link.get("Type", "").endswith("/customXml")
                        and link.get("Target", "").replace("\\", "/")
                        in {"../customXml/item1.xml", "/customXml/item1.xml"}
                        for link in links
                    )
        report["passed"] = all(checks.values())
    except Exception as exc:
        report.update(passed=False, error=f"{type(exc).__name__}: {exc}")
    return report


def run_trial(command, case, spec, directory, rows, reference):
    import psutil

    directory.mkdir(parents=True)
    output, result = directory / "output.xlsx", directory / "result.json"
    substitutions = {
        "{input}": spec["input"] or "-",
        "{output}": str(output.resolve()),
        "{result}": str(result.resolve()),
    }
    argv = [substitutions.get(arg, arg) for arg in command["argv"]]
    env = dict(os.environ, **command.get("env", {}))
    temporary = directory / "temp"
    temporary.mkdir()
    env.update(TEMP=str(temporary.resolve()), TMP=str(temporary.resolve()))
    stdout, stderr = directory / "stdout.txt", directory / "stderr.txt"
    peak = 0
    stop = threading.Event()
    with stdout.open("wb") as out, stderr.open("wb") as err:
        start = time.perf_counter()
        process = subprocess.Popen(argv, stdout=out, stderr=err, env=env, cwd=directory)

        def sample():
            nonlocal peak
            while not stop.wait(0.01):
                try:
                    parent = psutil.Process(process.pid)
                    peak = max(
                        peak,
                        sum(
                            p.memory_info().rss
                            for p in [parent, *parent.children(recursive=True)]
                            if p.is_running()
                        ),
                    )
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass

        sampler = threading.Thread(target=sample, daemon=True)
        sampler.start()
        timeout = False
        try:
            code = process.wait(timeout=180)
        except subprocess.TimeoutExpired:
            timeout = True
            for child in psutil.Process(process.pid).children(recursive=True):
                child.kill()
            process.kill()
            code = process.wait()
        elapsed = time.perf_counter() - start
        stop.set()
        sampler.join()
    verdict = (
        judge(case, spec["input"], output, result, rows, reference)
        if code == 0
        else {"passed": False}
    )
    return {
        "seconds": elapsed,
        "peak_process_tree_rss_mib": peak / 1024**2,
        "exit_code": code,
        "timeout": timeout,
        "stdout_bytes": stdout.stat().st_size,
        "stderr_bytes": stderr.stat().st_size,
        "verdict": verdict,
    }


def run(directory, registries, output, repeats=3):
    if repeats < 1:
        raise ValueError("repeats must be positive")
    manifest = json.loads((directory / "tasks.json").read_text(encoding="utf-8"))
    tasks = manifest["tasks"]
    inputs = {Path(spec["input"]) for spec in tasks.values() if spec["input"]}

    def check_inputs():
        for path in inputs:
            if hashlib.sha256(path.read_bytes()).hexdigest() != manifest["input_hashes"][path.name]:
                raise ValueError(f"Input changed since preparation: {path}")

    check_inputs()
    provenance = {
        name: {
            file.name: hashlib.sha256(file.read_bytes()).hexdigest()
            for file in Path(path).parent.iterdir()
            if file.is_file() and file.suffix in {".py", ".mjs", ".json", ".md"}
        }
        for name, path in registries.items()
    }
    versions = {}
    for package in ["duckdb", "openpyxl", "lxml", "psutil", "xlsxwriter", "pandas"]:
        try:
            versions[package] = version(package)
        except PackageNotFoundError:
            versions[package] = "not installed in evaluator environment"
    engine_hashes = {
        path.name: hashlib.sha256(path.read_bytes()).hexdigest()
        for path in (Path(__file__).parents[1] / "src" / "sheetjet").glob("*.py")
    }
    originals = {
        case: cell_fingerprints(spec["input"], expected_changes(case, manifest["rows_per_period"]))
        for case, spec in tasks.items()
        if spec["input"]
    }
    submissions = {
        name: json.loads(Path(path).read_text(encoding="utf-8-sig"))
        for name, path in registries.items()
    }
    jobs = [
        (name, case, trial) for name in submissions for case in tasks for trial in range(repeats)
    ]
    random.Random(20261005).shuffle(jobs)
    trials = []
    output.mkdir(parents=True, exist_ok=True)
    for name, case, trial in jobs:
        result = run_trial(
            submissions[name][case],
            case,
            tasks[case],
            output / f"{name}-{case}-{trial}",
            manifest["rows_per_period"],
            originals.get(case),
        )
        result.update(skill=name, case=case, trial=trial)
        trials.append(result)
        check_inputs()
        print(
            f"{name} {case} {trial + 1}/{repeats}: {result['seconds']:.3f}s, passed={result['verdict']['passed']}",
            flush=True,
        )
        # A failed process does not erase the other tasks or become an artificial timing win.
        (output / "trials.json").write_text(json.dumps(trials, indent=2), encoding="utf-8")
    medians = []
    for name in submissions:
        for case in tasks:
            selected = [t for t in trials if t["skill"] == name and t["case"] == case]
            medians.append(
                {
                    "skill": name,
                    "case": case,
                    "passed": sum(t["verdict"]["passed"] for t in selected),
                    "trials": len(selected),
                    "seconds": statistics.median(t["seconds"] for t in selected),
                    "peak_process_tree_rss_mib": statistics.median(
                        t["peak_process_tree_rss_mib"] for t in selected
                    ),
                }
            )
    report = {
        "platform": platform.platform(),
        "evaluator_python": platform.python_version(),
        "evaluator_dependency_versions": versions,
        "sheetjet_source_hashes": engine_hashes,
        "submission_hashes": provenance,
        "input_hashes": manifest["input_hashes"],
        "rows_per_period": manifest["rows_per_period"],
        "repeats": repeats,
        "agent_generations_per_skill": 1,
        "actual_model_tokens_measured": False,
        "methodology": "Independent same-model agents authored each submission from its assigned skill, then builders were frozen. Timings are fresh-process execution of frozen builders, not model reasoning time. Serial fixed-randomized trials; OS caches not flushed. Process-tree RSS sampled every 10ms may double-count shared pages and miss peaks. Independent openpyxl artifact judging is outside timed execution. Skill breadth and agent success-rate generalization are not established by one generation per skill. No universal skill ranking or token-saving claim.",
        "medians": medians,
        "trials": trials,
    }
    (output / "report.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(f"Saved {output / 'report.json'}")


def prepare(directory, rows=25000):
    import xlsxwriter

    directory.mkdir(parents=True, exist_ok=True)
    inputs = directory / "inputs"
    inputs.mkdir(exist_ok=True)
    large = inputs / "periods.xlsx"
    base = inputs / "base-periods.xlsx"
    fixture(base, rows, 4, shared_strings=True)
    # Extend the existing package so generated data remains identical to the parser benchmark.
    summary = f'<worksheet xmlns="{NS}"><dimension ref="A1:B5"/><sheetData>'
    for rn, (label, amount) in enumerate(
        [("Region", "Revenue snapshot"), ("APAC", 0), ("EMEA", 0), ("Americas", 0), ("Other", 0)], 1
    ):
        summary += f'<row r="{rn}"><c r="A{rn}" t="inlineStr"><is><t>{label}</t></is></c>'
        summary += (
            f'<c r="B{rn}" t="inlineStr"><is><t>{amount}</t></is></c>'
            if isinstance(amount, str)
            else f'<c r="B{rn}"><v>{amount}</v></c>'
        ) + "</row>"
    summary += "</sheetData></worksheet>"
    with ZipFile(base) as src, ZipFile(large, "w", compression=ZIP_DEFLATED) as dst:
        for info in src.infolist():
            raw = src.read(info)
            if info.filename == "xl/workbook.xml":
                node = ET.fromstring(raw)
                ET.SubElement(
                    node.find(Q + "sheets"),
                    Q + "sheet",
                    name="Summary",
                    sheetId="7",
                    attrib={
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id": "rIdSummary"
                    },
                )
                raw = ET.tostring(node)
            elif info.filename == "xl/_rels/workbook.xml.rels":
                node = ET.fromstring(raw)
                namespace = "{http://schemas.openxmlformats.org/package/2006/relationships}"
                for rid, kind, target in [
                    ("rIdSummary", "worksheet", "worksheets/sheet7.xml"),
                    ("rIdCustom", "customXml", "../customXml/item1.xml"),
                ]:
                    ET.SubElement(
                        node,
                        namespace + "Relationship",
                        Id=rid,
                        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
                        + kind,
                        Target=target,
                    )
                raw = ET.tostring(node)
            elif info.filename == "[Content_Types].xml":
                node = ET.fromstring(raw)
                ET.SubElement(
                    node,
                    "{http://schemas.openxmlformats.org/package/2006/content-types}Override",
                    PartName="/xl/worksheets/sheet7.xml",
                    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                )
                raw = ET.tostring(node)
            elif info.filename == "xl/worksheets/sheet5.xml":
                node = ET.fromstring(raw)
                cell = ET.SubElement(node.find(Q + "sheetData/" + Q + "row"), Q + "c", r="C1")
                ET.SubElement(cell, Q + "f").text = "B1*100"
                ET.SubElement(cell, Q + "v").text = "8"
                node.find(Q + "dimension").set("ref", "A1:C1")
                raw = ET.tostring(node)
            dst.writestr(info, raw)
        dst.writestr("xl/worksheets/sheet7.xml", summary)
        dst.writestr(
            "customXml/item1.xml",
            '<meta xmlns="urn:sheetjet:evaluation"><record>retain-this-metadata</record></meta>',
        )
    wide = inputs / "wide.xlsx"
    with xlsxwriter.Workbook(wide, {"constant_memory": True}) as book:
        sheet = book.add_worksheet("Wide")
        sheet.write_row(0, 0, ["Region", *[f"Field{i}" for i in range(1, 2499)], "Amount"])
        for i in range(1, 41):
            sheet.write_row(i, 0, ["APAC" if i % 2 == 0 else "EMEA", *([i] * 2498), i * 3])
        sheet = book.add_worksheet("Summary")
        sheet.write_row(0, 0, ["Metric", "Snapshot"])
        sheet.write_row(1, 0, ["APAC amount", 0])
    common = (
        "For edits, preserve all unrequested cell values, formulas, styles, sheet structure, hidden states, "
        "and any custom XML metadata. Do not recalculate existing formulas or refresh external content; "
        "preserve formula text and report when recalculation is needed. Write to a new output file. "
        "Return a compact result.json of at most 6000 UTF-8 bytes. Do not dump worksheet contents. "
        "Computed numeric snapshots below are explicitly requested as static values, not new formulas."
    )
    tasks = {
        "edit": {
            "input": str(large.resolve()),
            "instruction": "Change Assumptions!B1 from 0.08 to 0.12. Leave everything else intact. Return changed cell, previous/new values and recalculation status.",
        },
        "aggregate": {
            "input": str(large.resolve()),
            "instruction": "Sum Revenue by Region using only Period01 and Period03. Write static snapshot totals into existing Summary!B2:B5, aligned to its existing Region labels in A2:A5. Ignore other periods and hidden notes. Return the four totals and selected sheet names.",
        },
        "wide": {
            "input": str(wide.resolve()),
            "instruction": "Wide has 2500 columns and a header in row 1. Sum Amount for rows whose Region is APAC, using rows 2:41. Write the numeric snapshot into existing Summary!B2. Return only the total and source coordinates; keep wide headers and data out of the response.",
        },
        "create": {
            "input": None,
            "instruction": "Create a small workbook with one sheet named Budget. A1:C1 must be Month, Revenue, Cost. A2:C4 must be Jan/100/60, Feb/120/70, Mar/90/50. Add D1 Profit and D2:D4 as Excel formulas =B2-C2 through =B4-C4. Put Total in A5, with SUM formulas for B5:D5. Add a column chart for monthly Revenue and Cost, excluding totals. Bold the header, use two-decimal number formats for B2:D5, freeze the header row, and set readable column widths. Formula caches are not required; Excel can calculate on opening. Return a compact description of the created workbook.",
        },
    }
    manifest = {
        "common": common,
        "tasks": tasks,
        "rows_per_period": rows,
        "input_hashes": {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in [large, wide]},
    }
    (directory / "tasks.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    print(json.dumps({"manifest": str(directory / "tasks.json"), "tasks": list(tasks)}))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--prepare", type=Path)
    parser.add_argument("--rows", type=int, default=25000)
    parser.add_argument("--run", type=Path)
    parser.add_argument("--registry", action="append", default=[], help="skill-name=commands.json")
    parser.add_argument("--output", type=Path, default=Path("benchmark-output/skill-trials"))
    parser.add_argument("--repeats", type=int, default=3)
    args = parser.parse_args()
    if args.prepare:
        prepare(args.prepare, args.rows)
    if args.run:
        run(args.run, dict(item.split("=", 1) for item in args.registry), args.output, args.repeats)


if __name__ == "__main__":
    main()
