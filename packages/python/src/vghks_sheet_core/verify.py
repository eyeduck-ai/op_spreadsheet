"""Offline parity checks against real GAS exports, with no Google credentials."""
import argparse
import json
import subprocess
from datetime import date
from importlib.resources import files
from pathlib import Path

from .calendar import arrival, calendar_day, headers, in_calendar_range, scan
from .contract import CONTRACT
from .schema import DAY_MARKER, HOSPITALS, canonical_header, full_date, monthly_row, sheet_headers


def fixture(name):
    return json.loads(files("vghks_sheet_core").joinpath("data/cases/" + name).read_text(encoding="utf-8"))


def literal(value):
    return {"userEnteredValue": {"numberValue" if isinstance(value, (int, float)) else "stringValue": value}}


def verify(gas_repo=None, node="node"):
    parsing, formatting = fixture("schema_parsing.json"), fixture("sheet_contract.json")
    workflows, archives = fixture("workflow_contract.json"), fixture("gas_archives.json")
    if gas_repo:
        source = json.loads((Path(gas_repo) / "contracts/sheet-contract.json").read_text(encoding="utf-8"))
        if source != CONTRACT:
            raise ValueError("Installed core specification differs from the candidate GAS source")
        for script, function, expected in (("parser_contract", "parserContract", parsing),
                ("sheet_contract", "sheetContract", formatting), ("workflow_contract", "workflowContract", workflows),
                ("archive_contract", "archiveContract", archives)):
            result = subprocess.run([node, "-e", "process.stdout.write(JSON.stringify(require(process.argv[1])[process.argv[2]]()))",
                                     str(Path(gas_repo).resolve() / "tests" / (script + ".js")), function],
                                    capture_output=True, text=True, check=True)
            if json.loads(result.stdout) != expected:
                raise ValueError(f"Candidate GAS export differs: {script}; review and update the contract cases")
    count = 0
    for index, case in enumerate(parsing["date_cases"]):
        if full_date(literal(case["input"])) != case["iso"]:
            raise ValueError(f"Date case {index} differs")
        count += 1
    labels = formatting["monthly_headers"]
    sheet = {"properties": {"title": "202610"}, "rows": [[literal(label) for label in labels]]}
    cols = sheet_headers(sheet)
    if headers(sheet)[1] != cols or set(HOSPITALS) != set(formatting["hospitals"]) or DAY_MARKER != formatting["header_marker"]:
        raise ValueError("Field and hospital contract differs")
    kind = {"DATE_HEADER": "day", "HYBRID_CONFLICT": "invalid_header", "PATIENT": "patient", "BLANK": "blank"}
    for index, case in enumerate(parsing["row_cases"]):
        values = {"日期／報到時間": case["date"], "病歷號": case["marker"], "醫院": case["hospital"], "CalendarEventId": case.get("event", "")}
        if monthly_row([literal(values.get(label, "")) for label in labels], cols)["kind"] != kind[case["kind"]]:
            raise ValueError(f"Row case {index} differs")
        count += 1
    recognized = {"TIME": "日期／報到時間", "DIAGNOSIS": "診斷", "EVENT_ID": "CalendarEventId"}
    for index, case in enumerate(parsing["header_cases"]):
        cell = literal(case["header"])
        if case.get("formula"):
            cell = {"userEnteredValue": {"formulaValue": case["formula"]}, "formattedValue": case["header"]}
        actual = canonical_header(cell)
        if (actual if actual == recognized[case["key"]] else "") != case["canonical"]:
            raise ValueError(f"Header case {index} differs")
        count += 1
    for index, case in enumerate(parsing["time_cases"]):
        try:
            text, note = arrival({"time_cell": literal(case["input"]), "time": str(case["input"]), "sheet": "202610"})
            invalid = False
        except ValueError:
            text, note, invalid = "", "", True
        if (text, note, invalid) != (case["time"], case["note"], case["invalid"]):
            raise ValueError(f"Time case {index} differs")
        count += 1
    for phase in ("creation", "moved"):
        book = {"sheets": [{"properties": {"title": tab["name"], "sheetId": tab["id"]},
                           "rows": [[literal(v) for v in row] for row in tab["rows"]]} for tab in workflows[phase]]}
        snapshot = scan(book)
        found = [item for item in snapshot["patients"] if item["uid"] == workflows["patient_sync_id"]]
        if snapshot["issues"] or len(found) != 1:
            raise ValueError(f"Workflow {phase} has ambiguous identities")
        item = found[0]
        if ((item["date"], item["hospital"]) != (workflows[f"{phase}_date"], workflows[f"{phase}_hospital"])
                or item["event_id"] != workflows["patient_event_id"] or arrival(item) != ("08:30", "")
                or item["fields"]["人工備註"] != "人工欄位保留"
                or item["fields"].get("醫院", "") != workflows["patient_hospital_cells"][phase]):
            raise ValueError(f"Workflow {phase} changes route, binding or manual content")
        count += 1
    identities = None
    for phase in ("before", "after_september", "during_backfill", "after_backfill"):
        snapshot = scan(archives[phase])
        if snapshot["issues"] and phase != "during_backfill":
            raise ValueError(f"Archive {phase} has invalid identities")
        if identities is not None and snapshot["identities"] != identities:
            raise ValueError(f"Archive {phase} loses identities used to protect events")
        identities = snapshot["identities"]
        count += 1
    scope = fixture("calendar_scope.json")
    cutoff = date.fromisoformat(scope["today"])
    for index, case in enumerate(scope["cases"]):
        day = calendar_day(case["event"])
        if ((day.isoformat() if day else None) != case["day"]
                or in_calendar_range(case["event"], cutoff) != case["eligible"]):
            raise ValueError(f"Calendar date scope case {index} differs")
        count += 1
    return count


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--gas-repo", type=Path)
    parser.add_argument("--node", default="node")
    args = parser.parse_args()
    print(f"Shared sheet core passed: {verify(args.gas_repo, args.node)} synthetic cases")


if __name__ == "__main__":
    main()
