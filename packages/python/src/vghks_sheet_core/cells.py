"""Typed cell utilities shared by planners and monitor; no application dependencies."""
import copy

from .common import digest, norm
from .schema import cell_text as cell_text

CELL_FIELDS = "userEnteredValue,userEnteredFormat,dataValidation,note,textFormatRuns"

def patient_key(value):
    text = norm(value)
    return str(int(text)) if text.isdigit() else text

def native_cell(cell):
    return {k: copy.deepcopy(v) for k, v in cell.items() if k in CELL_FIELDS.split(",")}

def fingerprint(book):
    value = copy.deepcopy(book)
    # Reclaiming another confirmed batch's witness is not a clinical change.
    value["developerMetadata"] = [m for m in value.get("developerMetadata", [])
                                  if m.get("metadataKey") not in {"vghks.opd.update", "vghks.sheet.management"}]
    for sheet in value["sheets"]:
        sheet["rows"] = [[native_cell(c) for c in r] for r in sheet.get("rows", [])]
    return digest(value)

def cell_at(row, index):
    return row[index] if index < len(row) else {}

def column_name(index):
    value = ""
    while index >= 0:
        index, rem = divmod(index, 26)
        value = chr(65 + rem) + value
        index -= 1
    return value
