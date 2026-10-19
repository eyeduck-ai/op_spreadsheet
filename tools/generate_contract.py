"""Generate GAS CONFIG and installed Python data from the canonical contract.

Run without flags after intentional edits; CI uses --check and never rewrites.
"""
import argparse
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

def outputs():
    raw = (ROOT / "contracts/sheet-contract.json").read_text(encoding="utf-8")
    spec = json.loads(raw)
    code = (ROOT / "code.js").read_text(encoding="utf-8")
    version = re.search(r"[\"']?VERSION[\"']?\s*:\s*[\"']([^\"']+)", code).group(1)
    config = {"VERSION": version, **spec["gas_config"]}
    generated = "const CONFIG = " + json.dumps(config, ensure_ascii=False, indent=2) + ";"
    changed, count = re.subn(r"const CONFIG = \{.*?^\};", lambda _: generated, code, count=1, flags=re.S | re.M)
    if count != 1:
        raise ValueError("Expected exactly one CONFIG block")
    data = ROOT / "packages/python/src/vghks_sheet_core/data"
    result = {ROOT / "code.js": changed, data / "sheet-contract.json": raw}
    for case in (ROOT / "contracts/cases").glob("*.json"):
        result[data / "cases" / case.name] = case.read_text(encoding="utf-8")
    return result

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    stale = []
    for path, value in outputs().items():
        if not path.exists() or path.read_text(encoding="utf-8") != value:
            stale.append(str(path.relative_to(ROOT)))
            if not args.check:
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(value, encoding="utf-8")
    if stale and args.check:
        raise SystemExit("Regenerate contract outputs: " + ", ".join(stale))
    print("Contract generation verified" if args.check else "Contract outputs generated")

if __name__ == "__main__":
    main()
