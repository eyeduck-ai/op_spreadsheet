"""Check the deployed main-program parser against this candidate GAS contract."""
import argparse
import importlib
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--bot-repo", type=Path, required=True)
    args = parser.parse_args()
    sys.path.insert(0, str(args.bot_repo.resolve()))
    checks = importlib.import_module("tools.verify_sheet_contract")
    cases = ROOT / "contracts/cases"
    def read(name):
        return json.loads((cases / name).read_text(encoding="utf-8"))
    count = checks.check_contract(read("schema_parsing.json"), read("sheet_contract.json"))
    count += checks.check_workflows(read("workflow_contract.json"))
    count += checks.check_archives(read("gas_archives.json"))
    print(f"Main program and Calendar contract passed against candidate GAS: {count} cases")


if __name__ == "__main__":
    main()
