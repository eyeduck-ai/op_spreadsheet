"""Check the deployed main-program parser against this candidate GAS contract."""
import argparse
import importlib
import json
import sys
import tomllib
from importlib.metadata import version
from importlib.resources import files
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--bot-repo", type=Path, required=True)
    args = parser.parse_args()
    expected = json.loads((ROOT / "contracts/sheet-contract.json").read_text(encoding="utf-8"))
    installed = json.loads(files("vghks_sheet_core").joinpath("data/sheet-contract.json").read_text(encoding="utf-8"))
    package_version = tomllib.loads((ROOT / "packages/python/pyproject.toml").read_text())["project"]["version"]
    if installed != expected or version("vghks-sheet-core") != package_version:
        raise ValueError("Installed shared core differs from the candidate GAS specification; check consumer dependency pins")
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
