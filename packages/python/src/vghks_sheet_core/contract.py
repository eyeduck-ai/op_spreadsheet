"""The installed contract is bundled, immutable across a deployment."""
import json
from importlib.resources import files

from . import __version__

CONTRACT = json.loads(files(__package__).joinpath("data/sheet-contract.json").read_text(encoding="utf-8"))
CONTRACT_VERSION = CONTRACT["contract_version"]
if CONTRACT_VERSION.split(".")[:2] != __version__.split(".")[:2]:
    raise RuntimeError("Unsupported installed sheet contract; upgrade and verify the Python package.")
