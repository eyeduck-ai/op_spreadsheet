"""Canonical normalization and hashing; retain existing persistent hashes."""
import hashlib
import json
import re
import unicodedata
from datetime import datetime, timedelta, timezone

TAIPEI = timezone(timedelta(hours=8))

def norm(value):
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", str(value))).strip()

def digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True,
                                     separators=(",", ":")).encode()).hexdigest()

def timestamp():
    return datetime.now(TAIPEI).isoformat(timespec="microseconds")
