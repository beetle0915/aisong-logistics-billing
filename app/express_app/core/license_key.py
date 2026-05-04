"""Offline startup key generation and verification."""

from __future__ import annotations

import hashlib
import hmac
import re
import time


LICENSE_KEY_PROTOCOL_VERSION = "v1"
LICENSE_KEY_APP_ID = "aisong-hash-key"
LICENSE_KEY_WINDOW_SECONDS = 300
LICENSE_KEY_SECRET = "AISONG_HASH_KEY_OFFLINE_SECRET_2026_05"
LICENSE_KEY_PATTERN = re.compile(r"^[0-9A-F]{4}-[0-9A-F]{4}$")


def normalize_license_key(raw_key: str) -> str:
    compact = re.sub(r"[^0-9A-Fa-f]", "", raw_key or "").upper()
    if len(compact) != 8:
        return ""
    formatted = f"{compact[:4]}-{compact[4:]}"
    return formatted if LICENSE_KEY_PATTERN.match(formatted) else ""


def generate_license_key(epoch_seconds: int | float | None = None) -> str:
    current_seconds = int(time.time() if epoch_seconds is None else epoch_seconds)
    time_slot = current_seconds // LICENSE_KEY_WINDOW_SECONDS
    message = f"{LICENSE_KEY_PROTOCOL_VERSION}|{LICENSE_KEY_APP_ID}|{time_slot}"
    digest = hmac.new(
        LICENSE_KEY_SECRET.encode("utf-8"),
        message.encode("utf-8"),
        hashlib.sha256,
    ).hexdigest().upper()
    compact = digest[:8]
    return f"{compact[:4]}-{compact[4:]}"


def verify_license_key(
    raw_key: str,
    epoch_seconds: int | float | None = None,
    *,
    window_tolerance: int = 1,
) -> bool:
    key = normalize_license_key(raw_key)
    if not key:
        return False

    current_seconds = int(time.time() if epoch_seconds is None else epoch_seconds)
    current_slot = current_seconds // LICENSE_KEY_WINDOW_SECONDS
    for offset in range(-window_tolerance, window_tolerance + 1):
        slot_seconds = (current_slot + offset) * LICENSE_KEY_WINDOW_SECONDS
        if hmac.compare_digest(generate_license_key(slot_seconds), key):
            return True
    return False
