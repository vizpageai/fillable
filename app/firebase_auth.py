from __future__ import annotations

import base64
import json
import time


class FirebaseAuthError(RuntimeError):
    pass


def jwt_expiry_epoch(token: str) -> int | None:
    try:
        parts = token.split(".")
        if len(parts) < 2:
            return None
        payload = parts[1]
        padding = "=" * (-len(payload) % 4)
        decoded = base64.urlsafe_b64decode(payload + padding)
        data = json.loads(decoded.decode("utf-8", errors="ignore"))
        exp = data.get("exp")
        return int(exp) if exp else None
    except Exception:
        return None


def is_token_expired(exp_epoch: int | None, skew_seconds: int = 60) -> bool:
    if not exp_epoch:
        return True
    return time.time() >= (exp_epoch - skew_seconds)
