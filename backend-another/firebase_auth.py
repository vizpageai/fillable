from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

import json
import firebase_admin
from firebase_admin import auth, credentials

from config import SETTINGS


@dataclass(frozen=True)
class FirebaseUser:
    uid: str
    email: Optional[str]
    email_verified: bool


_initialized = False


def init_firebase() -> None:
    global _initialized
    if _initialized:
        return
    cred_dict = SETTINGS.firebase_credentials_dict()
    project_id = SETTINGS.firebase_project_id
    if cred_dict:
        project_id = project_id or str(cred_dict.get("project_id", "")).strip()
        cred = credentials.Certificate(cred_dict)
        firebase_admin.initialize_app(cred, {"projectId": project_id} if project_id else None)
        _initialized = True
        return
    if SETTINGS.firebase_credentials_path:
        project_id = project_id or _project_id_from_file(SETTINGS.firebase_credentials_path)
        cred = credentials.Certificate(SETTINGS.firebase_credentials_path)
        firebase_admin.initialize_app(cred, {"projectId": project_id} if project_id else None)
        _initialized = True
        return
    if SETTINGS.firebase_google_credentials_path:
        project_id = project_id or _project_id_from_file(SETTINGS.firebase_google_credentials_path)
        cred = credentials.Certificate(SETTINGS.firebase_google_credentials_path)
        firebase_admin.initialize_app(cred, {"projectId": project_id} if project_id else None)
        _initialized = True
        return
    firebase_admin.initialize_app(None, {"projectId": project_id} if project_id else None)
    _initialized = True


def _project_id_from_file(path: str) -> str:
    try:
        data = json.loads(open(path, "r", encoding="utf-8").read())
        return str(data.get("project_id", "")).strip()
    except Exception:
        return ""


def verify_firebase_token(token: str) -> FirebaseUser:
    init_firebase()
    decoded = auth.verify_id_token(token, check_revoked=False)
    uid = str(decoded.get("uid", "") or decoded.get("user_id", "")).strip()
    if not uid:
        raise ValueError("Invalid Firebase token: missing uid.")
    email = decoded.get("email")
    email_verified = bool(decoded.get("email_verified", False))
    if SETTINGS.require_email_verified and not email_verified:
        raise ValueError("Email not verified.")
    return FirebaseUser(uid=uid, email=email, email_verified=email_verified)
