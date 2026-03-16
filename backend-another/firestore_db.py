from __future__ import annotations

from datetime import datetime, timezone, timedelta
from typing import Optional, Tuple

from firebase_admin import firestore


def _client() -> firestore.Client:
    return firestore.client()


def _doc(uid: str) -> firestore.DocumentReference:
    return _client().collection("credits").document(uid)


def get_balance(uid: str) -> Tuple[float, Optional[str]]:
    snap = _doc(uid).get()
    if not snap.exists:
        return 0.0, None
    data = snap.to_dict() or {}
    balance = float(data.get("balance", 0.0) or 0.0)
    email = data.get("email")
    return balance, email


def set_email(uid: str, email: str) -> None:
    if not email:
        return
    _doc(uid).set(
        {
            "email": email,
            "updated_at": datetime.now(timezone.utc),
        },
        merge=True,
    )


def add_credits(uid: str, amount: float, email: str | None = None) -> float:
    if amount <= 0:
        return get_balance(uid)[0]
    updates = {
        "balance": firestore.Increment(float(amount)),
        "updated_at": datetime.now(timezone.utc),
    }
    if email:
        updates["email"] = email
    _doc(uid).set(updates, merge=True)
    return get_balance(uid)[0]


def deduct_credits(uid: str, amount: float) -> float:
    if amount <= 0:
        return get_balance(uid)[0]

    doc_ref = _doc(uid)

    @firestore.transactional
    def _tx_update(transaction: firestore.Transaction) -> float:
        snap = doc_ref.get(transaction=transaction)
        data = snap.to_dict() if snap.exists else {}
        balance = float((data or {}).get("balance", 0.0) or 0.0)
        new_balance = max(balance - float(amount), 0.0)
        transaction.set(
            doc_ref,
            {
                "balance": new_balance,
                "updated_at": datetime.now(timezone.utc),
            },
            merge=True,
        )
        return new_balance

    return _tx_update(_client().transaction())


def store_auth_session(session_id: str, payload: dict, ttl_seconds: int = 600) -> None:
    now = datetime.now(timezone.utc)
    _client().collection("auth_sessions").document(session_id).set(
        {
            "payload": payload,
            "created_at": now,
            "expires_at": now + timedelta(seconds=ttl_seconds),
        }
    )


def get_auth_session(session_id: str) -> Optional[dict]:
    doc_ref = _client().collection("auth_sessions").document(session_id)
    snap = doc_ref.get()
    if not snap.exists:
        return None
    data = snap.to_dict() or {}
    expires_at = data.get("expires_at")
    if expires_at and hasattr(expires_at, "timestamp"):
        if datetime.now(timezone.utc).timestamp() > expires_at.timestamp():
            doc_ref.delete()
            return None
    payload = data.get("payload")
    if payload:
        doc_ref.delete()
    return payload
