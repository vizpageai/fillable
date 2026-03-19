from __future__ import annotations

from typing import Any, Dict, Optional

from pydantic import BaseModel


class AuthRequest(BaseModel):
    email: str
    password: str


class RefreshRequest(BaseModel):
    refresh_token: str


class AuthResponse(BaseModel):
    id_token: str
    refresh_token: str
    expires_in: int
    email: str
    uid: str


class CheckoutSessionResponse(BaseModel):
    url: str
    id: str


class EntitlementResponse(BaseModel):
    active: bool
    subscriptions: list[Dict[str, Any]]
    credits: float


class CreditsResponse(BaseModel):
    credits: float


class ProxyRequest(BaseModel):
    payload: Dict[str, Any]
    model_fallback: Optional[str] = None


class FeedbackRequest(BaseModel):
    name: str
    email: str
    message: str
