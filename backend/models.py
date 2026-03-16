from __future__ import annotations

from typing import Any, Dict, List, Optional

from pydantic import BaseModel, Field


class CollectionsTokenResponse(BaseModel):
    access_token: str
    expires_in: int
    token_type: str = "Bearer"
    audience: str


class EntitlementRequest(BaseModel):
    user_store_id: str = Field(..., description="User Store ID (Collections ID) from StoreContext.")
    product_ids: Optional[List[str]] = None
    sku_ids: Optional[List[str]] = None
    sandbox: Optional[str] = None


class EntitlementResponse(BaseModel):
    active: bool
    matched_product_ids: List[str]
    items: List[Dict[str, Any]]


class OpenAIProxyRequest(BaseModel):
    user_store_id: Optional[str] = None
    sandbox: Optional[str] = None
    payload: Dict[str, Any]
