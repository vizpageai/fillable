from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import httpx

from .config import SETTINGS, require_values


_token_cache: Dict[str, Tuple[str, float]] = {}


@dataclass(frozen=True)
class EntitlementCheck:
    active: bool
    matched_product_ids: List[str]
    raw_items: List[dict]


def _get_cached_token(resource: str) -> str | None:
    token_entry = _token_cache.get(resource)
    if not token_entry:
        return None
    token, exp_at = token_entry
    if time.time() >= exp_at:
        _token_cache.pop(resource, None)
        return None
    return token


def _cache_token(resource: str, token: str, expires_in: int) -> None:
    exp_at = time.time() + max(expires_in - 90, 30)
    _token_cache[resource] = (token, exp_at)


def _aad_token(resource: str) -> str:
    cached = _get_cached_token(resource)
    if cached:
        return cached

    require_values(
        [SETTINGS.aad_client_id, SETTINGS.aad_client_secret, SETTINGS.resolve_token_url()],
        label="AAD client credentials",
    )
    token_url = SETTINGS.resolve_token_url()
    payload = {
        "grant_type": "client_credentials",
        "client_id": SETTINGS.aad_client_id,
        "client_secret": SETTINGS.aad_client_secret,
        "resource": resource,
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    with httpx.Client(timeout=30) as client:
        response = client.post(token_url, data=payload, headers=headers)
    if response.status_code >= 400:
        raise RuntimeError(f"AAD token error {response.status_code}: {response.text[:800]}")
    data = response.json()
    token = str(data.get("access_token", "")).strip()
    if not token:
        raise RuntimeError(f"AAD token response missing access_token: {data}")
    expires_in = int(data.get("expires_in", 3600))
    _cache_token(resource, token, expires_in)
    return token


def get_create_collections_token() -> str:
    return _aad_token(SETTINGS.store_create_collections_audience)


def _collections_service_token() -> str:
    return _aad_token(SETTINGS.store_audience)


def _normalize_product_pairs(
    product_ids: List[str],
    sku_ids: Optional[List[str]],
) -> List[dict]:
    pairs = []
    if sku_ids:
        for idx, product_id in enumerate(product_ids):
            sku_id = sku_ids[idx] if idx < len(sku_ids) else ""
            entry = {"productId": product_id}
            if sku_id:
                entry["skuId"] = sku_id
            pairs.append(entry)
    else:
        for product_id in product_ids:
            pairs.append({"productId": product_id})
    return pairs


def _request_collections(payload: dict) -> dict:
    token = _collections_service_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "User-Agent": SETTINGS.store_user_agent,
    }
    with httpx.Client(timeout=30) as client:
        response = client.post(SETTINGS.store_collections_endpoint, json=payload, headers=headers)
    if response.status_code >= 400:
        raise RuntimeError(
            f"Store collections error {response.status_code}: {response.text[:1200]}"
        )
    data = response.json()
    if not isinstance(data, dict):
        raise RuntimeError("Store collections response was not a JSON object.")
    return data


def query_entitlements(
    *,
    user_store_id: str,
    product_ids: List[str],
    sku_ids: Optional[List[str]] = None,
    sandbox: str | None = None,
) -> List[dict]:
    pairs = _normalize_product_pairs(product_ids, sku_ids)
    items: List[dict] = []
    continuation_token = ""

    while True:
        payload = {
            "beneficiaries": [
                {
                    "identitytype": "b2b",
                    "identityValue": user_store_id,
                    "localTicketReference": "",
                }
            ],
            "productSkuIds": pairs,
            "excludeDuplicates": True,
            "maxPageSize": 200,
        }
        if continuation_token:
            payload["continuationToken"] = continuation_token
        sandbox_value = sandbox or SETTINGS.store_sandbox
        if sandbox_value:
            payload["sbx"] = sandbox_value

        data = _request_collections(payload)
        page_items = data.get("items") or []
        if isinstance(page_items, list):
            items.extend(page_items)

        continuation_token = str(data.get("continuationToken", "")).strip()
        if not continuation_token:
            break

    return items


def check_entitlement(
    *,
    user_store_id: str,
    product_ids: List[str],
    sku_ids: Optional[List[str]] = None,
    sandbox: str | None = None,
) -> EntitlementCheck:
    if not product_ids:
        raise ValueError("No product IDs provided for entitlement check.")
    items = query_entitlements(
        user_store_id=user_store_id,
        product_ids=product_ids,
        sku_ids=sku_ids,
        sandbox=sandbox,
    )
    matched = []
    for item in items:
        product_id = str(item.get("productId", "")).strip()
        status = str(item.get("status", "")).strip().lower()
        if product_id and product_id in product_ids and status == "active":
            matched.append(product_id)
    return EntitlementCheck(active=bool(matched), matched_product_ids=matched, raw_items=items)
