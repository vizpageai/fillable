from __future__ import annotations

from fastapi import FastAPI, Header, HTTPException

from .config import SETTINGS
from .models import (
    CollectionsTokenResponse,
    EntitlementRequest,
    EntitlementResponse,
    OpenAIProxyRequest,
)
from .openai_proxy import call_openai
from .store_services import check_entitlement, get_create_collections_token


app = FastAPI(title="Fillable Store Backend", version="1.0.0")


def _require_app_secret(x_app_secret: str | None) -> None:
    if not SETTINGS.app_shared_secret:
        return
    if not x_app_secret or x_app_secret.strip() != SETTINGS.app_shared_secret:
        raise HTTPException(status_code=401, detail="Invalid app secret.")


@app.get("/health")
def health() -> dict:
    return {"ok": True}


@app.post("/v1/store/collections-token", response_model=CollectionsTokenResponse)
def collections_token(x_app_secret: str | None = Header(default=None)) -> CollectionsTokenResponse:
    _require_app_secret(x_app_secret)
    token = get_create_collections_token()
    return CollectionsTokenResponse(
        access_token=token,
        expires_in=3600,
        audience=SETTINGS.store_create_collections_audience,
    )


@app.post("/v1/store/entitlement-check", response_model=EntitlementResponse)
def entitlement_check(
    request: EntitlementRequest,
    x_app_secret: str | None = Header(default=None),
) -> EntitlementResponse:
    _require_app_secret(x_app_secret)
    product_ids = request.product_ids or SETTINGS.store_product_ids
    sku_ids = request.sku_ids or SETTINGS.store_sku_ids
    if not product_ids:
        raise HTTPException(status_code=400, detail="No product IDs configured or provided.")
    result = check_entitlement(
        user_store_id=request.user_store_id,
        product_ids=product_ids,
        sku_ids=sku_ids or None,
        sandbox=request.sandbox,
    )
    return EntitlementResponse(
        active=result.active,
        matched_product_ids=result.matched_product_ids,
        items=result.raw_items,
    )


@app.post("/v1/openai-proxy")
def openai_proxy(
    request: OpenAIProxyRequest,
    x_user_store_id: str | None = Header(default=None),
    x_app_secret: str | None = Header(default=None),
) -> dict:
    _require_app_secret(x_app_secret)
    user_store_id = (x_user_store_id or request.user_store_id or "").strip()
    if SETTINGS.require_entitlement:
        if not user_store_id:
            raise HTTPException(status_code=401, detail="Missing user_store_id for entitlement check.")
        product_ids = SETTINGS.store_product_ids
        if not product_ids:
            raise HTTPException(status_code=500, detail="STORE_PRODUCT_IDS not configured.")
        entitlement = check_entitlement(
            user_store_id=user_store_id,
            product_ids=product_ids,
            sku_ids=SETTINGS.store_sku_ids or None,
            sandbox=request.sandbox,
        )
        if not entitlement.active:
            raise HTTPException(status_code=402, detail="Subscription inactive.")

    payload = dict(request.payload)
    if "model" not in payload:
        payload["model"] = SETTINGS.openai_model_default
    return call_openai(payload)
