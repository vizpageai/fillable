from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Iterable, List


def _split_csv(value: str) -> List[str]:
    return [item.strip() for item in value.split(",") if item.strip()]


def _truthy(value: str) -> bool:
    return value.strip().lower() in {"1", "true", "yes", "on"}


@dataclass(frozen=True)
class Settings:
    host: str = os.getenv("HOST", "0.0.0.0")
    port: int = int(os.getenv("PORT", "8080"))
    log_level: str = os.getenv("LOG_LEVEL", "info")

    openai_api_key: str = os.getenv("OPENAI_API_KEY", "")
    openai_base_url: str = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1").rstrip("/")
    openai_model_default: str = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")

    aad_tenant_id: str = os.getenv("AAD_TENANT_ID", "")
    aad_client_id: str = os.getenv("AAD_CLIENT_ID", "")
    aad_client_secret: str = os.getenv("AAD_CLIENT_SECRET", "")
    aad_token_url: str = os.getenv("AAD_TOKEN_URL", "")

    store_collections_endpoint: str = os.getenv(
        "STORE_COLLECTIONS_ENDPOINT",
        "https://collections.mp.microsoft.com/v9.0/collections/publisherQuery",
    )
    store_sandbox: str = os.getenv("STORE_SANDBOX", "").strip()
    store_product_ids: List[str] = field(
        default_factory=lambda: _split_csv(os.getenv("STORE_PRODUCT_IDS", ""))
    )
    store_sku_ids: List[str] = field(
        default_factory=lambda: _split_csv(os.getenv("STORE_SKU_IDS", ""))
    )
    store_user_agent: str = os.getenv("STORE_USER_AGENT", "FillableStoreBackend/1.0")
    store_audience: str = os.getenv("STORE_AUDIENCE", "https://onestore.microsoft.com")
    store_create_collections_audience: str = os.getenv(
        "STORE_CREATE_COLLECTIONS_AUDIENCE",
        "https://onestore.microsoft.com/b2b/keys/create/collections",
    )

    app_shared_secret: str = os.getenv("APP_SHARED_SECRET", "")
    require_entitlement: bool = _truthy(os.getenv("REQUIRE_ENTITLEMENT", "true"))

    def resolve_token_url(self) -> str:
        if self.aad_token_url.strip():
            return self.aad_token_url.strip()
        if not self.aad_tenant_id:
            return ""
        return f"https://login.microsoftonline.com/{self.aad_tenant_id}/oauth2/token"


SETTINGS = Settings()


def require_values(values: Iterable[str], *, label: str) -> None:
    missing = [name for name in values if not name]
    if missing:
        raise ValueError(f"Missing required {label} configuration.")
