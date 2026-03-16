# Fillable Store Backend

FastAPI service that validates Microsoft Store subscription entitlements and proxies OpenAI
requests using your OpenAI API key. This is intended for MSIX-packaged apps that sell a
subscription add-on through the Microsoft Store while keeping the "bring your own API key"
flow free in the client.

## Required setup (Store + Entra ID)

1. Create a Store subscription add-on in Partner Center and note its `productId` (and `skuId` if needed).
2. Register an Entra ID (Azure AD) app for Store services access.
3. In Partner Center, link the Entra ID app to your Store app.
4. Create a client secret for the Entra ID app.

The backend uses the Store services Collections API. The client should use
`StoreContext.GetCustomerCollectionsIdAsync` to obtain a User Store ID (Collections ID) by
calling `/v1/store/collections-token` and passing the returned token into the StoreContext call.
Then send the User Store ID to the backend for entitlement checks and OpenAI proxy calls.

## Environment

Set these in your server environment:

```
OPENAI_API_KEY=sk-...
OPENAI_BASE_URL=https://api.openai.com/v1
OPENAI_MODEL=gpt-4.1-mini

AAD_TENANT_ID=...
AAD_CLIENT_ID=...
AAD_CLIENT_SECRET=...

STORE_PRODUCT_IDS=9NXXXXXXXXXX
STORE_SKU_IDS=   # optional, comma-separated
STORE_COLLECTIONS_ENDPOINT=https://collections.mp.microsoft.com/v9.0/collections/publisherQuery
STORE_CREATE_COLLECTIONS_AUDIENCE=https://onestore.microsoft.com/b2b/keys/create/collections
STORE_AUDIENCE=https://onestore.microsoft.com
STORE_SANDBOX=   # optional for Store sandbox

APP_SHARED_SECRET=your-shared-secret
REQUIRE_ENTITLEMENT=true
```

## Run locally

```
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
uvicorn backend.main:app --host 0.0.0.0 --port 8080
```

## Endpoints

- `POST /v1/store/collections-token`
  - Returns a short-lived token the client uses with `StoreContext.GetCustomerCollectionsIdAsync`.
- `POST /v1/store/entitlement-check`
  - Body: `{ "user_store_id": "...", "product_ids": ["..."] }`
- `POST /v1/openai-proxy`
  - Body: `{ "payload": { ...chat-completions... }, "user_store_id": "..." }`

If `APP_SHARED_SECRET` is set, include `X-App-Secret` on all calls.
