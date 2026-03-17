from __future__ import annotations

import json

from fastapi import FastAPI, Header, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse
import stripe

from auth_service import AuthServiceError, refresh_id_token, sign_in_with_password
from config import SETTINGS
from firestore_db import (
    add_credits,
    claim_webhook_event,
    deduct_credits,
    get_auth_session,
    get_balance,
    set_email,
    store_auth_session,
)
from firebase_auth import init_firebase, verify_firebase_token
from models import (
    AuthRequest,
    AuthResponse,
    CheckoutSessionResponse,
    CreditsResponse,
    EntitlementResponse,
    ProxyRequest,
    RefreshRequest,
)
from openai_proxy import call_openai
from stripe_service import create_checkout_session, init_stripe
from firebase_admin import auth as firebase_admin_auth


app = FastAPI(title="Fillable Firebase + Stripe Backend", version="1.0.0")



def _require_user(authorization: str | None) -> tuple[str, str | None]:
    if not authorization or not authorization.lower().startswith("bearer "):
        raise HTTPException(status_code=401, detail="Missing Authorization bearer token.")
    token = authorization.split(" ", 1)[1].strip()
    try:
        user = verify_firebase_token(token)
    except Exception as exc:
        raise HTTPException(status_code=401, detail=str(exc)) from exc
    if user.email:
        set_email(user.uid, user.email)
    return user.uid, user.email


def _validate_password(password: str) -> None:
    if len(password) < 8:
        raise HTTPException(status_code=400, detail="Password must be at least 8 characters.")
    if password.lower() == password or password.upper() == password:
        raise HTTPException(status_code=400, detail="Password must include upper and lower case letters.")
    if not any(ch.isdigit() for ch in password):
        raise HTTPException(status_code=400, detail="Password must include a number.")
    if not any(not ch.isalnum() for ch in password):
        raise HTTPException(status_code=400, detail="Password must include a special character.")


@app.get("/health")
def health() -> dict:
    return {"ok": True}


@app.get("/billing/success", response_class=HTMLResponse)
def billing_success() -> str:
    return """
    <!doctype html>
    <html><head><meta charset="utf-8"><title>Success</title></head>
    <body>
      <h2>Payment successful</h2>
      <p>You can close this window.</p>
      <script>setTimeout(() => window.close(), 1200);</script>
    </body></html>
    """


@app.get("/billing/cancel", response_class=HTMLResponse)
def billing_cancel() -> str:
    return """
    <!doctype html>
    <html><head><meta charset="utf-8"><title>Cancelled</title></head>
    <body>
      <h2>Payment cancelled</h2>
      <p>You can close this window.</p>
      <script>setTimeout(() => window.close(), 1200);</script>
    </body></html>
    """


@app.post("/v1/auth/register", response_model=AuthResponse)
def register_user(request: AuthRequest) -> AuthResponse:
    init_firebase()
    email = request.email.strip().lower()
    password = request.password.strip()
    if not email or not password:
        raise HTTPException(status_code=400, detail="Email and password are required.")
    _validate_password(password)
    try:
        user = firebase_admin_auth.create_user(email=email, password=password)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    try:
        session = sign_in_with_password(email, password)
    except AuthServiceError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return AuthResponse(
        id_token=str(session.get("idToken", "")),
        refresh_token=str(session.get("refreshToken", "")),
        expires_in=int(session.get("expiresIn", 0) or 0),
        email=str(session.get("email", email)),
        uid=str(session.get("localId", user.uid)),
    )


@app.post("/v1/auth/login", response_model=AuthResponse)
def login_user(request: AuthRequest) -> AuthResponse:
    email = request.email.strip().lower()
    password = request.password.strip()
    if not email or not password:
        raise HTTPException(status_code=400, detail="Email and password are required.")
    try:
        session = sign_in_with_password(email, password)
    except AuthServiceError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return AuthResponse(
        id_token=str(session.get("idToken", "")),
        refresh_token=str(session.get("refreshToken", "")),
        expires_in=int(session.get("expiresIn", 0) or 0),
        email=str(session.get("email", email)),
        uid=str(session.get("localId", "")),
    )


@app.post("/v1/auth/refresh", response_model=AuthResponse)
def refresh_user(request: RefreshRequest) -> AuthResponse:
    refresh_token = request.refresh_token.strip()
    if not refresh_token:
        raise HTTPException(status_code=400, detail="Refresh token is required.")
    try:
        session = refresh_id_token(refresh_token)
    except AuthServiceError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return AuthResponse(
        id_token=str(session.get("id_token", "")),
        refresh_token=str(session.get("refresh_token", "")),
        expires_in=int(session.get("expires_in", 0) or 0),
        email="",
        uid=str(session.get("user_id", "")),
    )

@app.post("/v1/billing/create-checkout-session", response_model=CheckoutSessionResponse)
def create_checkout(authorization: str | None = Header(default=None)) -> CheckoutSessionResponse:
    uid, email = _require_user(authorization)
    session = create_checkout_session(uid=uid, email=email)
    return CheckoutSessionResponse(url=str(session.url), id=str(session.id))


@app.get("/v1/credits", response_model=CreditsResponse)
def credits(authorization: str | None = Header(default=None)) -> CreditsResponse:
    uid, _ = _require_user(authorization)
    balance, _ = get_balance(uid)
    return CreditsResponse(credits=balance)


@app.get("/v1/entitlement", response_model=EntitlementResponse)
def entitlement(authorization: str | None = Header(default=None)) -> EntitlementResponse:
    uid, email = _require_user(authorization)
    balance, _ = get_balance(uid)
    active = balance > 0
    return EntitlementResponse(active=active, subscriptions=[], credits=balance)


@app.post("/v1/openai-proxy")
def openai_proxy(
    request: ProxyRequest,
    authorization: str | None = Header(default=None),
) -> dict:
    uid, email = _require_user(authorization)
    balance, _ = get_balance(uid)
    if balance <= 0:
        raise HTTPException(status_code=402, detail="Insufficient credits.")
    payload = dict(request.payload)
    if "model" not in payload:
        payload["model"] = request.model_fallback or "gpt-4.1-mini"
    response = call_openai(payload)
    usage = response.get("usage") if isinstance(response, dict) else None
    output_tokens = 0
    if isinstance(usage, dict):
        output_tokens = int(
            usage.get("completion_tokens")
            or usage.get("output_tokens")
            or usage.get("total_tokens")
            or 0
        )
    if output_tokens <= 0:
        try:
            content = response["choices"][0]["message"]["content"]
            output_tokens = max(1, int(len(str(content)) / 4))
        except Exception:
            output_tokens = 1
    credits_used = output_tokens * 0.0002
    remaining = deduct_credits(uid, credits_used)
    response["credits_used"] = round(credits_used, 6)
    response["credits_remaining"] = round(remaining, 6)
    return response


@app.post("/webhooks/stripe")
async def stripe_webhook(request: Request) -> JSONResponse:
    payload = await request.body()
    sig = request.headers.get("stripe-signature", "")
    if not SETTINGS.stripe_webhook_secret:
        raise HTTPException(status_code=500, detail="STRIPE_WEBHOOK_SECRET not configured.")
    try:
        event = stripe.Webhook.construct_event(payload, sig, SETTINGS.stripe_webhook_secret)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Invalid signature: {exc}") from exc

    init_stripe()
    event_type = event.get("type", "")
    event_id = str(event.get("id", "")).strip()
    if event_id and not claim_webhook_event(event_id):
        return JSONResponse({"received": True, "deduped": True})
    if event_type == "checkout.session.completed":
        session = event.get("data", {}).get("object", {}) or {}
        firebase_uid = str(session.get("client_reference_id", "")).strip()
        customer_id = session.get("customer")
        amount_total = session.get("amount_total") or session.get("amount_subtotal") or 0
        if firebase_uid and customer_id:
            try:
                stripe.Customer.modify(customer_id, metadata={"firebase_uid": firebase_uid})
            except Exception:
                pass
        if firebase_uid and amount_total:
            credits = (float(amount_total) / 100.0) * 10.0
            add_credits(firebase_uid, credits)
    return JSONResponse({"received": True})
@app.get("/auth/google", response_class=HTMLResponse)
def auth_google(session_id: str) -> str:
    if not session_id or len(session_id) < 8:
        raise HTTPException(status_code=400, detail="Missing session_id.")
    config = SETTINGS.firebase_web_config()
    if not config.get("apiKey") or not config.get("projectId"):
        raise HTTPException(status_code=500, detail="Firebase web config missing.")
    config_json = json.dumps(config)
    return f"""<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Sign in</title>
  <script src="https://www.gstatic.com/firebasejs/9.22.2/firebase-app-compat.js"></script>
  <script src="https://www.gstatic.com/firebasejs/9.22.2/firebase-auth-compat.js"></script>
</head>
<body>
  <p>Signing you in with Google...</p>
  <script>
    const firebaseConfig = {config_json};
    firebase.initializeApp(firebaseConfig);
    const auth = firebase.auth();
    const provider = new firebase.auth.GoogleAuthProvider();
    auth.signInWithPopup(provider).then(async (result) => {{
      const user = result.user;
      const idToken = await user.getIdToken();
      const payload = {{
        session_id: "{session_id}",
        id_token: idToken,
        refresh_token: user.refreshToken || "",
        email: user.email || "",
        uid: user.uid || ""
      }};
      await fetch("/v1/auth/complete", {{
        method: "POST",
        headers: {{ "Content-Type": "application/json" }},
        body: JSON.stringify(payload)
      }});
      document.body.innerText = "Signed in. You can close this window.";
      setTimeout(() => window.close(), 1200);
    }}).catch((err) => {{
      document.body.innerText = "Sign-in failed: " + err.message;
    }});
  </script>
</body>
</html>"""


@app.post("/v1/auth/complete")
def auth_complete(payload: dict) -> dict:
    session_id = str(payload.get("session_id", "")).strip()
    id_token = str(payload.get("id_token", "")).strip()
    refresh_token = str(payload.get("refresh_token", "")).strip()
    email = str(payload.get("email", "")).strip()
    uid = str(payload.get("uid", "")).strip()
    if not session_id or not id_token or not uid:
        raise HTTPException(status_code=400, detail="Missing auth payload.")
    try:
        user = verify_firebase_token(id_token)
    except Exception as exc:
        raise HTTPException(status_code=401, detail=str(exc)) from exc
    if user.uid != uid:
        raise HTTPException(status_code=401, detail="Token mismatch.")
    store_auth_session(
        session_id,
        {
            "status": "ok",
            "id_token": id_token,
            "refresh_token": refresh_token,
            "email": email or user.email or "",
            "uid": uid,
        },
    )
    return {"ok": True}


@app.get("/v1/auth/poll")
def auth_poll(session_id: str) -> dict:
    payload = get_auth_session(session_id)
    if not payload:
        return {"status": "pending"}
    return payload
