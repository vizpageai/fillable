# Backend (Firebase Auth + Stripe Subscriptions)

FastAPI backend that authenticates users with Firebase ID tokens and gates access
to your OpenAI proxy by a credit balance. Users buy credits via Stripe.

## Requirements

- Firebase project (Google login + Microsoft login enabled)
- Stripe one-time product + price (e.g., $10 for 100 credits)
- OpenAI API key (server-side)

## Environment

```
OPENAI_API_KEY=***REMOVED***
OPENAI_BASE_URL=https://api.openai.com/v1

STRIPE_SECRET_KEY=sk_live_...
STRIPE_WEBHOOK_SECRET=whsec_...
STRIPE_PRICE_ID=price_...
STRIPE_SUCCESS_URL=https://your-domain/billing/success
STRIPE_CANCEL_URL=https://your-domain/billing/cancel
STRIPE_PORTAL_RETURN_URL=https://your-app/account

# Firebase service account JSON or path
FIREBASE_CREDENTIALS_JSON=
# or:
GOOGLE_APPLICATION_CREDENTIALS=C:\Users\E20263395\Apps\vizpage-backend\courai-firebase-adminsdk-fbsvc-0a3f09f1ba.json

FIREBASE_PROJECT_ID=courai
REQUIRE_EMAIL_VERIFIED=false
```

$apiKey = "AIzaSyBZzNCcFybRrIxQilLrWfopI22LxGe7d1g"
  $email = "hxk111@yeah.net"
  $password = "123456"

  $body = @{
    email = $email
    password = $password
    returnSecureToken = $true
  } | ConvertTo-Json

  $resp = Invoke-RestMethod -Method Post `
    -Uri "https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=$apiKey" `
    -ContentType "application/json" `
    -Body $body

  $idToken = $resp.idToken


## Run locally

```
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
uvicorn main:app --host 0.0.0.0 --port 8081
```

## API usage (from client)

All endpoints require `Authorization: Bearer <firebase_id_token>`.

- `POST /v1/billing/create-checkout-session`
  - Returns Stripe Checkout URL for a one-time credit purchase.
- `GET /v1/credits`
  - Returns `{ credits: <number> }`.
- `POST /v1/openai-proxy`
  - Body: `{ "payload": { ...chat-completions... } }`

Stripe webhook endpoint:
- `POST /webhooks/stripe`
  - Used for Stripe events. This backend currently checks entitlement live via Stripe, so no persistence is required.
