# Subscription Backend API Contract

Use this backend when app mode is `app_subscription`.

## Security model
- Store your OpenAI API key only on backend servers.
- Client never receives your OpenAI key.
- Client sends a short-lived subscription token (issued by your auth service).
- Backend validates Microsoft Store subscription entitlement before serving AI requests.

## Endpoint
`POST /v1/openai-proxy`

Headers:
- `Authorization: Bearer <subscription_token>`
- `Content-Type: application/json`

Request body:
```json
{
  "model": "gpt-4.1-mini",
  "prompt": "string",
  "response_format": "json_object"
}
```

Success response:
```json
{
  "output_text": "{\"values\":{\"FIELD\":\"value\"}}"
}
```

Error response:
```json
{
  "error": "subscription_inactive"
}
```

## Validation rules
- Reject if token is invalid or expired.
- Reject if entitlement is inactive or canceled.
- Enforce per-user rate and token limits.
- Log request IDs and usage for audit.

## Microsoft Store policy alignment
- Digital subscription should be purchased/managed according to current Store terms for your app type.
- If using Store add-ons/subscriptions, entitlement checks must be authoritative on backend.
- Clearly disclose billing model in Store listing and in-app settings.
