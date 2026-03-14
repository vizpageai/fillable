from __future__ import annotations

from typing import Optional, Tuple

import stripe

from config import SETTINGS


def init_stripe() -> None:
    if not SETTINGS.stripe_secret_key:
        raise ValueError("STRIPE_SECRET_KEY is required.")
    stripe.api_key = SETTINGS.stripe_secret_key


def create_checkout_session(
    *,
    uid: str,
    email: Optional[str],
) -> stripe.checkout.Session:
    init_stripe()
    if not SETTINGS.stripe_price_id:
        raise ValueError("STRIPE_PRICE_ID is required.")
    if not SETTINGS.stripe_success_url or not SETTINGS.stripe_cancel_url:
        raise ValueError("STRIPE_SUCCESS_URL and STRIPE_CANCEL_URL are required.")
    return stripe.checkout.Session.create(
        mode="payment",
        line_items=[{"price": SETTINGS.stripe_price_id, "quantity": 1}],
        success_url=SETTINGS.stripe_success_url,
        cancel_url=SETTINGS.stripe_cancel_url,
        client_reference_id=uid,
        customer_email=email or None,
        metadata={"firebase_uid": uid},
        payment_intent_data={"metadata": {"firebase_uid": uid}},
    )


def _find_customer_for_uid(uid: str) -> Optional[stripe.Customer]:
    init_stripe()
    try:
        result = stripe.Customer.search(query=f"metadata['firebase_uid']:'{uid}'", limit=1)
        data = list(result.get("data", []))
        return data[0] if data else None
    except Exception:
        pass
    return None


def _find_customer_for_email(email: str) -> Optional[stripe.Customer]:
    init_stripe()
    if not email:
        return None
    try:
        result = stripe.Customer.search(query=f"email:'{email}'", limit=1)
        data = list(result.get("data", []))
        return data[0] if data else None
    except Exception:
        return None


def _find_subscription_for_uid(uid: str) -> Optional[dict]:
    init_stripe()
    try:
        result = stripe.Subscription.search(
            query=f"metadata['firebase_uid']:'{uid}'",
            limit=1,
        )
        data = list(result.get("data", []))
        return data[0] if data else None
    except Exception:
        return None


def get_subscription_status(uid: str, email: str | None = None) -> Tuple[bool, list[dict]]:
    init_stripe()
    customer = _find_customer_for_uid(uid)
    items: list[dict] = []
    if customer:
        subs = stripe.Subscription.list(customer=customer["id"], status="all", limit=50)
        items = list(subs.get("data", []))
    else:
        sub = _find_subscription_for_uid(uid)
        if sub:
            items = [sub]
        elif email:
            customer = _find_customer_for_email(email)
            if customer:
                subs = stripe.Subscription.list(customer=customer["id"], status="all", limit=50)
                items = list(subs.get("data", []))
    active = any(s.get("status") in {"active", "trialing"} for s in items)
    return active, items
