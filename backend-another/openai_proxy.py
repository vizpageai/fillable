from __future__ import annotations

import httpx

from config import SETTINGS


def call_openai(payload: dict) -> dict:
    if not SETTINGS.openai_api_key:
        raise ValueError("OPENAI_API_KEY is required.")
    url = f"{SETTINGS.openai_base_url}/chat/completions"
    headers = {
        "Authorization": f"Bearer {SETTINGS.openai_api_key}",
        "Content-Type": "application/json",
    }
    with httpx.Client(timeout=120) as client:
        response = client.post(url, json=payload, headers=headers)
    if response.status_code >= 400:
        raise RuntimeError(f"OpenAI error {response.status_code}: {response.text[:1200]}")
    data = response.json()
    if not isinstance(data, dict):
        raise RuntimeError("OpenAI response was not a JSON object.")
    return data
