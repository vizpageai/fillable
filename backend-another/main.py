from __future__ import annotations

import base64
import html
import io
import json
from pathlib import Path

from fastapi import FastAPI, Header, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse
import qrcode
from qrcode.image.svg import SvgPathImage
import stripe

from auth_service import AuthServiceError, refresh_id_token, sign_in_with_password
from config import SETTINGS
from firestore_db import (
    add_credits,
    claim_webhook_event,
    deduct_credits,
    get_auth_session,
    get_balance,
    save_feedback_submission,
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
    FeedbackRequest,
    ProxyRequest,
    RefreshRequest,
)
from openai_proxy import call_openai
from stripe_service import create_checkout_session, init_stripe
from firebase_admin import auth as firebase_admin_auth


app = FastAPI(title="Fillable Firebase + Stripe Backend", version="1.0.0")

STORE_URL = "https://apps.microsoft.com/detail/9pnr62jk2j1m?hl=en-US&gl=US"
APP_NAME = "Fillable-doc"
CONTACT_NAME = "Alex Huo"
CONTACT_EMAIL = "alexhuo@vizpageai.com"


def _svg_icon_data_uri() -> str:
    icon_path = Path(__file__).resolve().parents[1] / "fillableicon.svg"
    if not icon_path.exists():
        return ""
    svg_text = icon_path.read_text(encoding="utf-8")
    encoded = base64.b64encode(svg_text.encode("utf-8")).decode("ascii")
    return f"data:image/svg+xml;base64,{encoded}"


def _store_qr_data_uri() -> str:
    qr = qrcode.QRCode(border=1, box_size=8)
    qr.add_data(STORE_URL)
    qr.make(fit=True)
    img = qr.make_image(image_factory=SvgPathImage)
    output = io.BytesIO()
    img.save(output)
    encoded = base64.b64encode(output.getvalue()).decode("ascii")
    return f"data:image/svg+xml;base64,{encoded}"


def _marketing_page(message: str = "") -> str:
    icon_data_uri = _svg_icon_data_uri()
    qr_data_uri = _store_qr_data_uri()
    safe_message = html.escape(message)

    status_markup = (
        f'<div class="notice-bar"><svg width="16" height="16" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="7" stroke="#1a7f4e" stroke-width="1.5"/><path d="M5 8l2 2 4-4" stroke="#1a7f4e" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>{safe_message}</div>'
        if safe_message else ""
    )

    icon_markup = (
        f'<img src="{icon_data_uri}" alt="{APP_NAME} icon" class="brand-icon" />'
        if icon_data_uri
        else '<div class="brand-icon brand-icon-fallback"><svg viewBox="0 0 32 32" fill="none" xmlns="http://www.w3.org/2000/svg"><rect x="4" y="2" width="17" height="22" rx="2.5" fill="#f0f4ff"/><rect x="7" y="7" width="11" height="1.5" rx=".75" fill="#3b5bdb"/><rect x="7" y="11" width="8" height="1.5" rx=".75" fill="#3b5bdb" opacity=".5"/><rect x="7" y="15" width="9.5" height="1.5" rx=".75" fill="#3b5bdb" opacity=".5"/><circle cx="24" cy="24" r="7" fill="#2f9e44"/><path d="M21 24l2.2 2.2 3.8-3.8" stroke="white" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/></svg></div>'
    )

    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{APP_NAME} — AI Document Filling for Windows</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@300;400;500;600;700;800&family=Lora:ital,wght@0,400;0,600;1,400&display=swap" rel="stylesheet">
  <style>
    :root {{
      --white:   #ffffff;
      --off:     #f7f7f5;
      --border:  #e8e8e4;
      --ink:     #111110;
      --ink-2:   #3a3a38;
      --ink-3:   #6b6b68;
      --ink-4:   #a8a8a4;
      --blue:    #1d4ed8;
      --blue-lt: #eff3ff;
      --blue-md: #dbeafe;
      --green:   #2f9e44;
      --shadow-sm: 0 1px 3px rgba(0,0,0,.06), 0 1px 2px rgba(0,0,0,.04);
      --shadow-md: 0 4px 16px rgba(0,0,0,.07), 0 1px 4px rgba(0,0,0,.04);
      --shadow-lg: 0 12px 40px rgba(0,0,0,.09), 0 2px 8px rgba(0,0,0,.04);
      --r: 14px;
    }}

    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    html {{ scroll-behavior: smooth; -webkit-font-smoothing: antialiased; }}

    body {{
      font-family: 'Manrope', sans-serif;
      background: var(--white);
      color: var(--ink);
      font-size: 16px;
      line-height: 1.6;
    }}

    a {{ color: inherit; text-decoration: none; }}

    /* ── Wrapper ── */
    .wrap {{
      width: min(1080px, calc(100% - 48px));
      margin: 0 auto;
    }}

    /* ── Notice bar ── */
    .notice-bar {{
      background: #f0faf4;
      border-bottom: 1px solid #c3e6cc;
      color: #1a7f4e;
      font-size: .82rem;
      font-weight: 600;
      padding: 10px 24px;
      display: flex;
      align-items: center;
      gap: 8px;
      justify-content: center;
    }}

    /* ── Nav ── */
    nav {{
      position: sticky;
      top: 0;
      z-index: 100;
      background: rgba(255,255,255,0.92);
      backdrop-filter: blur(16px);
      -webkit-backdrop-filter: blur(16px);
      border-bottom: 1px solid var(--border);
    }}
    .nav-inner {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      height: 60px;
    }}
    .nav-logo {{
      display: flex;
      align-items: center;
      gap: 10px;
    }}
    .brand-icon {{
      width: 32px;
      height: 32px;
      border-radius: 8px;
    }}
    .brand-icon-fallback {{
      width: 32px;
      height: 32px;
      border-radius: 8px;
      background: var(--blue-lt);
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 5px;
    }}
    .nav-name {{
      font-size: .9rem;
      font-weight: 700;
      letter-spacing: -.01em;
      color: var(--ink);
    }}
    .nav-links {{
      display: flex;
      align-items: center;
      gap: 4px;
    }}
    .nav-link {{
      font-size: .82rem;
      font-weight: 500;
      color: var(--ink-3);
      padding: 6px 14px;
      border-radius: 8px;
      transition: color .15s, background .15s;
    }}
    .nav-link:hover {{ color: var(--ink); background: var(--off); }}
    .nav-cta {{
      font-size: .82rem;
      font-weight: 700;
      color: white;
      background: var(--ink);
      padding: 8px 18px;
      border-radius: 999px;
      transition: opacity .15s;
    }}
    .nav-cta:hover {{ opacity: .82; }}

    /* ── Hero ── */
    .hero {{
      padding: 100px 0 80px;
      text-align: center;
    }}
    .hero-label {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
      font-size: .72rem;
      font-weight: 700;
      letter-spacing: .1em;
      text-transform: uppercase;
      color: var(--blue);
      background: var(--blue-lt);
      border: 1px solid var(--blue-md);
      padding: 5px 14px;
      border-radius: 999px;
      margin-bottom: 32px;
    }}
    .hero-label svg {{ flex-shrink: 0; }}
    h1 {{
      font-family: 'Manrope', sans-serif;
      font-size: clamp(2.6rem, 5.5vw, 4.8rem);
      font-weight: 800;
      line-height: 1.05;
      letter-spacing: -.04em;
      color: var(--ink);
      max-width: 820px;
      margin: 0 auto 28px;
    }}
    h1 em {{
      font-family: 'Lora', Georgia, serif;
      font-style: italic;
      font-weight: 400;
      color: var(--blue);
    }}
    .hero-sub {{
      font-size: 1.1rem;
      font-weight: 400;
      color: var(--ink-3);
      max-width: 520px;
      margin: 0 auto 44px;
      line-height: 1.7;
    }}
    .hero-actions {{
      display: flex;
      justify-content: center;
      align-items: center;
      gap: 12px;
      flex-wrap: wrap;
      margin-bottom: 64px;
    }}
    .btn-primary {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      background: var(--ink);
      color: white;
      font-family: 'Manrope', sans-serif;
      font-size: .9rem;
      font-weight: 700;
      padding: 14px 28px;
      border-radius: 999px;
      border: none;
      cursor: pointer;
      transition: opacity .15s, transform .15s;
      box-shadow: 0 2px 8px rgba(0,0,0,.18);
    }}
    .btn-primary:hover {{ opacity: .84; transform: translateY(-1px); }}
    .btn-secondary {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      background: var(--white);
      color: var(--ink-2);
      font-family: 'Manrope', sans-serif;
      font-size: .9rem;
      font-weight: 600;
      padding: 14px 28px;
      border-radius: 999px;
      border: 1.5px solid var(--border);
      cursor: pointer;
      transition: border-color .15s, background .15s;
    }}
    .btn-secondary:hover {{ border-color: #c0c0bc; background: var(--off); }}

    /* ── Hero visual: doc mockup ── */
    .hero-visual {{
      max-width: 720px;
      margin: 0 auto;
      background: var(--off);
      border: 1px solid var(--border);
      border-radius: 20px;
      overflow: hidden;
      box-shadow: var(--shadow-lg), 0 0 0 1px rgba(0,0,0,.03);
    }}
    .mockup-bar {{
      background: var(--white);
      border-bottom: 1px solid var(--border);
      padding: 12px 18px;
      display: flex;
      align-items: center;
      gap: 8px;
    }}
    .dot {{ width: 10px; height: 10px; border-radius: 50%; }}
    .dot-r {{ background: #ff5f57; }}
    .dot-y {{ background: #febc2e; }}
    .dot-g {{ background: #28c840; }}
    .mockup-title {{
      font-size: .75rem;
      font-weight: 600;
      color: var(--ink-4);
      margin-left: 8px;
    }}
    .mockup-body {{
      padding: 32px 36px 28px;
      display: grid;
      grid-template-columns: 1fr 200px;
      gap: 24px;
      align-items: start;
    }}
    .mockup-doc {{
      background: white;
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 22px 24px;
      box-shadow: var(--shadow-sm);
    }}
    .doc-heading {{ width: 60%; height: 10px; background: var(--ink); border-radius: 5px; margin-bottom: 16px; }}
    .doc-line {{ height: 7px; border-radius: 4px; background: var(--border); margin-bottom: 9px; }}
    .doc-line.w-full {{ width: 100%; }}
    .doc-line.w-9  {{ width: 90%; }}
    .doc-line.w-7  {{ width: 70%; }}
    .doc-line.w-8  {{ width: 80%; }}
    .doc-field {{
      margin-top: 18px;
      border: 1.5px dashed #93c5fd;
      border-radius: 6px;
      padding: 8px 12px;
      display: flex;
      align-items: center;
      gap: 8px;
    }}
    .doc-field-label {{ font-size: .65rem; font-weight: 700; color: var(--blue); text-transform: uppercase; letter-spacing: .06em; }}
    .doc-field-val {{ font-size: .75rem; color: var(--ink-3); font-family: 'Lora', serif; font-style: italic; }}
    .doc-ai-badge {{
      display: inline-flex;
      align-items: center;
      gap: 5px;
      font-size: .6rem;
      font-weight: 700;
      color: var(--green);
      background: #f0faf4;
      border: 1px solid #c3e6cc;
      padding: 3px 8px;
      border-radius: 999px;
      margin-left: auto;
    }}
    .mockup-sidebar {{
      display: flex;
      flex-direction: column;
      gap: 10px;
    }}
    .side-card {{
      background: white;
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 14px 16px;
      box-shadow: var(--shadow-sm);
    }}
    .side-card-label {{
      font-size: .65rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: .08em;
      color: var(--ink-4);
      margin-bottom: 8px;
    }}
    .side-pill {{
      display: inline-block;
      font-size: .68rem;
      font-weight: 700;
      padding: 3px 8px;
      border-radius: 5px;
      margin: 2px;
      text-transform: uppercase;
      letter-spacing: .04em;
    }}
    .pill-b {{ background: var(--blue-lt); color: var(--blue); }}
    .pill-g {{ background: #f0faf4; color: var(--green); }}
    .pill-o {{ background: #fff7ed; color: #c2410c; }}
    .side-prog-label {{ font-size: .7rem; color: var(--ink-3); margin-bottom: 6px; }}
    .side-prog-bar {{ height: 5px; background: var(--border); border-radius: 999px; overflow: hidden; }}
    .side-prog-fill {{ height: 100%; background: var(--blue); border-radius: 999px; width: 72%; }}

    /* ── Stats strip ── */
    .stats-strip {{
      border-top: 1px solid var(--border);
      border-bottom: 1px solid var(--border);
      padding: 32px 0;
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 0;
    }}
    .stat-item {{
      text-align: center;
      padding: 0 24px;
      border-right: 1px solid var(--border);
    }}
    .stat-item:last-child {{ border-right: none; }}
    .stat-num {{
      font-size: 2rem;
      font-weight: 800;
      letter-spacing: -.04em;
      color: var(--ink);
      line-height: 1;
      margin-bottom: 6px;
    }}
    .stat-label {{ font-size: .78rem; color: var(--ink-3); font-weight: 500; }}

    /* ── Section common ── */
    section {{ padding: 96px 0; }}
    .section-eyebrow {{
      font-size: .72rem;
      font-weight: 700;
      letter-spacing: .1em;
      text-transform: uppercase;
      color: var(--blue);
      margin-bottom: 14px;
    }}
    h2 {{
      font-size: clamp(1.9rem, 3.5vw, 2.8rem);
      font-weight: 800;
      letter-spacing: -.04em;
      line-height: 1.1;
      color: var(--ink);
      margin-bottom: 16px;
    }}
    h2 em {{
      font-family: 'Lora', serif;
      font-style: italic;
      font-weight: 400;
    }}
    .section-sub {{
      font-size: 1rem;
      color: var(--ink-3);
      line-height: 1.7;
      max-width: 480px;
    }}
    .section-header {{ margin-bottom: 56px; }}

    /* ── Features ── */
    .features-section {{ background: var(--off); }}
    .features-grid {{
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 2px;
      background: var(--border);
      border: 1px solid var(--border);
      border-radius: 20px;
      overflow: hidden;
    }}
    .feat {{
      background: var(--white);
      padding: 40px;
      transition: background .2s;
    }}
    .feat:hover {{ background: #fafaf8; }}
    .feat-icon {{
      width: 44px;
      height: 44px;
      border-radius: 12px;
      background: var(--off);
      border: 1px solid var(--border);
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 20px;
      margin-bottom: 22px;
    }}
    .feat-title {{
      font-size: 1rem;
      font-weight: 700;
      color: var(--ink);
      margin-bottom: 10px;
      letter-spacing: -.01em;
    }}
    .feat-desc {{
      font-size: .88rem;
      color: var(--ink-3);
      line-height: 1.7;
    }}

    /* ── How it works ── */
    .steps-grid {{
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 32px;
      counter-reset: steps;
    }}
    .step {{ counter-increment: steps; }}
    .step-num {{
      font-size: .72rem;
      font-weight: 700;
      color: var(--ink-4);
      letter-spacing: .06em;
      text-transform: uppercase;
      margin-bottom: 14px;
    }}
    .step-num::before {{
      content: "0" counter(steps);
    }}
    .step-divider {{
      width: 32px;
      height: 2px;
      background: var(--ink);
      margin-bottom: 18px;
    }}
    .step-title {{
      font-size: .95rem;
      font-weight: 700;
      color: var(--ink);
      margin-bottom: 10px;
      letter-spacing: -.01em;
    }}
    .step-desc {{
      font-size: .84rem;
      color: var(--ink-3);
      line-height: 1.65;
    }}

    /* ── QR / Download section ── */
    .download-section {{
      background: var(--ink);
      color: white;
      border-radius: 28px;
      padding: 72px;
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 64px;
      align-items: center;
      margin: 0 0 96px;
    }}
    .dl-eyebrow {{
      font-size: .72rem;
      font-weight: 700;
      letter-spacing: .1em;
      text-transform: uppercase;
      color: rgba(255,255,255,.45);
      margin-bottom: 14px;
    }}
    .dl-title {{
      font-family: 'Manrope', sans-serif;
      font-size: clamp(1.8rem, 3vw, 2.6rem);
      font-weight: 800;
      letter-spacing: -.04em;
      line-height: 1.1;
      color: white;
      margin-bottom: 16px;
    }}
    .dl-title em {{
      font-family: 'Lora', serif;
      font-style: italic;
      font-weight: 400;
      color: rgba(255,255,255,.6);
    }}
    .dl-sub {{
      font-size: .95rem;
      color: rgba(255,255,255,.5);
      line-height: 1.7;
      max-width: 420px;
      margin-bottom: 32px;
    }}
    .btn-white {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      background: white;
      color: var(--ink);
      font-family: 'Manrope', sans-serif;
      font-size: .9rem;
      font-weight: 700;
      padding: 14px 28px;
      border-radius: 999px;
      border: none;
      cursor: pointer;
      transition: opacity .15s;
    }}
    .btn-white:hover {{ opacity: .88; }}
    .qr-block {{
      text-align: center;
      flex-shrink: 0;
    }}
    .qr-frame {{
      background: white;
      border-radius: 18px;
      padding: 18px;
      display: inline-block;
      box-shadow: 0 8px 32px rgba(0,0,0,.25);
    }}
    .qr-frame img {{
      width: 148px;
      height: 148px;
      display: block;
    }}
    .qr-caption {{
      font-size: .72rem;
      font-weight: 600;
      color: rgba(255,255,255,.4);
      margin-top: 14px;
      text-transform: uppercase;
      letter-spacing: .08em;
    }}

    /* ── Two-col utility ── */
    .two-col {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 64px;
      align-items: start;
    }}

    /* ── FAQ ── */
    .faq-list {{ display: grid; gap: 1px; background: var(--border); border: 1px solid var(--border); border-radius: 16px; overflow: hidden; }}
    .faq-item {{ background: var(--white); padding: 0; }}
    .faq-summary {{
      list-style: none;
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 22px 24px;
      cursor: pointer;
      font-size: .9rem;
      font-weight: 600;
      color: var(--ink);
      user-select: none;
      gap: 16px;
    }}
    .faq-summary::-webkit-details-marker {{ display: none; }}
    .faq-summary:hover {{ background: var(--off); }}
    .faq-chevron {{
      flex-shrink: 0;
      width: 20px; height: 20px;
      border-radius: 50%;
      background: var(--off);
      border: 1px solid var(--border);
      display: flex; align-items: center; justify-content: center;
      transition: transform .2s, background .2s;
    }}
    details[open] .faq-chevron {{ transform: rotate(45deg); background: var(--ink); border-color: var(--ink); }}
    details[open] .faq-chevron svg {{ stroke: white; }}
    .faq-answer {{
      padding: 0 24px 22px;
      font-size: .88rem;
      color: var(--ink-3);
      line-height: 1.7;
      border-top: 1px solid var(--border);
      padding-top: 16px;
    }}

    /* ── Contact ── */
    .contact-block {{
      display: grid;
      gap: 12px;
      margin-bottom: 36px;
    }}
    .contact-row {{
      display: flex;
      align-items: center;
      gap: 14px;
    }}
    .contact-icon {{
      width: 38px; height: 38px;
      border-radius: 10px;
      background: var(--off);
      border: 1px solid var(--border);
      display: flex; align-items: center; justify-content: center;
      font-size: 15px;
      flex-shrink: 0;
    }}
    .contact-text strong {{
      display: block;
      font-size: .7rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: .06em;
      color: var(--ink-4);
      margin-bottom: 1px;
    }}
    .contact-text span, .contact-text a {{
      font-size: .88rem;
      color: var(--ink-2);
    }}
    .contact-text a:hover {{ color: var(--blue); }}

    /* ── Form ── */
    .form {{ display: grid; gap: 14px; }}
    .form-row {{ display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }}
    .field {{ display: flex; flex-direction: column; gap: 5px; }}
    .field label {{
      font-size: .72rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: .07em;
      color: var(--ink-3);
    }}
    .field input, .field textarea {{
      background: var(--off);
      border: 1.5px solid var(--border);
      border-radius: 10px;
      padding: 12px 16px;
      font: inherit;
      font-size: .9rem;
      color: var(--ink);
      outline: none;
      transition: border-color .15s, background .15s;
    }}
    .field input:focus, .field textarea:focus {{
      border-color: var(--ink);
      background: white;
    }}
    .field input::placeholder, .field textarea::placeholder {{ color: var(--ink-4); }}
    .field textarea {{ min-height: 110px; resize: vertical; }}
    #feedback-status {{
      font-size: .82rem;
      color: var(--green);
      min-height: 1.2em;
    }}

    /* ── Footer ── */
    footer {{
      border-top: 1px solid var(--border);
      padding: 28px 0;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }}
    .footer-left {{ font-size: .8rem; color: var(--ink-4); }}
    .footer-links {{ display: flex; gap: 20px; }}
    .footer-links a {{
      font-size: .8rem;
      color: var(--ink-4);
      transition: color .15s;
    }}
    .footer-links a:hover {{ color: var(--ink); }}

    /* ── Animations ── */
    @keyframes fadeUp {{
      from {{ opacity: 0; transform: translateY(20px); }}
      to   {{ opacity: 1; transform: translateY(0); }}
    }}
    .hero > * {{
      animation: fadeUp .55s ease both;
    }}
    .hero-label   {{ animation-delay: .0s; }}
    h1            {{ animation-delay: .08s; }}
    .hero-sub     {{ animation-delay: .16s; }}
    .hero-actions {{ animation-delay: .22s; }}
    .hero-visual  {{ animation-delay: .30s; }}

    /* ── Responsive ── */
    @media (max-width: 860px) {{
      .two-col {{ grid-template-columns: 1fr; gap: 48px; }}
      .features-grid {{ grid-template-columns: 1fr; }}
      .steps-grid {{ grid-template-columns: 1fr 1fr; gap: 32px; }}
      .download-section {{ grid-template-columns: 1fr; padding: 48px; gap: 40px; }}
      .stats-strip {{ grid-template-columns: 1fr 1fr; }}
      .stat-item:nth-child(2) {{ border-right: none; }}
      .stat-item:nth-child(3) {{ border-top: 1px solid var(--border); }}
      .stat-item:nth-child(4) {{ border-top: 1px solid var(--border); border-right: none; }}
      .mockup-body {{ grid-template-columns: 1fr; }}
      .mockup-sidebar {{ flex-direction: row; }}
      nav .nav-links {{ display: none; }}
    }}
    @media (max-width: 540px) {{
      .hero {{ padding: 64px 0 48px; }}
      h1 {{ font-size: 2.4rem; }}
      .steps-grid {{ grid-template-columns: 1fr; }}
      .form-row {{ grid-template-columns: 1fr; }}
      .download-section {{ padding: 36px 28px; }}
      footer {{ flex-direction: column; gap: 12px; text-align: center; }}
    }}
  </style>
</head>
<body>

{status_markup}

<!-- Nav -->
<nav>
  <div class="wrap nav-inner">
    <a class="nav-logo" href="#">
      {icon_markup}
      <span class="nav-name">{APP_NAME}</span>
    </a>
    <div class="nav-links">
      <a class="nav-link" href="#features">Features</a>
      <a class="nav-link" href="#how-it-works">How it works</a>
      <a class="nav-link" href="#faq">FAQ</a>
      <a class="nav-link" href="#contact">Contact</a>
      <a class="nav-cta" href="{STORE_URL}" target="_blank" rel="noopener">Get the app</a>
    </div>
  </div>
</nav>

<!-- Hero -->
<section class="hero">
  <div class="wrap">
    <div class="hero-label">
      <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><rect x="1" y="1" width="5.5" height="5.5" rx="1" fill="#1d4ed8"/><rect x="7.5" y="1" width="5.5" height="5.5" rx="1" fill="#1d4ed8" opacity=".4"/><rect x="1" y="7.5" width="5.5" height="5.5" rx="1" fill="#1d4ed8" opacity=".4"/><rect x="7.5" y="7.5" width="5.5" height="5.5" rx="1" fill="#1d4ed8" opacity=".7"/></svg>
      Available on Microsoft Store
    </div>

    <h1>Fill any document,<br><em>in seconds.</em></h1>

    <p class="hero-sub">
      AI-powered template generation and document filling for Word, PowerPoint, and PDF — with batch automation built for Windows.
    </p>

    <div class="hero-actions">
      <a class="btn-primary" href="{STORE_URL}" target="_blank" rel="noopener">
        <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor"><path d="M0 0h7.5v7.5H0zm8.5 0H16v7.5H8.5zM0 8.5h7.5V16H0zm8.5 0H16V16H8.5z"/></svg>
        Download for Windows
      </a>
      <a class="btn-secondary" href="#features">
        See all features
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M2.5 7h9M7.5 3l4 4-4 4"/></svg>
      </a>
    </div>

    <!-- Doc mockup -->
    <div class="hero-visual">
      <div class="mockup-bar">
        <div class="dot dot-r"></div>
        <div class="dot dot-y"></div>
        <div class="dot dot-g"></div>
        <span class="mockup-title">Contract_Template.docx — Fillable-doc</span>
      </div>
      <div class="mockup-body">
        <div class="mockup-doc">
          <div class="doc-heading"></div>
          <div class="doc-line w-full"></div>
          <div class="doc-line w-9"></div>
          <div class="doc-line w-8"></div>
          <div class="doc-field">
            <span class="doc-field-label">Client Name</span>
            <span class="doc-field-val">Acme Corporation</span>
            <span class="doc-ai-badge">
              <svg width="8" height="8" viewBox="0 0 8 8" fill="none"><circle cx="4" cy="4" r="3" fill="#2f9e44"/></svg>
              AI filled
            </span>
          </div>
          <div style="margin-top:12px">
            <div class="doc-line w-full"></div>
            <div class="doc-line w-7"></div>
          </div>
          <div class="doc-field" style="margin-top:10px">
            <span class="doc-field-label">Date</span>
            <span class="doc-field-val">March 18, 2026</span>
            <span class="doc-ai-badge">
              <svg width="8" height="8" viewBox="0 0 8 8" fill="none"><circle cx="4" cy="4" r="3" fill="#2f9e44"/></svg>
              AI filled
            </span>
          </div>
        </div>
        <div class="mockup-sidebar">
          <div class="side-card">
            <div class="side-card-label">File types</div>
            <span class="side-pill pill-b">DOCX</span>
            <span class="side-pill pill-b">PPTX</span>
            <span class="side-pill pill-g">PDF</span>
            <span class="side-pill pill-o">CSV</span>
          </div>
          <div class="side-card">
            <div class="side-card-label">Batch progress</div>
            <div class="side-prog-label">72 / 100 records</div>
            <div class="side-prog-bar"><div class="side-prog-fill"></div></div>
          </div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- Stats -->
<div class="wrap">
  <div class="stats-strip">
    <div class="stat-item">
      <div class="stat-num">3</div>
      <div class="stat-label">Output formats</div>
    </div>
    <div class="stat-item">
      <div class="stat-num">∞</div>
      <div class="stat-label">Batch records</div>
    </div>
    <div class="stat-item">
      <div class="stat-num">AI</div>
      <div class="stat-label">Assisted filling</div>
    </div>
    <div class="stat-item">
      <div class="stat-num">1-click</div>
      <div class="stat-label">Context menu</div>
    </div>
  </div>
</div>

<!-- Features -->
<section class="features-section" id="features">
  <div class="wrap">
    <div class="section-header">
      <div class="section-eyebrow">Capabilities</div>
      <h2>Everything you need to<br><em>automate documents</em></h2>
      <p class="section-sub">A desktop-first workflow covering template creation, AI filling, and bulk output — no repetitive work.</p>
    </div>
    <div class="features-grid">
      <div class="feat">
        <div class="feat-icon">📄</div>
        <div class="feat-title">Template Generation</div>
        <p class="feat-desc">Convert any DOCX, PPTX, or PDF into a reusable fillable template. Your layout and structure are preserved exactly while fields become dynamic.</p>
      </div>
      <div class="feat">
        <div class="feat-icon">🤖</div>
        <div class="feat-title">AI-Assisted Filling</div>
        <p class="feat-desc">Guided prompts and contextual reference files let AI generate polished, coherent document content that understands your template structure.</p>
      </div>
      <div class="feat">
        <div class="feat-icon">⚡</div>
        <div class="feat-title">Batch Automation</div>
        <p class="feat-desc">Supply a CSV or JSON dataset and produce all documents in a single run — ideal for invoices, contracts, reports, and personalised letters.</p>
      </div>
      <div class="feat">
        <div class="feat-icon">🖱️</div>
        <div class="feat-title">Desktop-First Workflow</div>
        <p class="feat-desc">Right-click context menu actions let you trigger fills directly from File Explorer. Document-first editing for real production use.</p>
      </div>
    </div>
  </div>
</section>

<!-- How it works -->
<section id="how-it-works">
  <div class="wrap">
    <div class="section-header">
      <div class="section-eyebrow">Workflow</div>
      <h2>From file to finished<br><em>in four steps</em></h2>
    </div>
    <div class="steps-grid">
      <div class="step">
        <div class="step-num"></div>
        <div class="step-divider"></div>
        <div class="step-title">Import your source</div>
        <p class="step-desc">Open any DOCX, PPTX, or PDF as your starting layout. The app reads your structure and preserves all formatting.</p>
      </div>
      <div class="step">
        <div class="step-num"></div>
        <div class="step-divider"></div>
        <div class="step-title">Define the fields</div>
        <p class="step-desc">Mark placeholders and configure which sections AI should fill, and which come from structured data.</p>
      </div>
      <div class="step">
        <div class="step-num"></div>
        <div class="step-divider"></div>
        <div class="step-title">Provide your data</div>
        <p class="step-desc">Attach a reference file, type a prompt, or point to a CSV dataset. Works for one record or thousands.</p>
      </div>
      <div class="step">
        <div class="step-num"></div>
        <div class="step-divider"></div>
        <div class="step-title">Export documents</div>
        <p class="step-desc">Get perfectly filled, formatted documents ready to send — individually or in one batch export.</p>
      </div>
    </div>
  </div>
</section>

<!-- Download CTA -->
<div class="wrap">
  <div class="download-section">
    <div>
      <div class="dl-eyebrow">Microsoft Store</div>
      <div class="dl-title">Ready to stop filling<br><em>documents by hand?</em></div>
      <p class="dl-sub">Download Fillable-doc on Windows and let AI handle the repetitive work — from a single document to thousands.</p>
      <a class="btn-white" href="{STORE_URL}" target="_blank" rel="noopener">
        <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor"><path d="M0 0h7.5v7.5H0zm8.5 0H16v7.5H8.5zM0 8.5h7.5V16H0zm8.5 0H16V16H8.5z"/></svg>
        Get it on Microsoft Store
      </a>
    </div>
    <div class="qr-block">
      <div class="qr-frame">
        <img src="{qr_data_uri}" alt="QR code" />
      </div>
      <div class="qr-caption">Scan to open Store</div>
    </div>
  </div>
</div>

<!-- FAQ + Contact -->
<section id="faq" style="padding-top:0">
  <div class="wrap two-col">
    <div>
      <div class="section-eyebrow">FAQ</div>
      <h2 style="margin-bottom:32px">Common<br><em>questions</em></h2>
      <div class="faq-list">
        <div class="faq-item">
          <details>
            <summary class="faq-summary">
              What file types does Fillable-doc support?
              <span class="faq-chevron"><svg width="10" height="10" viewBox="0 0 10 10" fill="none" stroke="#6b6b68" stroke-width="1.6" stroke-linecap="round"><path d="M2 5h6M5 2l3 3-3 3"/></svg></span>
            </summary>
            <div class="faq-answer">It is built for DOCX, PPTX, and PDF output, with CSV and JSON as structured input sources for batch filling.</div>
          </details>
        </div>
        <div class="faq-item">
          <details>
            <summary class="faq-summary">
              Can I manually edit templates before filling?
              <span class="faq-chevron"><svg width="10" height="10" viewBox="0 0 10 10" fill="none" stroke="#6b6b68" stroke-width="1.6" stroke-linecap="round"><path d="M2 5h6M5 2l3 3-3 3"/></svg></span>
            </summary>
            <div class="faq-answer">Yes. The document-first workflow lets you fully refine any template before running a fill operation.</div>
          </details>
        </div>
        <div class="faq-item">
          <details>
            <summary class="faq-summary">
              How does batch generation work?
              <span class="faq-chevron"><svg width="10" height="10" viewBox="0 0 10 10" fill="none" stroke="#6b6b68" stroke-width="1.6" stroke-linecap="round"><path d="M2 5h6M5 2l3 3-3 3"/></svg></span>
            </summary>
            <div class="faq-answer">Supply a CSV or JSON file with multiple rows. Fillable-doc generates one complete document per record in a single operation.</div>
          </details>
        </div>
        <div class="faq-item">
          <details>
            <summary class="faq-summary">
              How does billing work?
              <span class="faq-chevron"><svg width="10" height="10" viewBox="0 0 10 10" fill="none" stroke="#6b6b68" stroke-width="1.6" stroke-linecap="round"><path d="M2 5h6M5 2l3 3-3 3"/></svg></span>
            </summary>
            <div class="faq-answer">Credits are purchased in-app and consumed per AI fill operation based on output length. Unused credits roll over.</div>
          </details>
        </div>
      </div>
    </div>

    <div id="contact">
      <div class="section-eyebrow">Contact</div>
      <h2 style="margin-bottom:24px">Get in touch</h2>
      <div class="contact-block">
        <div class="contact-row">
          <div class="contact-icon">👤</div>
          <div class="contact-text">
            <strong>Developer</strong>
            <span>{CONTACT_NAME}</span>
          </div>
        </div>
        <div class="contact-row">
          <div class="contact-icon">✉️</div>
          <div class="contact-text">
            <strong>Email</strong>
            <a href="mailto:{CONTACT_EMAIL}">{CONTACT_EMAIL}</a>
          </div>
        </div>
      </div>
      <form class="form" id="feedback-form">
        <div class="form-row">
          <div class="field">
            <label>Name</label>
            <input type="text" name="name" placeholder="Your name" required />
          </div>
          <div class="field">
            <label>Email</label>
            <input type="email" name="email" placeholder="you@example.com" required />
          </div>
        </div>
        <div class="field">
          <label>Message</label>
          <textarea name="message" placeholder="What's working, what's missing, or what should change..." required></textarea>
        </div>
        <button type="submit" class="btn-primary" style="justify-self:start">Send message</button>
        <div id="feedback-status"></div>
      </form>
    </div>
  </div>
</section>

<!-- Footer -->
<div class="wrap">
  <footer>
    <span class="footer-left">© 2025 VizpageAI · {APP_NAME}</span>
    <div class="footer-links">
      <a href="{STORE_URL}" target="_blank" rel="noopener">Microsoft Store</a>
      <a href="mailto:{CONTACT_EMAIL}">Support</a>
    </div>
  </footer>
</div>

<script>
  const form = document.getElementById("feedback-form");
  const status = document.getElementById("feedback-status");
  form.addEventListener("submit", async (e) => {{
    e.preventDefault();
    const btn = form.querySelector("button[type=submit]");
    btn.textContent = "Sending…";
    btn.disabled = true;
    const data = Object.fromEntries(new FormData(form).entries());
    try {{
      const res = await fetch("/feedback", {{
        method: "POST",
        headers: {{ "Content-Type": "application/json" }},
        body: JSON.stringify(data)
      }});
      const payload = await res.json();
      if (!res.ok) throw new Error(payload.detail || "Could not submit.");
      form.reset();
      status.textContent = "✓ Message received — thanks!";
      btn.textContent = "Send message";
      btn.disabled = false;
    }} catch (err) {{
      status.textContent = err.message || "Submission failed. Please email us directly.";
      status.style.color = "#c0392b";
      btn.textContent = "Send message";
      btn.disabled = false;
    }}
  }});
</script>
</body>
</html>"""


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


@app.get("/", response_class=HTMLResponse)
def landing_page() -> str:
    return _marketing_page()


@app.post("/feedback")
def submit_feedback(payload: FeedbackRequest) -> dict:
    name = payload.name.strip()
    email = payload.email.strip()
    message = payload.message.strip()
    if not name or not email or not message:
        raise HTTPException(status_code=400, detail="Name, email, and message are required.")
    submission_id = save_feedback_submission(name=name, email=email, message=message)
    return {"ok": True, "id": submission_id}


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
    credits_used = output_tokens * 0.0003
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
        session_id = str(session.get("id", "")).strip()
        if session_id and not claim_webhook_event(f"session:{session_id}"):
            return JSONResponse({"received": True, "deduped": True, "scope": "session"})
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
        {{
            "status": "ok",
            "id_token": id_token,
            "refresh_token": refresh_token,
            "email": email or user.email or "",
            "uid": uid,
        }},
    )
    return {{"ok": True}}


@app.get("/v1/auth/poll")
def auth_poll(session_id: str) -> dict:
    payload = get_auth_session(session_id)
    if not payload:
        return {{"status": "pending"}}
    return payload