#!/usr/bin/env python3
"""
fetch_packages.py — Connects to IMAP (Gmail or Yahoo), pulls shipping emails
from INBOX ONLY, uses Claude to extract structured package data, and saves to
packages.json.
"""

import os
import re
import json
import html
import imaplib
import email
import email.header
import hashlib
import argparse
from datetime import datetime, timezone
from pathlib import Path
from email.utils import parsedate_to_datetime

import anthropic


# ── Config ─────────────────────────────────────────────────────────────────────

IMAP_PROFILES = {
    "gmail": {
        "host": "imap.gmail.com",
        "port": 993,
        "note": "Use an App Password (myaccount.google.com/apppasswords)",
    },
    "yahoo": {
        "host": "imap.mail.yahoo.com",
        "port": 993,
        "note": "Use an App Password (Yahoo account security settings)",
    },
    "outlook": {
        "host": "outlook.office365.com",
        "port": 993,
        "note": "Use your regular password or app password",
    },
}

EMAIL_USER = os.environ.get("EMAIL_USER", "")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "")
EMAIL_PROVIDER = os.environ.get("EMAIL_PROVIDER", "gmail").lower()
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

PACKAGES_FILE = Path(__file__).parent / "packages.json"
SEEN_FILE = Path(__file__).parent / ".seen_ids.json"

SEARCH_SUBJECTS = [
    "shipped",
    "shipping",
    "tracking",
    "out for delivery",
    "delivered",
    "package",
    "order confirmed",
    "order shipped",
    "your order",
    "dispatch",
    "on its way",
    "arriving",
    "UPS",
    "FedEx",
    "USPS",
    "DHL",
    "Amazon",
]

# Change this later if you want a different recent window
SEARCH_SINCE = "20-Mar-2026"

STATUS_RANK = {
    "ordered": 0,
    "shipped": 1,
    "in_transit": 2,
    "out_for_delivery": 3,
    "delivered": 4,
    "delayed": 2,
    "exception": 3,
    "unknown": -1,
}


# ── Persistence ────────────────────────────────────────────────────────────────

def load_packages() -> dict:
    if PACKAGES_FILE.exists():
        with open(PACKAGES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"packages": [], "last_updated": None}


def save_packages(data: dict) -> None:
    data["last_updated"] = datetime.now(timezone.utc).isoformat()
    with open(PACKAGES_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def load_seen() -> set:
    if SEEN_FILE.exists():
        with open(SEEN_FILE, "r", encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def save_seen(seen: set) -> None:
    with open(SEEN_FILE, "w", encoding="utf-8") as f:
        json.dump(sorted(list(seen)), f, indent=2)


def email_id(msg_id: str, subject: str) -> str:
    return hashlib.md5(f"{msg_id}{subject}".encode("utf-8")).hexdigest()


# ── Email parsing helpers ──────────────────────────────────────────────────────

def decode_header_value(val: str) -> str:
    if not val:
        return ""
    parts = email.header.decode_header(val)
    result = []
    for part, enc in parts:
        if isinstance(part, bytes):
            result.append(part.decode(enc or "utf-8", errors="replace"))
        else:
            result.append(part)
    return "".join(result).strip()


def clean_email_body(raw_html: str) -> str:
    """
    Convert HTML email into readable text while preserving links as:
    Link Text [LINK: https://...]
    """
    if not raw_html:
        return ""

    text = raw_html
    text = text.replace("=\r\n", "").replace("=\n", "")
    text = html.unescape(text)

    # Preserve anchor text + href
    text = re.sub(
        r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>',
        lambda m: f"{re.sub(r'<[^>]+>', ' ', m.group(2)).strip()} [LINK: {m.group(1)}]",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )

    # Helpful line breaks before stripping tags
    text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'</p\s*>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'</div\s*>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'</tr\s*>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'</li\s*>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'</td\s*>', ' ', text, flags=re.IGNORECASE)
    text = re.sub(r'</h[1-6]\s*>', '\n', text, flags=re.IGNORECASE)

    # Strip remaining tags
    text = re.sub(r'<[^>]+>', ' ', text, flags=re.DOTALL)

    # Normalize whitespace
    text = text.replace("\xa0", " ")
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n\s*\n+', '\n\n', text)
    text = re.sub(r' +\n', '\n', text)
    text = re.sub(r'\n +', '\n', text)

    return text.strip()


def extract_focus_section(text: str) -> str:
    """
    Pull the most useful snippets for Claude so it focuses on the actual
    shipping details instead of the entire messy email.
    """
    if not text:
        return ""

    text_lower = text.lower()

    keywords = [
        "just shipped",
        "your order has shipped",
        "on the move",
        "estimated arrival",
        "track package",
        "tracking number",
        "carrier:",
        "shipped via",
        "order #",
        "qty:",
        "item #",
        "shipping to:",
        "delivered",
        "better trucks",
        "black diamond",
        "walmart",
        "fleet farm",
        "nobull",
        "rei co-op",
        "product",
        "item",
    ]

    snippets = []

    for kw in keywords:
        idx = text_lower.find(kw)
        if idx != -1:
            start = max(0, idx - 250)
            end = min(len(text), idx + 1200)
            snippets.append(text[start:end])

    head = text[:2500]

    if snippets:
        combined = head + "\n\n---FOCUSED CONTENT---\n\n" + "\n\n---\n\n".join(snippets[:8])
    else:
        combined = head

    combined = re.sub(r'\s+', ' ', combined).strip()
    return combined[:9000]


def get_email_body(msg) -> str:
    """Extract useful email text from plain text and/or HTML."""
    plain_body = ""
    html_body = ""

    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get("Content-Disposition", ""))

            if "attachment" in cd.lower():
                continue

            if ct == "text/plain" and not plain_body:
                try:
                    plain_body = part.get_payload(decode=True).decode(
                        part.get_content_charset() or "utf-8",
                        errors="replace"
                    )
                except Exception:
                    pass

            elif ct == "text/html" and not html_body:
                try:
                    html_body = part.get_payload(decode=True).decode(
                        part.get_content_charset() or "utf-8",
                        errors="replace"
                    )
                except Exception:
                    pass
    else:
        try:
            payload = msg.get_payload(decode=True).decode(
                msg.get_content_charset() or "utf-8",
                errors="replace"
            )
            if msg.get_content_type() == "text/html":
                html_body = payload
            else:
                plain_body = payload
        except Exception:
            pass

    plain_body = (plain_body or "").strip()
    html_text = clean_email_body(html_body) if html_body else ""

    # Prefer HTML when plain text is weak/generic
    if len(plain_body) < 500 and html_text:
        body = html_text
    elif plain_body and html_text:
        body = plain_body + "\n\nHTML VERSION:\n" + html_text
    else:
        body = plain_body or html_text

    return body[:20000].strip()


# ── IMAP fetching ──────────────────────────────────────────────────────────────

def fetch_emails_from_folder(mail, folder: str, max_emails: int) -> list[dict]:
    try:
        status, _ = mail.select(f'"{folder}"', readonly=True)
        if status != "OK":
            print(f"    ⚠ Could not open folder: {folder}")
            return []
    except Exception as e:
        print(f"    ⚠ Error selecting folder {folder}: {e}")
        return []

    all_ids = set()

    for kw in SEARCH_SUBJECTS[:10]:
        try:
            _, data = mail.search(None, f'(SINCE "{SEARCH_SINCE}" SUBJECT "{kw}")')
            if data and data[0]:
                all_ids.update(data[0].split())
        except Exception:
            pass

    emails = []

    for num in sorted(all_ids)[-max_emails:]:
        try:
            _, msg_data = mail.fetch(num, "(RFC822)")
            raw = msg_data[0][1]
            msg = email.message_from_bytes(raw)

            subject = decode_header_value(msg.get("Subject", ""))
            sender = decode_header_value(msg.get("From", ""))
            date_raw = msg.get("Date", "")
            msg_id = msg.get("Message-ID", str(num))
            body = get_email_body(msg)

            if not subject and not body:
                continue

            try:
                received_at = parsedate_to_datetime(date_raw).isoformat()
            except Exception:
                received_at = date_raw

            emails.append({
                "id": email_id(msg_id, subject),
                "subject": subject,
                "sender": sender,
                "received_at": received_at,
                "body": body,
            })

        except Exception as e:
            print(f"    ⚠ Error reading email {num}: {e}")

    return emails


def fetch_shipping_emails(max_emails: int = 50) -> list[dict]:
    profile = IMAP_PROFILES.get(EMAIL_PROVIDER)
    if not profile:
        raise ValueError(f"Unknown provider: {EMAIL_PROVIDER}. Choose: {list(IMAP_PROFILES)}")

    print(f"  Connecting to {profile['host']}…")
    mail = imaplib.IMAP4_SSL(profile["host"], profile["port"])
    mail.login(EMAIL_USER, EMAIL_PASSWORD)

    print("  📥 Searching Inbox only…")
    inbox_emails = fetch_emails_from_folder(mail, "INBOX", max_emails)
    print(f"     Found {len(inbox_emails)} candidate emails.")

    mail.logout()

    print(f"  Combined: {len(inbox_emails)} unique emails from Inbox only.\n")
    return inbox_emails


# ── Claude parsing ─────────────────────────────────────────────────────────────

def parse_with_claude(emails: list[dict]) -> list[dict]:
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    packages = []
    batch_size = 5
    today_str = datetime.now().strftime("%Y-%m-%d")

    for i in range(0, len(emails), batch_size):
        batch = emails[i:i + batch_size]
        print(f"  🤖 Claude parsing emails {i + 1}–{min(i + batch_size, len(emails))}…")

        email_blocks = "\n\n---\n\n".join(
            f"EMAIL {j + 1}\n"
            f"From: {e['sender']}\n"
            f"Subject: {e['subject']}\n"
            f"Received: {e['received_at']}\n"
            f"ID: {e['id']}\n\n"
            f"IMPORTANT SECTION:\n{extract_focus_section(e['body'])}"
            for j, e in enumerate(batch)
        )

        prompt = f"""You are a package tracking parser. Extract shipping and delivery information from these emails.

Only include emails that are truly about physical product orders, shipping, tracking, or delivery.
Skip newsletters, promotions, gift cards, memberships, digital purchases, and pickup-only orders.

Emails:
{email_blocks}

Respond with a JSON array only. No markdown.

Each object must have exactly these fields:
{{
  "email_id": "<the ID from the email header above>",
  "retailer": "<store or sender name, e.g. Amazon, REI Co-op, Walmart.com, Fleet Farm>",
  "description": "<actual product name from the email body>",
  "carrier": "<UPS | FedEx | USPS | DHL | Amazon Logistics | Better Trucks | Other | Unknown>",
  "tracking_number": "<tracking number string or null>",
  "tracking_url": "<direct tracking URL if present or null>",
  "status": "<one of: ordered | shipped | in_transit | out_for_delivery | delivered | delayed | exception | unknown>",
  "status_detail": "<short human-readable status>",
  "estimated_delivery": "<YYYY-MM-DD or null>",
  "delivered_at": "<ISO datetime if delivered, else null>",
  "order_number": "<order number string or null>",
  "item_cost": <float or null>,
  "order_total": <float or null>,
  "currency": "<USD | CAD | EUR | etc., or null>",
  "received_at": "<copy the received_at from the email>",
  "last_updated": "{today_str}"
}}

Description rules:
- Extract the real product name from the email body whenever possible.
- Good descriptions are specific product titles like:
  - "Black Diamond Alpine Carbon Cork Trek Poles"
  - "Patagonia Torrentshell 3L Jacket"
  - "ThermoPro TP828BW Wireless Meat Thermometer"
- Bad descriptions are generic phrases like:
  - "Package"
  - "Your order"
  - "Order shipped"
  - "Outdoor gear"
  - "Items shipped"
  - "Track your order"
- Do NOT use "Package" as the description.
- Look carefully for line items, product titles, item name blocks, SKU lines, qty lines, size/color lines, or text near product images.
- If multiple items are listed, choose the first specific item name, or list the first two short product names separated by a comma.
- If no exact item name can be found anywhere, use a short fallback like "<retailer> order". Never use "Package".

Tracking rules:
- If the email contains a direct tracking link, return it in tracking_url.
- If the email contains a tracking number but no direct tracking link, still return the tracking_number.
- If the email says "Track package" and a URL is present in text like [LINK: ...], use that URL.
- Extract tracking numbers even if they look like BTP_014409X03JT.
- If the carrier can be inferred from the email, set it accordingly.
- Better Trucks is a valid carrier.

Filtering rules:
- Only include physical shipped items.
- Skip pickup-only orders.
- Skip store memberships, gift cards, digital items, and promotional emails.

Formatting rules:
- Strip currency symbols and return numbers only.
- Use null for missing values.
"""

        msg = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            messages=[{"role": "user", "content": prompt}],
        )

        raw = msg.content[0].text.strip()

        if raw.startswith("```"):
            raw = raw.split("\n", 1)[1].rsplit("```", 1)[0].strip()

        try:
            batch_packages = json.loads(raw)
            if isinstance(batch_packages, list):
                packages.extend(batch_packages)
                print(f"     → Extracted {len(batch_packages)} package(s)")
            else:
                print("  ⚠ Claude response was not a JSON array.")
        except json.JSONDecodeError as e:
            print(f"  ⚠ JSON parse error: {e}")
            print(f"     Raw: {raw[:300]}")

    return packages


# ── Heuristic extraction helpers ───────────────────────────────────────────────

def extract_product_from_text(text: str) -> str | None:
    """
    Hard fallback extraction for product names when Claude returns 'Package'.
    """
    if not text:
        return None

    # Common branded patterns
    brand_patterns = [
        r"(Black Diamond[^\n]{0,120})",
        r"(ThermoPro[^\n]{0,120})",
        r"(Patagonia[^\n]{0,120})",
        r"(Columbia[^\n]{0,120})",
        r"(Salomon[^\n]{0,120})",
        r"(HOKA[^\n]{0,120})",
        r"(Garmin[^\n]{0,120})",
        r"(YETI[^\n]{0,120})",
    ]

    for pattern in brand_patterns:
        m = re.search(pattern, text, flags=re.IGNORECASE)
        if m:
            value = re.sub(r'\s+', ' ', m.group(1)).strip(" -:,.")
            if len(value) > 8:
                return value

    # Look for product block before qty/item #
    generic_patterns = [
        r"([A-Z][A-Za-z0-9,&()\/\-\.\' ]{12,120})\s+Qty:",
        r"([A-Z][A-Za-z0-9,&()\/\-\.\' ]{12,120})\s+Item\s*#",
        r"([A-Z][A-Za-z0-9,&()\/\-\.\' ]{12,120})\s+One Size",
        r"([A-Z][A-Za-z0-9,&()\/\-\.\' ]{12,120})\s+Size",
    ]

    for pattern in generic_patterns:
        m = re.search(pattern, text, flags=re.IGNORECASE)
        if m:
            value = re.sub(r'\s+', ' ', m.group(1)).strip(" -:,.")
            if len(value) > 8:
                return value

    return None


def extract_tracking_number_from_text(text: str) -> str | None:
    if not text:
        return None

    patterns = [
        r"Tracking number:\s*([A-Z0-9_]+)",
        r"tracking number\s*[:#]?\s*([A-Z0-9_]+)",
        r"\b(BTP_[A-Z0-9]+)\b",
        r"\b(1Z[0-9A-Z]+)\b",  # UPS-like
        r"\b([0-9]{12,22})\b",  # numeric tracking-ish
    ]

    for pattern in patterns:
        m = re.search(pattern, text, flags=re.IGNORECASE)
        if m:
            return m.group(1).strip()

    return None


def extract_tracking_url_from_text(text: str) -> str | None:
    if not text:
        return None

    # Prefer "Track package [LINK: ...]"
    m = re.search(r"Track package\s*\[LINK:\s*(https?://[^\]\s]+)\]", text, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip()

    # Generic LINK capture
    m = re.search(r"\[LINK:\s*(https?://[^\]\s]+)\]", text, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip()

    return None


def infer_carrier_from_text(text: str) -> str | None:
    if not text:
        return None

    lowered = text.lower()

    if "better trucks" in lowered:
        return "Better Trucks"
    if "ups" in lowered:
        return "UPS"
    if "fedex" in lowered:
        return "FedEx"
    if "usps" in lowered:
        return "USPS"
    if "dhl" in lowered:
        return "DHL"

    return None


def build_tracking_url(carrier: str | None, tracking_number: str | None) -> str | None:
    if not tracking_number:
        return None

    carrier_text = (carrier or "").lower()
    tn = tracking_number.strip()

    if "ups" in carrier_text:
        return f"https://www.ups.com/track?tracknum={tn}"
    if "fedex" in carrier_text:
        return f"https://www.fedex.com/fedextrack/?trknbr={tn}"
    if "usps" in carrier_text:
        return f"https://tools.usps.com/go/TrackConfirmAction?tLabels={tn}"
    if "dhl" in carrier_text:
        return f"https://www.dhl.com/us-en/home/tracking.html?tracking-id={tn}"

    return None


# ── Cleanup / enrichment ───────────────────────────────────────────────────────

def clean_extracted_packages(packages: list[dict], email_lookup: dict[str, dict]) -> list[dict]:
    cleaned = []

    weak_descriptions = {
        "",
        "package",
        "your order",
        "order",
        "product",
        "shipment",
        "items",
        "items shipped",
        "track your order",
        "thank you for your order",
        "your order has shipped",
        "thanks for your order",
    }

    for pkg in packages:
        email_id_value = pkg.get("email_id")
        source_email = email_lookup.get(email_id_value, {})
        raw_text = source_email.get("body", "")

        retailer = (pkg.get("retailer") or "").strip()
        description = (pkg.get("description") or "").strip()
        status_detail = (pkg.get("status_detail") or "").lower()
        tracking_number = pkg.get("tracking_number")
        tracking_url = pkg.get("tracking_url")
        status = (pkg.get("status") or "unknown").strip().lower()

        # Skip pickup-only
        if "pickup" in status_detail or "pickup" in raw_text.lower():
            continue

        # Hard-fix weak descriptions
        if description.lower() in weak_descriptions:
            extracted_product = extract_product_from_text(raw_text)
            if extracted_product:
                pkg["description"] = extracted_product
            else:
                pkg["description"] = f"{retailer} order" if retailer else "Order"

        # Fill tracking number if Claude missed it
        if not tracking_number:
            extracted_tracking = extract_tracking_number_from_text(raw_text)
            if extracted_tracking:
                pkg["tracking_number"] = extracted_tracking
                tracking_number = extracted_tracking

        # Fill tracking URL if Claude missed it
        if not tracking_url:
            extracted_url = extract_tracking_url_from_text(raw_text)
            if extracted_url:
                pkg["tracking_url"] = extracted_url
                tracking_url = extracted_url

        # Fill carrier if Claude missed it
        if not pkg.get("carrier") or pkg.get("carrier") == "Unknown":
            extracted_carrier = infer_carrier_from_text(raw_text)
            if extracted_carrier:
                pkg["carrier"] = extracted_carrier

        # Fallback carrier URL build if possible
        if not pkg.get("tracking_url") and pkg.get("tracking_number"):
            built_url = build_tracking_url(pkg.get("carrier"), pkg.get("tracking_number"))
            if built_url:
                pkg["tracking_url"] = built_url

        # Drop weak junk entries with no useful tracking info
        final_desc = (pkg.get("description") or "").strip().lower()
        if (
            final_desc in weak_descriptions
            and not pkg.get("tracking_number")
            and not pkg.get("tracking_url")
            and status in {"ordered", "unknown"}
        ):
            continue

        cleaned.append(pkg)

    return cleaned


# ── Merge logic ────────────────────────────────────────────────────────────────

def merge_packages(existing: list[dict], new_packages: list[dict]) -> list[dict]:
    by_tracking = {
        p["tracking_number"]: i
        for i, p in enumerate(existing)
        if p.get("tracking_number")
    }
    by_order = {
        p["order_number"]: i
        for i, p in enumerate(existing)
        if p.get("order_number")
    }
    by_email = {
        p["email_id"]: i
        for i, p in enumerate(existing)
        if p.get("email_id")
    }

    result = list(existing)

    for pkg in new_packages:
        tn = pkg.get("tracking_number")
        on = pkg.get("order_number")
        eid = pkg.get("email_id")

        idx = by_tracking.get(tn) if tn else None
        if idx is None:
            idx = by_order.get(on) if on else None
        if idx is None:
            idx = by_email.get(eid) if eid else None

        if idx is not None:
            old_rank = STATUS_RANK.get(result[idx].get("status", "unknown"), -1)
            new_rank = STATUS_RANK.get(pkg.get("status", "unknown"), -1)

            if new_rank >= old_rank:
                result[idx] = pkg
        else:
            result.append(pkg)
            new_idx = len(result) - 1

            if tn:
                by_tracking[tn] = new_idx
            if on:
                by_order[on] = new_idx
            if eid:
                by_email[eid] = new_idx

    def sort_key(p):
        delivered = 1 if p.get("status") == "delivered" else 0
        eta = p.get("estimated_delivery") or "9999-99-99"
        return (delivered, eta)

    result.sort(key=sort_key)
    return result


# ── Main ───────────────────────────────────────────────────────────────────────

def run(dry_run: bool = False, max_emails: int = 50) -> None:
    if not EMAIL_USER or not EMAIL_PASSWORD:
        print("ERROR: Set EMAIL_USER and EMAIL_PASSWORD environment variables.")
        print(f"Note for {EMAIL_PROVIDER}: {IMAP_PROFILES.get(EMAIL_PROVIDER, {}).get('note', '')}")
        return

    if not ANTHROPIC_API_KEY:
        print("ERROR: Set ANTHROPIC_API_KEY environment variable.")
        return

    print(f"\n{'=' * 56}")
    print(f"  Package Tracker — {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"  Provider: {EMAIL_PROVIDER}  User: {EMAIL_USER}")
    print(f"{'=' * 56}\n")

    data = load_packages()
    seen = load_seen()

    print("📬 Fetching shipping emails…")
    all_emails = fetch_shipping_emails(max_emails)

    new_emails = [e for e in all_emails if e["id"] not in seen]
    print(f"  {len(new_emails)} new emails to process (skipping {len(all_emails) - len(new_emails)} already seen)\n")

    if not new_emails:
        print("Nothing new. packages.json is up to date.\n")
        return

    print("📦 Extracting package data with Claude…")
    new_packages = parse_with_claude(new_emails)

    email_lookup = {e["id"]: e for e in new_emails}
    new_packages = clean_extracted_packages(new_packages, email_lookup)

    if dry_run:
        print("\n🧪 Dry run — extracted packages (not saved):")
        for p in new_packages:
            print(f"  {p.get('retailer', 'Unknown'):<20} {p.get('status', 'unknown'):<18} {p.get('description', '')}")
        return

    data["packages"] = merge_packages(data.get("packages", []), new_packages)
    save_packages(data)

    seen.update(e["id"] for e in new_emails)
    save_seen(seen)

    print(f"\n✅ Saved {len(data['packages'])} total packages to packages.json")
    print(f"   Active: {sum(1 for p in data['packages'] if p.get('status') != 'delivered')}")
    print(f"   Delivered: {sum(1 for p in data['packages'] if p.get('status') == 'delivered')}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Fetch and parse shipping emails")
    parser.add_argument("--dry-run", action="store_true", help="Parse but don't save")
    parser.add_argument("--max-emails", type=int, default=50, help="Max emails to fetch")
    args = parser.parse_args()
    run(dry_run=args.dry_run, max_emails=args.max_emails)
