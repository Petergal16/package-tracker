#!/usr/bin/env python3
"""
fetch_packages.py — Connects to IMAP (Gmail or Yahoo), pulls shipping emails
from INBOX ONLY, uses Claude to extract structured package data, and saves to
packages.json.
"""

import os
import json
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


# ── IMAP helpers ───────────────────────────────────────────────────────────────

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


def get_email_body(msg) -> str:
    """Extract plain text body from email, falling back to stripped HTML."""
    body = ""

    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get("Content-Disposition", ""))

            if ct == "text/plain" and "attachment" not in cd.lower():
                try:
                    body = part.get_payload(decode=True).decode(
                        part.get_content_charset() or "utf-8",
                        errors="replace"
                    )
                    if body:
                        break
                except Exception:
                    pass

        if not body:
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    try:
                        import re
                        html = part.get_payload(decode=True).decode(
                            part.get_content_charset() or "utf-8",
                            errors="replace"
                        )
                        body = re.sub(r"<[^>]+>", " ", html)
                        body = re.sub(r"\s+", " ", body).strip()
                        if body:
                            break
                    except Exception:
                        pass
    else:
        try:
            body = msg.get_payload(decode=True).decode(
                msg.get_content_charset() or "utf-8",
                errors="replace"
            )
        except Exception:
            body = ""

    return body[:6000].strip()


def fetch_emails_from_folder(mail, folder: str, max_emails: int) -> list[dict]:
    """Search one IMAP folder for likely shipping emails."""
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
            _, data = mail.search(None, f'(SUBJECT "{kw}")')
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

            # Skip obviously useless emails
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
    """Connect to IMAP and search INBOX ONLY."""
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
    """
    Send batches of emails to Claude and extract structured package data.
    Returns list of package dicts.
    """
    client = anthropic.Anthropic()
    packages = []

    batch_size = 5

    for i in range(0, len(emails), batch_size):
        batch = emails[i:i + batch_size]
        print(f"  🤖 Claude parsing emails {i + 1}–{min(i + batch_size, len(emails))}…")

        email_blocks = "\n\n---\n\n".join(
            f"EMAIL {j + 1}\n"
            f"From: {e['sender']}\n"
            f"Subject: {e['subject']}\n"
            f"Received: {e['received_at']}\n"
            f"ID: {e['id']}\n\n"
            f"Body:\n{e['body']}"
            for j, e in enumerate(batch)
        )

        today_str = datetime.now().strftime("%Y-%m-%d")

        prompt = f"""You are a package tracking parser. Extract shipping and delivery information from these emails.

Only include emails that are truly about physical product orders, shipping, tracking, or delivery.
Skip newsletters, promotions, gift cards, memberships, digital purchases, and pickup-only orders.

Emails:
{email_blocks}

Respond with a JSON array only. No markdown.

Each object must have exactly these fields:
{{
  "email_id": "<the ID from the email header above>",
  "retailer": "<store or sender name, e.g. Amazon, REI, Backcountry>",
  "description": "<specific product name if present; avoid generic phrases like 'your order' or 'gear order'>",
  "carrier": "<UPS | FedEx | USPS | DHL | Amazon Logistics | Other | Unknown>",
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

Rules:
- Prefer the actual item/product name from the email body.
- Never use generic descriptions like "your order", "gear order", "package", "items shipped".
- If multiple items are listed, use the first 1–2 real item names.
- If no real product name exists, use null-like behavior by falling back to a cleaned version of the subject without words like shipped, order, tracking, delivery, package.
- Only include physical shipped items.
- If there is no tracking link in the email, return null.
- If there is no tracking number, return null.
- Strip currency symbols and return numbers only.
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


# ── Merge logic ────────────────────────────────────────────────────────────────

def merge_packages(existing: list[dict], new_packages: list[dict]) -> list[dict]:
    """
    Merge new packages into existing list.
    Prefer dedupe by tracking_number, then order_number, then email_id.
    Keep the most advanced status.
    """
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
