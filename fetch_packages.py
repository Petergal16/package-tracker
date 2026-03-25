#!/usr/bin/env python3
"""
fetch_packages.py — Connects to IMAP (Gmail or Yahoo), pulls shipping emails,
uses Claude to extract structured package data, saves to packages.json.
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
        "trash": "[Gmail]/Trash",
    },
    "yahoo": {
        "host": "imap.mail.yahoo.com",
        "port": 993,
        "note": "Use an App Password (account.security.yahoo.com)",
        "trash": "Trash",
    },
    "outlook": {
        "host": "outlook.office365.com",
        "port": 993,
        "note": "Use your regular password or app password",
        "trash": "Deleted Items",
    },
}

EMAIL_USER     = os.environ.get("EMAIL_USER", "")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "")  # App Password
EMAIL_PROVIDER = os.environ.get("EMAIL_PROVIDER", "gmail").lower()  # gmail | yahoo | outlook

PACKAGES_FILE = Path(__file__).parent / "packages.json"
SEEN_FILE     = Path(__file__).parent / ".seen_ids.json"

# Keywords to search in subject — covers most carriers
SEARCH_SUBJECTS = [
    "shipped", "shipping", "tracking", "out for delivery",
    "delivered", "package", "order confirmed", "order shipped",
    "your order", "dispatch", "on its way", "arriving",
    "UPS", "FedEx", "USPS", "DHL", "Amazon",
]

# ── Persistence ────────────────────────────────────────────────────────────────

def load_packages() -> dict:
    if PACKAGES_FILE.exists():
        with open(PACKAGES_FILE) as f:
            return json.load(f)
    return {"packages": [], "last_updated": None}

def save_packages(data: dict):
    data["last_updated"] = datetime.now(timezone.utc).isoformat()
    with open(PACKAGES_FILE, "w") as f:
        json.dump(data, f, indent=2)

def load_seen() -> set:
    if SEEN_FILE.exists():
        with open(SEEN_FILE) as f:
            return set(json.load(f))
    return set()

def save_seen(seen: set):
    with open(SEEN_FILE, "w") as f:
        json.dump(list(seen), f)

def email_id(msg_id: str, subject: str) -> str:
    return hashlib.md5(f"{msg_id}{subject}".encode()).hexdigest()

# ── IMAP ───────────────────────────────────────────────────────────────────────

def decode_header_value(val: str) -> str:
    parts = email.header.decode_header(val)
    result = []
    for part, enc in parts:
        if isinstance(part, bytes):
            result.append(part.decode(enc or "utf-8", errors="replace"))
        else:
            result.append(part)
    return "".join(result)

def get_email_body(msg) -> str:
    """Extract plain text body from email, falling back to HTML."""
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get("Content-Disposition", ""))
            if ct == "text/plain" and "attachment" not in cd:
                try:
                    body = part.get_payload(decode=True).decode(
                        part.get_content_charset() or "utf-8", errors="replace"
                    )
                    break
                except Exception:
                    pass
        if not body:
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    try:
                        import re
                        html = part.get_payload(decode=True).decode(
                            part.get_content_charset() or "utf-8", errors="replace"
                        )
                        body = re.sub(r"<[^>]+>", " ", html)
                        body = re.sub(r"\s+", " ", body).strip()
                        break
                    except Exception:
                        pass
    else:
        try:
            body = msg.get_payload(decode=True).decode(
                msg.get_content_charset() or "utf-8", errors="replace"
            )
        except Exception:
            body = ""
    return body[:6000]  # Cap to avoid token overflow

def fetch_emails_from_folder(mail, folder: str, max_emails: int) -> list[dict]:
    """Search a single IMAP folder for shipping emails. Returns list of email dicts."""
    try:
        status, _ = mail.select(f'"{folder}"', readonly=True)
        if status != 'OK':
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

            subject  = decode_header_value(msg.get("Subject", ""))
            sender   = decode_header_value(msg.get("From", ""))
            date_raw = msg.get("Date", "")
            msg_id   = msg.get("Message-ID", str(num))
            body     = get_email_body(msg)

            try:
                received_at = parsedate_to_datetime(date_raw).isoformat()
            except Exception:
                received_at = date_raw

            emails.append({
                "id":          email_id(msg_id, subject),
                "subject":     subject,
                "sender":      sender,
                "received_at": received_at,
                "body":        body,
            })
        except Exception as e:
            print(f"    ⚠ Error reading email {num}: {e}")

    return emails


def fetch_shipping_emails(max_emails: int = 50) -> list[dict]:
    """Connect to IMAP, search Inbox AND Trash, return combined deduplicated list."""
    profile = IMAP_PROFILES.get(EMAIL_PROVIDER)
    if not profile:
        raise ValueError(f"Unknown provider: {EMAIL_PROVIDER}. Choose: {list(IMAP_PROFILES)}")

    print(f"  Connecting to {profile['host']}…")
    mail = imaplib.IMAP4_SSL(profile["host"], profile["port"])
    mail.login(EMAIL_USER, EMAIL_PASSWORD)

    # Search Inbox
    print("  📥 Searching Inbox…")
    inbox_emails = fetch_emails_from_folder(mail, "INBOX", max_emails)
    print(f"     Found {len(inbox_emails)} candidate emails.")

    # Search Trash
    trash_folder = profile.get("trash", "Trash")
    print(f"  🗑️  Searching Trash ({trash_folder})…")
    trash_emails = fetch_emails_from_folder(mail, trash_folder, max_emails)
    print(f"     Found {len(trash_emails)} candidate emails.")

    mail.logout()

    # Deduplicate by email ID — inbox version takes priority
    seen_ids = {e["id"] for e in inbox_emails}
    unique_trash = [e for e in trash_emails if e["id"] not in seen_ids]

    all_emails = inbox_emails + unique_trash
    print(f"  Combined: {len(all_emails)} unique emails across both folders.\n")
    return all_emails

# ── Claude parsing ─────────────────────────────────────────────────────────────

def parse_with_claude(emails: list[dict]) -> list[dict]:
    """
    Send batches of emails to Claude and extract structured package data.
    Returns list of package dicts.
    """
    client   = anthropic.Anthropic()
    packages = []

    # Process in batches of 5 to stay within token limits
    batch_size = 5
    for i in range(0, len(emails), batch_size):
        batch = emails[i:i+batch_size]
        print(f"  🤖 Claude parsing emails {i+1}–{min(i+batch_size, len(emails))}…")

        email_blocks = "\n\n---\n\n".join(
            f"EMAIL {j+1}\nFrom: {e['sender']}\nSubject: {e['subject']}\n"
            f"Received: {e['received_at']}\nID: {e['id']}\n\nBody:\n{e['body']}"
            for j, e in enumerate(batch)
        )

        prompt = f"""You are a package tracking parser. Extract shipping/delivery information from these emails.

For EACH email that contains package/order/shipping information, output one JSON object.
Skip emails that are clearly not shipping-related (newsletters, promotions without tracking info, etc.).

Emails:
{email_blocks}

Respond with a JSON array only, no markdown. Each object must have exactly these fields:
{{
  "email_id": "<the ID from the email header above>",
  "retailer": "<store or sender name, e.g. 'Amazon', 'REI', 'Backcountry'>",
  "description": "<what was ordered, 1 short phrase, e.g. 'Black Diamond trekking poles'>",
  "carrier": "<UPS | FedEx | USPS | DHL | Amazon Logistics | Other | Unknown>",
  "tracking_number": "<tracking number string or null>",
  "tracking_url": "<direct tracking URL if present or null>",
  "status": "<one of: ordered | shipped | in_transit | out_for_delivery | delivered | delayed | exception | unknown>",
  "status_detail": "<short human-readable status, e.g. 'In transit — Chicago, IL'>",
  "estimated_delivery": "<YYYY-MM-DD or null>",
  "delivered_at": "<ISO datetime if delivered, else null>",
  "order_number": "<order number string or null>",
  "item_cost": <float — the price of just this item/line if shown, else null>,
  "order_total": <float — the full order total including tax/shipping if shown, else null>,
  "currency": "<USD | CAD | EUR | etc., or null>",
  "received_at": "<copy the received_at from the email>",
  "last_updated": "<today's date YYYY-MM-DD>"
}}

For costs: prefer item_cost when a single item price is listed. Use order_total for the full
order amount. If only one dollar figure appears in the email, put it in order_total.
Strip currency symbols — return numbers only (e.g. 149.95 not "$149.95").

If multiple emails refer to the same package (e.g. order confirmation + shipped + delivered),
merge them into ONE object using the most recent/complete status. Use the email_id of the
most recent email for that package.
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
            packages.extend(batch_packages)
            print(f"     → Extracted {len(batch_packages)} package(s)")
        except json.JSONDecodeError as e:
            print(f"  ⚠ JSON parse error: {e}")
            print(f"     Raw: {raw[:200]}")

    return packages

# ── Merge logic ────────────────────────────────────────────────────────────────

STATUS_RANK = {
    "ordered": 0, "shipped": 1, "in_transit": 2,
    "out_for_delivery": 3, "delivered": 4,
    "delayed": 2, "exception": 3, "unknown": -1,
}

def merge_packages(existing: list[dict], new_packages: list[dict]) -> list[dict]:
    """
    Merge new packages into existing list.
    Deduplicates by tracking_number (preferred) or order_number.
    Keeps most advanced status.
    """
    # Index existing by tracking number and order number
    by_tracking = {p["tracking_number"]: i for i, p in enumerate(existing)
                   if p.get("tracking_number")}
    by_order    = {p["order_number"]: i for i, p in enumerate(existing)
                   if p.get("order_number")}

    result = list(existing)

    for pkg in new_packages:
        tn  = pkg.get("tracking_number")
        on  = pkg.get("order_number")
        idx = by_tracking.get(tn) if tn else None
        if idx is None:
            idx = by_order.get(on) if on else None

        if idx is not None:
            # Update if new status is more advanced
            old_rank = STATUS_RANK.get(result[idx]["status"], -1)
            new_rank = STATUS_RANK.get(pkg["status"], -1)
            if new_rank >= old_rank:
                result[idx] = pkg
        else:
            result.append(pkg)
            if tn:
                by_tracking[tn] = len(result) - 1
            if on:
                by_order[on] = len(result) - 1

    # Sort: active first (not delivered), then by estimated delivery
    def sort_key(p):
        delivered = 1 if p["status"] == "delivered" else 0
        eta = p.get("estimated_delivery") or "9999-99-99"
        return (delivered, eta)

    result.sort(key=sort_key)
    return result

# ── Main ────────────────────────────────────────────────────────────────────────

def run(dry_run: bool = False, max_emails: int = 50):
    if not EMAIL_USER or not EMAIL_PASSWORD:
        print("ERROR: Set EMAIL_USER and EMAIL_PASSWORD environment variables.")
        print(f"Note for {EMAIL_PROVIDER}: {IMAP_PROFILES.get(EMAIL_PROVIDER, {}).get('note', '')}")
        return

    print(f"\n{'='*56}")
    print(f"  Package Tracker — {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"  Provider: {EMAIL_PROVIDER}  User: {EMAIL_USER}")
    print(f"{'='*56}\n")

    # Load state
    data    = load_packages()
    seen    = load_seen()

    # Fetch emails
    print("📬 Fetching shipping emails…")
    all_emails = fetch_shipping_emails(max_emails)

    # Filter already-seen
    new_emails = [e for e in all_emails if e["id"] not in seen]
    print(f"  {len(new_emails)} new emails to process (skipping {len(all_emails)-len(new_emails)} already seen)\n")

    if not new_emails:
        print("Nothing new. packages.json is up to date.\n")
        return

    # Parse with Claude
    print("📦 Extracting package data with Claude…")
    new_packages = parse_with_claude(new_emails)

    if dry_run:
        print("\n🧪 Dry run — extracted packages (not saved):")
        for p in new_packages:
            print(f"  {p['retailer']:<20} {p['status']:<18} {p.get('description','')}")
        return

    # Merge and save
    data["packages"] = merge_packages(data.get("packages", []), new_packages)
    save_packages(data)

    # Mark emails as seen
    seen.update(e["id"] for e in new_emails)
    save_seen(seen)

    print(f"\n✅ Saved {len(data['packages'])} total packages to packages.json")
    print(f"   Active: {sum(1 for p in data['packages'] if p['status'] != 'delivered')}")
    print(f"   Delivered: {sum(1 for p in data['packages'] if p['status'] == 'delivered')}\n")

if __name__ == "__main__":
    p = argparse.ArgumentParser(description="Fetch and parse shipping emails")
    p.add_argument("--dry-run",    action="store_true", help="Parse but don't save")
    p.add_argument("--max-emails", type=int, default=50, help="Max emails to fetch")
    args = p.parse_args()
    run(dry_run=args.dry_run, max_emails=args.max_emails)
