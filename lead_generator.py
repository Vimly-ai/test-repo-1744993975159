#!/usr/bin/env python3
"""
Expired Domain Lead Generator for Treasure Valley
Finds service businesses whose Google Business Profile website links are expired/parked/broken.
"""

import argparse
import csv
import json
import re
import time
import socket
import os
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from urllib.parse import urlparse

import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ─────────────────────────── CONFIG ───────────────────────────

CITIES = [
    "Boise",
    "Meridian",
    "Nampa",
    "Caldwell",
    "Eagle",
    "Star",
    "Kuna",
    "Garden City",
    "Middleton",
    "Emmett",
]

CATEGORIES = [
    "Plumber",
    "Plumbing contractor",
    "Electrician",
    "Electrical contractor",
    "Roofer",
    "Roofing contractor",
    "HVAC",
    "Heating and cooling",
    "Pest control",
    "Landscaping",
    "Lawn care",
    "Auto repair",
    "Mechanic",
    "Cleaning service",
    "House cleaning",
    "Janitorial service",
    "Handyman",
    "Painter",
    "Painting contractor",
    "Concrete contractor",
    "Fencing contractor",
    "Tree service",
    "Tree removal",
    "Garage door repair",
    "Appliance repair",
    "Carpet cleaning",
    "Window cleaning",
    "Pressure washing",
    "Power washing",
    "Locksmith",
    "Septic service",
    "Irrigation repair",
    "Sprinkler repair",
    "Flooring contractor",
    "Tile contractor",
    "Cabinet maker",
    "Kitchen remodel",
    "Bathroom remodel",
    "General contractor",
    "Foundation repair",
    "Gutter cleaning",
    "Gutter installation",
    "Drywall contractor",
    "Insulation contractor",
    "Siding contractor",
    "Deck builder",
    "Pool service",
    "Pool cleaning",
    "Moving company",
    "Junk removal",
    "Chimney sweep",
    "Glass repair",
    "Window repair",
    "Welding service",
    "Excavation contractor",
    "Paving contractor",
]

PLACES_API_URL = "https://places.googleapis.com/v1/places:searchText"
FIELD_MASK = (
    "places.displayName,"
    "places.formattedAddress,"
    "places.websiteUri,"
    "places.nationalPhoneNumber,"
    "places.rating,"
    "places.userRatingCount,"
    "places.googleMapsUri,"
    "places.businessStatus,"
    "places.types,"
    "places.id"
)

ALL_BUSINESSES_CSV = "all_businesses.csv"
WITH_WEBSITES_CSV = "businesses_with_websites.csv"
LANDER_RESULTS_CSV = "lander_results.csv"

PARKING_DOMAINS = {
    "dan.com", "afternic.com", "sedo.com", "hugedomains.com",
    "godaddy.com", "parkingcrew.com", "bodis.com", "namedrive.com",
    "parking.com", "domainmarket.com", "undeveloped.com", "epik.com",
    "above.com", "domainnamesales.com", "premiumdrops.com",
}

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/121.0.0.0 Safari/537.36"
)

# ─────────────────────────── STEP 1: SCRAPER ───────────────────────────

def scrape_places(api_key: str) -> list[dict]:
    """Pull businesses from Google Places API for all city+category combos."""
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": api_key,
        "X-Goog-FieldMask": FIELD_MASK,
    }

    seen_ids: set[str] = set()
    all_businesses: list[dict] = []

    total_combos = len(CITIES) * len(CATEGORIES)
    combo_num = 0

    for city in CITIES:
        for category in CATEGORIES:
            combo_num += 1
            query = f"{category} in {city}, Idaho"
            print(f"  [{combo_num}/{total_combos}] Searching: {query}", flush=True)

            page_token = None
            page = 0

            while True:
                body: dict = {"textQuery": query, "maxResultCount": 20}
                if page_token:
                    body["pageToken"] = page_token

                try:
                    resp = requests.post(PLACES_API_URL, headers=headers, json=body, timeout=15)
                    resp.raise_for_status()
                    data = resp.json()
                except requests.RequestException as e:
                    print(f"    ⚠ API error: {e}", flush=True)
                    break

                places = data.get("places", [])
                page += 1
                new_count = 0

                for place in places:
                    place_id = place.get("id", "")
                    if not place_id or place_id in seen_ids:
                        continue
                    seen_ids.add(place_id)

                    business = {
                        "place_id": place_id,
                        "name": place.get("displayName", {}).get("text", ""),
                        "address": place.get("formattedAddress", ""),
                        "phone": place.get("nationalPhoneNumber", ""),
                        "website": place.get("websiteUri", ""),
                        "rating": place.get("rating", ""),
                        "review_count": place.get("userRatingCount", 0),
                        "maps_url": place.get("googleMapsUri", ""),
                        "status": place.get("businessStatus", ""),
                        "types": ",".join(place.get("types", [])),
                        "category": category,
                        "city": city,
                    }
                    all_businesses.append(business)
                    new_count += 1

                if new_count:
                    print(f"    Page {page}: +{new_count} new businesses", flush=True)

                page_token = data.get("nextPageToken")
                if not page_token:
                    break

                time.sleep(0.15)  # brief pause between pages

            time.sleep(0.15)  # brief pause between combos

    return all_businesses


def save_businesses(businesses: list[dict]) -> list[dict]:
    """Save all businesses and filter to those with websites. Returns filtered list."""
    fieldnames = [
        "place_id", "name", "address", "phone", "website",
        "rating", "review_count", "maps_url", "status", "types",
        "category", "city",
    ]

    # Save all
    with open(ALL_BUSINESSES_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(businesses)
    print(f"\n  Saved {len(businesses)} total businesses → {ALL_BUSINESSES_CSV}", flush=True)

    # Filter: operational + has website
    with_websites = [
        b for b in businesses
        if b.get("website") and b.get("status") == "OPERATIONAL"
    ]

    with open(WITH_WEBSITES_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(with_websites)
    print(f"  Saved {len(with_websites)} businesses with websites → {WITH_WEBSITES_CSV}", flush=True)

    return with_websites


def load_businesses_with_websites() -> list[dict]:
    """Load the filtered CSV if it exists."""
    if not os.path.exists(WITH_WEBSITES_CSV):
        print(f"ERROR: {WITH_WEBSITES_CSV} not found. Run --step scrape first.", file=sys.stderr)
        sys.exit(1)

    with open(WITH_WEBSITES_CSV, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


# ─────────────────────────── STEP 2: LANDER DETECTION ───────────────────────────

def classify_lander(url: str, status_code: int, final_url: str,
                    content: str, size: int, error: str) -> tuple[str, str]:
    """
    Returns (tier, lander_type).
    tier: 'tier1', 'tier2', 'active'
    """
    low = content.lower() if content else ""
    final_domain = urlparse(final_url).netloc.lower().lstrip("www.") if final_url else ""

    # DNS / connection failure
    if error:
        if "dns" in error.lower() or "resolve" in error.lower() or "nodename" in error.lower():
            return "tier1", "DNS Failure"
        if "timeout" in error.lower() or "timed out" in error.lower():
            return "tier2", "Connection Timeout"
        return "tier2", f"Connection Error"

    # Redirect to known parking domain
    if final_domain:
        for pd in PARKING_DOMAINS:
            if final_domain == pd or final_domain.endswith("." + pd):
                return "tier1", "Redirect to Parking"

    # HTTP errors
    if status_code and status_code >= 400:
        if status_code == 403:
            return "tier2", f"HTTP 403 Forbidden"
        return "tier2", f"HTTP {status_code} Error"

    # GoDaddy for sale
    if (
        'window.location.href="/lander"' in content
        or "/lander" in final_url
        or "forsale" in final_url.lower()
        or "this domain is for sale" in low
    ):
        return "tier1", "GoDaddy For Sale"

    # GoDaddy parked
    if "parked free" in low or "godaddy.com/parking" in low:
        return "tier1", "GoDaddy Parked"

    # HugeDomains
    if "hugedomains.com" in low or "hugedomains.com" in final_url.lower():
        return "tier1", "HugeDomains"

    # Sedo
    if "sedo.com" in low or "sedoparking" in low:
        return "tier1", "Sedo Parked"

    # Namecheap
    if "namecheap" in low and ("parked" in low or "parking" in low):
        return "tier1", "Namecheap Parked"

    # Wix expired
    if status_code == 503 and ("wix" in low or "_wix" in low):
        return "tier1", "Wix Expired"
    if "this domain has flown away" in low:
        return "tier1", "Wix Expired"

    # Generic domain for sale
    generic_sale_patterns = [
        "domain is for sale",
        "buy this domain",
        "domain for sale",
        "make an offer",
        "purchase this domain",
        "this domain may be for sale",
    ]
    if any(p in low for p in generic_sale_patterns):
        return "tier1", "Domain For Sale"

    # Squarespace expired
    if "squarespace" in low and ("expired" in low or "not found" in low):
        return "tier1", "Squarespace Expired"

    # Account suspended / hosting expired
    if "account suspended" in low or "account has been suspended" in low or "hosting expired" in low:
        return "tier2", "Hosting Suspended"

    # BlueHost default
    if "bluehost" in low and ("under construction" in low or "default" in low or "parked" in low):
        return "tier2", "BlueHost Default"

    # DreamHost parked
    if "dreamhost" in low and ("parked" in low or "default" in low):
        return "tier2", "DreamHost Parked"

    # HostGator default
    if "hostgator" in low and ("default" in low or "suspended" in low or "under construction" in low):
        return "tier2", "HostGator Default"

    # 5xx server errors
    if status_code and status_code >= 500:
        return "tier2", f"HTTP {status_code} Server Error"

    # Suspiciously small page
    if size and size < 500 and status_code == 200:
        return "tier2", "Suspiciously Small Page"

    return "active", "Active Site"


def check_url(business: dict) -> dict:
    """Fetch one URL and return lander classification."""
    url = business["website"]
    result = {**business, "status_code": "", "final_url": url, "page_size": 0,
              "lander_type": "", "tier": "", "check_error": "", "notes": ""}

    try:
        resp = requests.get(
            url,
            timeout=10,
            allow_redirects=True,
            headers={"User-Agent": USER_AGENT},
            stream=True,
        )
        # Read up to 50KB of content
        content_bytes = b""
        for chunk in resp.iter_content(chunk_size=8192):
            content_bytes += chunk
            if len(content_bytes) >= 51200:
                break

        content = content_bytes.decode("utf-8", errors="replace")
        final_url = resp.url
        status_code = resp.status_code
        size = len(content_bytes)

        tier, lander_type = classify_lander(url, status_code, final_url, content, size, "")

        result.update({
            "status_code": status_code,
            "final_url": final_url,
            "page_size": size,
            "tier": tier,
            "lander_type": lander_type,
        })

    except requests.exceptions.ConnectionError as e:
        err = str(e)
        tier, lander_type = classify_lander(url, None, url, "", 0, err)
        result.update({"tier": tier, "lander_type": lander_type, "check_error": err[:120]})

    except requests.exceptions.Timeout:
        tier, lander_type = classify_lander(url, None, url, "", 0, "timeout")
        result.update({"tier": tier, "lander_type": lander_type, "check_error": "Timeout"})

    except Exception as e:
        err = str(e)
        result.update({"tier": "tier2", "lander_type": "Check Error", "check_error": err[:120]})

    return result


def check_all_landers(businesses: list[dict], workers: int = 10) -> list[dict]:
    """Check all websites concurrently. Returns list with tier/lander_type fields."""
    results: list[dict] = []
    total = len(businesses)
    done = 0

    print(f"\n  Checking {total} URLs with {workers} workers...", flush=True)

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(check_url, b): b for b in businesses}
        for future in as_completed(futures):
            result = future.result()
            results.append(result)
            done += 1
            tier_label = {"tier1": "🔥", "tier2": "⚠", "active": "✅"}.get(result["tier"], "?")
            if done % 25 == 0 or done == total:
                print(f"  [{done}/{total}] last: {tier_label} {result['lander_type'][:40]} — {result['name'][:40]}", flush=True)

    return results


def save_lander_results(results: list[dict]) -> None:
    if not results:
        return
    fieldnames = list(results[0].keys())
    with open(LANDER_RESULTS_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(results)
    print(f"\n  Saved lander results → {LANDER_RESULTS_CSV}", flush=True)


def load_lander_results() -> list[dict]:
    if not os.path.exists(LANDER_RESULTS_CSV):
        print(f"ERROR: {LANDER_RESULTS_CSV} not found. Run --step check first.", file=sys.stderr)
        sys.exit(1)
    with open(LANDER_RESULTS_CSV, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


# ─────────────────────────── STEP 3: EXCEL REPORT ───────────────────────────

def _hex_fill(hex_color: str) -> PatternFill:
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

HEADER_FILL = _hex_fill("2E4057")   # dark blue
TIER1_FILL  = _hex_fill("FFCCCC")   # light red
TIER2_FILL  = _hex_fill("FFF3CC")   # light yellow
ACTIVE_FILL = _hex_fill("CCFFCC")   # light green
WHITE_FILL  = _hex_fill("FFFFFF")
ALT_FILL    = _hex_fill("F7F7F7")

LEAD_COLUMNS = [
    ("Business Name",    "name"),
    ("Category",         "category"),
    ("Phone",            "phone"),
    ("Address",          "address"),
    ("City",             "city"),
    ("Rating",           "rating"),
    ("Review Count",     "review_count"),
    ("Google Maps Link", "maps_url"),
    ("Website URL",      "website"),
    ("Lander Type",      "lander_type"),
    ("Final URL",        "final_url"),
    ("HTTP Status",      "status_code"),
    ("Page Size (bytes)","page_size"),
    ("WHOIS Expiry",     "whois_expiry"),
    ("Days Until Expiry","whois_days_left"),
    ("Notes",            "notes"),
]


def _write_header(ws, columns: list[tuple[str, str]]) -> None:
    for col_idx, (header, _) in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)


def _write_rows(ws, records: list[dict], columns: list[tuple[str, str]], row_fill_fn) -> None:
    for row_idx, record in enumerate(records, start=2):
        fill = row_fill_fn(record, row_idx)
        for col_idx, (_, key) in enumerate(columns, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=record.get(key, ""))
            cell.fill = fill
            cell.alignment = Alignment(wrap_text=False)


def _auto_fit(ws) -> None:
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                val = str(cell.value) if cell.value else ""
                max_len = max(max_len, len(val))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + 4, 60)


def _tier_fill(record, row_idx):
    tier = record.get("tier", "")
    if tier == "tier1":
        return TIER1_FILL
    if tier == "tier2":
        return TIER2_FILL
    return ACTIVE_FILL if tier == "active" else (WHITE_FILL if row_idx % 2 == 0 else ALT_FILL)


def generate_report(results: list[dict], output_path: str) -> None:
    tier1 = [r for r in results if r.get("tier") == "tier1"]
    tier2 = [r for r in results if r.get("tier") == "tier2"]
    active = [r for r in results if r.get("tier") == "active"]
    expiring_soon = [
        r for r in results
        if r.get("whois_days_left") not in ("", None)
        and isinstance(r.get("whois_days_left"), (int, float))
        and 0 <= int(r["whois_days_left"]) <= 90
    ]

    # Sort leads by review count desc
    def _rc(r):
        try:
            return int(float(r.get("review_count", 0) or 0))
        except Exception:
            return 0

    tier1_sorted = sorted(tier1, key=_rc, reverse=True)
    tier2_sorted = sorted(tier2, key=_rc, reverse=True)

    wb = openpyxl.Workbook()

    # ── Sheet 1: Summary ──────────────────────────────────────────
    ws_sum = wb.active
    ws_sum.title = "Summary"

    def _sum_row(label, value, row):
        ws_sum.cell(row=row, column=1, value=label).font = Font(bold=True)
        ws_sum.cell(row=row, column=2, value=value)

    row = 1
    ws_sum.cell(row=row, column=1, value="Treasure Valley Expired Domain Lead Report").font = Font(bold=True, size=14)
    ws_sum.cell(row=row, column=1).fill = HEADER_FILL
    ws_sum.cell(row=row, column=1).font = Font(bold=True, size=14, color="FFFFFF")
    ws_sum.merge_cells(f"A{row}:B{row}")
    row += 1
    ws_sum.cell(row=row, column=1, value=f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}").font = Font(italic=True)
    row += 2

    _sum_row("Total Businesses Scanned", len(results), row); row += 1
    _sum_row("🔥 Confirmed Landers (Tier 1)", len(tier1), row); row += 1
    _sum_row("⚠ Suspicious (Tier 2)", len(tier2), row); row += 1
    _sum_row("✅ Active Sites", len(active), row); row += 1
    _sum_row("📅 Expiring Within 90 Days (any tier)", len(expiring_soon), row); row += 2

    whois_expired = [r for r in results if str(r.get("whois_days_left", "")).lstrip("-").isdigit()
                     and int(r.get("whois_days_left", 1)) < 0]
    _sum_row("  WHOIS Already Expired", len(whois_expired), row); row += 1
    within_30 = [r for r in expiring_soon if int(r["whois_days_left"]) <= 30]
    within_60 = [r for r in expiring_soon if 30 < int(r["whois_days_left"]) <= 60]
    within_90 = [r for r in expiring_soon if 60 < int(r["whois_days_left"]) <= 90]
    _sum_row("  Expiring within 30 days", len(within_30), row); row += 1
    _sum_row("  Expiring within 31–60 days", len(within_60), row); row += 1
    _sum_row("  Expiring within 61–90 days", len(within_90), row); row += 2

    ws_sum.cell(row=row, column=1, value="By Lander Type").font = Font(bold=True, size=12)
    row += 1
    from collections import Counter
    lander_counts = Counter(r.get("lander_type", "Unknown") for r in results if r.get("tier") in ("tier1", "tier2"))
    for ltype, count in lander_counts.most_common():
        _sum_row(ltype, count, row); row += 1
    row += 1

    ws_sum.cell(row=row, column=1, value="By City").font = Font(bold=True, size=12)
    row += 1
    city_counts = Counter(r.get("city", "Unknown") for r in results if r.get("tier") in ("tier1", "tier2"))
    for city, count in city_counts.most_common():
        _sum_row(city, count, row); row += 1
    row += 1

    ws_sum.cell(row=row, column=1, value="By Category (Top 20)").font = Font(bold=True, size=12)
    row += 1
    cat_counts = Counter(r.get("category", "Unknown") for r in results if r.get("tier") in ("tier1", "tier2"))
    for cat, count in cat_counts.most_common(20):
        _sum_row(cat, count, row); row += 1

    ws_sum.column_dimensions["A"].width = 40
    ws_sum.column_dimensions["B"].width = 20

    # ── Sheet 2: Confirmed Leads (Tier 1) ─────────────────────────
    ws_t1 = wb.create_sheet("Confirmed Leads")
    _write_header(ws_t1, LEAD_COLUMNS)
    _write_rows(ws_t1, tier1_sorted, LEAD_COLUMNS, lambda r, i: TIER1_FILL)
    ws_t1.freeze_panes = "A2"
    _auto_fit(ws_t1)

    # ── Sheet 3: Suspicious (Tier 2) ──────────────────────────────
    ws_t2 = wb.create_sheet("Suspicious")
    _write_header(ws_t2, LEAD_COLUMNS)
    _write_rows(ws_t2, tier2_sorted, LEAD_COLUMNS, lambda r, i: TIER2_FILL)
    ws_t2.freeze_panes = "A2"
    _auto_fit(ws_t2)

    # ── Sheet 4: Expiring Soon ────────────────────────────────────
    EXPIRY_ORANGE = _hex_fill("FFE0B2")  # orange tint for expiring soon
    ws_exp = wb.create_sheet("Expiring Soon")
    _write_header(ws_exp, LEAD_COLUMNS)
    expiring_sorted = sorted(
        expiring_soon,
        key=lambda r: int(r.get("whois_days_left", 9999)),
    )
    _write_rows(ws_exp, expiring_sorted, LEAD_COLUMNS, lambda r, i: EXPIRY_ORANGE)
    ws_exp.freeze_panes = "A2"
    _auto_fit(ws_exp)

    # ── Sheet 5: All Businesses ────────────────────────────────────
    ws_all = wb.create_sheet("All Businesses")
    all_columns = LEAD_COLUMNS + [
        ("Tier", "tier"),
        ("Types", "types"),
        ("Place ID", "place_id"),
        ("Business Status", "status"),
        ("Check Error", "check_error"),
    ]
    _write_header(ws_all, all_columns)
    _write_rows(ws_all, results, all_columns, _tier_fill)
    ws_all.freeze_panes = "A2"
    _auto_fit(ws_all)

    wb.save(output_path)
    print(f"\n  ✅ Report saved → {output_path}", flush=True)
    print(f"     🔥 Tier 1 Confirmed Leads : {len(tier1)}", flush=True)
    print(f"     ⚠  Tier 2 Suspicious      : {len(tier2)}", flush=True)
    print(f"     ✅ Active Sites            : {len(active)}", flush=True)


# ─────────────────────────── WHOIS CHECK ───────────────────────────

def check_whois_expiry(domain: str) -> tuple[str, int | None]:
    """
    Look up WHOIS expiry for a domain.
    Returns (expiry_date_str, days_left) where days_left is:
      - negative  → already expired
      - 0+        → days until expiry
      - None      → could not determine
    """
    try:
        import whois as pythonwhois
        w = pythonwhois.query(domain)
        exp = w.expiration_date
        if not exp:
            return "", None
        if isinstance(exp, list):
            exp = exp[0]
        now = datetime.utcnow()
        # Strip timezone if present so comparison works
        if hasattr(exp, "tzinfo") and exp.tzinfo is not None:
            from datetime import timezone
            now = datetime.now(timezone.utc)
        days_left = (exp - now).days
        return exp.strftime("%Y-%m-%d"), days_left
    except ImportError:
        pass  # python-whois not installed
    except Exception:
        pass
    return "", None


def _whois_one(record: dict, warn_days: int) -> dict:
    """Run WHOIS for a single record and return updated copy."""
    url = record.get("website", "")
    if not url:
        return record
    domain = urlparse(url).netloc.lstrip("www.")
    if not domain:
        return record

    expiry_str, days_left = check_whois_expiry(domain)
    if not expiry_str:
        return record

    record = dict(record)  # shallow copy — don't mutate original
    record["whois_expiry"] = expiry_str
    record["whois_days_left"] = days_left if days_left is not None else ""

    if days_left is None:
        return record

    if days_left < 0:
        # Domain already expired per WHOIS
        record["notes"] = f"WHOIS: EXPIRED {expiry_str} ({abs(days_left)}d ago)"
        # Elevate active sites — the domain has technically lapsed
        if record.get("tier") == "active":
            record["tier"] = "tier1"
            record["lander_type"] = "WHOIS Expired"
        elif record.get("tier") == "tier2":
            # Keep tier2 but flag in notes (lander check may have missed it)
            pass
    elif days_left <= warn_days:
        label = (
            f"Expires in {days_left}d ({expiry_str})"
            if days_left > 0
            else f"Expires TODAY ({expiry_str})"
        )
        record["notes"] = f"WHOIS: {label}"
        # Elevate active sites that are expiring soon to tier2
        if record.get("tier") == "active":
            record["tier"] = "tier2"
            record["lander_type"] = f"Expiring Soon ({days_left}d)"
    else:
        record["notes"] = f"WHOIS: Expires {expiry_str} ({days_left}d)"

    return record


def enrich_with_whois(results: list[dict], warn_days: int = 90,
                      workers: int = 10) -> list[dict]:
    """
    Run WHOIS lookups on ALL records (not just existing leads).
    Active sites expiring within warn_days → elevated to tier2.
    Active sites already expired per WHOIS → elevated to tier1.
    Uses a thread pool so 2,000 lookups don't take forever.
    """
    print(
        f"\n  Running WHOIS lookups on ALL {len(results)} domains "
        f"(warn window: {warn_days} days, {workers} workers)...",
        flush=True,
    )

    updated: list[dict] = [None] * len(results)  # type: ignore
    done = 0
    elevated_t1 = 0
    elevated_t2 = 0

    with ThreadPoolExecutor(max_workers=workers) as executor:
        future_to_idx = {
            executor.submit(_whois_one, r, warn_days): i
            for i, r in enumerate(results)
        }
        for future in as_completed(future_to_idx):
            idx = future_to_idx[future]
            rec = future.result()
            updated[idx] = rec
            done += 1

            old_tier = results[idx].get("tier", "")
            new_tier = rec.get("tier", "")
            if new_tier == "tier1" and old_tier != "tier1":
                elevated_t1 += 1
            elif new_tier == "tier2" and old_tier not in ("tier1", "tier2"):
                elevated_t2 += 1

            if done % 100 == 0 or done == len(results):
                print(
                    f"  WHOIS [{done}/{len(results)}] "
                    f"+{elevated_t1} new tier1, +{elevated_t2} new tier2",
                    flush=True,
                )

    print(
        f"\n  WHOIS complete — elevated {elevated_t1} to Tier 1, "
        f"{elevated_t2} to Tier 2 (expiring within {warn_days}d).",
        flush=True,
    )
    return updated


# ─────────────────────────── MAIN ───────────────────────────

def parse_args():
    parser = argparse.ArgumentParser(
        description="Expired Domain Lead Generator for Treasure Valley service businesses"
    )
    parser.add_argument("--api-key", help="Google Places API key (required for scrape step)")
    parser.add_argument(
        "--step",
        choices=["scrape", "check", "report", "all"],
        default="all",
        help="Which step(s) to run (default: all)",
    )
    parser.add_argument(
        "--output",
        default="treasure_valley_leads.xlsx",
        help="Output Excel file path (default: treasure_valley_leads.xlsx)",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=10,
        help="Concurrent workers for URL checking (default: 10)",
    )
    parser.add_argument(
        "--whois",
        action="store_true",
        help="Run WHOIS expiry lookups on ALL businesses (requires 'python-whois' package)",
    )
    parser.add_argument(
        "--whois-days",
        type=int,
        default=90,
        help="Flag domains expiring within this many days as leads (default: 90)",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    step = args.step

    print("=" * 60, flush=True)
    print("  Treasure Valley Expired Domain Lead Generator", flush=True)
    print("=" * 60, flush=True)

    # ── SCRAPE ───────────────────────────────────────────────────
    if step in ("scrape", "all"):
        if not args.api_key:
            print("ERROR: --api-key is required for the scrape step.", file=sys.stderr)
            sys.exit(1)
        print(f"\n[STEP 1] Scraping Google Places API...", flush=True)
        businesses = scrape_places(args.api_key)
        with_websites = save_businesses(businesses)
        print(
            f"\n  Total unique businesses : {len(businesses)}\n"
            f"  With websites (OPERATIONAL): {len(with_websites)}",
            flush=True,
        )

    # ── CHECK ────────────────────────────────────────────────────
    if step in ("check", "all"):
        print(f"\n[STEP 2] Checking websites for expired/parked domains...", flush=True)
        businesses_to_check = load_businesses_with_websites()
        results = check_all_landers(businesses_to_check, workers=args.workers)

        # Ensure WHOIS fields exist on every record before saving
        for r in results:
            r.setdefault("whois_expiry", "")
            r.setdefault("whois_days_left", "")

        if args.whois:
            results = enrich_with_whois(results, warn_days=args.whois_days, workers=args.workers)

        save_lander_results(results)

        tier1_count = sum(1 for r in results if r.get("tier") == "tier1")
        tier2_count = sum(1 for r in results if r.get("tier") == "tier2")
        active_count = sum(1 for r in results if r.get("tier") == "active")
        expiring_count = sum(
            1 for r in results
            if str(r.get("whois_days_left", "")).lstrip("-").isdigit()
            and 0 <= int(r["whois_days_left"]) <= args.whois_days
        )
        print(
            f"\n  Results:\n"
            f"    🔥 Tier 1 (Confirmed Landers) : {tier1_count}\n"
            f"    ⚠  Tier 2 (Suspicious)        : {tier2_count}\n"
            f"    ✅ Active                      : {active_count}\n"
            f"    📅 Expiring ≤{args.whois_days}d (WHOIS) : {expiring_count}",
            flush=True,
        )

    # ── REPORT ───────────────────────────────────────────────────
    if step in ("report", "all"):
        print(f"\n[STEP 3] Generating Excel report → {args.output}...", flush=True)
        results = load_lander_results()
        generate_report(results, args.output)

    print("\n✅ Done!", flush=True)


if __name__ == "__main__":
    main()
