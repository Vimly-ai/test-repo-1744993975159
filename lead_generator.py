#!/usr/bin/env python3
"""
Expired Domain Lead Generator for Treasure Valley

Finds service businesses in the Boise/Treasure Valley area whose Google Business
Profile website links point to expired, parked, or broken domains.

Usage:
    # Run everything end-to-end:
    python lead_generator.py --api-key YOUR_KEY --output treasure_valley_leads.xlsx

    # Run in stages:
    python lead_generator.py --api-key YOUR_KEY --step scrape
    python lead_generator.py --step check
    python lead_generator.py --step report
"""

import argparse
import csv
import json
import os
import re
import ssl
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse

import requests
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Top 5 Treasure Valley cities by population / market size
CITIES = [
    "Boise",
    "Meridian",
    "Nampa",
    "Caldwell",
    "Eagle",
]

# 56 service niches (local-pack focus)
CATEGORIES = [
    "Plumber",
    "Plumbing",
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
    "Janitorial",
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
    "Irrigation",
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

# How many results to keep per query (top local-pack positions)
LOCAL_PACK_LIMIT = 7

ALL_BUSINESSES_CSV = "all_businesses.csv"
WEBSITES_CSV = "businesses_with_websites.csv"
NO_WEBSITE_CSV = "businesses_no_website.csv"
CHECKED_CSV = "businesses_checked.csv"
DEFAULT_OUTPUT = "treasure_valley_leads.xlsx"

PLACES_API_URL = "https://places.googleapis.com/v1/places:searchText"
FIELD_MASK = (
    "places.displayName,places.formattedAddress,places.websiteUri,"
    "places.nationalPhoneNumber,places.rating,places.userRatingCount,"
    "places.googleMapsUri,places.businessStatus,places.types,places.id,"
    "nextPageToken"
)

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)

# Known parking/sale domains
PARKING_DOMAINS = {
    "dan.com",
    "afternic.com",
    "sedo.com",
    "hugedomains.com",
    "bodis.com",
    "undeveloped.com",
    "flippa.com",
    "godaddy.com",
    "parkingcrew.net",
    "sedoparking.com",
    "domainmarket.com",
    "buydomains.com",
}

CSV_FIELDS = [
    "place_id",
    "business_name",
    "category",
    "search_category",
    "phone",
    "address",
    "city",
    "rating",
    "review_count",
    "google_maps_url",
    "website_url",
    "business_status",
    "types",
]

CHECKED_FIELDS = CSV_FIELDS + [
    "http_status",
    "final_url",
    "lander_type",
    "tier",
    "page_size",
    "notes",
]


# ---------------------------------------------------------------------------
# Step 1: Google Places API Scraper
# ---------------------------------------------------------------------------

def extract_city_from_address(address):
    """Try to extract the city name from a formatted address."""
    if not address:
        return ""
    for city in CITIES:
        if city.lower() in address.lower():
            return city
    # Fallback: try to parse city from address pattern "..., City, ID ..."
    parts = [p.strip() for p in address.split(",")]
    for part in parts:
        cleaned = re.sub(r"\s+\d{5}.*", "", part).strip()
        if cleaned and cleaned not in ("ID", "Idaho"):
            for city in CITIES:
                if city.lower() == cleaned.lower():
                    return city
    return ""


def search_places(api_key, query, page_token=None):
    """Execute a single Places API text search request."""
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": api_key,
        "X-Goog-FieldMask": FIELD_MASK,
    }
    body = {"textQuery": query, "languageCode": "en"}
    if page_token:
        body["pageToken"] = page_token

    resp = requests.post(PLACES_API_URL, headers=headers, json=body, timeout=30)
    resp.raise_for_status()
    return resp.json()


def parse_place(place, search_category):
    """Parse a single place result into a flat dict."""
    display_name = place.get("displayName", {})
    name = display_name.get("text", "") if isinstance(display_name, dict) else str(display_name)
    address = place.get("formattedAddress", "")
    types_list = place.get("types", [])

    return {
        "place_id": place.get("id", ""),
        "business_name": name,
        "category": types_list[0] if types_list else search_category,
        "search_category": search_category,
        "phone": place.get("nationalPhoneNumber", ""),
        "address": address,
        "city": extract_city_from_address(address),
        "rating": place.get("rating", ""),
        "review_count": place.get("userRatingCount", 0),
        "google_maps_url": place.get("googleMapsUri", ""),
        "website_url": place.get("websiteUri", ""),
        "business_status": place.get("businessStatus", ""),
        "types": "|".join(types_list),
    }


def scrape_places(api_key):
    """Scrape top local-pack businesses from the Google Places API.

    Only keeps the top LOCAL_PACK_LIMIT results per query (simulating the
    top positions in the Google local pack).  Does NOT paginate — we only
    want businesses that rank highly.
    """
    all_businesses = {}  # place_id -> business dict
    total_queries = len(CITIES) * len(CATEGORIES)
    completed = 0

    print(f"\n{'='*60}")
    print("STEP 1: Scraping Google Places API (top {0} local pack)".format(LOCAL_PACK_LIMIT))
    print(f"{'='*60}")
    print(f"Cities: {len(CITIES)}")
    print(f"Niches: {len(CATEGORIES)}")
    print(f"Total queries: {total_queries}")
    print(f"Max results per query: {LOCAL_PACK_LIMIT}")
    print()

    for city in CITIES:
        for category in CATEGORIES:
            query = f"{category} in {city}, Idaho"
            completed += 1

            try:
                # Single request — no pagination. We only want top-ranking
                # businesses (positions 1-7 in the local pack).
                data = search_places(api_key, query)
                places = data.get("places", [])

                # Keep only the top LOCAL_PACK_LIMIT results
                places = places[:LOCAL_PACK_LIMIT]

                new_count = 0
                for place in places:
                    biz = parse_place(place, category)

                    # Only keep OPERATIONAL businesses
                    if biz["business_status"] != "OPERATIONAL":
                        continue

                    pid = biz["place_id"]
                    if pid and pid not in all_businesses:
                        all_businesses[pid] = biz
                        new_count += 1

                progress = f"[{completed}/{total_queries}]"
                print(
                    f"  {progress} {query} — "
                    f"{len(places)} results, "
                    f"{new_count} new  (total unique: {len(all_businesses)})"
                )

            except requests.exceptions.HTTPError as e:
                print(f"  [ERROR] {query}: HTTP {e.response.status_code} — {e}")
            except Exception as e:
                print(f"  [ERROR] {query}: {e}")

            time.sleep(0.15)

    # Split into: with websites / without websites
    all_list = list(all_businesses.values())
    with_websites = [b for b in all_list if b.get("website_url")]
    no_websites = [b for b in all_list if not b.get("website_url")]

    print(f"\n{'─'*60}")
    print(f"Total unique businesses found: {len(all_list)}")
    print(f"Businesses WITH website URLs:  {len(with_websites)}")
    print(f"Businesses WITHOUT websites:   {len(no_websites)}")
    print(f"{'─'*60}\n")

    # Save CSVs
    _write_csv(ALL_BUSINESSES_CSV, CSV_FIELDS, all_list)
    _write_csv(WEBSITES_CSV, CSV_FIELDS, with_websites)
    _write_csv(NO_WEBSITE_CSV, CSV_FIELDS, no_websites)

    print(f"Saved {ALL_BUSINESSES_CSV}  ({len(all_list)} rows)")
    print(f"Saved {WEBSITES_CSV}  ({len(with_websites)} rows)")
    print(f"Saved {NO_WEBSITE_CSV}  ({len(no_websites)} rows)")

    return with_websites


# ---------------------------------------------------------------------------
# Step 2: Lander / Expired Domain Detection
# ---------------------------------------------------------------------------

def classify_site(url):
    """
    Fetch a URL and classify whether it's an active site or a lander/expired domain.
    Returns a dict with detection results.
    """
    result = {
        "http_status": "",
        "final_url": "",
        "lander_type": "Active",
        "tier": "Active",
        "page_size": 0,
        "notes": "",
    }

    # Ensure URL has a scheme
    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    try:
        resp = requests.get(
            url,
            timeout=10,
            allow_redirects=True,
            headers={"User-Agent": USER_AGENT},
            verify=True,
        )
        result["http_status"] = resp.status_code
        result["final_url"] = resp.url
        result["page_size"] = len(resp.content)

        content = resp.text.lower() if resp.text else ""
        final_domain = urlparse(resp.url).netloc.lower()

        # --- Check redirects to known parking domains ---
        for parking in PARKING_DOMAINS:
            if parking in final_domain:
                result["lander_type"] = f"Redirect to parking ({parking})"
                result["tier"] = "Tier 1"
                result["notes"] = f"Redirected to {resp.url}"
                return result

        # --- GoDaddy Forsale ---
        if (
            'window.location.href="/lander"' in content
            or "/forsale" in resp.url.lower()
            or "this domain is for sale" in content
        ):
            result["lander_type"] = "GoDaddy Forsale"
            result["tier"] = "Tier 1"
            return result

        # --- GoDaddy Parked ---
        if "parked free" in content or "godaddy.com/parking" in content:
            result["lander_type"] = "GoDaddy Parked"
            result["tier"] = "Tier 1"
            return result

        # --- HugeDomains ---
        if "hugedomains.com" in content or "hugedomains.com" in final_domain:
            result["lander_type"] = "HugeDomains"
            result["tier"] = "Tier 1"
            return result

        # --- Sedo ---
        if "sedo.com" in content or "sedoparking" in content:
            result["lander_type"] = "Sedo Parked"
            result["tier"] = "Tier 1"
            return result

        # --- Wix Expired ---
        if (
            (resp.status_code == 503 and ("wix" in content or "wixstatic" in content))
            or "this domain has flown away" in content
        ):
            result["lander_type"] = "Wix Expired"
            result["tier"] = "Tier 1"
            return result

        # --- Domain For Sale (Generic) ---
        sale_phrases = [
            "domain is for sale",
            "buy this domain",
            "domain for sale",
            "make an offer",
            "this website is for sale",
        ]
        for phrase in sale_phrases:
            if phrase in content:
                result["lander_type"] = "Domain For Sale"
                result["tier"] = "Tier 1"
                result["notes"] = f"Matched: '{phrase}'"
                return result

        # --- Namecheap Parked ---
        if "namecheap" in content and ("parked" in content or "parking page" in content):
            result["lander_type"] = "Namecheap Parked"
            result["tier"] = "Tier 2"
            return result

        # --- BlueHost Default ---
        if "bluehost" in content and (
            "this site is under construction" in content
            or "coming soon" in content
            or "welcome to bluehost" in content
        ):
            result["lander_type"] = "BlueHost Default"
            result["tier"] = "Tier 2"
            return result

        # --- DreamHost Parked ---
        if "dreamhost" in content and (
            "parked" in content or "default" in content or "coming soon" in content
        ):
            result["lander_type"] = "DreamHost Parked"
            result["tier"] = "Tier 2"
            return result

        # --- HostGator Default ---
        if "hostgator" in content and (
            "default" in content or "coming soon" in content or "under construction" in content
        ):
            result["lander_type"] = "HostGator Default"
            result["tier"] = "Tier 2"
            return result

        # --- Squarespace Expired ---
        if "squarespace" in content and (
            "expired" in content or "site is expired" in content
        ):
            result["lander_type"] = "Squarespace Expired"
            result["tier"] = "Tier 1"
            return result

        # --- Hosting Suspended ---
        suspended_phrases = [
            "account suspended",
            "account has been suspended",
            "hosting expired",
            "this account has been suspended",
        ]
        for phrase in suspended_phrases:
            if phrase in content:
                result["lander_type"] = "Hosting Suspended"
                result["tier"] = "Tier 2"
                result["notes"] = f"Matched: '{phrase}'"
                return result

        # --- HTTP Errors (not 403) ---
        if resp.status_code >= 400 and resp.status_code != 403:
            result["lander_type"] = f"HTTP Error ({resp.status_code})"
            result["tier"] = "Tier 2"
            return result

        # --- Suspiciously Small Page ---
        if len(resp.content) < 500:
            # Check if it looks like a real redirect or minimal page
            if not any(
                keyword in content
                for keyword in ["business", "service", "contact", "phone", "call", "about"]
            ):
                result["lander_type"] = "Suspiciously Small Page"
                result["tier"] = "Tier 2"
                result["notes"] = f"Page size: {len(resp.content)} bytes"
                return result

        # If we got here, it's an active site
        return result

    except requests.exceptions.SSLError:
        # Retry without SSL verification
        try:
            resp = requests.get(
                url,
                timeout=10,
                allow_redirects=True,
                headers={"User-Agent": USER_AGENT},
                verify=False,
            )
            result["http_status"] = resp.status_code
            result["final_url"] = resp.url
            result["page_size"] = len(resp.content)
            result["notes"] = "SSL error (proceeded without verification)"

            content = resp.text.lower() if resp.text else ""
            final_domain = urlparse(resp.url).netloc.lower()

            # Re-run the same checks on the SSL-error retry
            for parking in PARKING_DOMAINS:
                if parking in final_domain:
                    result["lander_type"] = f"Redirect to parking ({parking})"
                    result["tier"] = "Tier 1"
                    return result

            if len(resp.content) < 500:
                if not any(
                    kw in content
                    for kw in ["business", "service", "contact", "phone", "call", "about"]
                ):
                    result["lander_type"] = "Suspiciously Small Page"
                    result["tier"] = "Tier 2"
                    return result

            return result

        except Exception as e2:
            result["lander_type"] = "SSL/Connection Error"
            result["tier"] = "Tier 2"
            result["notes"] = str(e2)[:200]
            return result

    except requests.exceptions.ConnectionError as e:
        err_str = str(e).lower()
        if "name or service not known" in err_str or "nodename nor servname" in err_str or "getaddrinfo failed" in err_str:
            result["lander_type"] = "DNS Failure"
            result["tier"] = "Tier 1"
        else:
            result["lander_type"] = "Connection Error"
            result["tier"] = "Tier 2"
        result["notes"] = str(e)[:200]
        return result

    except requests.exceptions.Timeout:
        result["lander_type"] = "Connection Timeout"
        result["tier"] = "Tier 2"
        result["notes"] = "Request timed out after 10 seconds"
        return result

    except Exception as e:
        result["lander_type"] = "Error"
        result["tier"] = "Tier 2"
        result["notes"] = str(e)[:200]
        return result


def check_single_business(biz):
    """Check a single business URL and return the updated dict."""
    url = biz.get("website_url", "")
    if not url:
        biz.update({
            "http_status": "",
            "final_url": "",
            "lander_type": "No URL",
            "tier": "",
            "page_size": 0,
            "notes": "",
        })
        return biz

    result = classify_site(url)
    biz.update(result)
    return biz


def check_websites(businesses=None, workers=10):
    """Check all business websites for lander/expired status."""
    if businesses is None:
        businesses = _read_csv(WEBSITES_CSV, CSV_FIELDS)
        if not businesses:
            print(f"ERROR: No businesses found. Run --step scrape first or check {WEBSITES_CSV}")
            return []

    total = len(businesses)
    print(f"\n{'='*60}")
    print("STEP 2: Checking Websites for Expired/Parked Domains")
    print(f"{'='*60}")
    print(f"Total URLs to check: {total}")
    print(f"Workers: {workers}")
    print()

    checked = []
    tier1_count = 0
    tier2_count = 0
    active_count = 0
    completed = 0

    with ThreadPoolExecutor(max_workers=workers) as executor:
        future_to_biz = {
            executor.submit(check_single_business, biz.copy()): biz
            for biz in businesses
        }

        for future in as_completed(future_to_biz):
            completed += 1
            try:
                result = future.result()
                checked.append(result)

                tier = result.get("tier", "Active")
                if tier == "Tier 1":
                    tier1_count += 1
                    marker = "🔥"
                elif tier == "Tier 2":
                    tier2_count += 1
                    marker = "⚠️ "
                else:
                    active_count += 1
                    marker = "✅"

                print(
                    f"  [{completed}/{total}] {marker} {result['business_name'][:40]:<40} "
                    f"| {result.get('lander_type', 'Active')}"
                )
            except Exception as e:
                completed_biz = future_to_biz[future]
                completed_biz.update({
                    "http_status": "",
                    "final_url": "",
                    "lander_type": "Check Error",
                    "tier": "Tier 2",
                    "page_size": 0,
                    "notes": str(e)[:200],
                })
                checked.append(completed_biz)
                tier2_count += 1
                print(f"  [{completed}/{total}] ❌ {completed_biz['business_name'][:40]:<40} | Error: {e}")

            # Save intermediate results every 50 checks
            if completed % 50 == 0:
                _write_csv(CHECKED_CSV, CHECKED_FIELDS, checked)

    # Final save
    _write_csv(CHECKED_CSV, CHECKED_FIELDS, checked)

    print(f"\n{'─'*60}")
    print(f"Results:")
    print(f"  🔥 Tier 1 (Confirmed Landers): {tier1_count}")
    print(f"  ⚠️  Tier 2 (Suspicious):        {tier2_count}")
    print(f"  ✅ Active Sites:                {active_count}")
    print(f"{'─'*60}")
    print(f"Saved {CHECKED_CSV} ({len(checked)} rows)\n")

    return checked


# ---------------------------------------------------------------------------
# Step 3: Excel Report
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
TIER1_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
TIER2_FILL = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
ACTIVE_FILL = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")

REPORT_COLUMNS = [
    "Business Name",
    "Category",
    "Phone Number",
    "Address",
    "City",
    "Rating",
    "Review Count",
    "Google Maps Link",
    "Website URL",
    "Lander Type",
    "Tier",
    "Final URL",
    "Notes",
]


def _biz_to_report_row(biz):
    """Convert a business dict to a report row tuple."""
    return (
        biz.get("business_name", ""),
        biz.get("category", ""),
        biz.get("phone", ""),
        biz.get("address", ""),
        biz.get("city", ""),
        biz.get("rating", ""),
        biz.get("review_count", 0),
        biz.get("google_maps_url", ""),
        biz.get("website_url", ""),
        biz.get("lander_type", ""),
        biz.get("tier", ""),
        biz.get("final_url", ""),
        biz.get("notes", ""),
    )


def _style_header(ws):
    """Apply header styling to the first row."""
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"


def _auto_width(ws):
    """Auto-fit column widths based on content."""
    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        col_letter = get_column_letter(col_idx)
        for row in ws.iter_rows(min_col=col_idx, max_col=col_idx, values_only=False):
            for cell in row:
                val = str(cell.value) if cell.value else ""
                max_len = max(max_len, len(val))
        # Cap at 50 chars
        adjusted = min(max_len + 2, 50)
        ws.column_dimensions[col_letter].width = adjusted


def _color_rows(ws, tier_col_idx):
    """Color-code rows based on tier value."""
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        tier_val = str(row[tier_col_idx].value).strip() if row[tier_col_idx].value else ""
        if tier_val == "Tier 1":
            fill = TIER1_FILL
        elif tier_val == "Tier 2":
            fill = TIER2_FILL
        elif tier_val == "Active":
            fill = ACTIVE_FILL
        else:
            continue
        for cell in row:
            cell.fill = fill


def generate_report(businesses=None, output_file=DEFAULT_OUTPUT):
    """Generate the Excel report.

    ``businesses`` contains checked businesses (those that had a website URL).
    The no-website list is loaded separately from NO_WEBSITE_CSV.
    """
    if businesses is None:
        businesses = _read_csv(CHECKED_CSV, CHECKED_FIELDS)
        if not businesses:
            print(f"ERROR: No checked data found. Run --step check first or check {CHECKED_CSV}")
            return

    # Load businesses that had no website link in their GBP
    no_website_biz = _read_csv(NO_WEBSITE_CSV, CSV_FIELDS)

    print(f"\n{'='*60}")
    print("STEP 3: Generating Excel Report")
    print(f"{'='*60}")

    tier1 = [b for b in businesses if b.get("tier") == "Tier 1"]
    tier2 = [b for b in businesses if b.get("tier") == "Tier 2"]
    active = [b for b in businesses if b.get("tier") == "Active"]

    # Sort leads by review count descending (highest reviews = best leads)
    def sort_key(b):
        try:
            return -int(float(b.get("review_count", 0)))
        except (ValueError, TypeError):
            return 0

    tier1.sort(key=sort_key)
    tier2.sort(key=sort_key)

    wb = openpyxl.Workbook()

    # ---- Sheet 1: Summary ----
    ws_summary = wb.active
    ws_summary.title = "Summary"
    ws_summary.sheet_properties.tabColor = "2F5496"

    summary_header_font = Font(bold=True, size=14, color="2F5496")
    section_font = Font(bold=True, size=12)
    normal_font = Font(size=11)

    row = 1
    ws_summary.cell(row=row, column=1, value="Treasure Valley Expired Domain Lead Report").font = summary_header_font
    row += 2

    # Overview stats
    ws_summary.cell(row=row, column=1, value="Overview").font = section_font
    row += 1
    total_all = len(businesses) + len(no_website_biz)
    stats = [
        ("Total businesses scraped (top {0} local pack)".format(LOCAL_PACK_LIMIT), total_all),
        ("Businesses with GBP website link", len(businesses)),
        ("Businesses WITHOUT website link", len(no_website_biz)),
        ("Confirmed landers (Tier 1)", len(tier1)),
        ("Suspicious (Tier 2)", len(tier2)),
        ("Active sites", len(active)),
    ]
    for label, value in stats:
        ws_summary.cell(row=row, column=1, value=label).font = normal_font
        ws_summary.cell(row=row, column=2, value=value).font = normal_font
        row += 1

    row += 1

    # Breakdown by lander type
    ws_summary.cell(row=row, column=1, value="Breakdown by Lander Type").font = section_font
    row += 1
    lander_counts = {}
    for b in businesses:
        lt = b.get("lander_type", "Active")
        lander_counts[lt] = lander_counts.get(lt, 0) + 1
    for lt, count in sorted(lander_counts.items(), key=lambda x: -x[1]):
        ws_summary.cell(row=row, column=1, value=lt).font = normal_font
        ws_summary.cell(row=row, column=2, value=count).font = normal_font
        row += 1

    row += 1

    # Breakdown by category (leads only)
    ws_summary.cell(row=row, column=1, value="Leads by Business Category").font = section_font
    row += 1
    cat_counts = {}
    for b in tier1 + tier2:
        cat = b.get("category", "Unknown")
        cat_counts[cat] = cat_counts.get(cat, 0) + 1
    for cat, count in sorted(cat_counts.items(), key=lambda x: -x[1]):
        ws_summary.cell(row=row, column=1, value=cat).font = normal_font
        ws_summary.cell(row=row, column=2, value=count).font = normal_font
        row += 1

    row += 1

    # Breakdown by city (leads only)
    ws_summary.cell(row=row, column=1, value="Leads by City").font = section_font
    row += 1
    city_counts = {}
    for b in tier1 + tier2:
        city = b.get("city", "Unknown") or "Unknown"
        city_counts[city] = city_counts.get(city, 0) + 1
    for city, count in sorted(city_counts.items(), key=lambda x: -x[1]):
        ws_summary.cell(row=row, column=1, value=city).font = normal_font
        ws_summary.cell(row=row, column=2, value=count).font = normal_font
        row += 1

    ws_summary.column_dimensions["A"].width = 45
    ws_summary.column_dimensions["B"].width = 15

    # ---- Sheet 2: Confirmed Leads (Tier 1) ----
    ws_leads = wb.create_sheet("Confirmed Leads")
    ws_leads.sheet_properties.tabColor = "FF0000"
    ws_leads.append(REPORT_COLUMNS)
    for biz in tier1:
        ws_leads.append(_biz_to_report_row(biz))
    _style_header(ws_leads)
    _auto_width(ws_leads)
    tier_col = REPORT_COLUMNS.index("Tier")
    _color_rows(ws_leads, tier_col)

    # ---- Sheet 3: Suspicious (Tier 2) ----
    ws_suspicious = wb.create_sheet("Suspicious")
    ws_suspicious.sheet_properties.tabColor = "FFD700"
    ws_suspicious.append(REPORT_COLUMNS)
    for biz in tier2:
        ws_suspicious.append(_biz_to_report_row(biz))
    _style_header(ws_suspicious)
    _auto_width(ws_suspicious)
    _color_rows(ws_suspicious, tier_col)

    # ---- Sheet 4: No Website ----
    # Businesses ranking in the local pack but with NO website link in GBP.
    # These are a separate outreach list — not Tier 1/Tier 2 leads.
    NO_WEBSITE_COLUMNS = [
        "Business Name",
        "Category",
        "Phone Number",
        "Address",
        "City",
        "Rating",
        "Review Count",
        "Google Maps Link",
        "Notes",
    ]

    def _no_website_row(biz):
        return (
            biz.get("business_name", ""),
            biz.get("category", ""),
            biz.get("phone", ""),
            biz.get("address", ""),
            biz.get("city", ""),
            biz.get("rating", ""),
            biz.get("review_count", 0),
            biz.get("google_maps_url", ""),
            "No website link on GBP — never had a site or removed it",
        )

    no_website_biz_sorted = sorted(no_website_biz, key=sort_key)
    ws_nosite = wb.create_sheet("No Website")
    ws_nosite.sheet_properties.tabColor = "808080"
    ws_nosite.append(NO_WEBSITE_COLUMNS)
    for biz in no_website_biz_sorted:
        ws_nosite.append(_no_website_row(biz))
    _style_header(ws_nosite)
    _auto_width(ws_nosite)

    # ---- Sheet 5: All Businesses ----
    ws_all = wb.create_sheet("All Businesses")
    ws_all.sheet_properties.tabColor = "2F5496"
    ws_all.append(REPORT_COLUMNS)
    for biz in businesses:
        ws_all.append(_biz_to_report_row(biz))
    _style_header(ws_all)
    _auto_width(ws_all)
    _color_rows(ws_all, tier_col)

    wb.save(output_file)

    print(f"\nReport saved to: {output_file}")
    print(f"  Sheet 'Summary':         Overview stats and breakdowns")
    print(f"  Sheet 'Confirmed Leads': {len(tier1)} Tier 1 leads (sorted by reviews)")
    print(f"  Sheet 'Suspicious':      {len(tier2)} Tier 2 leads")
    print(f"  Sheet 'No Website':      {len(no_website_biz)} businesses without GBP website")
    print(f"  Sheet 'All Businesses':  {len(businesses)} total checked entries")
    print()


# ---------------------------------------------------------------------------
# CSV Helpers
# ---------------------------------------------------------------------------

def _write_csv(filepath, fields, rows):
    """Write a list of dicts to a CSV file."""
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def _read_csv(filepath, fields):
    """Read a CSV file into a list of dicts."""
    if not os.path.exists(filepath):
        return []
    rows = []
    with open(filepath, "r", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# Main Entry Point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Expired Domain Lead Generator for Treasure Valley",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Run everything end-to-end:
  python lead_generator.py --api-key YOUR_KEY --output treasure_valley_leads.xlsx

  # Run in stages:
  python lead_generator.py --api-key YOUR_KEY --step scrape   # Pull businesses
  python lead_generator.py --step check                        # Check URLs
  python lead_generator.py --step report                       # Generate report
        """,
    )
    parser.add_argument(
        "--api-key",
        help="Google Places API key (required for 'scrape' step)",
    )
    parser.add_argument(
        "--step",
        choices=["scrape", "check", "report", "all"],
        default="all",
        help="Which step to run (default: all)",
    )
    parser.add_argument(
        "--output",
        default=DEFAULT_OUTPUT,
        help=f"Output Excel filename (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=10,
        help="Number of concurrent workers for URL checking (default: 10)",
    )

    args = parser.parse_args()

    print("=" * 60)
    print("  Expired Domain Lead Generator — Treasure Valley")
    print("=" * 60)

    businesses = None

    if args.step in ("scrape", "all"):
        if not args.api_key:
            parser.error("--api-key is required for the 'scrape' step")
        businesses = scrape_places(args.api_key)

    if args.step in ("check", "all"):
        businesses = check_websites(businesses, workers=args.workers)

    if args.step in ("report", "all"):
        generate_report(businesses, output_file=args.output)

    print("\nDone!")


if __name__ == "__main__":
    main()
