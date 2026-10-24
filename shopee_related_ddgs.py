"""
Shopee SG Related Products Finder - DuckDuckGo (ddgs) Version
================================================================
Fungsi:
- Baca kolom keyword (default J = Nama Produk Tambahan) dari GSheet
- Search produk Shopee SG via DuckDuckGo (gratis, tanpa API key)
- Ambil 3-4 link produk Shopee SG teratas
- Tulis ke kolom link (default T = Link Produk Sejenis)

Environment variables:
- SPREADSHEET_ID          (wajib)
- GOOGLE_SERVICE_ACCOUNT_JSON (wajib)
- SHEET_NAME              (opsional, default "Tracker")
- KEYWORD_COL             (opsional, default "J")
- LINK_COL                (opsional, default "T")
- MAX_ROWS                (opsional, default "50")
- DELAY_SECONDS            (opsional, default "5")
- MAX_CONSECUTIVE_FAILS    (opsional, default "10")

Catatan:
- Tidak butuh API key sama sekali (SERPER_API_KEY dihapus)
- Karena ini scraping non-official, delay antar request dibuat lebih
  besar (default 3 detik) untuk menghindari rate-limit/block dari DDG
- Kalau DDG rate-limit (muncul error timeout/429), script otomatis
  retry dengan backoff sebelum skip baris
- Kolom keyword/link sekarang configurable lewat env var, supaya
  satu script yang sama bisa dipakai untuk beberapa Tracker dengan
  layout kolom yang berbeda (mis. Tracker lama vs Tracker 2)
"""

import os
import json
import time
import random
from datetime import datetime

from ddgs import DDGS
from ddgs.exceptions import DDGSException, RatelimitException

from google.oauth2 import service_account
from googleapiclient.discovery import build

# ============================================================
# KONFIGURASI
# ============================================================

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")
SHEET_NAME     = os.environ.get("SHEET_NAME", "Tracker")
KEYWORD_COL    = os.environ.get("KEYWORD_COL", "J").strip().upper()
LINK_COL       = os.environ.get("LINK_COL", "T").strip().upper()
MAX_ROWS       = int(os.environ.get("MAX_ROWS", "50"))

SHEETS_SCOPES  = ["https://www.googleapis.com/auth/spreadsheets"]

MAX_RESULTS   = 4    # Ambil max 4 link per keyword
DELAY_SECONDS = float(os.environ.get("DELAY_SECONDS", "5"))  # Jeda dasar antar request (detik)
JITTER_SECONDS = 3.0  # Ditambah acak 0-3 detik supaya pola tidak seperti bot
# Circuit breaker: kalau sekian baris BERTURUT-TURUT gagal, DuckDuckGo hampir pasti
# sedang memblokir/throttle IP ini. Hentikan run (jangan buang menit GitHub Actions).
MAX_CONSECUTIVE_FAILS = int(os.environ.get("MAX_CONSECUTIVE_FAILS", "10"))
MAX_RETRIES   = 3    # Retry kalau kena rate-limit


def col_letter_to_index(letter: str) -> int:
    """Konversi huruf kolom (A, B, ..., Z, AA, AB, ...) ke index 0-based."""
    letter = letter.strip().upper()
    result = 0
    for ch in letter:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1


KEYWORD_IDX = col_letter_to_index(KEYWORD_COL)
LINK_IDX    = col_letter_to_index(LINK_COL)
# Lebar baris minimal yang perlu dibaca supaya index kolom keyword & link aman diakses
ROW_WIDTH   = max(KEYWORD_IDX, LINK_IDX) + 1

# ============================================================
# GOOGLE SHEETS HELPERS
# ============================================================

def get_sheets_credentials():
    sa_info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    return service_account.Credentials.from_service_account_info(
        sa_info, scopes=SHEETS_SCOPES
    )

def get_sheets_service():
    return build("sheets", "v4", credentials=get_sheets_credentials())

def get_pending_rows(sheets_service):
    """
    Ambil baris dari sheet yang:
    - Kolom keyword (KEYWORD_COL) tidak kosong
    - Kolom link (LINK_COL) masih kosong
    """
    sheet_meta = sheets_service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID,
        ranges=[f"{SHEET_NAME}!A:A"],
        fields="sheets.properties.gridProperties.rowCount"
    ).execute()
    total_rows = sheet_meta["sheets"][0]["properties"]["gridProperties"]["rowCount"]

    last_col = chr(ord("A") + ROW_WIDTH - 1) if ROW_WIDTH <= 26 else LINK_COL
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A2:{last_col}{total_rows}",
    ).execute()
    rows = result.get("values", [])

    pending = []
    for i, row in enumerate(rows):
        row = row + [""] * (ROW_WIDTH - len(row))
        keyword = row[KEYWORD_IDX].strip()
        link    = row[LINK_IDX].strip()

        if keyword and not link:
            pending.append((i + 2, keyword))

    print(f"[GSHEET] Sheet='{SHEET_NAME}' Total baris: {total_rows}, {len(pending)} baris pending ditemukan "
          f"(keyword col={KEYWORD_COL}, link col={LINK_COL})")
    return pending


def update_link_sejenis(sheets_service, row_index, links: list):
    """Tulis link produk sejenis ke kolom LINK_COL."""
    value = "\n".join(links)
    sheets_service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!{LINK_COL}{row_index}",
        valueInputOption="USER_ENTERED",
        body={"values": [[value]]},
    ).execute()
    print(f"[GSHEET] Row {row_index} → {len(links)} link ditulis ke kolom {LINK_COL}")

# ============================================================
# DUCKDUCKGO SEARCH (ddgs) — GRATIS, TANPA API KEY
# ============================================================

def search_shopee_via_ddg(keyword: str, max_results: int = 4) -> list:
    """
    Search produk Shopee SG via DuckDuckGo menggunakan ddgs.
    Returns list of Shopee product links.
    """
    query = f"site:shopee.sg {keyword}"

    organic = []
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            results = DDGS().text(query, region="sg-en", max_results=10)
            organic = list(results)
            break
        except RatelimitException:
            wait = 5 * attempt
            print(f"[RATE-LIMIT] Kena limit, retry {attempt}/{MAX_RETRIES} dalam {wait}s...")
            time.sleep(wait)
        except DDGSException as e:
            print(f"[ERROR] DDG gagal: {e}")
            return []
        except Exception as e:
            print(f"[ERROR] DDG gagal (unexpected): {e}")
            return []

    if not organic:
        print(f"[SEARCH] Tidak ada hasil untuk: {keyword}")
        return []

    print(f"[SEARCH] '{keyword}' → {len(organic)} hasil DuckDuckGo")

    skip_patterns = [
        "/web/",
        "/list/",
        "sv.shopee.sg",
        "/search",
        "/category/",
        "/mall/",
        "/campaign/",
        "shopee.sg/app",
    ]

    links = []

    # Pass 1: prioritaskan link dengan "-i." (link produk langsung)
    for item in organic:
        link = item.get("href", "")
        if "shopee.sg" not in link:
            continue
        if any(p in link for p in skip_patterns):
            print(f"[SKIP] {link[:60]}...")
            continue
        if "-i." in link:
            clean_link = link.split("?")[0] if "?" in link else link
            if clean_link not in links:
                links.append(clean_link)
                print(f"[OK] {clean_link}")
            if len(links) >= max_results:
                break

    # Pass 2: kalau kurang, ambil link shopee.sg valid lainnya
    if len(links) < max_results:
        for item in organic:
            link = item.get("href", "")
            if "shopee.sg" not in link:
                continue
            if any(p in link for p in skip_patterns):
                continue
            clean_link = link.split("?")[0] if "?" in link else link
            if clean_link not in links:
                links.append(clean_link)
                print(f"[OK+] {clean_link}")
            if len(links) >= max_results:
                break

    return links

# ============================================================
# MAIN PIPELINE
# ============================================================

def run_related_finder(max_rows: int = None):
    if max_rows is None:
        max_rows = MAX_ROWS

    print(f"\n[START] Shopee Related Finder (DuckDuckGo/ddgs) — {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"[CONFIG] Spreadsheet={SPREADSHEET_ID}, Sheet={SHEET_NAME}, "
          f"KeywordCol={KEYWORD_COL}, LinkCol={LINK_COL}, MaxRows={max_rows}")

    sheets_service = get_sheets_service()
    pending_rows   = get_pending_rows(sheets_service)

    if not pending_rows:
        print("[INFO] Tidak ada baris pending.")
        return

    processed = 0
    success = 0
    consecutive_fails = 0
    for row_index, keyword in pending_rows[:max_rows]:
        print(f"\n[ROW {row_index}] Keyword: '{keyword}'")

        links = search_shopee_via_ddg(keyword, max_results=MAX_RESULTS)

        if links:
            update_link_sejenis(sheets_service, row_index, links)
            success += 1
            consecutive_fails = 0
        else:
            print(f"[SKIP] Row {row_index}: tidak ada hasil")
            consecutive_fails += 1

        processed += 1

        if consecutive_fails >= MAX_CONSECUTIVE_FAILS:
            print(f"\n[STOP] {consecutive_fails} baris berturut-turut gagal. "
                  f"Kemungkinan besar DuckDuckGo memblokir/throttle IP ini. "
                  f"Run dihentikan lebih awal; baris yang gagal akan dicoba lagi di run berikutnya.")
            break

        # Jeda acak antar request supaya tidak terdeteksi sebagai bot
        if processed < len(pending_rows[:max_rows]):
            time.sleep(DELAY_SECONDS + random.uniform(0, JITTER_SECONDS))

    print(f"\n[DONE] Diproses {processed} baris, berhasil {success}, gagal {processed - success}.")


if __name__ == "__main__":
    # Mulai dengan max_rows kecil dulu (misal 30-50) untuk test stabilitas
    # sebelum naikin ke jumlah besar, karena DDG lebih rawan rate-limit
    # dibanding paid API. Ganti lewat env var MAX_ROWS kalau perlu.
    run_related_finder()
