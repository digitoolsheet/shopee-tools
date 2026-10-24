"""
Shopee SG Product Category Classifier - Hybrid Version
==========================================================
Fungsi:
- Baca kolom Nama Produk (posisi diatur via env PRODUCT_COL) dari GSheet Tracker
- STAGE 1: Coba klasifikasi pakai keyword matching (instant, gratis, no limit)
- STAGE 2: Kalau tidak ketemu keyword yang cocok, fallback ke Groq AI
- Tulis hasil ke kolom Kategori (posisi diatur via env CATEGORY_COL)

Environment variables:
- GROQ_API_KEY       (buat fallback AI, daftar gratis di console.groq.com)
- SPREADSHEET_ID
- GOOGLE_SERVICE_ACCOUNT_JSON
- SHEET_NAME               (default "Tracker")
- PRODUCT_COL              (default "I")  -> huruf kolom Nama Produk
- CATEGORY_COL             (default "X")  -> huruf kolom Kategori

Perubahan vs versi sebelumnya:
- FIX: PRODUCT_COL / CATEGORY_COL sekarang BENAR-BENAR dipakai (dibaca dari
  environment variable dan dikonversi ke index kolom), tidak lagi hardcode
  ke I/X.
- FIX: deteksi baris "pending" (perlu diklasifikasi/reklasifikasi) sekarang
  bukan cuma kalau kolom Kategori kosong atau "Miscellaneous" saja, tapi
  juga kalau isinya BUKAN salah satu dari 26 kategori resmi di
  SHOPEE_CATEGORIES. Nilai kategori lama (termasuk yang masih berbahasa
  Indonesia / nama kategori versi lama) ikut direklasifikasi ulang.
- FIX PENTING: SHOPEE_CATEGORIES diganti dari 23 kategori versi LAMA
  (termasuk "Muslim Fashion" yang sebenarnya bukan kategori resmi Shopee SG)
  ke 26 kategori RESMI Shopee SG yang sama dengan yang dipakai tab `list`
  Tracker 2. KEYWORD_MAP dipecah mengikuti kategori baru ini:
    - "Muslim Fashion" dihapus -> keyword-nya disebar: hijab/jilbab/gamis/
      abaya/mukena/kerudung -> "Women's Apparel"; koko/sarung -> "Men's Wear"
    - "Bags & Luggage" -> koper/luggage -> "Travel & Luggage"; sisanya
      (ransel, tote bag, sling bag, dompet, dll) -> "Women's Bags"
    - "Health & Beauty" -> vitamin/suplemen/collagen -> "Health & Wellness";
      sisanya (skincare, makeup, shampoo, dll) -> "Beauty & Personal Care"
    - "Babies & Kids" + "Toys, Hobbies & Books" -> perlengkapan & mainan bayi
      -> "Toys Kids & Babies"; baju bayi -> "Kids Fashion"; buku/board game
      -> "Hobbies & Books"
    - "Women's Fashion" -> "Women's Apparel"; "Men's Fashion" -> "Men's Wear"
    - "Cameras" -> "Cameras & Drones"; "Groceries" -> "Food & Beverages";
      "Pet Food & Supplies" -> "Pet Supplies"; "Tickets, Vouchers & Services"
      -> "Dining Travel & Services"; "Others" -> "Miscellaneous"
- Range baca sheet tetap DINAMIS (ambil rowCount aktual), bukan hardcode
  "A2:X5000" lagi — jadi baris baru di luar 5000 (atau berapa pun) tetap
  ke-scan.
"""

import os
import json
import time
import re
from datetime import datetime

import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ============================================================
# KONFIGURASI
# ============================================================

GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")
SHEET_NAME = os.environ.get("SHEET_NAME", "Tracker")
PRODUCT_COL_LETTER = os.environ.get("PRODUCT_COL", "I").strip().upper()
CATEGORY_COL_LETTER = os.environ.get("CATEGORY_COL", "X").strip().upper()

SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

GROQ_MODEL = "llama-3.3-70b-versatile"
GROQ_URL = "https://api.groq.com/openai/v1/chat/completions"

AI_BATCH_SIZE = 20  # jumlah produk per 1 kali panggilan AI (buat yang gak ketemu keyword)
DELAY_BETWEEN_AI_CALLS = 2.0
MAX_RETRIES = 5
RETRY_BASE_DELAY = 10


def col_letter_to_index(letter: str) -> int:
    """Konversi huruf kolom (A, B, ..., Z, AA, AB, ...) ke index 0-based."""
    letter = letter.strip().upper()
    result = 0
    for ch in letter:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1


PRODUCT_COL_IDX = col_letter_to_index(PRODUCT_COL_LETTER)
CATEGORY_COL_IDX = col_letter_to_index(CATEGORY_COL_LETTER)
MAX_COL_IDX = max(PRODUCT_COL_IDX, CATEGORY_COL_IDX)
READ_RANGE_LAST_COL_LETTER = CATEGORY_COL_LETTER if CATEGORY_COL_IDX >= PRODUCT_COL_IDX else PRODUCT_COL_LETTER

# 26 kategori RESMI Shopee SG (sama dengan yang dipakai tab `list` Tracker 2)
SHOPEE_CATEGORIES = [
    "Women's Apparel",
    "Men's Wear",
    "Women's Shoes",
    "Men's Shoes",
    "Women's Bags",
    "Men's Bags",
    "Travel & Luggage",
    "Watches",
    "Jewellery & Accessories",
    "Mobile & Gadgets",
    "Computers & Peripherals",
    "Cameras & Drones",
    "Home Appliances",
    "Home & Living",
    "Beauty & Personal Care",
    "Health & Wellness",
    "Kids Fashion",
    "Toys Kids & Babies",
    "Hobbies & Books",
    "Sports & Outdoors",
    "Automotive",
    "Food & Beverages",
    "Pet Supplies",
    "Video Games",
    "Dining Travel & Services",
    "Miscellaneous",
]

CATEGORIES_STR = ", ".join(SHOPEE_CATEGORIES)
VALID_CATEGORIES_SET = set(SHOPEE_CATEGORIES)

# Nilai yang dianggap "belum diklasifikasi dengan benar" dan perlu direklasifikasi:
# - kosong
# - ATAU bukan salah satu dari 26 kategori resmi di atas (termasuk nilai lama/
#   peninggalan proses klasifikasi jaman dulu, misalnya masih berbahasa
#   Indonesia atau nama kategori versi lama seperti "Muslim Fashion")
def is_pending_category(value: str) -> bool:
    value = (value or "").strip()
    if not value:
        return True
    if value not in VALID_CATEGORIES_SET:
        return True
    return False

# ============================================================
# STAGE 1: KEYWORD MATCHING (instant, gratis, no limit)
# ============================================================
# Urutan penting: keyword yang lebih spesifik ditaruh duluan supaya
# tidak "ketebak" kategori umum yang salah.

KEYWORD_MAP = {
    "Health & Wellness": [
        "vitamin", "supplement", "suplemen", "collagen", "kolagen",
    ],
    "Beauty & Personal Care": [
        "serum", "skincare", "skin care", "moisturizer", "sunscreen", "cream", "krim wajah",
        "lipstick", "lip", "makeup", "make up", "shampoo", "sampo", "conditioner",
        "sabun", "soap", "parfum", "perfume", "deodorant", "toothpaste", "pasta gigi",
        "face mask", "masker wajah", "acne", "jerawat", "whitening", "essence",
        "toner", "hair oil", "minyak rambut", "body wash", "body lotion",
    ],
    "Mobile & Gadgets": [
        "iphone", "samsung galaxy", "handphone", "hp ", "smartphone", "charger",
        "cable usb", "kabel data", "power bank", "powerbank", "earphone",
        "earbuds", "headset", "bluetooth", "case hp", "casing hp", "tempered glass",
        "screen protector", "phone holder", "tripod hp",
    ],
    "Computers & Peripherals": [
        "laptop", "keyboard", "mouse", "monitor", "webcam", "flashdisk",
        "hardisk", "ssd", "ram laptop", "printer", "router wifi", "modem",
    ],
    "Cameras & Drones": [
        "kamera", "camera", "dslr", "mirrorless", "lensa kamera", "gopro",
        "action cam", "tripod kamera", "drone",
    ],
    "Home Appliances": [
        "kulkas", "refrigerator", "mesin cuci", "washing machine", "ac split",
        "air conditioner", "rice cooker", "blender", "mixer", "microwave",
        "kompor", "stove", "vacuum cleaner", "penyedot debu", "air fryer",
        "dispenser", "water heater", "kipas angin", "fan electric", "setrika", "iron",
        "bread maker", "waffle maker", "electric kettle", "hair dryer", "air purifier",
    ],
    "Home & Living": [
        "sprei", "bed sheet", "bantal", "pillow", "guling", "selimut", "blanket",
        "gorden", "curtain", "karpet", "carpet", "rak", "shelf", "lemari",
        "wardrobe", "meja", "table", "kursi", "chair", "lampu", "lamp",
        "tempat sampah", "trash bin", "gantungan baju", "hanger", "vas bunga",
        "hiasan dinding", "wall decor", "storage box", "kotak penyimpanan",
        "mattress", "led light", "essential oil", "clothes rack", "mop",
        "sofa cover", "spice jar", "dehumidifier", "diffuser",
    ],
    "Toys Kids & Babies": [
        "popok", "diaper", "susu formula", "baby food", "mainan bayi",
        "baby toy", "stroller", "car seat", "botol susu", "baby bottle",
        "mpasi", "breast pump", "pompa asi",
        "baby carrier", "baby balm", "mamil gold", "aptamil", "dumex", "enfagrow",
        "growing up milk",
        "mainan", "toy", "action figure", "puzzle", "lego", "boneka", "doll", "popmart",
    ],
    "Kids Fashion": [
        "baju bayi", "baby clothes",
    ],
    "Hobbies & Books": [
        "buku", "book", "novel", "komik", "comic", "alat lukis", "painting kit",
        "kartu remi", "playing card", "board game", "display box",
    ],
    "Sports & Outdoors": [
        "sepatu lari", "running shoes", "matras yoga", "yoga mat", "dumbbell",
        "resistance band", "sepeda", "bicycle", "tenda", "tent", "ransel hiking",
        "hiking bag", "raket", "racket", "bola", "ball sport", "treadmill",
        "gym", "protein shaker", "lacrosse ball", "massage",
    ],
    "Automotive": [
        "oli mobil", "engine oil", "aki mobil", "car battery", "ban motor",
        "tire", "helm", "helmet", "sarung tangan motor", "riding gloves",
        "car freshener", "parfum mobil", "dashcam", "car charger",
    ],
    "Food & Beverages": [
        "beras", "rice", "minyak goreng", "cooking oil", "gula", "sugar",
        "kopi", "coffee", "teh", "tea", "snack", "biskuit", "biscuit",
        "mie instan", "instant noodle", "saus", "sauce", "bumbu", "seasoning",
        "chocolate", "coklat", "susu uht", "air mineral", "mineral water",
        "freeze dried", "beer", "frozen", "grass fed beef",
    ],
    "Pet Supplies": [
        "makanan kucing", "cat food", "makanan anjing", "dog food", "pasir kucing",
        "cat litter", "kandang hewan", "pet cage", "aquarium", "akuarium",
        "pet bed", "pet donut",
    ],
    "Dining Travel & Services": [
        "voucher", "e-ticket", "tiket", "pulsa", "top up game", "top up",
        "gift card", "subscription",
    ],
    "Video Games": [
        "ps5", "ps4", "playstation", "xbox", "nintendo switch", "steam wallet",
        "gaming chair", "gaming mouse", "controller game",
    ],
    "Watches": [
        "jam tangan", "wristwatch", "smartwatch", "apple watch", "fashion trendy watch",
    ],
    "Jewellery & Accessories": [
        "kalung", "necklace", "gelang", "bracelet", "anting", "earring",
        "cincin", "ring perhiasan", "brooch", "bros",
    ],
    "Travel & Luggage": [
        "koper", "luggage",
    ],
    "Women's Bags": [
        "tas ransel", "backpack", "tote bag", "sling bag",
        "dompet", "wallet", "tas selempang", "lunch bag", "insulated lunch",
        "mesh zip bag",
    ],
    "Women's Shoes": [
        "sepatu wanita", "women shoes", "heels", "sandal wanita", "wedges",
    ],
    "Men's Shoes": [
        "sepatu pria", "men shoes", "sneakers pria", "sandal pria",
    ],
    "Women's Apparel": [
        "dress wanita", "blouse", "rok", "skirt", "baju wanita", "kemeja wanita",
        "celana wanita", "dress", "pyjamas", "padded top", "pajamas",
        "hijab", "jilbab", "gamis", "abaya", "mukena", "kerudung",
    ],
    "Men's Wear": [
        "kemeja pria", "kaos pria", "celana pria", "jaket pria", "baju pria",
        "hoodie men", "sweatshirt hoodie", "koko", "sarung",
    ],
}


def classify_by_keyword(nama_produk: str) -> str | None:
    """
    Cek apakah nama produk mengandung keyword yang dikenal.
    Return nama kategori kalau ketemu, None kalau tidak ada yang cocok.
    """
    nama_lower = nama_produk.lower()
    for kategori, keywords in KEYWORD_MAP.items():
        for kw in keywords:
            if kw in nama_lower:
                return kategori
    return None

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


def get_sheet_row_count(sheets_service) -> int:
    """Ambil jumlah baris aktual di sheet Tracker (dinamis, bukan hardcode)."""
    sheet_meta = sheets_service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID,
        ranges=[f"{SHEET_NAME}!A:A"],
        fields="sheets.properties.gridProperties.rowCount",
    ).execute()
    return sheet_meta["sheets"][0]["properties"]["gridProperties"]["rowCount"]


def get_pending_rows(sheets_service):
    """
    Ambil baris dari Tracker yang:
    - Kolom Nama Produk (PRODUCT_COL) tidak kosong
    - Kolom Kategori (CATEGORY_COL) masih kosong ATAU isinya bukan salah satu
      dari 26 kategori resmi (termasuk nilai lama/peninggalan proses
      klasifikasi jaman dulu yang mungkin masih berbahasa Indonesia atau
      pakai nama kategori versi lama)
    """
    total_rows = get_sheet_row_count(sheets_service)

    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A2:{READ_RANGE_LAST_COL_LETTER}{total_rows}",
    ).execute()
    rows = result.get("values", [])

    pending = []
    for i, row in enumerate(rows):
        row = row + [""] * (MAX_COL_IDX + 1 - len(row))
        nama_produk = row[PRODUCT_COL_IDX].strip()
        kategori = row[CATEGORY_COL_IDX].strip()

        if nama_produk and is_pending_category(kategori):
            pending.append((i + 2, nama_produk))

    print(f"[GSHEET] Sheet={SHEET_NAME}, Produk={PRODUCT_COL_LETTER}, Kategori={CATEGORY_COL_LETTER}")
    print(f"[GSHEET] Total baris di sheet: {total_rows}, {len(pending)} baris pending ditemukan")
    return pending


def update_kategori_batch(sheets_service, updates: list):
    if not updates:
        return
    data = [
        {"range": f"{SHEET_NAME}!{CATEGORY_COL_LETTER}{row_index}", "values": [[kategori]]}
        for row_index, kategori in updates
    ]
    body = {"valueInputOption": "USER_ENTERED", "data": data}
    sheets_service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body=body
    ).execute()
    print(f"[GSHEET] {len(updates)} kategori ditulis (batch) ke kolom {CATEGORY_COL_LETTER}")

# ============================================================
# STAGE 2: GROQ AI FALLBACK (untuk yang tidak ketemu keyword)
# ============================================================

def classify_products_via_ai(products: list) -> dict:
    """
    products: list of (row_index, nama_produk)
    Returns: dict {row_index: kategori}
    """
    if not GROQ_API_KEY:
        print("[WARN] GROQ_API_KEY tidak ada, semua sisa produk di-fallback ke 'Miscellaneous'")
        return {row_index: "Miscellaneous" for row_index, _ in products}

    numbered_list = "\n".join(
        f"{i+1}. {nama}" for i, (_, nama) in enumerate(products)
    )

    prompt = f"""Kamu adalah classifier kategori produk e-commerce Shopee Singapore.

Daftar kategori yang BOLEH dipakai (pilih SATU per produk, HARUS persis salah satu dari daftar ini):
{CATEGORIES_STR}

Berikut daftar nama produk yang perlu diklasifikasikan:
{numbered_list}

Balas HANYA dalam format JSON array, urutan harus sama persis dengan nomor di atas.
Contoh format balasan:
["Home Appliances", "Women's Apparel", "Beauty & Personal Care"]

Jangan tambahkan teks lain selain JSON array tersebut."""

    payload = {
        "model": GROQ_MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.1,
        "response_format": {"type": "json_object"},
    }
    headers = {
        "Authorization": f"Bearer {GROQ_API_KEY}",
        "Content-Type": "application/json",
    }

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.post(GROQ_URL, json=payload, headers=headers, timeout=60)

            if resp.status_code == 429:
                wait_time = RETRY_BASE_DELAY * attempt
                print(f"[RATE LIMIT] 429 diterima, percobaan {attempt}/{MAX_RETRIES}, tunggu {wait_time}s...")
                time.sleep(wait_time)
                continue

            resp.raise_for_status()
            data = resp.json()
            text = data["choices"][0]["message"]["content"]

            # Groq/Llama kadang bungkus array dalam object -> handle dua kemungkinan
            parsed = json.loads(text)
            if isinstance(parsed, dict):
                # cari key pertama yang isinya list
                categories = next((v for v in parsed.values() if isinstance(v, list)), [])
            else:
                categories = parsed

            result = {}
            for i, (row_index, nama) in enumerate(products):
                if i < len(categories):
                    kategori = categories[i]
                    if kategori not in SHOPEE_CATEGORIES:
                        kategori = "Miscellaneous"
                    result[row_index] = kategori
                else:
                    result[row_index] = "Miscellaneous"
            return result

        except Exception as e:
            print(f"[ERROR] Groq AI gagal (percobaan {attempt}/{MAX_RETRIES}): {e}")
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_BASE_DELAY)

    print(f"[FAILED] Gagal setelah {MAX_RETRIES} percobaan, fallback semua ke 'Miscellaneous'")
    return {row_index: "Miscellaneous" for row_index, _ in products}

# ============================================================
# MAIN PIPELINE
# ============================================================

def run_classifier(max_rows: int = 2000):
    print(f"\n[START] Shopee Category Classifier (Hybrid) — {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    sheets_service = get_sheets_service()
    pending_rows = get_pending_rows(sheets_service)

    if not pending_rows:
        print("[INFO] Tidak ada baris pending.")
        return

    to_process = pending_rows[:max_rows]

    # === STAGE 1: keyword matching ===
    keyword_results = []
    needs_ai = []

    for row_index, nama in to_process:
        kategori = classify_by_keyword(nama)
        if kategori:
            keyword_results.append((row_index, kategori))
        else:
            needs_ai.append((row_index, nama))

    print(f"[STAGE 1] Keyword matching: {len(keyword_results)} produk terklasifikasi instant")
    print(f"[STAGE 1] Sisa butuh AI: {len(needs_ai)} produk")

    # Tulis hasil keyword matching langsung (batch besar, gratis, cepat)
    for i in range(0, len(keyword_results), 100):
        chunk = keyword_results[i:i + 100]
        update_kategori_batch(sheets_service, chunk)

    # === STAGE 2: AI fallback buat sisanya ===
    ai_processed = 0
    for batch_start in range(0, len(needs_ai), AI_BATCH_SIZE):
        batch = needs_ai[batch_start:batch_start + AI_BATCH_SIZE]
        print(f"\n[STAGE 2 - BATCH] Memproses {len(batch)} produk via AI (row {batch[0][0]}-{batch[-1][0]})...")

        results = classify_products_via_ai(batch)

        if results:
            updates = list(results.items())
            update_kategori_batch(sheets_service, updates)
            for row_index, kategori in results.items():
                nama = dict(batch)[row_index]
                print(f"[OK] Row {row_index}: '{nama[:40]}' -> {kategori}")
            ai_processed += len(results)

        if batch_start + AI_BATCH_SIZE < len(needs_ai):
            time.sleep(DELAY_BETWEEN_AI_CALLS)

    print(f"\n[DONE] Total diproses: {len(keyword_results)} (keyword) + {ai_processed} (AI) = {len(keyword_results) + ai_processed} dari {len(to_process)} baris.")


if __name__ == "__main__":
    run_classifier(max_rows=2000)
