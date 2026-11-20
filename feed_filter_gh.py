# ============================================================
# feed_filter_gh.py  (versi GitHub Actions)
# Shopee SG Product Feed -> sheet "Shopee Video Generator" (tab produk)
#
# Dijalankan oleh .github/workflows/feed_filter.yml.
# Repo ini PUBLIK: tidak ada link CSV, ID sheet, atau kredensial di file ini.
#   - SHEET_URL dan GOOGLE_SERVICE_ACCOUNT_JSON  -> GitHub Secrets
#   - link CSV, kategori, filter                  -> tab "setting" di sheet
# ============================================================

import json
import os
import re
import sys

import pandas as pd

# ---------- PENGATURAN DEFAULT (ditimpa oleh tab "setting" di sheet) ----------
_DEFAULT = {
    "CSV_URLS": [],          # diisi dari tab setting (baris CSV_URLS)
    "SHEET_URL": os.environ.get("SHEET_URL", ""),
    "TAB_NAME": "produk",
    "KATEGORI": ["women", "beauty", "personal care", "jewel", "health"],
    "KEYWORDS": [],
    "EXCLUDE": [],
    "MIN_RATING": 4.7,
    "MIN_SOLD": 500,
    "MIN_HARGA": 3,
    "MAX_HARGA": 60,
    "MIN_STOCK": 10,
    "JUMLAH_AMBIL": 50,
    "CHUNK": 200_000,
}

CFG = dict(_DEFAULT)
LIST_KEYS = {"CSV_URLS", "KATEGORI", "KEYWORDS", "EXCLUDE"}

KOLOM = ["itemid", "title", "sale_price", "price", "discount_percentage",
         "item_rating", "item_sold", "like", "stock", "shop_name",
         "global_category1", "global_category2", "global_category3",
         "image_link", "description", "product_link", "product_short link"]
ANGKA = ["sale_price", "price", "discount_percentage",
         "item_rating", "item_sold", "like", "stock"]
BUANG = ["gwp", "not for sale", "free gift", "sample"]


# ---------- 1. BACA & SARING FEED ----------
def saring_feed():
    hasil = []
    for n, url in enumerate(CFG["CSV_URLS"], 1):
        print(f"\nMembaca feed {n}/{len(CFG['CSV_URLS'])} ...")
        for i, ch in enumerate(pd.read_csv(url, usecols=KOLOM, chunksize=CFG["CHUNK"],
                                           dtype=str, on_bad_lines="skip")):
            for c in ANGKA:
                ch[c] = pd.to_numeric(ch[c], errors="coerce")

            f = ch[(ch["item_rating"] >= CFG["MIN_RATING"]) &
                   (ch["item_sold"] >= CFG["MIN_SOLD"]) &
                   (ch["sale_price"].between(CFG["MIN_HARGA"], CFG["MAX_HARGA"])) &
                   (ch["stock"] >= CFG["MIN_STOCK"])]

            if CFG["KATEGORI"]:
                kat = f["global_category1"].str.lower().fillna("")
                f = f[kat.str.contains("|".join(CFG["KATEGORI"]), regex=True)]
            if CFG["KEYWORDS"]:
                jd = f["title"].str.lower().fillna("")
                f = f[jd.str.contains("|".join(CFG["KEYWORDS"]), regex=True)]
            buang = BUANG + [x.lower() for x in CFG["EXCLUDE"]]
            jd = f["title"].str.lower().fillna("")
            f = f[~jd.str.contains("|".join(map(re.escape, buang)), regex=True)]

            hasil.append(f)
            print(f"  chunk {i + 1}: {len(f)} lolos")

    df = pd.concat(hasil, ignore_index=True).drop_duplicates("itemid")
    print("\nTotal lolos filter:", len(df))

    # Skor: laris + rating + like + diskon
    df["skor"] = (df["item_sold"].rank(pct=True) * 0.5 +
                  df["item_rating"].rank(pct=True) * 0.2 +
                  df["like"].rank(pct=True) * 0.2 +
                  df["discount_percentage"].fillna(0).rank(pct=True) * 0.1)
    return df.sort_values("skor", ascending=False)


# ---------- 2. LOGIN GOOGLE SHEETS ----------
def buka_sheet():
    import gspread
    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not sa_json:
        sys.exit("GOOGLE_SERVICE_ACCOUNT_JSON belum diisi di GitHub Secrets.")
    if not CFG["SHEET_URL"]:
        sys.exit("SHEET_URL belum diisi di GitHub Secrets.")
    info = json.loads(sa_json)
    gc = gspread.service_account_from_dict(info)
    print("Login pakai service account:", info.get("client_email"))
    sh = gc.open_by_url(CFG["SHEET_URL"])
    baca_setting(sh)
    return sh.worksheet(CFG["TAB_NAME"])


def baca_setting(sh):
    """Timpa CFG dengan isi tab 'setting' (A = nama, B = nilai)."""
    try:
        rows = sh.worksheet("setting").get_all_values()
    except Exception:
        print("Tab 'setting' tidak ada, pakai pengaturan default.")
        return
    for row in rows:
        if len(row) < 2 or row[0].strip() not in CFG:
            continue
        key, val = row[0].strip(), row[1].strip()
        if key in LIST_KEYS:
            CFG[key] = [x.strip() for x in val.replace("\n", ",").split(",") if x.strip()]
        elif key in ("SHEET_URL", "TAB_NAME"):
            CFG[key] = val or CFG[key]
        elif val:
            val = val.replace(",", ".")          # sheet lokal Indonesia pakai koma desimal
            CFG[key] = float(val) if "." in val else int(val)
    print("Pengaturan:", {k: v for k, v in CFG.items() if k not in ("CSV_URLS", "SHEET_URL")})


# ---------- 3. TULIS KE SHEET ----------
def tulis(ws, df):
    # Kolom tambahan di kanan (L-R) supaya kolom A-K tidak bergeser
    ws.update(values=[["Item ID", "Kategori", "Terjual", "Rating", "Diskon %",
                       "Toko", "Link Asli"]], range_name="L1:R1")

    sudah = set(ws.col_values(12)[1:])                  # Item ID yang sudah ada
    df = df[~df["itemid"].isin(sudah)].head(CFG["JUMLAH_AMBIL"])

    baris_awal = len(ws.col_values(2)) + 1              # baris kosong pertama (kolom B)
    rows = []
    for _, r in df.iterrows():
        desk = r["description"] if isinstance(r["description"], str) else ""
        r = r.where(pd.notna(r), "")
        kategori = " > ".join(x for x in [r["global_category1"], r["global_category2"],
                                          r["global_category3"]] if x)
        rows.append([
            r["product_short link"], r["title"], r["sale_price"], desk[:500],
            r["image_link"], "", "", "", "", "belum",                       # B-K
            r["itemid"], kategori, int(r["item_sold"]), r["item_rating"],
            r["discount_percentage"], r["shop_name"], r["product_link"],   # L-R
        ])

    if rows:
        ws.update(values=rows, range_name=f"B{baris_awal}",
                  value_input_option="USER_ENTERED")
    print(f"\n{len(rows)} produk baru ditulis mulai baris {baris_awal}")


ws = buka_sheet()          # login dulu, supaya error akses ketahuan sebelum proses lama
if not CFG["CSV_URLS"]:
    sys.exit("CSV_URLS kosong: isi link CSV Product Feed di tab setting.")
df = saring_feed()
tulis(ws, df)
print("Selesai.")
