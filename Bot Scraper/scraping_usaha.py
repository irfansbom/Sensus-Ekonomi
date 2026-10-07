#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
scraping_usaha.py

Verifikasi usaha berdasarkan:
1. Google Maps       -> sinyal awal + verifikasi nama/alamat
2. Website           -> discovery + aktivitas <= 90 hari
3. JobStreet         -> sinyal rekrutmen
4. Glints            -> sinyal rekrutmen
5. REKAP (tahap akhir) -> 2 file Excel: rekap_UBUM.xlsx (flag_usaha UM/UB)
                          dan rekap_UMK.xlsx (flag_usaha UMK)

Input CSV minimal:
    nama_perusahaan,alamat

Kolom opsional:
    nama_kab, flag_usaha (wajib bila ingin rekap terpisah UBUM/UMK)

Contoh:
    nama_perusahaan,alamat
    "LENKO SURYA PERKASA,","JL RAYA CIHAURBEUTI DSN DESA"

Run (scraping + rekap otomatis di akhir):
    python scraping_usaha.py input.csv hasil.csv

Atur nama file rekap:
    python scraping_usaha.py input.csv hasil.csv \
        --rekap-ubum rekap_UBUM.xlsx --rekap-umk rekap_UMK.xlsx

Tanpa rekap:
    python scraping_usaha.py input.csv hasil.csv --no-rekap

Rekap saja dari CSV hasil yang sudah ada (tanpa scraping):
    python scraping_usaha.py --rekap-only hasil.csv

Resume:
    Jalankan command yang sama lagi.

Mulai dari nol:
    python scraping_usaha.py input.csv hasil.csv --mulai-ulang

ATURAN summary_category pada REKAP:
  GATE : maps_ditemukan = TRUE, kemiripan nama > 0.80, kemiripan alamat > 0.80
  recency  : 'kurang' bila ulasan Maps terakhir < 90 hari, selain itu 'lebih'
  jobstreet: lowongan_cocok_lokasi | lowongan_lokasi_lain | tanpa_lowongan
  -> summary_category = maps_{recency}_{jobstreet}; tidak lolos GATE -> kosong.
"""

import argparse
import asyncio
import csv
import json
import os
import random
import re
import shutil
import unicodedata
from datetime import datetime, timedelta
from urllib.parse import quote_plus, urlparse, parse_qs, unquote

from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font


# ============================================================
# CONFIG
# ============================================================

ACTIVITY_DAYS = 90
PAGE_TIMEOUT = 25
WAIT_TIMEOUT = 10

MAPS_DELAY = (2.5, 5.0)
SEARCH_DELAY = (1.5, 3.5)
SITE_DELAY = (1.0, 2.5)
JOB_DELAY = (2.0, 4.0)

# Sumber yang memblokir akses tidak dicoba berulang kali dalam satu run.
BLOCKED_JOB_SITES = set()

# Konkurensi (jumlah perusahaan yang diproses paralel dalam satu browser).
# Target Google (Maps/Search) agresif memblokir; nilai kecil lebih aman.
# Turunkan ke 1 bila sering kena CAPTCHA.
CONCURRENCY_DEFAULT = 3

# Retry per perusahaan dengan exponential backoff (2**attempt detik).
MAX_RETRIES_DEFAULT = 3

SOCIAL_HOSTS = {
    "instagram.com",
    "facebook.com",
    "linkedin.com",
    "tiktok.com",
    "youtube.com",
    "x.com",
    "twitter.com",
}

SEARCH_HOSTS = {
    "google.com",
    "google.co.id",
    "bing.com",
    "search.yahoo.com",
}

IGNORE_HOSTS = SEARCH_HOSTS | {
    "jobstreet.co.id",
    "id.jobstreet.com",
    "glints.com",
    "id.glints.com",
    "wikipedia.org",
}


# ============================================================
# TEXT NORMALIZATION
# ============================================================

NAME_STOPWORDS = {
    "pt", "cv", "tb", "ud", "pd", "persero", "toko",
    "store", "shop", "official", "indonesia"
}

ADDRESS_STOPWORDS = {
    "jl", "jln", "jalan", "no", "nomor", "rt", "rw",
    "kel", "kelurahan", "kec", "kecamatan",
    "kab", "kabupaten", "kota", "prov", "provinsi",
    "indonesia", "blok", "bl", "lantai", "lt",
    "gedung", "ruko", "komplek", "kompleks",
    "desa", "dsn", "dusun"
}


def norm(s):
    if s is None:
        return ""

    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower()
    s = s.replace("&", " dan ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def compact(s):
    return re.sub(r"[^a-z0-9]", "", norm(s))


def tokens(s, stopwords=None):
    stopwords = stopwords or set()
    return {
        x for x in norm(s).split()
        if len(x) >= 2 and x not in stopwords
    }


def name_similarity(a, b):
    A = tokens(a, NAME_STOPWORDS)
    B = tokens(b, NAME_STOPWORDS)

    if not A or not B:
        return 0.0

    return len(A & B) / max(len(A), len(B))


def address_similarity(a, b):
    A = tokens(a, ADDRESS_STOPWORDS)
    B = tokens(b, ADDRESS_STOPWORDS)

    if not A or not B:
        return 0.0

    score = len(A & B) / max(len(A), len(B))

    na = set(re.findall(r"\b\d+[a-z]?\b", norm(a)))
    nb = set(re.findall(r"\b\d+[a-z]?\b", norm(b)))

    if na and nb:
        if na & nb:
            score += 0.15
        else:
            score -= 0.20

    return max(0.0, min(1.0, score))


def match_name(input_name, found_name):
    a = compact(input_name)
    b = compact(found_name)

    if not a or not b:
        return "UNCERTAIN", 0.0

    if a == b:
        return "MATCH", 1.0

    sim = name_similarity(input_name, found_name)

    if a in b or b in a:
        return "MATCH", max(sim, 0.85)

    if sim >= 0.75:
        return "MATCH", sim

    if sim >= 0.45:
        return "UNCERTAIN", sim

    return "NOT_MATCH", sim


def match_address(input_address, found_address, input_kab=""):
    if not found_address:
        return "UNCERTAIN", address_similarity(input_address, found_address)

    a = compact(input_address)
    b = compact(found_address)
    kab_tokens = tokens(input_kab, ADDRESS_STOPWORDS)
    found_tokens = tokens(found_address, ADDRESS_STOPWORDS)

    if kab_tokens and kab_tokens & found_tokens:
        return "MATCH", max(address_similarity(input_address, found_address), 0.85)

    if not input_address:
        return "UNCERTAIN", 0.0

    if a == b or a in b or b in a:
        return "MATCH", 1.0

    sim = address_similarity(input_address, found_address)

    if sim >= 0.78:
        return "MATCH", sim

    if sim <= 0.20:
        return "NOT_MATCH", sim

    return "UNCERTAIN", sim


# ============================================================
# DATES
# ============================================================

def relative_days(text):
    if not text:
        return None

    t = norm(text)

    if "baru saja" in t or "just now" in t:
        return 0.0

    if "kemarin" in t:
        return 1.0

    if "seminggu" in t:
        return 7.0

    if "sebulan" in t:
        return 30.0

    if "setahun" in t:
        return 365.0

    patterns = [
        (r"(\d+)\s*menit", 1 / 1440),
        (r"(\d+)\s*jam", 1 / 24),
        (r"(\d+)\s*hari", 1),
        (r"(\d+)\s*minggu", 7),
        (r"(\d+)\s*bulan", 30),
        (r"(\d+)\s*tahun", 365),
        (r"(\d+)\s*minute", 1 / 1440),
        (r"(\d+)\s*hour", 1 / 24),
        (r"(\d+)\s*day", 1),
        (r"(\d+)\s*week", 7),
        (r"(\d+)\s*month", 30),
        (r"(\d+)\s*year", 365),
    ]

    for pattern, multiplier in patterns:
        m = re.search(pattern, t)
        if m:
            return int(m.group(1)) * multiplier

    return None


def date_text(days):
    if days is None:
        return ""

    d = datetime.now() - timedelta(days=float(days))
    return d.strftime("%Y-%m-%d")


def find_recent_timestamp(text):
    if not text:
        return None, ""

    patterns = [
        r"baru saja",
        r"kemarin",
        r"sehari",
        r"seminggu",
        r"sebulan",
        r"setahun",
        r"\d+\s*(?:menit|jam|hari|minggu|bulan|tahun)\s*(?:yang\s*)?lalu",
        r"\d+\s*(?:minute|hour|day|week|month|year)s?\s*ago",
    ]

    for p in patterns:
        m = re.search(p, text, re.I)
        if m:
            value = relative_days(m.group(0))
            if value is not None:
                return value, m.group(0)

    return None, ""


# ============================================================
# DRIVER (Playwright async_api)
# ============================================================

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

class By:
    """Pengganti selenium.webdriver.common.by.By untuk API find_elements."""
    CSS_SELECTOR = "css"
    XPATH = "xpath"


class _Element:
    """Bungkus Playwright ElementHandle agar punya .click() seperti Selenium.

    Metode I/O bersifat async (await) mengikuti playwright.async_api.
    """

    def __init__(self, handle):
        self._handle = handle

    @property
    def handle(self):
        return self._handle

    async def click(self):
        await self._handle.click(timeout=WAIT_TIMEOUT * 1000)

    async def get_text(self, *args, **kwargs):
        try:
            return await self._handle.inner_text()
        except Exception:
            return ""


class BrowserPool:
    """Kelola SATU browser + context Playwright async, dipakai ulang lintas page.

    Dibuat sekali lewat `await BrowserPool.create(...)`. Tiap perusahaan
    mengambil page baru lewat `await pool.new_session()` lalu menutupnya.
    Tidak ada auto-install: bila Playwright/browser belum siap, create()
    berhenti dengan instruksi manual.
    """

    def __init__(self, pw, browser, context):
        self._pw = pw
        self._browser = browser
        self._context = context

    @classmethod
    async def create(cls, headless=True, channel=None):
        try:
            from playwright.async_api import async_playwright
        except ImportError as e:
            raise SystemExit(
                "Playwright belum terpasang. Pasang SEKALI secara manual:\n"
                "    pip install -U playwright beautifulsoup4 openpyxl\n"
                "    playwright install chromium   "
                "# atau pakai --channel chrome/msedge\n"
                f"Detail error: {e}"
            )

        pw = await async_playwright().start()

        launch_kwargs = {
            "headless": headless,
            "args": [
                "--disable-gpu",
                "--disable-dev-shm-usage",
                "--no-sandbox",
                "--window-size=1440,1000",
                "--lang=id-ID",
            ],
        }

        # Pemilihan browser (TANPA mengunduh apa pun):
        #   1. CHROME_BIN -> executable Chrome/Chromium sistem (path manual)
        #   2. channel    -> Chrome/Edge terpasang ("chrome", "msedge", dst.)
        #   3. default    -> Chromium bawaan Playwright (perlu
        #                    `playwright install chromium` sekali)
        chrome_bin = os.environ.get("CHROME_BIN")
        channel = channel or os.environ.get("CHROME_CHANNEL")
        if chrome_bin:
            launch_kwargs["executable_path"] = chrome_bin
        elif channel:
            launch_kwargs["channel"] = channel

        try:
            browser = await pw.chromium.launch(**launch_kwargs)
        except Exception as e:
            try:
                await pw.stop()
            except Exception:
                pass
            raise SystemExit(
                "Gagal meluncurkan browser (tidak ada auto-install). "
                "Pastikan salah satu berikut:\n"
                "  - jalankan `playwright install chromium` sekali, ATAU\n"
                "  - pakai --channel chrome / --channel msedge "
                "(browser terpasang), ATAU\n"
                "  - set CHROME_BIN ke path executable Chrome/Chromium.\n"
                f"Detail error: {e}"
            )

        context = await browser.new_context(
            locale="id-ID",
            viewport={"width": 1440, "height": 1000},
            user_agent=USER_AGENT,
        )
        return cls(pw, browser, context)

    async def new_session(self):
        """Buka page baru dan bungkus jadi PageSession (satu per perusahaan)."""
        page = await self._context.new_page()
        return PageSession(page)

    async def quit(self):
        for closer in (
            getattr(self, "_context", None),
            getattr(self, "_browser", None),
        ):
            try:
                if closer is not None:
                    await closer.close()
            except Exception:
                pass
        try:
            if self._pw is not None:
                await self._pw.stop()
        except Exception:
            pass


class PageSession:

    def __init__(self, page):
        self._page = page
        self._timeout_ms = PAGE_TIMEOUT * 1000
        self._page.set_default_navigation_timeout(self._timeout_ms)
        self._page.set_default_timeout(WAIT_TIMEOUT * 1000)

    @property
    def page(self):
        return self._page

    def set_page_load_timeout(self, seconds):
        self._timeout_ms = int(float(seconds) * 1000)
        self._page.set_default_navigation_timeout(self._timeout_ms)

    async def get(self, url):
        # Raise saat timeout/gagal navigasi, seperti Selenium .get().
        await self._page.goto(
            url,
            timeout=self._timeout_ms,
            wait_until="domcontentloaded",
        )

    async def content(self):
        try:
            return await self._page.content()
        except Exception:
            return ""

    async def wait_for_css(self, selector, timeout=None):
        ms = int((timeout if timeout is not None else WAIT_TIMEOUT) * 1000)
        await self._page.wait_for_selector(selector, timeout=ms, state="attached")

    async def find_elements(self, by, selector):
        if by == By.XPATH:
            selector = "xpath=" + selector
        try:
            handles = await self._page.query_selector_all(selector)
        except Exception:
            return []
        return [_Element(h) for h in handles]

    async def execute_script(self, script, *args):
        # Emulasi execute_script Selenium untuk skrip ber-argumen elemen,
        # mis. execute_script("arguments[0].click();", el).
        js = "(args) => { " + script.replace("arguments", "args") + " }"
        pw_args = [
            a.handle if isinstance(a, _Element) else a
            for a in args
        ]
        return await self._page.evaluate(js, pw_args)

    async def close(self):
        try:
            await self._page.close()
        except Exception:
            pass


async def create_browser(headless=True, channel=None):
    return await BrowserPool.create(headless=headless, channel=channel)


async def open_url(driver, url, wait_selector=None):
    try:
        await driver.get(url)

        if wait_selector:
            try:
                await driver.wait_for_css(wait_selector, WAIT_TIMEOUT)
            except Exception:
                pass

        return True

    except Exception:
        return False


async def soup(driver):
    return BeautifulSoup(await driver.content(), "html.parser")


async def text_page(driver):
    try:
        return (await soup(driver)).get_text(" ", strip=True)
    except Exception:
        return ""


async def blocked(driver):
    text = (await driver.content()).lower()

    markers = [
        "unusual traffic",
        "captcha",
        "recaptcha",
        "verify you are human",
        "access denied",
        "robot check",
        "403 forbidden",
        "error 403",
        "just a moment...",
        "cloudflare"
    ]

    return any(x in text for x in markers)


# ============================================================
# GOOGLE MAPS
# ============================================================

def _status_from_snippet(text):
    """Tentukan status dari potongan teks kecil (mis. widget jam operasional)."""
    raw = (text or "").lower()
    if not raw:
        return ""

    if "tutup permanen" in raw or "permanently closed" in raw:
        return "TUTUP_PERMANEN"
    if "tutup sementara" in raw or "temporarily closed" in raw:
        return "TUTUP_SEMENTARA"
    if "buka 24 jam" in raw or "open 24 hours" in raw:
        return "BUKA"

    # Pola pil status Maps: "<status> . <jam berikutnya>", mis.
    # "Buka . Tutup pukul 17.00" (buka sekarang) atau
    # "Tutup . Buka pukul 08.00" (tutup sekarang). Titik pemisah bisa
    # berupa U+00B7 atau U+22C5.
    for pattern, value in [
        (r"buka\s*[·⋅]", "BUKA"),
        (r"open\s*[·⋅]", "BUKA"),
        (r"tutup\s*[·⋅]", "TUTUP"),
        (r"closed\s*[·⋅]", "TUTUP"),
        (r"\bbuka sekarang\b", "BUKA"),
        (r"\bopen now\b", "BUKA"),
        (r"\bclosed now\b", "TUTUP"),
    ]:
        if re.search(pattern, raw):
            return value

    return ""


def detect_business_status(soup_obj, full_text):
    """Deteksi status operasional dari halaman Google Maps.

    Nilai yang dikembalikan:
        TUTUP_PERMANEN   -> tutup permanen
        TUTUP_SEMENTARA  -> tutup sementara
        BUKA             -> sedang buka
        TUTUP            -> sedang tutup (di luar jam), tetapi masih beroperasi
        ""               -> tidak terbaca
    """
    raw = (full_text or "").lower()

    # Frasa penutupan bersifat unik sehingga aman dicek di seluruh teks halaman.
    if "tutup permanen" in raw or "permanently closed" in raw:
        return "TUTUP_PERMANEN"
    if "tutup sementara" in raw or "temporarily closed" in raw:
        return "TUTUP_SEMENTARA"

    # Status buka/tutup sekarang: batasi ke elemen jam operasional supaya tidak
    # tertukar dengan kata 'buka'/'tutup' yang bertebaran di ulasan.
    selectors = [
        "span.ZDu9vd",
        "div.o0Svhf",
        "div.OqCZI",
        "[aria-label*='Buka']",
        "[aria-label*='Tutup']",
        "[aria-label*='Open']",
        "[aria-label*='Closed']",
        "[jsaction*='openhours']",
        "[data-item-id*='oh']",
    ]
    for selector in selectors:
        try:
            elements = soup_obj.select(selector)
        except Exception:
            continue
        for el in elements:
            for candidate in (el.get("aria-label"), el.get_text(" ", strip=True)):
                status = _status_from_snippet(candidate)
                if status:
                    return status

    # Fallback: kata buka/tutup yang berdekatan dengan 'pukul'.
    m = re.search(r"\b(buka|tutup)\b[^.]{0,40}pukul", raw)
    if m:
        return "BUKA" if m.group(1) == "buka" else "TUTUP"

    return ""


async def maps_search(driver, name, address, kab=""):
    result = {
        "maps_found": False,
        "maps_name": "",
        "maps_address": "",
        "maps_category": "",
        "maps_phone": "",
        "maps_website": "",
        "maps_url": "",
        "maps_business_status": "",
        "maps_rating": "",
        "maps_reviews_count": "",
        "maps_name_match": "UNCERTAIN",
        "maps_name_similarity": 0.0,
        "maps_address_match": "UNCERTAIN",
        "maps_address_similarity": 0.0,
        "maps_review_latest_days": None,
        "maps_review_latest_text": "",
        "maps_note": "",
        "maps_blocked": False,
    }

    query = re.sub(r"\s+", " ", f"{name} {address} {kab}").strip()
    maps_url = "https://www.google.com/maps/search/" + quote_plus(query)

    result["maps_url"] = maps_url

    if not await open_url(driver, maps_url):
        result["maps_note"] = "Gagal membuka Google Maps"
        return result

    await asyncio.sleep(random.uniform(*MAPS_DELAY))

    if await blocked(driver):
        result["maps_blocked"] = True
        result["maps_note"] = "Google Maps meminta verifikasi/CAPTCHA"
        return result

    # Kandidat pertama.
    try:
        links = await driver.find_elements(By.CSS_SELECTOR, "a.hfpxzc")
        if links:
            await links[0].click()
            await asyncio.sleep(random.uniform(2, 4))
    except Exception:
        pass

    await asyncio.sleep(1)

    s = await soup(driver)

    name_el = (
        s.select_one("h1.DUwDvf")
        or s.select_one("h1.fontHeadlineLarge")
        or s.select_one("div.fontHeadlineLarge")
        or s.select_one("h1")
    )

    if name_el:
        result["maps_name"] = name_el.get_text(" ", strip=True)
        result["maps_found"] = True

    # Address.
    for selector in [
        "[data-item-id='address']",
        "button[data-item-id='address']",
        "a[data-item-id='address']",
    ]:
        el = s.select_one(selector)
        if el:
            result["maps_address"] = el.get_text(" ", strip=True)
            break

    # Website.
    el = s.select_one("a[data-item-id='authority']")
    if el and el.get("href"):
        result["maps_website"] = el["href"]

    # Phone.
    el = s.select_one("[data-item-id^='phone:tel:']")
    if el:
        result["maps_phone"] = el.get_text(" ", strip=True)

    # Rating.
    el = s.select_one("div.F7nice span[aria-hidden='true']")
    if el:
        result["maps_rating"] = el.get_text(" ", strip=True)

    full_text = s.get_text(" ", strip=True)

    # Review count.
    m = re.search(r"([\d.,]+)\s*(?:ulasan|reviews)", full_text, re.I)
    if m:
        result["maps_reviews_count"] = m.group(1)

    # Status bisnis: BUKA / TUTUP (sekarang) / TUTUP_SEMENTARA / TUTUP_PERMANEN.
    result["maps_business_status"] = detect_business_status(s, full_text)

    # Category.
    try:
        el = s.select_one("button.DkEaL")
        if el:
            result["maps_category"] = el.get_text(" ", strip=True)
    except Exception:
        pass

    result["maps_name_match"], result["maps_name_similarity"] = match_name(
        name,
        result["maps_name"]
    )

    result["maps_address_match"], result["maps_address_similarity"] = match_address(
        address,
        result["maps_address"],
        kab,
    )

    if not result["maps_found"]:
        result["maps_note"] = "Nama usaha tidak terbaca"

    elif result["maps_name_match"] == "NOT_MATCH":
        result["maps_note"] = "Kandidat Maps kemungkinan bukan usaha yang dimaksud"

    elif result["maps_address_match"] == "NOT_MATCH":
        result["maps_note"] = "Usaha ditemukan tetapi alamat berbeda"

    elif result["maps_address_match"] == "UNCERTAIN":
        result["maps_note"] = "Usaha ditemukan tetapi alamat belum cukup kuat"

    return result


async def maps_latest_review(driver):
    """Buka tab Ulasan, urutkan 'Terbaru', ambil review teratas.

    Mengembalikan (hari, teks_waktu, isi_review). Best-effort: bila tab/sortir
    tidak tersedia, jatuh ke pembacaan waktu relatif apa adanya (isi kosong).
    """
    isi = ""

    try:
        buttons = await driver.find_elements(
            By.XPATH,
            "//button[contains(.,'Ulasan') or contains(.,'Reviews') "
            "or contains(@aria-label,'Ulasan') "
            "or contains(@aria-label,'Reviews')]"
        )

        if buttons:
            await driver.execute_script("arguments[0].click();", buttons[0])
            try:
                await driver.wait_for_css("div.jftiEf, div.jJc9Ad", WAIT_TIMEOUT)
            except Exception:
                pass
            await asyncio.sleep(random.uniform(1.5, 2.5))

            # Urutkan berdasarkan 'Terbaru' bila tombol sortir tersedia.
            try:
                sort_buttons = await driver.find_elements(
                    By.XPATH,
                    "//button[contains(.,'Urutkan') or contains(.,'Sort') "
                    "or contains(@aria-label,'Urutkan') "
                    "or contains(@aria-label,'Sort')]"
                )
                if sort_buttons:
                    await driver.execute_script(
                        "arguments[0].click();", sort_buttons[0]
                    )
                    await asyncio.sleep(1)
                    newest = await driver.find_elements(
                        By.XPATH,
                        "//div[@role='menuitemradio']"
                        "[contains(.,'Terbaru') or contains(.,'Newest')]"
                    )
                    if newest:
                        await driver.execute_script(
                            "arguments[0].click();", newest[0]
                        )
                        await asyncio.sleep(random.uniform(1.5, 2.5))
            except Exception:
                pass
    except Exception:
        pass

    s = await soup(driver)

    # Isi review teratas (class 'wiI7pd' umum dipakai untuk teks ulasan).
    top = s.select_one("div.jftiEf, div.jJc9Ad")
    if top:
        isi_el = top.select_one("span.wiI7pd")
        if isi_el:
            isi = isi_el.get_text(" ", strip=True)

    reviews = s.select("div.jftiEf, div.jJc9Ad")

    candidates = []
    for review in reviews[:10]:
        date_elements = review.select(
            "span.rsqaWe, span.xRkPPb, "
            "[aria-label*='lalu'], [aria-label*='ago']"
        )
        texts = [el.get_text(" ", strip=True) for el in date_elements]
        texts.extend(review.stripped_strings)

        for candidate_text in texts:
            review_days, review_text = find_recent_timestamp(candidate_text)
            if review_days is not None:
                candidates.append((review_days, review_text))

    if not candidates:
        return None, "", isi

    best_days, best_text = min(candidates, key=lambda item: item[0])
    return best_days, best_text, isi


# ============================================================
# GOOGLE SEARCH DISCOVERY
# ============================================================

async def search_google(driver, query, limit=8):
    url = "https://www.google.com/search?q=" + quote_plus(query)

    if not await open_url(driver, url):
        return [], "ERROR"

    await asyncio.sleep(random.uniform(*SEARCH_DELAY))

    if await blocked(driver):
        return [], "BLOCKED"

    s = await soup(driver)
    urls = []

    for a in s.select("a[href]"):
        raw = a.get("href", "")

        if raw.startswith("/url?"):
            qs = parse_qs(urlparse(raw).query)
            raw = qs.get("q", qs.get("url", [""]))[0]

        raw = unquote(raw)
        if not raw.startswith("http"):
            continue

        h = urlparse(raw).netloc.lower()

        if h.startswith("www."):
            h = h[4:]

        if h in IGNORE_HOSTS:
            continue

        if raw not in urls:
            urls.append(raw)

        if len(urls) >= limit:
            break

    return urls, "OK"


def social_type(url):
    h = urlparse(url).netloc.lower()

    if "instagram.com" in h:
        return "instagram"
    if "facebook.com" in h:
        return "facebook"
    if "linkedin.com" in h:
        return "linkedin"
    if "tiktok.com" in h:
        return "tiktok"
    if "youtube.com" in h:
        return "youtube"
    if "twitter.com" in h or h == "x.com":
        return "x"

    return ""


async def discover_website(driver, name, address):
    result = {
        "website": "",
        "search_status": "NO_RESULT",
        "search_notes": "",
    }

    queries = [
        f'"{name}" "{address}"',
        f'"{name}" website',
    ]

    candidates = []

    for q in queries:
        urls, status = await search_google(driver, q)

        if status == "BLOCKED":
            result["search_status"] = "BLOCKED"
            result["search_notes"] = "Google search terkena CAPTCHA/blokir"
            break

        candidates.extend(urls)
        await asyncio.sleep(random.uniform(1, 2))

    candidates = list(dict.fromkeys(candidates))

    websites = [
        u for u in candidates
        if not social_type(u)
    ]

    # Website validation.
    best = None
    best_score = -1

    for u in websites[:5]:
        if not await open_url(driver, u):
            continue

        await asyncio.sleep(random.uniform(*SITE_DELAY))

        s = await soup(driver)
        title = s.title.get_text(" ", strip=True) if s.title else ""

        meta = s.select_one("meta[name='description']")
        description = meta.get("content", "") if meta else ""

        visible = s.get_text(" ", strip=True)[:20000]

        combined = " ".join([title, description, visible])

        nm, ns = match_name(name, combined)
        am, ass = match_address(address, combined)

        score = ns * 0.65 + ass * 0.35

        if nm == "NOT_MATCH":
            score -= 0.4

        if score > best_score:
            best_score = score
            best = (u, nm, am, ns, ass)

    if best:
        result["website"] = best[0]
        result["website_name_match"] = best[1]
        result["website_address_match"] = best[2]
        result["website_name_similarity"] = best[3]
        result["website_address_similarity"] = best[4]

    if result["website"]:
        result["search_status"] = "OK"

    return result


# ============================================================
# WEBSITE ACTIVITY
# ============================================================

async def website_activity(driver, website):
    result = {
        "status": "UNCERTAIN",
        "days": None,
        "text": "",
        "url": "",
        "note": "",
    }

    if not website:
        result["note"] = "Tidak ada website"
        return result

    base = website.rstrip("/")

    pages = [
        website,
        base + "/news",
        base + "/blog",
        base + "/berita",
        base + "/artikel",
        base + "/event",
        base + "/events",
        base + "/kegiatan",
        base + "/updates",
        base + "/contact",
        base + "/kontak",
    ]

    best_days = None
    best_text = ""
    best_url = ""

    for url in pages:
        if not await open_url(driver, url):
            continue

        await asyncio.sleep(random.uniform(*SITE_DELAY))

        txt = await text_page(driver)

        days, stamp = find_recent_timestamp(txt)

        if days is not None:
            if best_days is None or days < best_days:
                best_days = days
                best_text = stamp
                best_url = url

        # ISO dates in HTML.
        for y, m, d in re.findall(
            r"\b(20\d{2})[-/](\d{1,2})[-/](\d{1,2})\b",
            await driver.content()
        ):
            try:
                dt = datetime(int(y), int(m), int(d))
                age = (datetime.now() - dt).total_seconds() / 86400

                if age >= -1 and (best_days is None or age < best_days):
                    best_days = max(0, age)
                    best_text = f"{y}-{m}-{d}"
                    best_url = url
            except Exception:
                pass

    result["days"] = best_days
    result["text"] = best_text
    result["url"] = best_url

    if best_days is not None and best_days <= ACTIVITY_DAYS:
        result["status"] = "ACTIVE"
        result["note"] = "Aktivitas website <= 90 hari"
    elif best_days is not None:
        result["note"] = "Tanggal website lama; bukan bukti usaha tutup"
    else:
        result["note"] = "Tidak ada timestamp publik yang terbaca"

    return result


# ============================================================
# JOBSTREET / GLINTS
# ============================================================

def location_match(input_address, found_location, input_kab=""):
    if not found_location:
        return None

    found_tokens = tokens(found_location, ADDRESS_STOPWORDS)
    if not found_tokens:
        return None

    address_tokens = tokens(input_address, ADDRESS_STOPWORDS)
    kab_tokens = tokens(input_kab, ADDRESS_STOPWORDS)

    if kab_tokens:
        return bool(kab_tokens & found_tokens)

    if address_tokens:
        return bool(address_tokens & found_tokens)

    return None


def company_match(input_name, found_company, card_text=""):
    """Require an explicit company-name match in a job card."""
    source = found_company or card_text
    if not input_name or not source:
        return False

    input_compact = compact(input_name)
    source_compact = compact(source)

    if input_compact and input_compact in source_compact:
        return True

    input_tokens = tokens(input_name, NAME_STOPWORDS)
    source_tokens = tokens(source, NAME_STOPWORDS)

    return bool(input_tokens) and input_tokens.issubset(source_tokens)


async def glints_search_fallback(driver, name, address, kab=""):
    result = {
        "found": False,
        "title": "",
        "company": "",
        "location": "",
        "age_days": None,
        "age_text": "",
        "location_match": None,
        "url": "",
        "note": "",
    }

    query = f'site:glints.com/id/ "{name}" lowongan'
    url = "https://www.google.com/search?q=" + quote_plus(query)
    result["url"] = url

    if not await open_url(driver, url):
        result["note"] = "Fallback Google untuk Glints gagal dibuka"
        return result

    await asyncio.sleep(random.uniform(*SEARCH_DELAY))

    if await blocked(driver):
        result["note"] = "Google juga meminta verifikasi saat fallback Glints"
        return result

    soup_page = await soup(driver)
    best = None

    for link in soup_page.select("a[href]"):
        href = link.get("href", "")
        if href.startswith("/url?"):
            href = parse_qs(urlparse(href).query).get("q", [""])[0]
        href = unquote(href)
        host = urlparse(href).netloc.lower().removeprefix("www.")

        if host not in {"glints.com", "id.glints.com"}:
            continue
        if "/id/" not in href or not any(
            part in href for part in ("jobs", "opportunities")
        ):
            continue

        card = link.find_parent(["div", "article", "li"])
        text = card.get_text(" ", strip=True) if card else link.get_text(" ", strip=True)
        if not text:
            continue

        if not company_match(name, name, text):
            continue

        days, stamp = find_recent_timestamp(text)
        location = text
        loc_match = location_match(address, location, kab)
        score = 100 if loc_match is True else 50 if loc_match is None else 0
        if days is not None:
            score += max(0, 50 - min(days, 50))

        candidate = (
            score,
            link.get_text(" ", strip=True),
            name,
            location,
            days,
            stamp,
            loc_match,
            href,
        )
        if best is None or candidate[0] > best[0]:
            best = candidate

    if best:
        (
            _,
            result["title"],
            result["company"],
            result["location"],
            result["age_days"],
            result["age_text"],
            result["location_match"],
            result["url"],
        ) = best
        result["found"] = True
        result["note"] = "Ditemukan dari indeks Google karena akses langsung Glints diblokir"
    else:
        result["note"] = "Tidak ada hasil Glints pada indeks Google"

    return result


async def scrape_job_site(driver, site, name, address, kab=""):
    result = {
        "found": False,
        "title": "",
        "company": "",
        "location": "",
        "age_days": None,
        "age_text": "",
        "location_match": None,
        "url": "",
        "note": "",
    }

    if site == "jobstreet":
        url = f"https://id.jobstreet.com/id/jobs?keywords={quote_plus(name)}"
    else:
        url = f"https://glints.com/id/find-jobs/loker-jobs?keyword={quote_plus(name)}"

    result["url"] = url

    if site in BLOCKED_JOB_SITES:
        if site == "glints":
            return await glints_search_fallback(driver, name, address, kab)
        result["note"] = f"{site} dilewati: akses sudah diblokir pada run ini"
        return result

    if not await open_url(driver, url):
        result["note"] = f"Gagal membuka {site}"
        return result

    await asyncio.sleep(random.uniform(*JOB_DELAY))

    if await blocked(driver):
        BLOCKED_JOB_SITES.add(site)
        if site == "glints":
            fallback = await glints_search_fallback(driver, name, address, kab)
            if fallback["found"]:
                fallback["note"] += "; halaman langsung Glints terblokir"
            return fallback
        result["note"] = f"{site} memblokir akses (403/CAPTCHA/Cloudflare)"
        return result

    s = await soup(driver)

    if site == "jobstreet":
        cards = s.select(
            "article[data-automation='normalJob'], "
            "article[data-testid='job-card'], article"
        )
    else:
        cards = s.select(
            "div[class*='JobCard'], "
            "a[class*='opportunity-card'], "
            "div[class*='job-card']"
        )

    best = None

    for card in cards[:15]:
        text = card.get_text(" ", strip=True)

        if not text:
            continue

        title_el = (
            card.select_one("[data-automation='jobTitle']")
            or card.select_one("h2")
            or card.select_one("h3")
            or card.select_one("[class*='JobTitle']")
            or card.select_one("[class*='job-title']")
        )

        company_el = (
            card.select_one("[data-automation='jobCompany']")
            or card.select_one("[data-testid='company-name']")
            or card.select_one("[class*='CompanyName']")
            or card.select_one("[class*='company-name']")
            or card.select_one("[class*='companyName']")
        )

        location_el = (
            card.select_one("[data-automation='jobLocation']")
            or card.select_one("[class*='location']")
            or card.select_one("[class*='Location']")
        )

        title = title_el.get_text(" ", strip=True) if title_el else ""
        company = (
            company_el.get_text(" ", strip=True)
            if company_el else ""
        )
        location = (
            location_el.get_text(" ", strip=True)
            if location_el else ""
        )

        if not company_match(name, company, text):
            continue

        days, stamp = find_recent_timestamp(text)
        loc_match = location_match(address, location, kab)

        # Prioritaskan lokasi cocok lalu timestamp.
        score = (
            100 if loc_match is True else
            50 if loc_match is None else 0
        )

        if days is not None:
            score += max(0, 50 - min(days, 50))

        if best is None or score > best[0]:
            best = (
                score,
                title,
                company,
                location,
                days,
                stamp,
                loc_match,
            )

    if best:
        (
            _,
            result["title"],
            result["company"],
            result["location"],
            result["age_days"],
            result["age_text"],
            result["location_match"],
        ) = best

        result["found"] = True

        if result["location_match"] is False:
            result["note"] = "Lowongan ditemukan tetapi lokasi tidak cocok"
        elif result["location_match"] is True:
            result["note"] = "Lowongan ditemukan dengan nama perusahaan dan lokasi cocok"
        else:
            result["note"] = "Nama perusahaan cocok; lokasi lowongan tidak cukup jelas"

    else:
        result["note"] = "Tidak ada lowongan dengan nama perusahaan yang cocok"

    return result


# ============================================================
# FINAL STATUS
# ============================================================

def final_status(r):
    maps_name = r.get("maps_name_match", "UNCERTAIN")
    maps_addr = r.get("maps_address_match", "UNCERTAIN")
    maps_addr_similarity = r.get("maps_address_similarity")

    if maps_addr in (None, "", "UNCERTAIN") and r.get("maps_address"):
        maps_addr, maps_addr_similarity = match_address(
            r.get("alamat", ""),
            r.get("maps_address", ""),
            r.get("nmkab", ""),
        )
        r["maps_address_match"] = maps_addr
        r["maps_address_similarity"] = maps_addr_similarity

    web_name = r.get("website_name_match", "UNCERTAIN")
    web_addr = r.get("website_address_match", "UNCERTAIN")

    # ENTITY
    if maps_name == "MATCH":
        entity_status = "FOUND"
        entity_conf = float(r.get("maps_name_similarity") or 0)
    elif web_name == "MATCH":
        entity_status = "FOUND"
        entity_conf = float(r.get("website_name_similarity") or 0)
    else:
        entity_status = "UNCERTAIN"
        entity_conf = 0.0

    # ADDRESS
    if maps_name == "MATCH":
        address_status = maps_addr
    elif web_name == "MATCH":
        address_status = web_addr
    else:
        address_status = "UNCERTAIN"

    signals = []

    # Maps review.
    review_days = r.get("maps_review_latest_days")
    if (
        address_status == "MATCH"
        and review_days not in ("", None)
        and float(review_days) <= ACTIVITY_DAYS
    ):
        signals.append("Google Maps review <= 90 hari")

    # Maps open.
    maps_status = norm(r.get("maps_business_status", ""))
    if address_status == "MATCH":
        if "buka" in maps_status or "open" in maps_status:
            signals.append("Google Maps Open/Buka")

    # Website.
    web_days = r.get("website_activity_days")
    if (
        web_name == "MATCH"
        and web_addr in ("MATCH", "UNCERTAIN")
        and web_days not in ("", None)
        and float(web_days) <= ACTIVITY_DAYS
    ):
        signals.append("Website <= 90 hari")

    # JobStreet.
    js_days = r.get("jobstreet_age_days")
    if (
        address_status == "MATCH"
        and r.get("jobstreet_found") in (True, "True", "true", "1")
        and r.get("jobstreet_location_match") in (True, "True", "true", "1")
        and js_days not in ("", None)
        and float(js_days) <= ACTIVITY_DAYS
    ):
        signals.append("JobStreet <= 90 hari + lokasi cocok")

    # Glints.
    gl_days = r.get("glints_age_days")
    if (
        address_status == "MATCH"
        and r.get("glints_found") in (True, "True", "true", "1")
        and r.get("glints_location_match") in (True, "True", "true", "1")
        and gl_days not in ("", None)
        and float(gl_days) <= ACTIVITY_DAYS
    ):
        signals.append("Glints <= 90 hari + lokasi cocok")

    # Latest activity.
    activity_candidates = []

    for field, source in [
        ("maps_review_latest_days", "Google Maps"),
        ("website_activity_days", "Website"),
        ("jobstreet_age_days", "JobStreet"),
        ("glints_age_days", "Glints"),
    ]:
        value = r.get(field)

        if value not in ("", None):
            try:
                activity_candidates.append(
                    (float(value), source)
                )
            except Exception:
                pass

    activity_candidates.sort(key=lambda x: x[0])

    if activity_candidates:
        last_days, last_source = activity_candidates[0]
        last_date = date_text(last_days)
    else:
        last_days = None
        last_source = ""
        last_date = ""

    # Activity.
    if signals:
        activity = "ACTIVE"
    elif (
        address_status == "MATCH"
        and maps_name == "MATCH"
        and ("tutup" in maps_status or "closed" in maps_status)
    ):
        activity = "INACTIVE"
    else:
        activity = "UNCERTAIN"

    # Confidence.
    confidence = 0

    if entity_status == "FOUND":
        confidence += 20

    if address_status == "MATCH":
        confidence += 35
    elif address_status == "NOT_MATCH":
        confidence += 25

    confidence += min(45, len(signals) * 10)

    summary_category = output_category(r)
    category_scores = {
        "maps_kurang_lowongan_cocok_lokasi": 100,
        "maps_kurang_lowongan_lokasi_lain": 75,
        "maps_kurang_tanpa_lowongan": 40,
        "maps_lebih_lowongan_cocok_lokasi": 60,
    }
    if summary_category in category_scores:
        confidence = category_scores[summary_category]

    notes = []

    if address_status == "NOT_MATCH":
        notes.append(
            "Usaha ditemukan tetapi alamat Maps berbeda dari alamat input"
        )

    category_narratives = {
        "maps_kurang_lowongan_cocok_lokasi": (
            "Maps aktif < 90 hari dan JobStreet memiliki lowongan dengan lokasi cocok"
        ),
        "maps_kurang_lowongan_lokasi_lain": (
            "Maps aktif < 90 hari dan JobStreet memiliki lowongan di lokasi lain"
        ),
        "maps_kurang_tanpa_lowongan": (
            "Maps aktif < 90 hari dan tidak ditemukan lowongan JobStreet"
        ),
        "maps_lebih_lowongan_cocok_lokasi": (
            "Maps aktif > 90 hari dan JobStreet memiliki lowongan dengan lokasi cocok"
        ),
    }

    category_statuses = {
        "maps_kurang_lowongan_cocok_lokasi": "MAPS_AKTIF_KURANG_90_LOWONGAN_COCOK_LOKASI",
        "maps_kurang_lowongan_lokasi_lain": "MAPS_AKTIF_KURANG_90_LOWONGAN_LOKASI_LAIN",
        "maps_kurang_tanpa_lowongan": "MAPS_AKTIF_KURANG_90_TANPA_LOWONGAN",
        "maps_lebih_lowongan_cocok_lokasi": "MAPS_AKTIF_LEBIH_90_LOWONGAN_COCOK_LOKASI",
    }

    if summary_category in category_statuses:
        final = category_statuses[summary_category]
        notes.insert(0, category_narratives[summary_category])
    elif address_status == "NOT_MATCH":
        final = "DITEMUKAN_TAPI_ALAMAT_MAPS_TIDAK_COCOK"
    elif entity_status != "FOUND":
        final = "TIDAK_CUKUP_BUKTI_IDENTITAS_USAHA"
    elif activity == "ACTIVE" and address_status == "MATCH":
        final = "AKTIF_DI_ALAMAT_INPUT"
    elif activity == "INACTIVE" and address_status == "MATCH":
        final = "KEMUNGKINAN_TIDAK_AKTIF_DI_ALAMAT_INPUT"
    else:
        final = "PERLU_VERIFIKASI"

    if summary_category == "":
        notes.append(
            "Tidak masuk kategori summary: aktivitas Maps tepat 90 hari, "
            "data Maps/JobStreet belum lengkap, atau lokasi JobStreet belum jelas"
        )
    elif not signals:
        notes.append(
            "Tidak ditemukan bukti aktivitas <= 90 hari yang cukup kuat"
        )

    return {
        "entity_status": entity_status,
        "entity_confidence": round(entity_conf, 3),
        "address_status": address_status,
        "activity_status": activity,
        "last_activity_days": last_days,
        "last_activity_date": last_date,
        "last_activity_source": last_source,
        "activity_signal_count": len(signals),
        "activity_signals": " | ".join(signals),
        "confidence_score": min(100, confidence),
        "summary_category": summary_category,
        "final_status": final,
        "verification_note": " | ".join(notes),
    }


# ============================================================
# PROCESS ROW
# ============================================================

def get_col(row, names):
    for name in names:
        if name in row and str(row[name]).strip():
            return str(row[name]).strip()
    return ""


async def process_row(row, driver, args):
    name = get_col(
        row,
        ["nama_perusahaan", "nama", "nama_se", "nama_bpjph"]
    )

    address = get_col(
        row,
        ["alamat", "address", "alamat_se", "alamat_bpjph"]
    )

    kab = get_col(row, ["nmkab","nama_kab", "kabupaten", "kota"])

    row["nama_perusahaan"] = name
    row["alamat"] = address
    row["nama_kab"] = kab

    # ---------------- MAPS ----------------
    if not row.get("maps_processed") and not args.skip_maps:
        print("   -> Google Maps")

        m = await maps_search(driver, name, address, kab)

        row.update(m)
        row["maps_processed"] = "1"

        if m["maps_found"]:
            days, stamp, isi = await maps_latest_review(driver)

            row["maps_review_latest_days"] = days
            row["maps_review_latest_text"] = stamp
            row["maps_review_latest_isi"] = isi

    # ---------------- WEBSITE DISCOVERY ----------------
    if not row.get("search_processed") and not args.skip_search:
        print("   -> Website")

        d = await discover_website(driver, name, address)

        row.update(d)
        row["search_processed"] = "1"

    # Maps website lebih dipercaya sebagai kandidat awal.
    if (
        not row.get("website")
        and row.get("maps_website")
    ):
        row["website"] = row["maps_website"]

    # ---------------- WEBSITE ACTIVITY ----------------
    if (
        not row.get("website_activity_processed")
        and not args.skip_search
        and row.get("website")
    ):
        print("   -> Website activity")

        wa = await website_activity(
            driver,
            row["website"]
        )

        row["website_activity_status"] = wa["status"]
        row["website_activity_days"] = wa["days"]
        row["website_activity_date"] = date_text(wa["days"])
        row["website_activity_text"] = wa["text"]
        row["website_activity_url"] = wa["url"]
        row["website_activity_note"] = wa["note"]
        row["website_activity_processed"] = "1"

    # ---------------- JOBSTREET ----------------
    if (
        not row.get("jobstreet_processed")
        and not args.skip_lowongan
        and not args.skip_jobstreet
        and args.cek_lowongan in ("jobstreet", "keduanya")
    ):
        print("   -> JobStreet")

        js = await scrape_job_site(
            driver,
            "jobstreet",
            name,
            address,
            kab
        )

        row["jobstreet_found"] = str(js["found"])
        row["jobstreet_title"] = js["title"]
        row["jobstreet_company"] = js["company"]
        row["jobstreet_location"] = js["location"]
        row["jobstreet_age_days"] = js["age_days"]
        row["jobstreet_age_text"] = js["age_text"]
        row["jobstreet_date"] = date_text(js["age_days"])
        row["jobstreet_url"] = js["url"]
        row["jobstreet_location_match"] = js["location_match"]
        row["jobstreet_note"] = js["note"]

        row["jobstreet_processed"] = "1"

        await asyncio.sleep(random.uniform(*JOB_DELAY))

    # ---------------- GLINTS ----------------
    if (
        not row.get("glints_processed")
        and not args.skip_lowongan
        and not args.skip_glints
        and args.cek_lowongan in ("glints", "keduanya")
    ):
        print("   -> Glints")

        gl = await scrape_job_site(
            driver,
            "glints",
            name,
            address,
            kab
        )

        row["glints_found"] = str(gl["found"])
        row["glints_title"] = gl["title"]
        row["glints_company"] = gl["company"]
        row["glints_location"] = gl["location"]
        row["glints_age_days"] = gl["age_days"]
        row["glints_age_text"] = gl["age_text"]
        row["glints_date"] = date_text(gl["age_days"])
        row["glints_url"] = gl["url"]
        row["glints_location_match"] = gl["location_match"]
        row["glints_note"] = gl["note"]

        row["glints_processed"] = "1"

    if (
        (args.skip_lowongan or args.skip_jobstreet or row.get("jobstreet_processed"))
        and (args.skip_lowongan or args.skip_glints or row.get("glints_processed"))
    ):
        row["lowongan_processed"] = "1"

    # ---------------- FINAL ----------------
    row.update(final_status(row))
    row["checked_at"] = datetime.now().isoformat(timespec="seconds")

    return row


# ============================================================
# CSV
# ============================================================

INPUT_OUTPUT_FIELDS = [
    "nama",
    "alamat",
    "kd_prov",
    "kd_kab",
    "nmkab",
    "kd_kec",
    "kd_desa",
    "flag_usaha",
    "flag_keberadaan"
]

RESULT_OUTPUT_FIELDS = [
    "maps_found",
    "maps_name",
    "maps_address",
    "maps_address_similarity",
    "maps_url",
    "maps_business_status",
    "maps_review_latest_days",
    "maps_review_latest_text",
    "website",
    "website_activity_url",
    "website_name_match",
    "website_address_match",
    "website_activity_status",
    "website_activity_days",
    "website_activity_date",
    "jobstreet_found",
    "jobstreet_url",
    "jobstreet_company",
    "jobstreet_title",
    "jobstreet_location",
    "jobstreet_age_days",
    "jobstreet_date",
    "jobstreet_location_match",
    "jobstreet_note",
    "glints_found",
    "glints_url",
    "glints_company",
    "glints_title",
    "glints_location",
    "glints_age_days",
    "glints_date",
    "glints_location_match",
    "glints_note",
    "entity_status",
    "address_status",
    "activity_status",
    "last_activity_days",
    "last_activity_date",
    "last_activity_source",
    "activity_signal_count",
    "activity_signals",
    "confidence_score",
    "summary_category",
    "final_status",
    "verification_note",
    "checked_at",
]

OUTPUT_FIELDS = [
    "assignment_id",
    "idsbr",
    "nama",
    "alamat",
    "kd_prov",
    "kd_kab",
    "nmkab",
    "kd_kec",
    "kd_desa",
    "flag_keberadaan",
    "flag_usaha",
    "maps_ditemukan",
    "maps",
    "maps_nama",
    "maps_nama_similarity",
    "maps_alamat_teks",
    "maps_alamat_similarity",
    "maps_tlp",
    "maps_status_bisnis",
    "maps_review_terakhir_teks",
    "maps_review_terakhir_hari",
    "maps_review_terakhir_isi",
    "website",
    "jobstreet_ditemukan",
    "jobstreet_url",
    "jobstreet_nama_usaha",
    "jobstreet_judul_lowongan",
    "jobstreet_lokasi",
    "jobstreet_age_days",
    "jobstreet_tanggal",
    "jobstreet_lokasi_match",
    "summary_category",
    "checked_at",
]


def state_path(path):
    """Sidecar penyimpan state internal lengkap untuk resume.
    Sidecar ini menyimpan seluruh field internal agar resume tetap akurat.
    """
    return path + ".state.json"


def _blank(value):
    return "" if value is None else value


def apply_output_aliases(row):
    """Isi kolom output ringkas (OUTPUT_FIELDS) dari field internal."""
    # Google Maps
    row["maps_ditemukan"] = _blank(row.get("maps_found"))
    row["maps"] = _blank(row.get("maps_url"))
    row["maps_nama"] = _blank(row.get("maps_name"))
    row["maps_alamat_teks"] = _blank(row.get("maps_address"))
    row["maps_nama_similarity"] = _blank(row.get("maps_name_similarity"))
    row["maps_alamat_similarity"] = _blank(row.get("maps_address_similarity"))
    row["maps_tlp"] = _blank(row.get("maps_phone"))
    row["maps_status_bisnis"] = _blank(row.get("maps_business_status"))
    row["maps_review_terakhir_teks"] = _blank(row.get("maps_review_latest_text"))
    row["maps_review_terakhir_hari"] = _blank(row.get("maps_review_latest_days"))
    row["maps_review_terakhir_isi"] = _blank(row.get("maps_review_latest_isi"))

    # Website
    row["website"] = _blank(row.get("website"))

    # JobStreet
    row["jobstreet_ditemukan"] = _blank(row.get("jobstreet_found"))
    row["jobstreet_url"] = _blank(row.get("jobstreet_url"))
    row["jobstreet_nama_usaha"] = _blank(row.get("jobstreet_company"))
    row["jobstreet_judul_lowongan"] = _blank(row.get("jobstreet_title"))
    row["jobstreet_lokasi"] = _blank(row.get("jobstreet_location"))
    row["jobstreet_age_days"] = _blank(row.get("jobstreet_age_days"))
    row["jobstreet_tanggal"] = _blank(row.get("jobstreet_date"))
    row["jobstreet_lokasi_match"] = _blank(row.get("jobstreet_location_match"))

    # Ringkasan
    row["summary_category"] = _blank(row.get("summary_category"))
    return row


def _load_rows_for_report(csv_path):
    """Ambil baris untuk membangun XLSX: utamakan sidecar (field lengkap)."""
    sp = state_path(csv_path)
    if os.path.exists(sp):
        try:
            with open(sp, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def backup(path):
    if not os.path.exists(path):
        return ""

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target = path + ".backup_" + stamp
    shutil.copy2(path, target)
    return target


def read_existing(path):
    # Utamakan sidecar state (berisi seluruh field internal untuk resume).
    sp = state_path(path)
    if os.path.exists(sp):
        try:
            with open(sp, "r", encoding="utf-8") as f:
                data = json.load(f)
            result = {}
            for row in data:
                name = get_col(
                    row,
                    ["nama_perusahaan", "nama", "nama_se", "nama_bpjph"]
                )
                address = get_col(
                    row,
                    ["alamat", "address", "alamat_se", "alamat_bpjph"]
                )
                result[(name, address)] = row
            return result
        except Exception:
            pass

    if not os.path.exists(path):
        return {}

    result = {}

    with open(
        path,
        "r",
        encoding="utf-8-sig",
        newline=""
    ) as f:
        reader = csv.DictReader(f)

        for row in reader:
            name = get_col(
                row,
                ["nama_perusahaan", "nama", "nama_se", "nama_bpjph"]
            )
            address = get_col(
                row,
                ["alamat", "address", "alamat_se", "alamat_bpjph"]
            )
            key = (
                name,
                address
            )

            result[key] = row

    return result


def write_output(path, rows, fieldnames):
    # Selaraskan kolom output ringkas dengan state internal terbaru.
    for row in rows:
        apply_output_aliases(row)

    tmp = path + ".tmp"

    with open(
        tmp,
        "w",
        encoding="utf-8-sig",
        newline=""
    ) as f:
        writer = csv.DictWriter(
            f,
            fieldnames=fieldnames,
            extrasaction="ignore"
        )

        writer.writeheader()

        for row in rows:
            writer.writerow(row)

    os.replace(tmp, path)

    # Sidecar: simpan seluruh field internal agar resume tetap akurat.
    sp = state_path(path)
    tmp_state = sp + ".tmp"
    try:
        with open(tmp_state, "w", encoding="utf-8") as f:
            json.dump(rows, f, ensure_ascii=False, default=str)
        os.replace(tmp_state, sp)
    except Exception:
        pass


def as_bool(value):
    return value is True or str(value).strip().lower() in {
        "1", "true", "yes"
    }


def lowongan_location_category(row):
    jobstreet_found = str(row.get("jobstreet_found", "")).strip().lower()

    if jobstreet_found in {"true", "1", "yes"} and as_bool(
        row.get("jobstreet_location_match")
    ):
        return "lowongan_cocok_lokasi"
    if jobstreet_found in {"true", "1", "yes"} and str(
        row.get("jobstreet_location_match", "")
    ).strip().lower() in {"false", "0", "no"}:
        return "lowongan_lokasi_lain"
    if jobstreet_found in {"false", "0", "no"}:
        return "tanpa_lowongan"
    return ""


def output_category(row):
    if str(row.get("maps_address_match", "")).strip().upper() != "MATCH":
        return ""

    try:
        maps_days = float(row.get("maps_review_latest_days"))
    except (TypeError, ValueError):
        return ""

    job_category = lowongan_location_category(row)
    if not job_category:
        return ""

    if maps_days < ACTIVITY_DAYS:
        return f"maps_kurang_{job_category}"
    if maps_days > ACTIVITY_DAYS:
        return f"maps_lebih_{job_category}"
    return ""


def create_output_workbook(csv_path, xlsx_path, fieldnames):
    """Create an Excel workbook split into the requested status sheets."""
    rows = _load_rows_for_report(csv_path)

    workbook = Workbook()
    sheet_names = [
        "Diproses",
        "maps_<90_lowongan_cocok",
        "maps_<90_lowongan_lain",
        "maps_<90_tanpa_lowongan",
        "maps_>90_lowongan_cocok",
    ]
    sheet_categories = {
        "maps_<90_lowongan_cocok": "maps_kurang_lowongan_cocok_lokasi",
        "maps_<90_lowongan_lain": "maps_kurang_lowongan_lokasi_lain",
        "maps_<90_tanpa_lowongan": "maps_kurang_tanpa_lowongan",
        "maps_>90_lowongan_cocok": "maps_lebih_lowongan_cocok_lokasi",
    }
    sheets = {
        "Diproses": [row for row in rows if str(row.get("checked_at", "")).strip()]
    }

    for sheet_name, category in sheet_categories.items():
        sheets[sheet_name] = [
            row for row in sheets["Diproses"]
            if output_category(row) == category
        ]

    for index, sheet_name in enumerate(sheet_names):
        sheet = workbook.active if index == 0 else workbook.create_sheet()
        sheet.title = sheet_name
        sheet.append(fieldnames)
        for cell in sheet[1]:
            cell.font = Font(bold=True)

        for row in sheets[sheet_name]:
            sheet.append([row.get(field, "") for field in fieldnames])

    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column in sheet.columns:
            width = min(
                max(len(str(cell.value or "")) for cell in column) + 2,
                50,
            )
            sheet.column_dimensions[column[0].column_letter].width = width

    workbook.save(xlsx_path)


# ============================================================
# REKAP (tahap akhir): 2 file Excel UBUM & UMK
# ============================================================

REKAP_RECOMPUTE_SUMMARY = True   # hitung ulang summary_category dengan GATE
AMBANG_KEMIRIPAN = 0.80          # gate: nama & alamat harus > nilai ini
AMBANG_HARI_BARU = 90            # < 90 hari -> 'kurang', selain itu 'lebih'

FLAG_UMK = {"UMK"}
FLAG_UBUM = {"UM", "UB", "UBUM", "UMB"}

# Sheet UBUM: (nama sheet, nilai summary_category)
UBUM_SHEETS = [
    ("maps_<90_lowongan_cocok", "maps_kurang_lowongan_cocok_lokasi"),
    ("maps_<90_lowongan_lain", "maps_kurang_lowongan_lokasi_lain"),
    ("maps_<90_tanpa_lowongan", "maps_kurang_tanpa_lowongan"),
    ("maps_>90_lowongan_cocok", "maps_lebih_lowongan_cocok_lokasi"),
    # Aktifkan bila ingin SEMUA kategori dibuatkan sheet:
    # ("maps_>90_lowongan_lain", "maps_lebih_lowongan_lokasi_lain"),
    # ("maps_>90_tanpa_lowongan", "maps_lebih_tanpa_lowongan"),
]

# UBUM tanpa kolom kemiripan nama; UMK memakainya.
UBUM_COLUMNS = [c for c in OUTPUT_FIELDS if c != "maps_nama_similarity"]
UMK_COLUMNS = list(OUTPUT_FIELDS)


def _rekap_sim(a, b):
    """Kemiripan 0..1 (fallback bila kolom kemiripan kosong di CSV)."""
    a, b = " ".join(str(a or "").lower().split()), " ".join(str(b or "").lower().split())
    if not a or not b:
        return 0.0
    try:
        from rapidfuzz import fuzz
        return round(fuzz.token_set_ratio(a, b) / 100.0, 4)
    except Exception:
        from difflib import SequenceMatcher
        return round(SequenceMatcher(None, a, b).ratio(), 4)


def _to_bool(x):
    return str(x).strip().lower() in {"true", "1", "ya", "yes", "y", "t", "1.0"}


def _to_float(x):
    try:
        if x is None or (isinstance(x, str) and x.strip() == ""):
            return None
        return float(x)
    except (TypeError, ValueError):
        return None


def _rekap_summary_category(maps_found, nama_sim, alamat_sim, hari,
                            js_found, lokasi_match):
    """GATE + klasifikasi. Tidak lolos gate -> string kosong."""
    if not maps_found or nama_sim is None or alamat_sim is None:
        return ""
    if not (nama_sim > AMBANG_KEMIRIPAN and alamat_sim > AMBANG_KEMIRIPAN):
        return ""
    if hari is None:
        return ""
    recency = "kurang" if float(hari) < AMBANG_HARI_BARU else "lebih"
    if js_found and lokasi_match:
        pola = "lowongan_cocok_lokasi"
    elif js_found:
        pola = "lowongan_lokasi_lain"
    else:
        pola = "tanpa_lowongan"
    return f"maps_{recency}_{pola}"


def _siapkan_df_rekap(df):
    """Pastikan kolom bantu kemiripan/hari dan summary_category tersedia."""
    import pandas as pd

    df = df.copy().fillna("")

    def _kolom_sim(kolom, teks_maps, teks_input):
        if kolom in df.columns and df[kolom].astype(str).str.strip().ne("").any():
            return pd.to_numeric(df[kolom].apply(_to_float), errors="coerce")
        return pd.to_numeric(
            df.apply(lambda r: _rekap_sim(r.get(teks_maps, ""), r.get(teks_input, "")), axis=1),
            errors="coerce",
        )

    df["_alamat_sim_f"] = _kolom_sim("maps_alamat_similarity", "maps_alamat_teks", "alamat")
    df["_nama_sim_f"] = _kolom_sim("maps_nama_similarity", "maps_nama", "nama")
    if "maps_nama_similarity" not in df.columns or df["maps_nama_similarity"].astype(str).str.strip().eq("").all():
        df["maps_nama_similarity"] = df["_nama_sim_f"].apply(
            lambda v: "" if pd.isna(v) else round(float(v), 4))

    hari_col = df["maps_review_terakhir_hari"] if "maps_review_terakhir_hari" in df.columns else ""
    df["_hari_num"] = pd.to_numeric(
        pd.Series(hari_col, index=df.index).apply(_to_float), errors="coerce")

    perlu = (
        REKAP_RECOMPUTE_SUMMARY
        or "summary_category" not in df.columns
        or df["summary_category"].astype(str).str.strip().eq("").all()
    )
    if perlu:
        df["summary_category"] = df.apply(
            lambda r: _rekap_summary_category(
                maps_found=_to_bool(r.get("maps_ditemukan")),
                nama_sim=None if pd.isna(r["_nama_sim_f"]) else r["_nama_sim_f"],
                alamat_sim=None if pd.isna(r["_alamat_sim_f"]) else r["_alamat_sim_f"],
                hari=None if pd.isna(r["_hari_num"]) else r["_hari_num"],
                js_found=_to_bool(r.get("jobstreet_ditemukan")),
                lokasi_match=_to_bool(r.get("jobstreet_lokasi_match")),
            ),
            axis=1,
        )
    return df


def _tulis_sheet(writer, nama_sheet, data, kolom):
    cols = [c for c in kolom if c in data.columns]
    data[cols].to_excel(writer, sheet_name=nama_sheet[:31], index=False)
    ws = writer.sheets[nama_sheet[:31]]
    ws.freeze_panes = "A2"
    for cell in ws[1]:
        cell.font = Font(bold=True)
    if ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = ws.dimensions


def buat_rekap(csv_path, ubum_path, umk_path):
    """Baca CSV hasil scraping -> tulis rekap_UBUM.xlsx dan rekap_UMK.xlsx."""
    import pandas as pd

    df = pd.read_csv(csv_path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    if "flag_usaha" not in df.columns:
        print("REKAP dilewati: kolom 'flag_usaha' tidak ada di CSV hasil.")
        return

    df = _siapkan_df_rekap(df)
    flag = df["flag_usaha"].astype(str).str.upper().str.strip()
    umk = df[flag.isin(FLAG_UMK)].copy()
    ubum = df[flag.isin(FLAG_UBUM)].copy()

    # ---- UBUM ----
    with pd.ExcelWriter(ubum_path, engine="openpyxl") as w:
        _tulis_sheet(w, "Diproses", ubum, UBUM_COLUMNS)
        for sheet_name, cat in UBUM_SHEETS:
            _tulis_sheet(w, sheet_name, ubum[ubum["summary_category"] == cat], UBUM_COLUMNS)
    print(f"REKAP UBUM -> {ubum_path} ({len(ubum)} baris, {1 + len(UBUM_SHEETS)} sheet)")

    # ---- UMK ----
    gate = (umk["_nama_sim_f"] > AMBANG_KEMIRIPAN) & (umk["_alamat_sim_f"] > AMBANG_KEMIRIPAN)
    with pd.ExcelWriter(umk_path, engine="openpyxl") as w:
        _tulis_sheet(w, "diproses", umk, UMK_COLUMNS)
        _tulis_sheet(w, "ditemukan", umk[umk["maps_ditemukan"].apply(_to_bool)], UMK_COLUMNS)
        _tulis_sheet(w, "kemiripan > 0.80", umk[gate.fillna(False)], UMK_COLUMNS)
    print(f"REKAP UMK  -> {umk_path} ({len(umk)} baris, 3 sheet)")


def _jalankan_rekap(csv_path, ubum_path, umk_path):
    """Bungkus buat_rekap agar kegagalan rekap tidak menggugurkan hasil scraping."""
    try:
        buat_rekap(csv_path, ubum_path, umk_path)
    except ImportError as e:
        print(f"REKAP gagal: pustaka belum terpasang ({e}). "
              "Pasang: pip install pandas openpyxl rapidfuzz")
    except Exception as e:
        print(f"REKAP gagal: {type(e).__name__}: {e}")
        print(f"Jalankan ulang rekap saja: python scraping_usaha.py --rekap-only {csv_path}")


# ============================================================
# MAIN
# ============================================================

def _row_needs_processing(row, args):
    """True bila masih ada tahap yang belum selesai untuk baris ini (resume)."""
    done_values = {"1", "true", "yes"}

    maps_done = (
        args.skip_maps
        or str(row.get("maps_processed", "")).lower() in done_values
    )

    search_done = (
        args.skip_search
        or str(row.get("search_processed", "")).lower() in done_values
    )

    jobstreet_done = (
        args.skip_lowongan
        or args.skip_jobstreet
        or args.cek_lowongan not in ("jobstreet", "keduanya")
        or str(row.get("jobstreet_processed", "")).lower() in done_values
        or str(row.get("lowongan_processed", "")).lower() in done_values
    )

    glints_done = (
        args.skip_lowongan
        or args.skip_glints
        or args.cek_lowongan not in ("glints", "keduanya")
        or str(row.get("glints_processed", "")).lower() in done_values
        or str(row.get("lowongan_processed", "")).lower() in done_values
    )

    return not (maps_done and search_done and jobstreet_done and glints_done)


async def main():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "input_csv",
        nargs="?",
        help="CSV input (tidak perlu bila memakai --rekap-only)"
    )

    parser.add_argument(
        "output_csv",
        nargs="?",
        help="CSV output (tidak perlu bila memakai --rekap-only)"
    )

    parser.add_argument(
        "--rekap-ubum",
        default=None,
        help="Nama file rekap UM/UB (default: rekap_UBUM.xlsx di folder output_csv)"
    )

    parser.add_argument(
        "--rekap-umk",
        default=None,
        help="Nama file rekap UMK (default: rekap_UMK.xlsx di folder output_csv)"
    )

    parser.add_argument(
        "--no-rekap",
        action="store_true",
        help="Lewati pembuatan 2 file rekap di akhir proses"
    )

    parser.add_argument(
        "--rekap-only",
        metavar="HASIL_CSV",
        default=None,
        help="Lewati scraping; buat 2 file rekap dari CSV hasil yang sudah ada"
    )

    parser.add_argument(
        "--output-xlsx",
        default=None,
        help=(
            "Excel output dengan sheet hasil_detail dan resume. "
            "Default: nama output_csv dengan ekstensi .xlsx."
        ),
    )

    parser.add_argument(
        "--no-headless",
        action="store_true",
        help="Tampilkan browser"
    )

    parser.add_argument(
        "--channel",
        default=os.environ.get("CHROME_CHANNEL"),
        help=(
            "Pakai browser yang SUDAH terpasang tanpa mengunduh Chromium "
            "Playwright, mis. 'chrome' atau 'msedge'. Alternatif: set env "
            "CHROME_BIN ke path executable, atau CHROME_CHANNEL."
        ),
    )

    parser.add_argument(
        "--skip-maps",
        action="store_true"
    )

    parser.add_argument(
        "--skip-search",
        action="store_true"
    )

    parser.add_argument(
        "--skip-lowongan",
        action="store_true"
    )

    parser.add_argument(
        "--cek-lowongan",
        choices=("jobstreet", "glints", "keduanya"),
        default="keduanya",
        help="Pilih sumber lowongan yang diperiksa (default: keduanya)"
    )

    parser.add_argument(
        "--skip-jobstreet",
        action="store_true",
        help="Lewati pemeriksaan lowongan JobStreet"
    )

    parser.add_argument(
        "--skip-glints",
        action="store_true",
        help="Lewati pemeriksaan lowongan Glints"
    )

    parser.add_argument(
        "--mulai-ulang",
        action="store_true"
    )

    parser.add_argument(
        "--save-every",
        type=int,
        default=5
    )

    parser.add_argument(
        "--maks-baris",
        type=int,
        default=None,
        help=(
            "Batasi jumlah baris BARU yang diproses dalam sesi ini "
            "(baris yang sudah selesai dari run sebelumnya tetap ditulis "
            "apa adanya). Jalankan command yang sama lagi untuk lanjut "
            "ke baris berikutnya. Kalau tidak diisi, proses semua baris "
            "yang belum selesai."
        ),
    )

    parser.add_argument(
        "--konkurensi",
        type=int,
        default=CONCURRENCY_DEFAULT,
        help=(
            "Jumlah perusahaan yang diproses paralel dalam satu browser "
            f"(default {CONCURRENCY_DEFAULT}). Google agresif memblokir; "
            "turunkan ke 1 bila sering kena CAPTCHA."
        ),
    )

    parser.add_argument(
        "--maks-retry",
        type=int,
        default=MAX_RETRIES_DEFAULT,
        help=(
            "Jumlah percobaan per perusahaan dengan exponential backoff "
            f"(default {MAX_RETRIES_DEFAULT})."
        ),
    )

    args = parser.parse_args()

    def _rekap_paths(csv_path):
        folder = os.path.dirname(os.path.abspath(csv_path))
        ubum = args.rekap_ubum or os.path.join(folder, "rekap_UBUM.xlsx")
        umk = args.rekap_umk or os.path.join(folder, "rekap_UMK.xlsx")
        return ubum, umk

    # Mode rekap saja.
    if args.rekap_only:
        ubum_path, umk_path = _rekap_paths(args.rekap_only)
        _jalankan_rekap(args.rekap_only, ubum_path, umk_path)
        return

    if not args.input_csv or not args.output_csv:
        parser.error("input_csv dan output_csv wajib diisi (kecuali --rekap-only)")

    # Input.
    with open(
        args.input_csv,
        "r",
        encoding="utf-8-sig",
        newline=""
    ) as f:
        reader = csv.DictReader(f)
        input_rows = list(reader)

    if not input_rows:
        print("CSV kosong.")
        return

    # Existing checkpoint.
    if args.mulai_ulang:
        b = backup(args.output_csv)
        if b:
            print("Backup:", b)

        sp = state_path(args.output_csv)
        if os.path.exists(sp):
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            shutil.copy2(sp, sp + ".backup_" + stamp)
            os.remove(sp)

        existing = {}
    else:
        existing = read_existing(args.output_csv)

    rows = []

    for raw in input_rows:
        name = get_col(
            raw,
            ["nama_perusahaan", "nama", "nama_se", "nama_bpjph"]
        )

        address = get_col(
            raw,
            ["alamat", "address", "alamat_se", "alamat_bpjph"]
        )

        key = (name, address)

        if key in existing:
            row = existing[key]
        else:
            row = dict(raw)

        row["nama_perusahaan"] = name
        row["alamat"] = address
        row["nama_kab"] = get_col(
            raw,
            ["nama_kab", "kabupaten", "kota"]
        )

        rows.append(row)

    # Keep the report compact; processing flags remain internal row fields.
    fieldnames = OUTPUT_FIELDS

    print("=" * 70)
    print("ACTIVITY VERIFIER LOCAL (async)")
    print("=" * 70)
    print("Total:", len(rows))
    print("Maps:", not args.skip_maps)
    print("Website:", not args.skip_search)
    print(
        "JobStreet:",
        not args.skip_lowongan
        and not args.skip_jobstreet
        and args.cek_lowongan in ("jobstreet", "keduanya")
    )
    print(
        "Glints:",
        not args.skip_lowongan
        and not args.skip_glints
        and args.cek_lowongan in ("glints", "keduanya")
    )
    print("Headless:", not args.no_headless)
    if os.environ.get("CHROME_BIN"):
        print("Browser:", "CHROME_BIN=" + os.environ["CHROME_BIN"])
    elif args.channel:
        print("Browser:", f"channel={args.channel} (browser terpasang, tanpa unduh Chromium)")
    else:
        print("Browser:", "Chromium bawaan Playwright")
    print("Maks baris per sesi:", args.maks_baris if args.maks_baris else "semua yang belum selesai")
    print("Konkurensi:", args.konkurensi)
    print("Maks retry:", args.maks_retry)
    print("=" * 70)

    # Tentukan baris yang masih perlu diproses (resume). Baris yang sudah
    # selesai cukup di-refresh status turunannya (final_status).
    todo = []
    for index, row in enumerate(rows, 1):
        if _row_needs_processing(row, args):
            todo.append((index, row))
        else:
            row.update(final_status(row))

    if args.maks_baris is not None:
        # Batasi jumlah baris BARU untuk sesi ini; sisanya dilanjut run berikut.
        todo = todo[: args.maks_baris]

    print(f"Perlu diproses sesi ini: {len(todo)} dari {len(rows)}")

    browser = await create_browser(
        headless=not args.no_headless,
        channel=args.channel,
    )

    # Konkurensi terkendali + retry per perusahaan (gaya async pipeline).
    sem = asyncio.Semaphore(max(1, args.konkurensi))
    save_lock = asyncio.Lock()
    state = {"processed": 0, "ok": 0}

    async def worker(index, row):
        label = f"[{index}/{len(rows)}]"
        print(f"\n{label} {row['nama_perusahaan']}")
        print("Alamat:", row["alamat"])

        async with sem:
            for attempt in range(1, args.maks_retry + 1):
                # Satu page baru per percobaan (dipakai untuk semua tahap
                # verifikasi perusahaan ini), lalu ditutup di finally.
                session = await browser.new_session()
                try:
                    await process_row(row, session, args)
                    break

                except Exception as e:
                    print(
                        f"  {label} Attempt {attempt}/{args.maks_retry} "
                        f"gagal: {type(e).__name__}: {e}"
                    )
                    if attempt < args.maks_retry:
                        await asyncio.sleep(2 ** attempt)  # backoff
                    else:
                        row["verification_note"] = (
                            f"ERROR: {type(e).__name__}: {e}"
                        )
                        row["checked_at"] = datetime.now().isoformat(
                            timespec="seconds"
                        )
                        print(
                            f"  {label} Menyerah setelah "
                            f"{args.maks_retry} percobaan"
                        )

                finally:
                    await session.close()

        # Checkpoint aman lintas task (asyncio single-thread + lock).
        async with save_lock:
            state["processed"] += 1
            if str(row.get("checked_at", "")).strip() and not str(
                row.get("verification_note", "")
            ).startswith("ERROR:"):
                state["ok"] += 1
            if state["processed"] % args.save_every == 0:
                write_output(args.output_csv, rows, fieldnames)
                print("CHECKPOINT:", args.output_csv)

    try:
        await asyncio.gather(*(worker(idx, row) for idx, row in todo))

    except KeyboardInterrupt:
        print("\nDihentikan user. Menyimpan...")

    finally:
        write_output(
            args.output_csv,
            rows,
            fieldnames
        )

        try:
            await browser.quit()
        except Exception:
            pass

    processed = state["processed"]

    output_xlsx = args.output_xlsx or os.path.splitext(
        args.output_csv
    )[0] + ".xlsx"
    create_output_workbook(
        args.output_csv,
        output_xlsx,
        fieldnames,
    )

    # Summary.
    summary_categories = [
        (
            "Diproses",
            lambda row: bool(str(row.get("checked_at", "")).strip()),
        ),
        (
            "Maps aktif < 90 hari + lowongan cocok lokasi",
            lambda row: output_category(row)
            == "maps_kurang_lowongan_cocok_lokasi",
        ),
        (
            "Maps aktif < 90 hari + lowongan di lokasi lain",
            lambda row: output_category(row)
            == "maps_kurang_lowongan_lokasi_lain",
        ),
        (
            "Maps aktif < 90 hari + tidak ada lowongan",
            lambda row: output_category(row)
            == "maps_kurang_tanpa_lowongan",
        ),
        (
            "Maps > 90 hari + lowongan cocok lokasi",
            lambda row: output_category(row)
            == "maps_lebih_lowongan_cocok_lokasi",
        ),
    ]

    counts = [
        (label, sum(1 for row in rows if predicate(row)))
        for label, predicate in summary_categories
    ]

    belum_selesai = sum(
        1
        for row in rows
        if str(row.get("maps_processed", "")).lower() not in {"1", "true", "yes"}
        or str(row.get("search_processed", "")).lower() not in {"1", "true", "yes"}
        or (
            not args.skip_lowongan
            and not args.skip_jobstreet
            and args.cek_lowongan in ("jobstreet", "keduanya")
            and str(row.get("jobstreet_processed", "")).lower()
            not in {"1", "true", "yes"}
            and str(row.get("lowongan_processed", "")).lower()
            not in {"1", "true", "yes"}
        )
        or (
            not args.skip_lowongan
            and not args.skip_glints
            and args.cek_lowongan in ("glints", "keduanya")
            and str(row.get("glints_processed", "")).lower()
            not in {"1", "true", "yes"}
            and str(row.get("lowongan_processed", "")).lower()
            not in {"1", "true", "yes"}
        )
    )

    print()
    print("=" * 70)
    print("SELESAI")
    print("=" * 70)
    print("Output:", args.output_csv)
    print("Excel:", output_xlsx)
    print("Diproses sesi ini:", processed)

    for label, count in counts:
        print(f"{label}: {count}")

    if belum_selesai:
        print()
        print(f"Masih ada {belum_selesai} baris belum selesai diproses.")
        print("Jalankan command yang sama lagi untuk melanjutkan.")

    print("=" * 70)

    # ---------------- TAHAP AKHIR: REKAP UBUM & UMK ----------------
    if not args.no_rekap:
        print()
        print("REKAP (UBUM & UMK)")
        ubum_path, umk_path = _rekap_paths(args.output_csv)
        _jalankan_rekap(args.output_csv, ubum_path, umk_path)
        print("=" * 70)


if __name__ == "__main__":
    asyncio.run(main())
