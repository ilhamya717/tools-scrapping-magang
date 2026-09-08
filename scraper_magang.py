"""
Scraper Info Magang - IT/Software/Data
========================================
Script ini mengambil info lowongan magang dari beberapa sumber (dengan
beberapa query pencarian sekaligus, lihat QUERIES di bawah), lalu
memfilter hanya yang relevan dengan bidang IT/Software/Data/Cybersecurity,
dan menyimpan hasilnya ke file Excel.

Sumber yang didukung: Kalibrr, LinkedIn, Glints, Jobstreet.

CATATAN PENTING SEBELUM PAKAI:
1. Situs job portal (Glints, Kalibrr, Jobstreet, dll) sering mengubah
   struktur halaman/API mereka & punya proteksi anti-bot (Cloudflare, dsb).
   Jadi script ini mungkin perlu di-update kalau ada perubahan di sisi mereka.
2. Glints & Jobstreet full-JS render dan/atau ditutupi Cloudflare, jadi
   KEDUANYA wajib pakai Selenium (bukan cuma fallback) -> perlu install
   Google Chrome + chromedriver (otomatis lewat webdriver-manager).
   Kalibrr & LinkedIn bisa pakai request biasa (endpoint publik).
3. Selalu cek Terms of Service platform yang di-scrape. Gunakan dengan wajar
   (jangan request terlalu cepat/banyak) supaya IP tidak diblokir.
   -> KHUSUS LINKEDIN: User Agreement mereka EKSPLISIT melarang scraping
   otomatis (bukan cuma proteksi teknis kayak sumber lain). Endpoint yang
   dipakai di sini publik & tanpa login, tapi tetap ada risiko IP diblokir
   kalau dipakai dengan volume besar/sering. Defaultnya sengaja dibatasi
   kecil (lihat MAGANG_LINKEDIN_MAX_PAGES) - jangan dinaikkan sembarangan.
4. KEYWORDS mencakup istilah cybersecurity (SOC, pentester, red/blue team,
   dst), tapi keyword itu cuma MEMFILTER hasil pencarian - supaya lowongan
   itu benar-benar ketemu, query pencarian yang dikirim ke tiap situs juga
   harus relevan. Makanya QUERIES di bawah defaultnya sudah multi-query
   (bukan cuma "magang it"). Override lewat MAGANG_QUERIES kalau perlu.

Install dependency:
    pip install requests beautifulsoup4 pandas openpyxl selenium webdriver-manager
"""

import json
import logging
import os
import re
import time
from datetime import date, datetime, timedelta

import pandas as pd
import requests
from requests.adapters import HTTPAdapter, Retry

# =========================================================
# KONFIGURASI
# =========================================================
# Bisa dioverride lewat environment variable, contoh (PowerShell):
#   $env:MAGANG_QUERIES = "magang backend,magang soc analyst"
#   $env:MAGANG_MAX_PAGES = "5"
#   $env:MAGANG_MONTH_START = "8"
#   $env:MAGANG_MONTH_END = "9"

# QUERIES: daftar query pencarian yang dikirim ke tiap situs (satu run scraping
# dilakukan PER query, hasilnya digabung & di-dedup di akhir). Ini penting -
# KEYWORDS di bawah cuma memfilter hasil yang situs balikin, jadi kalau query
# pencariannya cuma "magang it", lowongan seperti "SOC Analyst Internship"
# (yang tidak mengandung kata "IT" di judul) kemungkinan besar tidak akan
# pernah muncul di hasil pencarian situs sejak awal.
#
# - MAGANG_QUERIES (comma-separated) -> override penuh daftar query.
# - MAGANG_QUERY (single, utk kompatibilitas lama) -> dipakai kalau
#   MAGANG_QUERIES tidak di-set DAN env var ini memang di-set user.
# - Kalau dua-duanya tidak di-set -> pakai default multi-query di bawah.
_DEFAULT_QUERIES = ["magang it", "magang cyber security", "magang data"]

if os.environ.get("MAGANG_QUERIES", "").strip():
    QUERIES = [q.strip() for q in os.environ["MAGANG_QUERIES"].split(",") if q.strip()]
elif os.environ.get("MAGANG_QUERY", "").strip():
    QUERIES = [os.environ["MAGANG_QUERY"].strip()]
else:
    QUERIES = _DEFAULT_QUERIES

QUERY = QUERIES[0]  # dipakai sebagai default param tunggal di beberapa fungsi/log ringkas
MAX_PAGES = int(os.environ.get("MAGANG_MAX_PAGES", "3"))
RAW_DUMP_DIR = os.environ.get("MAGANG_RAW_DIR", "")  # isi path kalau mau simpan raw response utk debug

# Filter tanggal posting: hanya ambil lowongan yang di-posting di rentang bulan ini
# (default Agustus-September tahun berjalan), supaya tidak kebanjiran postingan lama.
DATE_FILTER_YEAR = int(os.environ.get("MAGANG_YEAR", str(datetime.now().year)))
MONTH_START = int(os.environ.get("MAGANG_MONTH_START", "8"))   # Agustus
MONTH_END = int(os.environ.get("MAGANG_MONTH_END", "9"))       # September
# Lowongan yang tanggal postingnya tidak berhasil dideteksi -> default DIBUANG
# (supaya filter beneran ketat). Set "true" kalau mau tetap disertakan.
INCLUDE_UNKNOWN_DATE = os.environ.get("MAGANG_INCLUDE_UNKNOWN_DATE", "false").strip().lower() == "true"

# Kata kunci untuk filter jurusan IT / Software / Data.
# Dicek sebagai WHOLE WORD (bukan substring) supaya "it" tidak match "iterasi",
# "ai" tidak match "air", "qa" tidak match "qariah", dst.
KEYWORDS = [
    "it", "informatika", "software", "developer", "programmer",
    "data", "data science", "data analyst", "machine learning",
    "artificial intelligence", "ai", "backend", "frontend", "fullstack",
    "full stack", "web developer", "mobile developer", "sistem informasi",
    "teknik komputer", "cyber security", "devops", "cloud", "qa", "quality assurance",
    "database", "sql", "python", "java", "javascript",
    # -- Cybersecurity / Security Operations --
    "cybersecurity", "cyber", "keamanan siber", "information security", "infosec",
    "security analyst", "security engineer", "security researcher",
    "soc", "security operation center", "security operations center",
    "l1", "l2", "l3", "tier 1", "tier 2", "tier 3",
    "pentester", "penetration tester", "penetration testing", "pentest",
    "vulnerability assessment", "vapt", "ethical hacker", "ethical hacking",
    "red team", "red teamer", "blue team", "blue teamer", "purple team",
    "incident response", "threat hunting", "threat intelligence",
    "malware analyst", "malware analysis", "reverse engineering",
    "digital forensics", "forensic analyst", "siem", "network security",
    "application security", "appsec", "cloud security", "iam",
    "bug bounty", "ctf",
]

# Keyword yang wajib dianggap sebagai "kata utuh" (rawan false-positive kalau substring)
_STRICT_WORD_KEYWORDS = {
    "it", "ai", "qa", "sql", "cloud", "data", "cyber",
    "l1", "l2", "l3", "iam", "ctf",
    "soc",       # tanpa word-boundary nyangkut ke "social" (Social Media, dst)
    "database",  # sama generiknya dengan "data" - sering muncul di deskripsi non-IT
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "id-ID,id;q=0.9,en-US;q=0.8,en;q=0.7",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

OUTPUT_FILE = f"hasil_magang_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("scraper_magang")

# Tanda-tanda halaman yang sebenarnya adalah proteksi anti-bot/captcha,
# bukan hasil pencarian asli.
ANTI_BOT_MARKERS = [
    "checking your browser", "cf-browser-verification", "attention required",
    "verify you are human", "captcha", "access denied", "just a moment",
]


def build_session() -> requests.Session:
    """Session requests dengan retry+backoff otomatis untuk error transient."""
    session = requests.Session()
    session.headers.update(HEADERS)
    retries = Retry(
        total=3,
        backoff_factor=1.5,  # 0s, 1.5s, 3s, ...
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session


def looks_like_anti_bot(text: str) -> bool:
    lowered = text[:3000].lower()
    return any(marker in lowered for marker in ANTI_BOT_MARKERS)


def dump_raw(name: str, content: str):
    """Simpan raw response untuk debugging kalau MAGANG_RAW_DIR di-set."""
    if not RAW_DUMP_DIR:
        return
    try:
        os.makedirs(RAW_DUMP_DIR, exist_ok=True)
        path = os.path.join(RAW_DUMP_DIR, name)
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
    except Exception as e:
        log.warning("Gagal simpan raw dump %s: %s", name, e)


# Keyword untuk deteksi tipe kerja (Remote / Hybrid / On-site).
# Urutan penting: cek Hybrid dulu supaya kalimat kayak "hybrid (WFH & WFO)"
# tidak kepental jadi Remote gara-gara ada kata "wfh"/"remote" di situ juga.
REMOTE_MARKERS = ["remote", "wfh", "work from home", "fully remote", "kerja dari rumah"]
HYBRID_MARKERS = ["hybrid", "wfh & wfo", "wfh/wfo", "semi remote", "campuran"]
ONSITE_MARKERS = ["on-site", "onsite", "wfo", "work from office", "di kantor"]

_keyword_patterns = None


def _get_keyword_patterns():
    """Compile regex sekali saja, dipisah jadi 2 grup:
    - generic: keyword pendek/ambigu (whole word \\b...\\b) -> HANYA dicek di judul.
    - specific: frasa/keyword panjang -> dicek di judul + deskripsi.
    (Alasan pemisahan ada di docstring is_relevant().)"""
    global _keyword_patterns
    if _keyword_patterns is not None:
        return _keyword_patterns

    generic, specific = [], []
    for kw in KEYWORDS:
        kw_lower = kw.lower()
        if kw_lower in _STRICT_WORD_KEYWORDS:
            generic.append(re.compile(r"\b" + re.escape(kw_lower) + r"\b"))
        else:
            specific.append(re.compile(re.escape(kw_lower)))
    _keyword_patterns = (generic, specific)
    return _keyword_patterns


def is_relevant(title: str, description: str = "") -> bool:
    """Cek apakah lowongan relevan dengan IT/Software/Data/Cybersecurity.

    Keyword pendek & ambigu (it, data, cloud, ai, qa, sql, cyber, l1-l3, iam, ctf)
    HANYA dicek di JUDUL, bukan deskripsi -> deskripsi lowongan APAPUN (HR,
    admin, gudang, marketing, dst) hampir selalu menyebut kata "data"/"IT"/
    "cloud" secara insidental (mis. "input data pelanggan", "koordinasi
    dengan tim IT", "aplikasi berbasis cloud"), jadi kalau ikut dicek di
    deskripsi hasilnya kebanjiran false positive (sudah terverifikasi lewat
    run nyata - lihat catatan histori/commit).
    Keyword spesifik/frasa (data science, penetration tester, dst) tetap
    dicek di judul MAUPUN deskripsi karena jauh lebih jarang muncul kebetulan.
    """
    generic_patterns, specific_patterns = _get_keyword_patterns()
    title_lower = (title or "").lower()
    if not title_lower:
        return False

    if any(p.search(title_lower) for p in generic_patterns):
        return True

    combined = title_lower + " " + (description or "").lower()
    return any(p.search(combined) for p in specific_patterns)


def detect_work_type(*texts: str) -> str:
    """Deteksi tipe kerja dari judul/deskripsi/lokasi: Remote, Hybrid, On-site,
    atau "Tidak diketahui" kalau tidak ada penanda sama sekali."""
    combined = " ".join(t for t in texts if t).lower()
    if not combined:
        return "Tidak diketahui"

    def has_any(markers):
        return any(re.search(r"\b" + re.escape(m) + r"\b", combined) for m in markers)

    # Hybrid dicek duluan karena kalimat hybrid sering menyebut remote & on-site sekaligus.
    if has_any(HYBRID_MARKERS):
        return "Hybrid"
    if has_any(REMOTE_MARKERS):
        return "Remote"
    if has_any(ONSITE_MARKERS):
        return "On-site"
    return "Tidak diketahui"


# =========================================================
# FILTER TANGGAL POSTING (Agustus-September, dsb sesuai config)
# =========================================================
_RELATIVE_DATE_PATTERNS = [
    (re.compile(r"(\d+)\s*(?:hari|day)s?\s*(?:yang lalu|ago|lalu)", re.I), "days"),
    (re.compile(r"(\d+)\s*(?:jam|hour)s?\s*(?:yang lalu|ago|lalu)", re.I), "hours"),
    (re.compile(r"(\d+)\s*(?:minggu|week)s?\s*(?:yang lalu|ago|lalu)", re.I), "weeks"),
    (re.compile(r"(\d+)\s*(?:bulan|month)s?\s*(?:yang lalu|ago|lalu)", re.I), "months"),
]


def parse_posted_date(text: str = "", iso_hint: str = "") -> "date | None":
    """Coba tentukan tanggal posting lowongan dari:
    1. `iso_hint`: tanggal absolut/ISO dari API (mis. field created_at Kalibrr).
    2. `text`: teks bebas dari card (mis. "3 hari yang lalu", "Posted 2 days ago",
       "Hari ini", "Kemarin").
    Return None kalau tidak berhasil dideteksi sama sekali.
    """
    if iso_hint:
        try:
            return datetime.fromisoformat(iso_hint.replace("Z", "+00:00")).date()
        except Exception:
            try:
                return datetime.strptime(iso_hint[:10], "%Y-%m-%d").date()
            except Exception:
                pass

    if not text:
        return None
    t = text.lower()

    if any(m in t for m in ["hari ini", "today", "baru saja", "just now"]):
        return date.today()
    if any(m in t for m in ["kemarin", "yesterday"]):
        return date.today() - timedelta(days=1)

    for pattern, unit in _RELATIVE_DATE_PATTERNS:
        m = pattern.search(t)
        if m:
            n = int(m.group(1))
            if unit == "days":
                return date.today() - timedelta(days=n)
            if unit == "hours":
                return date.today()  # dalam hitungan jam -> anggap hari ini
            if unit == "weeks":
                return date.today() - timedelta(weeks=n)
            if unit == "months":
                return date.today() - timedelta(days=30 * n)

    return None


def is_within_target_period(posted: "date | None") -> bool:
    """Cek apakah tanggal posting ada di rentang bulan target (MONTH_START..MONTH_END,
    tahun DATE_FILTER_YEAR). Kalau tanggal tidak diketahui, ikuti INCLUDE_UNKNOWN_DATE."""
    if posted is None:
        return INCLUDE_UNKNOWN_DATE
    return posted.year == DATE_FILTER_YEAR and MONTH_START <= posted.month <= MONTH_END


# =========================================================
# SUMBER 1: KALIBRR (punya endpoint pencarian berbasis JSON)
# =========================================================
def scrape_kalibrr(session: requests.Session, query=QUERY, max_pages=MAX_PAGES):
    """
    Kalibrr punya endpoint pencarian job board yang mengembalikan JSON.
    Endpoint ini bisa berubah sewaktu-waktu, jadi kalau tidak jalan,
    cek ulang lewat browser (buka Network tab pas search di kalibrr.com).
    """
    results = []
    skipped_not_relevant = 0
    skipped_date = 0
    base_url = "https://www.kalibrr.com/kjs/job_board/search"

    for page in range(max_pages):
        limit = 30
        params = {
            "q": query,
            "country": "Indonesia",
            "employment_type": "Internship",
            "limit": limit,   # wajib -> tanpa ini API balas 400 "Missing query parameter 'limit'"
            "offset": page * limit,  # wajib juga -> API pakai offset, bukan "page"
        }
        try:
            resp = session.get(base_url, params=params, timeout=15)
            resp.raise_for_status()
            if looks_like_anti_bot(resp.text):
                log.warning("[Kalibrr] Halaman %s kelihatan seperti proteksi anti-bot, berhenti.", page)
                break
            data = resp.json()
        except requests.exceptions.RequestException as e:
            log.warning("[Kalibrr] Gagal ambil halaman %s (jaringan): %s", page, e)
            break
        except (ValueError, json.JSONDecodeError) as e:
            log.warning("[Kalibrr] Gagal parse JSON halaman %s: %s", page, e)
            dump_raw(f"kalibrr_page{page}.txt", resp.text)
            break

        jobs = data.get("jobs", [])
        if not jobs:
            log.info("[Kalibrr] Halaman %s kosong, anggap sudah habis.", page)
            break

        for job in jobs:
            title = job.get("name", "") or ""
            description = job.get("description", "") or ""
            # NB: objek "company" cuma punya code/description, BUKAN "name" -> nama
            # perusahaan sebenarnya ada di field top-level "company_name".
            company = job.get("company_name", "") or ""
            url = f"https://www.kalibrr.com/c/{job.get('company', {}).get('code','')}/jobs/{job.get('id','')}"
            posted_raw = job.get("created_at") or job.get("activation_date") or ""
            # NB: "city" nested di dalam address_components, bukan langsung di google_location.
            lokasi = job.get("google_location", {}).get("address_components", {}).get("city", "")
            posted_date = parse_posted_date(iso_hint=posted_raw)

            if not is_relevant(title, description):
                skipped_not_relevant += 1
                continue
            if not is_within_target_period(posted_date):
                skipped_date += 1
                continue

            # Kalibrr expose flag eksplisit is_hybrid/is_work_from_home -> lebih
            # akurat daripada nebak dari teks judul/deskripsi.
            if job.get("is_hybrid"):
                tipe_kerja = "Hybrid"
            elif job.get("is_work_from_home"):
                tipe_kerja = "Remote"
            else:
                tipe_kerja = detect_work_type(title, description, lokasi)
                if tipe_kerja == "Tidak diketahui":
                    tipe_kerja = "On-site"  # Kalibrr eksplisit set False/False utk kerja di kantor

            results.append({
                "sumber": "Kalibrr",
                "posisi": title,
                "perusahaan": company,
                "lokasi": lokasi,
                "tipe_kerja": tipe_kerja,
                "tanggal_posting": posted_date.isoformat() if posted_date else posted_raw,
                "link": url,
            })
        time.sleep(1.5)  # jangan spam request

    log.info(
        "[Kalibrr] q=%r: %s cocok, %s di-skip (bukan bidang IT/security), %s di-skip (di luar rentang tanggal)",
        query, len(results), skipped_not_relevant, skipped_date,
    )
    return results


# =========================================================
# SUMBER 2: LINKEDIN (guest job search endpoint - publik, tanpa login)
# =========================================================
# PENTING soal legal/ToS - BEDA dengan Kalibrr/Glints/Jobstreet:
# LinkedIn User Agreement (linkedin.com/legal/user-agreement) EKSPLISIT
# melarang scraping otomatis, bukan cuma proteksi teknis. Endpoint di bawah
# ini publik (tanpa login, dipakai luas oleh scraper open-source), datanya
# juga publik (job posting yang memang ditampilkan ke pengunjung anonim),
# tapi tetap melanggar ketentuan layanan mereka kalau dipakai otomatis.
# - Jangan jalankan dengan volume besar / terlalu sering -> IP kamu bisa
#   diblokir sementara oleh LinkedIn.
# - LINKEDIN_MAX_PAGES sengaja dibikin kecil by default & delay antar
#   halaman lebih panjang daripada sumber lain.
# - Kalau mau benar-benar aman secara legal, pertimbangkan pakai LinkedIn
#   Jobs API resmi (perlu partner access) alih-alih endpoint guest ini.
LINKEDIN_MAX_PAGES = int(os.environ.get("MAGANG_LINKEDIN_MAX_PAGES", "2"))
LINKEDIN_PAGE_SIZE = 10  # tetap dari API-nya, tidak bisa diubah lewat parameter


def scrape_linkedin(session: requests.Session, query=QUERY, max_pages=None):
    """
    Endpoint guest LinkedIn ("seeMoreJobPostings") mengembalikan HTML fragment
    berisi job card, tanpa perlu login. Field yang tersedia lengkap (judul,
    perusahaan, lokasi, link, tanggal ISO absolut di tag <time datetime>).
    """
    from bs4 import BeautifulSoup

    if max_pages is None:
        max_pages = LINKEDIN_MAX_PAGES

    results = []
    skipped_not_relevant = 0
    skipped_date = 0
    base_url = "https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search"

    for page in range(max_pages):
        params = {
            "keywords": query,
            "location": "Indonesia",
            "f_JT": "I",  # Internship
            "start": page * LINKEDIN_PAGE_SIZE,
        }
        try:
            resp = session.get(base_url, params=params, timeout=15)
            resp.raise_for_status()
        except requests.exceptions.RequestException as e:
            log.warning("[LinkedIn] Gagal ambil halaman %s: %s", page, e)
            break

        if looks_like_anti_bot(resp.text):
            log.warning("[LinkedIn] Halaman %s kelihatan seperti proteksi anti-bot, berhenti.", page)
            dump_raw(f"linkedin_page{page}.html", resp.text)
            break

        soup = BeautifulSoup(resp.text, "html.parser")
        cards = soup.select("div.base-card")
        if not cards:
            log.info("[LinkedIn] Halaman %s tidak ada job card, anggap sudah habis.", page)
            break

        for card in cards:
            title_tag = card.select_one(".base-search-card__title")
            title = title_tag.get_text(strip=True) if title_tag else ""

            company_tag = card.select_one(".base-search-card__subtitle")
            company = company_tag.get_text(strip=True) if company_tag else ""

            location_tag = card.select_one(".job-search-card__location")
            location = location_tag.get_text(strip=True) if location_tag else ""

            link_tag = card.select_one("a.base-card__full-link")
            link = link_tag["href"].split("?")[0] if link_tag and link_tag.get("href") else ""

            time_tag = card.select_one("time")
            posted_iso = time_tag.get("datetime") if time_tag else ""

            description = card.get_text(" ", strip=True)
            posted_date = parse_posted_date(iso_hint=posted_iso)

            if not title or not is_relevant(title, description):
                skipped_not_relevant += 1
                continue
            if not is_within_target_period(posted_date):
                skipped_date += 1
                continue

            results.append({
                "sumber": "LinkedIn",
                "posisi": title,
                "perusahaan": company,
                "lokasi": location,
                "tipe_kerja": detect_work_type(title, description, location),
                "tanggal_posting": posted_date.isoformat() if posted_date else "",
                "link": link,
            })

        # LinkedIn lebih sensitif thd rate-limit daripada sumber lain -> jeda lebih panjang.
        time.sleep(3)

    log.info(
        "[LinkedIn] q=%r: %s cocok, %s di-skip (bukan bidang IT/security), %s di-skip (di luar rentang tanggal)",
        query, len(results), skipped_not_relevant, skipped_date,
    )
    return results


# =========================================================
# HELPER SELENIUM (dipakai bareng oleh Glints & Jobstreet, dua-duanya
# butuh render JS supaya kontennya kebaca / supaya lolos proteksi Cloudflare)
# =========================================================
def _import_selenium():
    """Import lazy supaya scraper lain tetap jalan walau selenium belum terinstall."""
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait
    from webdriver_manager.chrome import ChromeDriverManager
    return webdriver, Options, Service, By, EC, WebDriverWait, ChromeDriverManager


def create_chrome_driver(source_name: str):
    """Bikin Chrome headless driver siap pakai, atau None kalau gagal
    (selenium belum terinstall / chromedriver gagal start)."""
    try:
        webdriver, Options, Service, By, EC, WebDriverWait, ChromeDriverManager = _import_selenium()
    except ImportError:
        log.info("[%s] Selenium belum terinstall, lewati sumber ini.", source_name)
        log.info("         Install dengan: pip install selenium webdriver-manager")
        return None, None

    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument(f"user-agent={HEADERS['User-Agent']}")

    # Cache lokasi chromedriver supaya tidak selalu hit network di setiap run.
    driver_cache_file = os.path.join(os.path.expanduser("~"), ".cache_scraper_magang_driver_path")
    driver_path = None
    if os.path.exists(driver_cache_file):
        try:
            with open(driver_cache_file, "r") as f:
                cached = f.read().strip()
            if cached and os.path.exists(cached):
                driver_path = cached
        except Exception:
            driver_path = None
    if not driver_path:
        driver_path = ChromeDriverManager().install()
        try:
            with open(driver_cache_file, "w") as f:
                f.write(driver_path)
        except Exception:
            pass

    try:
        driver = webdriver.Chrome(service=Service(driver_path), options=options)
    except Exception as e:
        log.error("[%s] Gagal start Chrome driver: %s", source_name, e)
        return None, None

    return driver, (By, EC, WebDriverWait)


def load_page_with_retry(driver, url, wait_selector, sel, source_name, page, timeout=15, tries=2):
    """`driver.get(url)` + tunggu elemen muncul, dicoba ulang `tries` kali kalau
    timeout/gagal (koneksi lambat, render telat, dsb) sebelum benar-benar nyerah.
    Return True kalau berhasil, False kalau semua percobaan gagal."""
    By, EC, WebDriverWait = sel
    for attempt in range(1, tries + 1):
        try:
            driver.get(url)
            WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, wait_selector))
            )
            return True
        except Exception as e:
            log.warning(
                "[%s] Percobaan %s/%s gagal load halaman %s: %s",
                source_name, attempt, tries, page, e,
            )
            if attempt < tries:
                time.sleep(2)
    return False


# =========================================================
# SUMBER 3: GLINTS (butuh render JS -> pakai Selenium)
# =========================================================
def scrape_glints(query=QUERY, max_pages=2):
    """
    Glints render kontennya lewat JavaScript, jadi requests biasa
    biasanya cuma dapat halaman kosong. Kita pakai Selenium headless.
    Kalau tidak mau install Selenium, fungsi ini bisa dilewati saja.
    """
    results = []
    skipped_not_relevant = 0
    skipped_date = 0

    driver, sel = create_chrome_driver("Glints")
    if driver is None:
        return results
    By, EC, WebDriverWait = sel

    try:
        for page in range(1, max_pages + 1):
            url = f"https://glints.com/id/opportunities/jobs/explore?keyword={query}&country=ID&page={page}"
            ok = load_page_with_retry(
                driver, url, "[class*='JobcardContainer'], body", sel, "Glints", page
            )
            if not ok:
                continue

            if looks_like_anti_bot(driver.page_source):
                log.warning("[Glints] Halaman %s kelihatan seperti proteksi anti-bot, berhenti.", page)
                break

            # PENTING: jangan pakai selector generik "[class*='JobCard']" -> itu
            # cocok dengan BEBERAPA wrapper bersarang punya nama sama untuk 1
            # lowongan yang sama (JobcardContainer, JobCardWrapper, CompactJobCard,
            # dst semua mengandung substring "JobCard"), jadi tiap lowongan
            # kehitung 3x. "JobcardContainer" adalah wrapper terluar yang unik.
            cards = driver.find_elements(By.CSS_SELECTOR, "[class*='JobcardContainer']")
            if not cards:
                log.info("[Glints] Halaman %s tidak ada job card, anggap sudah habis.", page)
                break

            for card in cards:
                # NB: Glints pakai styled-components dengan hash class acak (mis.
                # "CompactOpportunityCardsc__JobTitle-sc-dkg8my-11 kpFYNG"), jadi
                # kita match prefix nama komponennya (stabil), bukan hash-nya.
                # Diverifikasi langsung dari DOM asli per 2026-09-07 - kalau Glints
                # redesign, selector ini perlu diupdate lagi.
                # "JobCardTitleNoStyleAnchor" dipakai (bukan "JobTitle" polos) karena
                # "JobTitle" juga match wrapper "JobTitleSalaryWrapper" yang teksnya
                # ikut kebawa info gaji.
                try:
                    title_el = card.find_element(By.CSS_SELECTOR, "[class*='JobCardTitleNoStyleAnchor']")
                    title = title_el.text
                except Exception:
                    title_el = None
                    title = card.text.split("\n")[0] if card.text else ""

                company = ""
                for css_sel in ["[class*='CompanyLinkResolver']", "[class*='CompanyDetailContainer']"]:
                    try:
                        company = card.find_element(By.CSS_SELECTOR, css_sel).text
                        if company:
                            break
                    except Exception:
                        continue

                location = ""
                for css_sel in ["[class*='CardJobLocation']", "[class*='JobCardLocationNoStyleAnchor']"]:
                    try:
                        location = card.find_element(By.CSS_SELECTOR, css_sel).text
                        if location:
                            break
                    except Exception:
                        continue

                description = card.text  # fallback: seluruh teks card, dipakai buat filter tambahan

                posted_text = ""
                for css_sel in ["[class*='UpdatedTimeContainer']", "[class*='OpportunityMeta']"]:
                    try:
                        posted_text = card.find_element(By.CSS_SELECTOR, css_sel).text
                        if posted_text:
                            break
                    except Exception:
                        continue
                posted_date = parse_posted_date(text=posted_text or description)

                if not title or not is_relevant(title, description):
                    skipped_not_relevant += 1
                    continue
                if not is_within_target_period(posted_date):
                    skipped_date += 1
                    continue

                try:
                    link = title_el.get_attribute("href") if title_el else card.find_element(By.TAG_NAME, "a").get_attribute("href")
                except Exception:
                    link = url
                results.append({
                    "sumber": "Glints",
                    "posisi": title,
                    "perusahaan": company,
                    "lokasi": location,
                    "tipe_kerja": detect_work_type(title, description, location),
                    "tanggal_posting": posted_date.isoformat() if posted_date else "",
                    "link": link,
                })

            time.sleep(2)  # jangan spam request antar halaman
    except Exception as e:
        log.error("[Glints] Error saat scraping: %s", e)
    finally:
        driver.quit()

    log.info(
        "[Glints] q=%r: %s cocok, %s di-skip (bukan bidang IT/security), %s di-skip (di luar rentang tanggal)",
        query, len(results), skipped_not_relevant, skipped_date,
    )
    return results


# =========================================================
# SUMBER 4: JOBSTREET / SEEK ID (butuh render JS + tembus Cloudflare -> Selenium)
# =========================================================
def scrape_jobstreet(query=QUERY, max_pages=2):
    """
    Jobstreet Indonesia (bagian dari SEEK) full-JS render dan halamannya
    ditutupi Cloudflare -> request polos (`requests`) selalu dibalas 403
    "Just a moment..." apapun header-nya (sudah dicoba & dikonfirmasi).
    Jadi sumber ini, sama seperti Glints, pakai Selenium headless.
    Selector data-automation di bawah sudah diverifikasi langsung dari DOM
    render per 2026-09-07 - kalau Jobstreet redesign, update di sini.
    """
    results = []
    skipped_not_relevant = 0
    skipped_date = 0

    driver, sel = create_chrome_driver("Jobstreet")
    if driver is None:
        return results
    By, EC, WebDriverWait = sel

    try:
        for page in range(1, max_pages + 1):
            url = f"https://id.jobstreet.com/id/{query.replace(' ', '-')}-jobs?page={page}"
            ok = load_page_with_retry(driver, url, "article, body", sel, "Jobstreet", page, timeout=20)
            if not ok:
                continue

            if looks_like_anti_bot(driver.page_source):
                log.warning("[Jobstreet] Halaman %s kelihatan seperti proteksi anti-bot, berhenti.", page)
                dump_raw(f"jobstreet_page{page}.html", driver.page_source)
                break

            cards = driver.find_elements(By.CSS_SELECTOR, "article")
            if not cards:
                log.info("[Jobstreet] Halaman %s tidak ada job card, anggap sudah habis.", page)
                break

            for card in cards:
                try:
                    title_el = card.find_element(By.CSS_SELECTOR, "[data-automation='jobTitle']")
                    title = title_el.text
                    link = title_el.get_attribute("href")
                except Exception:
                    continue  # tanpa judul/link, entri ini tidak berguna

                company = ""
                try:
                    company = card.find_element(By.CSS_SELECTOR, "[data-automation='jobCompany']").text
                except Exception:
                    pass

                location = ""
                try:
                    location = card.find_element(By.CSS_SELECTOR, "[data-automation='jobCardLocation']").text
                except Exception:
                    pass

                posted_text = ""
                try:
                    posted_text = card.find_element(By.CSS_SELECTOR, "[data-automation='jobListingDate']").text
                except Exception:
                    pass

                description = card.text  # fallback: seluruh teks card, dipakai buat filter tambahan
                posted_date = parse_posted_date(text=posted_text or description)

                if not title or not is_relevant(title, description):
                    skipped_not_relevant += 1
                    continue
                if not is_within_target_period(posted_date):
                    skipped_date += 1
                    continue

                results.append({
                    "sumber": "Jobstreet",
                    "posisi": title,
                    "perusahaan": company,
                    "lokasi": location,
                    "tipe_kerja": detect_work_type(title, description, location),
                    "tanggal_posting": posted_date.isoformat() if posted_date else "",
                    "link": link,
                })

            time.sleep(2)  # jangan spam request antar halaman
    except Exception as e:
        log.error("[Jobstreet] Error saat scraping: %s", e)
    finally:
        driver.quit()

    log.info(
        "[Jobstreet] q=%r: %s cocok, %s di-skip (bukan bidang IT/security), %s di-skip (di luar rentang tanggal)",
        query, len(results), skipped_not_relevant, skipped_date,
    )
    return results


# =========================================================
# MAIN
# =========================================================
def main():
    all_results = []

    log.info("=== Mulai scraping info magang IT/Software/Data ===")
    log.info("Queries: %r | max_pages per query: %s", QUERIES, MAX_PAGES)
    log.info(
        "Filter tanggal posting: bulan %s-%s tahun %s (unknown date %s)",
        MONTH_START, MONTH_END, DATE_FILTER_YEAR,
        "disertakan" if INCLUDE_UNKNOWN_DATE else "dibuang",
    )

    session = build_session()

    # Tiap sumber dijalankan sekali PER query di QUERIES, hasilnya digabung &
    # di-dedup di akhir. Ini penting: KEYWORDS cuma memfilter hasil yang situs
    # balikin, jadi supaya lowongan cybersecurity/data juga ketemu, query
    # pencariannya sendiri harus relevan (bukan cuma "magang it").
    jobs = []
    for query in QUERIES:
        jobs.append((f"Kalibrr [{query}]", lambda q=query: scrape_kalibrr(session, query=q, max_pages=MAX_PAGES)))
        jobs.append((f"LinkedIn [{query}]", lambda q=query: scrape_linkedin(session, query=q)))
        jobs.append((f"Glints [{query}]", lambda q=query: scrape_glints(query=q, max_pages=MAX_PAGES)))
        jobs.append((f"Jobstreet [{query}]", lambda q=query: scrape_jobstreet(query=q, max_pages=MAX_PAGES)))

    for i, (name, func) in enumerate(jobs, start=1):
        log.info("[%s/%s] Mengambil data dari %s...", i, len(jobs), name)
        try:
            data = func()
            log.info("[%s] Dapat %s lowongan relevan.", name, len(data))
            all_results.extend(data)
        except Exception as e:
            # Satu sumber/query gagal tidak boleh menggagalkan yang lain.
            log.error("[%s] Gagal total, dilewati: %s", name, e)

    if not all_results:
        log.warning(
            "Tidak ada hasil yang berhasil diambil. "
            "Kemungkinan struktur situs berubah, proteksi anti-bot aktif, atau koneksi bermasalah."
        )
        return

    df = pd.DataFrame(all_results)
    # Normalisasi supaya dedup tidak kecolongan gara-gara whitespace/case beda,
    # atau link yang sama tapi beda query-string tracking (utm_referrer, traceInfo, dst).
    if "link" in df.columns:
        df["link_key"] = df["link"].fillna("").str.split("?").str[0].str.strip().str.lower()
    for col in ["posisi", "perusahaan"]:
        if col in df.columns:
            df[col + "_key"] = df[col].fillna("").str.strip().str.lower()
    dedup_cols = [c for c in ["posisi_key", "perusahaan_key", "link_key"] if c in df.columns]
    df = df.drop_duplicates(subset=dedup_cols).drop(columns=dedup_cols)

    df.to_excel(OUTPUT_FILE, index=False)

    log.info("Selesai! %s lowongan magang relevan disimpan ke: %s", len(df), OUTPUT_FILE)


if __name__ == "__main__":
    main()
