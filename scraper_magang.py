"""
Scraper Info Magang - IT/Software/Data/Cybersecurity
======================================================
Ambil lowongan magang dari Kalibrr, LinkedIn, Glints, dan Jobstreet
(multi-query, lihat QUERIES), filter yang relevan ke IT/Software/Data/
Cybersecurity, simpan ke Excel. Detail & cara pakai ada di README.md.

Catatan penting:
- Glints & Jobstreet wajib Selenium (full-JS render + Cloudflare).
  Kalibrr & LinkedIn cukup pakai request biasa (endpoint publik).
- LinkedIn: User Agreement mereka melarang scraping otomatis. Endpoint
  guest yang dipakai di sini publik & tanpa login, tapi tetap ada risiko
  IP diblokir kalau volume terlalu besar/sering - jangan naikkan
  MAGANG_LINKEDIN_MAX_PAGES sembarangan.
- Situs bisa berubah struktur kapan saja -> kalau satu sumber tiba-tiba
  0 hasil, cek ulang selector-nya lewat Inspect Element / Network tab.

Install: pip install -r requirements.txt
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
# KONFIGURASI (semua bisa dioverride lewat environment variable)
# =========================================================
_DEFAULT_QUERIES = ["magang it", "magang cyber security", "magang data"]

if os.environ.get("MAGANG_QUERIES", "").strip():
    QUERIES = [q.strip() for q in os.environ["MAGANG_QUERIES"].split(",") if q.strip()]
elif os.environ.get("MAGANG_QUERY", "").strip():
    QUERIES = [os.environ["MAGANG_QUERY"].strip()]
else:
    QUERIES = _DEFAULT_QUERIES

QUERY = QUERIES[0]  # dipakai sebagai default param tunggal di beberapa fungsi
MAX_PAGES = int(os.environ.get("MAGANG_MAX_PAGES", "3"))
RAW_DUMP_DIR = os.environ.get("MAGANG_RAW_DIR", "")

# Filter tanggal posting (default Agustus-September tahun berjalan)
DATE_FILTER_YEAR = int(os.environ.get("MAGANG_YEAR", str(datetime.now().year)))
MONTH_START = int(os.environ.get("MAGANG_MONTH_START", "8"))
MONTH_END = int(os.environ.get("MAGANG_MONTH_END", "9"))
INCLUDE_UNKNOWN_DATE = os.environ.get("MAGANG_INCLUDE_UNKNOWN_DATE", "false").strip().lower() == "true"

# Keyword filter IT/Software/Data/Cybersecurity
KEYWORDS = [
    "it", "informatika", "software", "developer", "programmer",
    "data", "data science", "data analyst", "machine learning",
    "artificial intelligence", "ai", "backend", "frontend", "fullstack",
    "full stack", "web developer", "mobile developer", "sistem informasi",
    "teknik komputer", "cyber security", "devops", "cloud", "qa", "quality assurance",
    "database", "sql", "python", "java", "javascript",
    # Cybersecurity / Security Operations
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

# Keyword pendek/ambigu -> wajib whole-word match & cuma dicek di judul (lihat is_relevant)
_STRICT_WORD_KEYWORDS = {
    "it", "ai", "qa", "sql", "cloud", "data", "cyber", "soc", "database",
    "l1", "l2", "l3", "iam", "ctf",
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

# Penanda halaman anti-bot/captcha, bukan hasil pencarian asli
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
        backoff_factor=1.5,
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
        with open(os.path.join(RAW_DUMP_DIR, name), "w", encoding="utf-8") as f:
            f.write(content)
    except Exception as e:
        log.warning("Gagal simpan raw dump %s: %s", name, e)


# Keyword deteksi tipe kerja. Hybrid dicek duluan karena kalimat "hybrid
# (WFH & WFO)" bisa kepental jadi Remote gara-gara ada kata "wfh" juga.
REMOTE_MARKERS = ["remote", "wfh", "work from home", "fully remote", "kerja dari rumah"]
HYBRID_MARKERS = ["hybrid", "wfh & wfo", "wfh/wfo", "semi remote", "campuran"]
ONSITE_MARKERS = ["on-site", "onsite", "wfo", "work from office", "di kantor"]

_keyword_patterns = None


def _get_keyword_patterns():
    """Compile regex sekali saja, dipisah 2 grup: generic (keyword pendek/ambigu,
    whole-word, cuma dicek di judul) dan specific (frasa, dicek di judul+deskripsi)."""
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
    """Cek relevansi ke IT/Software/Data/Cybersecurity.

    Keyword generik (it, data, cloud, dst) cuma dicek di judul - kalau ikut
    dicek di deskripsi, hampir semua lowongan (admin, HR, gudang, dst) ikut
    lolos karena mereka juga sering menyebut "data"/"IT"/"cloud" secara
    insidental. Keyword spesifik/frasa (data science, pentester, dst) tetap
    dicek di judul+deskripsi karena jarang muncul kebetulan.
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
    atau "Tidak diketahui" kalau tidak ada penanda."""
    combined = " ".join(t for t in texts if t).lower()
    if not combined:
        return "Tidak diketahui"

    def has_any(markers):
        return any(re.search(r"\b" + re.escape(m) + r"\b", combined) for m in markers)

    if has_any(HYBRID_MARKERS):
        return "Hybrid"
    if has_any(REMOTE_MARKERS):
        return "Remote"
    if has_any(ONSITE_MARKERS):
        return "On-site"
    return "Tidak diketahui"


# =========================================================
# FILTER TANGGAL POSTING
# =========================================================
_RELATIVE_DATE_PATTERNS = [
    (re.compile(r"(\d+)\s*(?:hari|day)s?\s*(?:yang lalu|ago|lalu)", re.I), "days"),
    (re.compile(r"(\d+)\s*(?:jam|hour)s?\s*(?:yang lalu|ago|lalu)", re.I), "hours"),
    (re.compile(r"(\d+)\s*(?:minggu|week)s?\s*(?:yang lalu|ago|lalu)", re.I), "weeks"),
    (re.compile(r"(\d+)\s*(?:bulan|month)s?\s*(?:yang lalu|ago|lalu)", re.I), "months"),
]


def parse_posted_date(text: str = "", iso_hint: str = "") -> "date | None":
    """Tentukan tanggal posting dari `iso_hint` (tanggal absolut dari API/HTML)
    atau `text` (relatif, mis. "3 hari yang lalu"). None kalau gagal dideteksi."""
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
                return date.today()
            if unit == "weeks":
                return date.today() - timedelta(weeks=n)
            if unit == "months":
                return date.today() - timedelta(days=30 * n)

    return None


def is_within_target_period(posted: "date | None") -> bool:
    """Cek apakah tanggal posting ada di rentang MONTH_START..MONTH_END/DATE_FILTER_YEAR.
    Tanggal tidak diketahui -> ikuti INCLUDE_UNKNOWN_DATE."""
    if posted is None:
        return INCLUDE_UNKNOWN_DATE
    return posted.year == DATE_FILTER_YEAR and MONTH_START <= posted.month <= MONTH_END


# =========================================================
# SUMBER 1: KALIBRR (endpoint pencarian JSON)
# =========================================================
def scrape_kalibrr(session: requests.Session, query=QUERY, max_pages=MAX_PAGES):
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
            "limit": limit,          # wajib, kalau tidak API balas 400
            "offset": page * limit,  # API pakai offset, bukan "page"
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
            company = job.get("company_name", "") or ""  # bukan job["company"]["name"] - field itu tidak ada
            url = f"https://www.kalibrr.com/c/{job.get('company', {}).get('code','')}/jobs/{job.get('id','')}"
            posted_raw = job.get("created_at") or job.get("activation_date") or ""
            lokasi = job.get("google_location", {}).get("address_components", {}).get("city", "")
            posted_date = parse_posted_date(iso_hint=posted_raw)

            if not is_relevant(title, description):
                skipped_not_relevant += 1
                continue
            if not is_within_target_period(posted_date):
                skipped_date += 1
                continue

            # Kalibrr expose flag eksplisit is_hybrid/is_work_from_home -> lebih akurat
            if job.get("is_hybrid"):
                tipe_kerja = "Hybrid"
            elif job.get("is_work_from_home"):
                tipe_kerja = "Remote"
            else:
                tipe_kerja = detect_work_type(title, description, lokasi)
                if tipe_kerja == "Tidak diketahui":
                    tipe_kerja = "On-site"

            results.append({
                "sumber": "Kalibrr",
                "posisi": title,
                "perusahaan": company,
                "lokasi": lokasi,
                "tipe_kerja": tipe_kerja,
                "tanggal_posting": posted_date.isoformat() if posted_date else posted_raw,
                "link": url,
            })
        time.sleep(1.5)

    log.info(
        "[Kalibrr] q=%r: %s cocok, %s di-skip (bukan bidang IT/security), %s di-skip (di luar rentang tanggal)",
        query, len(results), skipped_not_relevant, skipped_date,
    )
    return results


# =========================================================
# SUMBER 2: LINKEDIN (guest job search endpoint, publik tanpa login)
# =========================================================
# PENTING: User Agreement LinkedIn melarang scraping otomatis (beda dengan
# sumber lain yang cuma proteksi teknis). Endpoint ini publik & datanya
# publik, tapi tetap berisiko IP diblokir kalau dipakai agresif - jangan
# naikkan MAGANG_LINKEDIN_MAX_PAGES sembarangan.
LINKEDIN_MAX_PAGES = int(os.environ.get("MAGANG_LINKEDIN_MAX_PAGES", "2"))
LINKEDIN_PAGE_SIZE = 10  # tetap dari API-nya

def scrape_linkedin(session: requests.Session, query=QUERY, max_pages=None):
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

        time.sleep(3)  # LinkedIn lebih sensitif thd rate-limit -> jeda lebih panjang

    log.info(
        "[LinkedIn] q=%r: %s cocok, %s di-skip (bukan bidang IT/security), %s di-skip (di luar rentang tanggal)",
        query, len(results), skipped_not_relevant, skipped_date,
    )
    return results


# =========================================================
# HELPER SELENIUM (dipakai bareng Glints & Jobstreet)
# =========================================================
def _import_selenium():
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait
    from webdriver_manager.chrome import ChromeDriverManager
    return webdriver, Options, Service, By, EC, WebDriverWait, ChromeDriverManager


def create_chrome_driver(source_name: str):
    """Bikin Chrome headless driver siap pakai, atau (None, None) kalau gagal."""
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

    # Cache lokasi chromedriver supaya tidak hit network tiap run
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
    """driver.get(url) + tunggu elemen muncul, retry `tries` kali. Return True/False."""
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
# SUMBER 3: GLINTS (butuh render JS -> Selenium)
# =========================================================
def scrape_glints(query=QUERY, max_pages=2):
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

            # "[class*='JobCard']" match beberapa wrapper bersarang sekaligus (duplikat!)
            # -> pakai "JobcardContainer", wrapper terluar yang unik per lowongan.
            cards = driver.find_elements(By.CSS_SELECTOR, "[class*='JobcardContainer']")
            if not cards:
                log.info("[Glints] Halaman %s tidak ada job card, anggap sudah habis.", page)
                break

            for card in cards:
                # Class Glints pakai styled-components (hash acak) - match prefix
                # nama komponennya. "JobCardTitleNoStyleAnchor" dipakai (bukan
                # "JobTitle" polos) karena "JobTitle" juga match wrapper yang
                # teksnya kebawa info gaji.
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

                description = card.text

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

            time.sleep(2)
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
# SUMBER 4: JOBSTREET / SEEK ID (full-JS + Cloudflare -> Selenium)
# =========================================================
def scrape_jobstreet(query=QUERY, max_pages=2):
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

                description = card.text
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

            time.sleep(2)
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

    # Tiap sumber dijalankan per query di QUERIES, hasilnya digabung & di-dedup
    # di akhir - supaya lowongan cybersecurity/data juga ketemu (bukan cuma "magang it").
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
            log.error("[%s] Gagal total, dilewati: %s", name, e)

    if not all_results:
        log.warning(
            "Tidak ada hasil yang berhasil diambil. "
            "Kemungkinan struktur situs berubah, proteksi anti-bot aktif, atau koneksi bermasalah."
        )
        return

    df = pd.DataFrame(all_results)
    # Normalisasi biar dedup tidak kecolongan whitespace/case beda atau
    # link sama tapi beda query-string tracking (utm_referrer, traceInfo, dst)
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
