# Scraper Info Magang IT/Software/Data/Cybersecurity

Script Python buat ngumpulin info lowongan magang dari beberapa job portal (Kalibrr, LinkedIn, Glints, Jobstreet), filter yang relevan ke bidang IT/Software/Data/Cybersecurity, lalu simpan ke Excel.

## Fitur

- **4 sumber**: Kalibrr, LinkedIn, Glints, Jobstreet.
- **Multi-query search** — bisa cari beberapa kata kunci sekaligus (default: `magang it`, `magang cyber security`, `magang data`), hasil digabung & di-dedup otomatis.
- **Filter kata kunci luas**: IT/software/data science/data analyst + cybersecurity (SOC, pentester, red/blue team, dst) — dicek dengan whole-word matching biar tidak salah tangkap kata umum.
- **Filter tanggal posting** — default cuma ambil lowongan yang di-posting Agustus-September tahun berjalan, biar tidak kebanjiran postingan lama.
- **Deteksi tipe kerja** (Remote/Hybrid/On-site) dari teks judul/deskripsi/lokasi (Kalibrr malah pakai flag eksplisit dari API-nya, jadi paling akurat).
- Retry otomatis untuk error jaringan transient, deteksi halaman anti-bot/Cloudflare, logging jelas per sumber.

## Instalasi

```bash
pip install -r requirements.txt
```

Glints & Jobstreet butuh render JavaScript (full-JS + proteksi Cloudflare), jadi wajib punya **Google Chrome** terinstall — `chromedriver`-nya otomatis didownload lewat `webdriver-manager`, tidak perlu setup manual.

## Cara pakai

```bash
python scraper_magang.py
```

Hasilnya disimpan sebagai `hasil_magang_YYYYMMDD_HHMM.xlsx` di folder yang sama.

### Konfigurasi (environment variable)

| Variabel | Default | Keterangan |
|---|---|---|
| `MAGANG_QUERIES` | `magang it,magang cyber security,magang data` | Daftar query pencarian, dipisah koma. Override penuh daftar default. |
| `MAGANG_QUERY` | - | Alternatif single-query (dipakai kalau `MAGANG_QUERIES` tidak di-set). |
| `MAGANG_MAX_PAGES` | `3` | Jumlah halaman per query untuk Kalibrr/Glints/Jobstreet. |
| `MAGANG_LINKEDIN_MAX_PAGES` | `2` | Jumlah halaman khusus LinkedIn (dibuat lebih kecil, lihat catatan ToS di bawah). |
| `MAGANG_MONTH_START` / `MAGANG_MONTH_END` | `8` / `9` | Rentang bulan filter tanggal posting (Agustus-September). |
| `MAGANG_YEAR` | tahun berjalan | Tahun untuk filter tanggal posting. |
| `MAGANG_INCLUDE_UNKNOWN_DATE` | `false` | `true` kalau mau tetap menyertakan lowongan yang tanggal postingnya tidak berhasil dideteksi. |
| `MAGANG_RAW_DIR` | - | Isi path folder kalau mau simpan raw response buat debugging (misal selector situs berubah). |

Contoh (PowerShell):

```powershell
$env:MAGANG_QUERIES = "magang backend,magang data analyst"
$env:MAGANG_MAX_PAGES = "5"
python scraper_magang.py
```

## ⚠️ Catatan penting

1. **Struktur situs bisa berubah kapan saja.** Kalibrr/Glints/Jobstreet/LinkedIn sering update halaman/API/proteksi anti-bot mereka. Kalau script tiba-tiba dapat 0 hasil dari satu sumber, kemungkinan besar selector-nya perlu di-update (buka Inspect Element / Network tab di browser buat cek struktur terbaru).
2. **LinkedIn — perhatikan ToS.** User Agreement LinkedIn eksplisit melarang scraping otomatis (beda dengan 3 sumber lain yang cuma proteksi teknis). Endpoint yang dipakai di sini publik & tanpa login, tapi tetap ada risiko IP kamu diblokir sementara kalau dipakai dengan volume besar/sering. Default `MAGANG_LINKEDIN_MAX_PAGES` sengaja dibuat kecil — jangan dinaikkan sembarangan, dan jangan jalankan script berkali-kali dalam waktu singkat.
3. **Gunakan dengan wajar.** Semua sumber punya rate-limiting bawaan di script ini (delay antar request/halaman), tapi tetap jangan dijalankan terlalu sering supaya IP tidak diblokir platform manapun.
4. Jalankan di komputer/laptop sendiri yang punya akses internet normal ke situs-situs tersebut.

## Struktur output

Kolom di file Excel hasil scraping:

| Kolom | Keterangan |
|---|---|
| `sumber` | Kalibrr / LinkedIn / Glints / Jobstreet |
| `posisi` | Judul lowongan |
| `perusahaan` | Nama perusahaan |
| `lokasi` | Kota/wilayah |
| `tipe_kerja` | Remote / Hybrid / On-site / Tidak diketahui |
| `tanggal_posting` | Tanggal posting (ISO `YYYY-MM-DD`, kalau berhasil dideteksi) |
| `link` | URL ke halaman lowongan |
