"""
Automasi PURE UI-CLICKING di Superset SQL Lab (fasih-dashboard.bps.go.id)

Beda sama versi sebelumnya (yang manggil API langsung), script ini betul-betul:
1. Nulis ulang query di editor (Ace editor)
2. Klik tombol RUN
3. Baca teks "X rows returned"
4. Kalau > 0 -> klik DOWNLOAD TO EXCEL, tunggu file kedownload, rename & simpen
5. Kalau == 0 -> kecamatan ini abis, pindah ke kecamatan berikutnya (auto-increment +10)
6. Ulangi sampai semua kecamatan di list selesai

INI MENIRU PERSIS ALUR MANUAL LU DI BROWSER -- jadi paling "aman" dari sisi pola
traffic (keliatan kayak orang beneran klak-klik), tapi juga paling lambat & paling
rawan patah kalau ada perubahan kecil di tampilan dashboard.

WAJIB SEBELUM JALANIN:
1. Buka Chrome manual dengan remote debugging aktif:
   chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\\selenium\\chrome_profile"
2. Nyalain VPN, login ke fasih-dashboard, buka tab SQL Lab yang query-nya udah
   ready (schema tgr_fd68e454 kepilih).
3. Cek SELECTORS di bawah -- kalau ada tombol yang script gagal nemuin, klik kanan
   di elemen itu di browser -> Inspect -> sesuaikan.
4. Isi KECAMATAN_CODES (idealnya dari wilayah.json lu, JANGAN cuma pure +10 math
   kalau lu belum yakin semua 25 kecamatan Banyuwangi kodenya beneran +10 rapi).
"""

import os
import re
import json
import time
import glob
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ---------------- CONFIG ----------------
DEBUGGER_ADDRESS = "127.0.0.1:9222"
SQLLAB_URL = "https://fasih-dashboard.bps.go.id/superset/sqllab/"
LIMIT = 9000

SLEEP_MIN, SLEEP_MAX = 6, 12
SLEEP_MIN_KEC, SLEEP_MAX_KEC = 20, 40
REQUESTS_BEFORE_LONG_BREAK = 15
LONG_BREAK_MIN, LONG_BREAK_MAX = 120, 240

DOWNLOAD_DIR = os.path.abspath("output/downloads")
FINAL_DIR = os.path.abspath("output/final")
CHECKPOINT_FILE = "output/checkpoint_ui.json"

# List eksplisit 25 kecamatan Banyuwangi -- BUKAN pola +10 murni, ada hasil pemekaran
# (011, 071, 121, 171) yang nyempil di antara kode lain. Diambil langsung dari data
# resmi, jangan digenerate otomatis.
KECAMATAN = [
    ("010", "PESANGGARAN"),
    ("011", "SILIRAGUNG"),
    ("020", "BANGOREJO"),
    ("030", "PURWOHARJO"),
    ("040", "TEGALDLIMO"),
    ("050", "MUNCAR"),
    ("060", "CLURING"),
    ("070", "GAMBIRAN"),
    ("071", "TEGALSARI"),
    ("080", "GLENMORE"),
    ("090", "KALIBARU"),
    ("100", "GENTENG"),
    ("110", "SRONO"),
    ("120", "ROGOJAMPI"),
    ("121", "BLIMBINGSARI"),
    ("130", "KABAT"),
    ("140", "SINGOJURUH"),
    ("150", "SEMPU"),
    ("160", "SONGGON"),
    ("170", "GLAGAH"),
    ("171", "LICIN"),
    ("180", "BANYUWANGI"),
    ("190", "GIRI"),
    ("200", "KALIPURO"),
    ("210", "WONGSOREJO"),
]
KECAMATAN_CODES = [code for code, _ in KECAMATAN]

# Selector elemen -- GANTI kalau ternyata beda, cek via Inspect Element
SELECTORS = {
    "editor_textarea": "textarea.ace_text-input",
    "run_button_xpath": "//button[.//span[normalize-space()='Run'] or contains(., 'Run')]",
    "rows_returned_xpath": "//*[contains(text(), 'rows returned')]",
    "download_excel_xpath": (
        "//button[contains(translate(., "
        "'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), "
        "'DOWNLOAD TO EXCEL')]"
    ),
}

QUERY_TEMPLATE = """
SELECT
n.assignment_id, r.level_6_full_code, r.assignment_status_alias,
n.keberadaan_usaha_value, r.no_bang, n.nama_usaha, n.nama_usaha_edit,
n.nama_komersial, n.alamat_usaha_view, n.badan_usaha_value, n.keg_utama,
n.produk_sendiri_value, n.layanan_mamin_value, n.keg_penjualan_value,
n.keg_jasa_value, n.lokasi_usaha_value, n.input, n.proses, n.produk,
n.klasifikasi_value, n.kbli_genai_value, n.kbli_value, n.kategori,
r.geotag_latitude, r.geotag_longitude
FROM tgr_fd68e454.se2026_nested n
INNER JOIN tgr_fd68e454.root_table r ON n.assignment_id = r.assignment_id
INNER JOIN tgr_fd68e454.base_table_assignment a ON n.assignment_id = a.assignment_id
WHERE n.keberadaan_usaha_value IN (1, 2)
AND n.kbli_genai_value IS NOT NULL
AND a.is_active = TRUE
AND r.level_6_full_code LIKE '{kec_prefix}%'
ORDER BY r.level_6_full_code, r.no_bang, n.nama_usaha
LIMIT {limit} OFFSET {offset};
""".strip()


def print_banner(text: str, char: str = "="):
    width = max(50, min(78, len(text) + 8))
    line = char * width
    print(f"\n{line}")
    print(f"  {text}")
    print(f"{line}\n")
def load_checkpoint() -> dict:
    if os.path.exists(CHECKPOINT_FILE):
        with open(CHECKPOINT_FILE) as f:
            return json.load(f)
    return {"completed_kecamatan": [], "in_progress": None}


def save_checkpoint(cp: dict):
    with open(CHECKPOINT_FILE, "w") as f:
        json.dump(cp, f, indent=2)


# ---------------- BROWSER SETUP ----------------
def connect_browser():
    options = webdriver.ChromeOptions()
    options.debugger_address = DEBUGGER_ADDRESS
    driver = webdriver.Chrome(options=options)

    driver.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": DOWNLOAD_DIR,
    })

    # window sempit bikin Superset nge-collapse tombol Download ke menu "..."
    # (overflow) -- maximize dulu biar semua tombol tampil normal
    try:
        driver.maximize_window()
    except Exception as e:
        print(f"debug: gagal maximize window ({e}), lanjut aja")

    switch_to_sqllab_tab(driver)
    set_row_limit_dropdown(driver, target_text="10 000")
    return driver


def switch_to_sqllab_tab(driver):
    """Selenium attach ke Chrome yang punya banyak tab -- defaultnya dia aktif
    di tab pertama, BUKAN otomatis ke tab yang lagi keliatan di layar. Ini cari
    tab yang URL-nya mengandung 'sqllab' dan pindah ke situ."""
    found = False
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        print(f"cek tab: {driver.current_url}")
        if "sqllab" in driver.current_url:
            found = True
            print(f"-> ketemu tab SQL Lab, switch ke sini.")
            break

    if not found:
        raise RuntimeError(
            "Ga ketemu tab dengan URL mengandung 'sqllab' di antara semua tab yang "
            "kebuka. Pastiin tab fasih-dashboard.bps.go.id/superset/sqllab beneran "
            "kebuka di Chrome yang di-attach ini."
        )


# ---------------- EDITOR ----------------
def set_query_text(driver, sql_text: str):
    """PENTING: sebelumnya pakai send_keys per baris + Enter -- ini BERMASALAH
    karena SQL Lab punya autocomplete popup yang suka nyaplok Keys.ENTER
    (harusnya bikin newline, malah milih suggestion), bikin query kepotong
    di tengah. Solusi lebih reliable: akses instance Ace Editor langsung lewat
    property standar '.env.editor' yang nempel di elemen div.ace_editor, terus
    panggil setValue() -- ini bypass sepenuhnya masalah autocomplete/keyboard."""
    containers = [el for el in driver.find_elements(By.CSS_SELECTOR, "div.ace_editor") if el.is_displayed()]
    print(f"debug: ketemu {len(containers)} div.ace_editor yang visible")

    if not containers:
        raise RuntimeError(
            "Ga nemu div.ace_editor yang visible. Cek lagi via Inspect Element."
        )

    container = containers[0]

    success = driver.execute_script("""
        const el = arguments[0];
        const text = arguments[1];
        if (!el.env || !el.env.editor) { return false; }
        el.env.editor.setValue(text, -1);
        el.env.editor.clearSelection();
        return true;
    """, container, sql_text)

    print(f"debug: set via ace .env.editor.setValue() -> {'sukses' if success else 'GAGAL'}")

    if not success:
        raise RuntimeError(
            "Elemen div.ace_editor ketemu tapi ga punya property .env.editor -- "
            "kemungkinan versi Ace/Superset beda dari yang diasumsikan. Kasih tau "
            "gue, kita cari cara lain."
        )

    time.sleep(0.3)


def set_row_limit_dropdown(driver, target_text="10 000", max_retries=3):
    """UI Superset punya dropdown 'LIMIT' terpisah dari LIMIT di SQL kita --
    defaultnya suka ke-set ke 1000 dan itu yang menang (lihat warning kuning
    'The number of rows displayed is limited to X by the dropdown'). Set ke
    angka >= LIMIT config kita (9000) biar data ga ke-cap diam-diam.

    Ada retry karena abis reload halaman, komponennya kadang belum sepenuhnya
    'hidup' (event listener React belum ke-attach) pas langsung diklik."""
    for attempt in range(1, max_retries + 1):
        try:
            trigger = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(@class,'limitDropdown')]]"))
            )

            current_value = trigger.text.strip().replace("LIMIT:", "").strip()
            if current_value == target_text:
                print(f"debug: dropdown LIMIT udah '{target_text}', skip.")
                return

            trigger.click()
            time.sleep(0.5)

            option = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, f"//*[self::li or self::div][normalize-space()='{target_text}']")
                )
            )
            option.click()
            time.sleep(0.5)
            print(f"debug: row limit dropdown di-set ke '{target_text}' (percobaan {attempt})")
            return
        except TimeoutException:
            print(f"debug: percobaan {attempt}/{max_retries} set dropdown LIMIT gagal, retry...")
            time.sleep(2)

    raise RuntimeError(f"Ga berhasil set dropdown LIMIT ke '{target_text}' setelah {max_retries}x percobaan.")


def check_row_limit_warning(driver) -> bool:
    """Cek apakah warning 'limited to X by the dropdown' masih muncul --
    kalau iya, berarti dropdown ke-reset (misal abis ganti tab) dan harus
    di-set ulang sebelum lanjut, biar ga ada data ke-skip diam-diam."""
    elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'limited to') and contains(text(), 'dropdown')]")
    return len(elements) > 0 and any(e.is_displayed() for e in elements)


def reload_sqllab_after_download(driver):
    """Klik 'Download to Excel' di Superset ini FULL PAGE NAVIGATION ke endpoint
    file-nya, bukan background download -- abis download, tab kita kedampar di
    halaman kosong isinya file XLSX doang. Reload balik ke SQL Lab (action yang
    natural, kayak refresh browser biasa -- bukan hit API mentah), tunggu editor
    siap lagi, terus set ulang dropdown LIMIT (reset tiap reload)."""
    print("debug: reload balik ke SQL Lab setelah download...")
    driver.get(SQLLAB_URL)

    WebDriverWait(driver, 30).until(
        lambda d: len([el for el in d.find_elements(By.CSS_SELECTOR, "div.ace_editor") if el.is_displayed()]) > 0
    )
    time.sleep(1.5)  # kasih waktu React bener-bener selesai hydrate sebelum interaksi
    set_row_limit_dropdown(driver, target_text="10 000")
    time.sleep(1)


def click_run(driver):
    btn = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, SELECTORS["run_button_xpath"]))
    )
    btn.click()
    btn = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, SELECTORS["run_button_xpath"]))
    )
    btn.click()


def get_rows_returned(driver, timeout=90) -> int:
    """Nunggu teks 'X rows returned' ATAU alert 'The query returned no data' muncul."""
    def _read(d):
        # 1. Cek alert kuning "The query returned no data" (kalau data abis)
        try:
            no_data_alert = d.find_elements(By.XPATH, "//*[contains(text(), 'The query returned no data')]")
            if no_data_alert and any(alert.is_displayed() for alert in no_data_alert):
                return 0  # Langsung balikin 0, artinya data kecamatan ini selesai
        except Exception:
            pass

        # 2. Cek teks normal "X rows returned" (kalau data masih ada)
        try:
            el = d.find_element(By.XPATH, SELECTORS["rows_returned_xpath"])
            match = re.search(r"([\d,]+)\s+rows returned", el.text)
            if match:
                return int(match.group(1).replace(",", ""))
        except Exception:
            return None
        return None

    try:
        # Tunggu sampai salah satu dari 2 kondisi di atas ketemu
        result = WebDriverWait(driver, timeout).until(lambda d: _read(d) is not None)
        return _read(driver)
    except TimeoutException:
        raise TimeoutError("Nunggu hasil query kelamaan, ga muncul info baris atau warning no data.")


def click_download_excel(driver):
    """Cari elemen (button/li/div/a) yang teksnya mengandung 'download to excel'.
    Kalau nggak ketemu di percobaan pertama, kemungkinan tombolnya ke-collapse ke
    menu overflow '...' -- coba klik trigger overflow itu dulu, baru cari lagi."""
    def _find_button(d):
        xpath = "//button | //li[@role='menuitem'] | //div[@role='menuitem'] | //a"
        for el in d.find_elements(By.XPATH, xpath):
            try:
                if "download to excel" in el.text.strip().lower() and el.is_displayed():
                    return el
            except Exception:
                continue
        return None

    def _try_open_overflow_menu(d):
        """Cari tombol icon-only (biasanya 'more'/'...' -- Ant Design sering pakai
        aria-label 'more' atau class mengandung 'ellipsis') dan klik buat buka."""
        candidates = d.find_elements(
            By.XPATH,
            "//button[not(normalize-space(text()))][.//span[contains(@class,'ellipsis') "
            "or contains(@class,'more')]] | //button[@aria-label='More'] | "
            "//button[contains(@class,'ant-dropdown-trigger')][not(.//span[contains(@class,'limitDropdown')])]"
        )
        for c in candidates:
            try:
                if c.is_displayed():
                    c.click()
                    time.sleep(0.5)
                    return True
            except Exception:
                continue
        return False

    btn = None
    try:
        btn = WebDriverWait(driver, 15).until(lambda d: _find_button(d) or False)
    except TimeoutException:
        print("debug: belum ketemu, coba buka menu overflow '...' dulu")
        opened = _try_open_overflow_menu(driver)
        print(f"debug: overflow menu ke-klik? {opened}")
        if opened:
            try:
                btn = WebDriverWait(driver, 15).until(lambda d: _find_button(d) or False)
            except TimeoutException:
                btn = None

    if btn is None:
        all_texts = [b.text.strip() for b in driver.find_elements(By.TAG_NAME, "button") if b.text.strip()]
        print(f"debug: tombol yang ketemu di halaman: {all_texts}")
        raise RuntimeError("Ga nemu tombol 'Download to Excel' walau udah coba buka overflow menu.")

    btn.click()


def wait_for_new_download(before_files: set, timeout=60) -> str:
    """Nunggu file baru (bukan .crdownload) muncul di DOWNLOAD_DIR."""
    end = time.time() + timeout
    while time.time() < end:
        current = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*")))
        new_files = current - before_files
        finished = [f for f in new_files if not f.endswith(".crdownload")]
        if finished:
            return finished[0]
        time.sleep(1)
    raise TimeoutError("File download ga muncul dalam waktu yang ditentukan.")


# ---------------- MAIN LOOP ----------------
def scrape_kecamatan(driver, kec_code: str, kec_name: str, resume_offset: int, request_counter: list):
    kec_prefix = f"3510{kec_code}"
    offset = resume_offset
    total_rows = 0
    files_saved = 0

    while True:
        sql = QUERY_TEMPLATE.format(kec_prefix=kec_prefix, limit=LIMIT, offset=offset)
        print(f"  [{kec_code} {kec_name}] offset={offset} -> nulis query & run...")

        set_query_text(driver, sql)
        time.sleep(1)
        click_run(driver)

        try:
            n_rows = get_rows_returned(driver)
        except TimeoutError:
            print(f"  [{kec_code} {kec_name}] TIMEOUT nunggu hasil di offset={offset}. Stop, checkpoint disimpan.")
            return False, offset, total_rows, files_saved

        if check_row_limit_warning(driver):
            print(f"  [{kec_code} {kec_name}] STOP: dropdown LIMIT ke-reset (warning 'limited to X' muncul lagi) "
                  f"di offset={offset}. Set ulang dropdown-nya manual, jalanin ulang script buat resume.")
            return False, offset, total_rows, files_saved

        print(f"  [{kec_code} {kec_name}] {n_rows} rows returned")

        # kasih jeda ekstra biar tabel hasil (apalagi ribuan baris) beneran
        # selesai render sebelum nyari tombol download -- row count muncul
        # duluan, tapi tabel & tombolnya nyusul beberapa detik kemudian
        time.sleep(4)

        if n_rows == 0:
            print(f"  [{kec_code} {kec_name}] data abis di offset={offset}. total={total_rows} baris, {files_saved} file.")
            return True, offset, total_rows, files_saved

        before = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*")))
        click_download_excel(driver)
        try:
            downloaded = wait_for_new_download(before)
        except TimeoutError:
            print(f"  [{kec_code} {kec_name}] download gagal/timeout di offset={offset}. Stop, checkpoint disimpan.")
            return False, offset, total_rows, files_saved

        os.makedirs(FINAL_DIR, exist_ok=True)
        final_name = f"se2026_{kec_code}_{kec_name}_offset{offset}.xlsx"
        final_path = os.path.join(FINAL_DIR, final_name)
        os.replace(downloaded, final_path)
        print(f"  [{kec_code} {kec_name}] saved -> {final_name}")

        reload_sqllab_after_download(driver)

        rows_saved = n_rows  # dipakai buat nentuin offset berikutnya (bukan asumsi = LIMIT)
        total_rows += rows_saved
        files_saved += 1

        offset += rows_saved
        request_counter[0] += 1

        if rows_saved < LIMIT:
            print(f"  [{kec_code} {kec_name}] WARNING: rows_returned ({rows_saved}) < LIMIT config ({LIMIT}). "
                  f"Kemungkinan ke-cap dropdown 'LIMIT' di UI Superset -- cek manual, "
                  f"pastiin di-set >= {LIMIT} biar ga ada data kelewat.")

        if request_counter[0] % REQUESTS_BEFORE_LONG_BREAK == 0:
            long_break = random.uniform(LONG_BREAK_MIN, LONG_BREAK_MAX)
            print(f"  -- udah {request_counter[0]} query total, istirahat {long_break:.0f}s --")
            time.sleep(long_break)
        else:
            time.sleep(random.uniform(SLEEP_MIN, SLEEP_MAX))


def main():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    os.makedirs(FINAL_DIR, exist_ok=True)

    driver = connect_browser()
    checkpoint = load_checkpoint()
    completed = set(checkpoint["completed_kecamatan"])
    request_counter = [0]

    grand_total_rows = 0
    grand_total_files = 0
    t_start = time.time()

    print_banner(f"MULAI SCRAPING SE2026 -- {len(KECAMATAN)} kecamatan, "
                 f"{len(completed)} udah selesai sebelumnya", char="#")

    for kec_code, kec_name in KECAMATAN:
        if kec_code in completed:
            print(f"[SKIP] {kec_code} {kec_name} -- udah selesai (checkpoint)")
            continue

        resume_offset = 0
        if checkpoint.get("in_progress") and checkpoint["in_progress"]["kec_code"] == kec_code:
            resume_offset = checkpoint["in_progress"]["offset"]
            print_banner(f"LANJUT: {kec_code} {kec_name} (resume dari offset={resume_offset})")
        else:
            print_banner(f"MULAI KECAMATAN: {kec_code} {kec_name}")

        finished, last_offset, total_rows, files_saved = scrape_kecamatan(
            driver, kec_code, kec_name, resume_offset, request_counter
        )

        if finished:
            grand_total_rows += total_rows
            grand_total_files += files_saved
            completed.add(kec_code)
            checkpoint["completed_kecamatan"] = list(completed)
            checkpoint["in_progress"] = None
            save_checkpoint(checkpoint)

            elapsed_min = (time.time() - t_start) / 60
            print_banner(
                f"SELESAI: {kec_code} {kec_name}  |  {total_rows} baris, {files_saved} file  |  "
                f"progress {len(completed)}/{len(KECAMATAN)} kecamatan  |  {elapsed_min:.1f} menit berjalan",
                char="-",
            )
            time.sleep(random.uniform(SLEEP_MIN_KEC, SLEEP_MAX_KEC))
        else:
            checkpoint["in_progress"] = {"kec_code": kec_code, "offset": last_offset}
            save_checkpoint(checkpoint)
            print_banner(
                f"BERHENTI: {kec_code} {kec_name} di offset={last_offset}. "
                f"Jalanin ulang script buat resume.",
                char="!",
            )
            return

    elapsed_min = (time.time() - t_start) / 60
    print_banner(
        f"SEMUA KECAMATAN SELESAI -- total {grand_total_rows} baris, "
        f"{grand_total_files} file, {elapsed_min:.1f} menit",
        char="#",
    )


if __name__ == "__main__":
    main()
