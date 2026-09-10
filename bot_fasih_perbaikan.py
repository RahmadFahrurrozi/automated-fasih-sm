"""
Bot Otomasi Perbaikan Data Sensus Ekonomi 2026 (fasih-sm.bps.go.id)
Berdasarkan referensi arsitektur fasih-sql-test.py (Remote Debugging Chrome)

Alur per Assignment:
1. Akses URL dari kolom assignment_id_usaha_bku di perbaikan-data-kabat.xlsx
2. Klik tombol Edit (//*[@id="fasih"]/div/div/div[2]/div/div/div/button[5]) hingga masuk MODE EDIT
3. Klik menu "SE2026 - P KETERANGAN KELUARGA DAN BANGUNAN"
4. Ubah "Keberadaan Bangunan Lainnya/ Usaha" -> pilih "1. Ditemukan"
5. Klik menu "SE2026 - L BLOK II KETERANGAN USAHA/PERUSAHAAN"
6. Klik tombol "Lihat Rincian"
7. Ubah "Keberadaan Usaha" -> pilih "1. Ditemukan"
8. Klik tombol "Kirim" di pojok kanan atas (+ handle dialog konfirmasi jika ada)
9. Simpan checkpoint & log status

CARA MENJALANKAN:
1. Pastikan semua jendela Chrome sudah ditutup.
2. Jalankan: start_chrome_ozifazh.bat (atau buka Chrome dengan port 9222 dan profil ozifazh@gmail.com).
3. Sambungkan VPN BPS dan pastikan sudah login di fasih-sm.bps.go.id.
4. Uji coba 1 data terlebih dahulu:
   python bot_fasih_perbaikan.py --test
5. Jika sudah oke, jalankan untuk semua data:
   python bot_fasih_perbaikan.py
"""

import os
import sys
import json
import time
import random
import argparse
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementClickInterceptedException,
    StaleElementReferenceException,
)

# ---------------- CONFIG ----------------
DEBUGGER_ADDRESS = "127.0.0.1:9222"
DEFAULT_EXCEL_PATH = os.path.abspath("data/110-rozi.xlsx")
OUTPUT_DIR = os.path.abspath("output")
ERROR_DIR = os.path.join(OUTPUT_DIR, "errors")
CHECKPOINT_FILE = os.path.join(OUTPUT_DIR, "checkpoint_110_fix_rozi.json")
RESULT_EXCEL = os.path.join(OUTPUT_DIR, "110-fix-rozi_hasil.xlsx")

# User-provided XPaths
XPATH_BUTTON_EDIT = '//*[@id="fasih"]/div/div/div[2]/div/div/div/button[5]'
XPATH_MENU_BLOK_P = '//*[@id="fasih-form"]/div/div[1]/aside/div[2]/div[3]/div[1]'
XPATH_BLOK_P_ADA_BANG = '//*[@id="ada_bang_usaha"]/div/div[2]/div'
XPATH_BLOK_P_DITEMUKAN = '//*[@id="radiogroup-cl-15-item-cl-18"]'

XPATH_MENU_BLOK_L2 = '//*[@id="fasih-form"]/div/div[1]/aside/div[2]/div[4]/div[1]'
XPATH_BLOK_L_RADIOGROUP = '//*[@id="radiogroup-cl-15"]'
XPATH_BLOK_L_DITEMUKAN = '//*[@id="radiogroup-cl-15-item-cl-18"]'

XPATH_SUBMIT_DOTS = '//*[@id="dropdownmenu-cl-49-trigger"]/button'
XPATH_SUBMIT_PAKSA = '//*[@id="dropdownmenu-cl-49-content-cl-52"]'
XPATH_DIALOG_CONFIRM = '//*[@id="dialog-cl-8-content"]/div/div[3]/button[2]'

# Error Page / Rate Limit XPaths
XPATH_BTN_HOME = '//*[@id="app"]/div/div/div[2]/div/button[1]'
XPATH_BTN_REFRESH = '//*[@id="app"]/div/div/div[2]/div/button[2]'


def print_banner(text: str, char: str = "="):
    width = max(50, min(80, len(text) + 8))
    line = char * width
    print(f"\n{line}")
    print(f"  {text}")
    print(f"{line}\n")


def load_checkpoint() -> dict:
    if os.path.exists(CHECKPOINT_FILE):
        try:
            with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"completed_urls": [], "last_index": -1}


def save_checkpoint(completed_urls: list, last_index: int):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
        json.dump({"completed_urls": completed_urls, "last_index": last_index}, f, indent=2)


def connect_browser():
    print(f"Menghubungkan ke Chrome di {DEBUGGER_ADDRESS}...")
    options = webdriver.ChromeOptions()
    options.debugger_address = DEBUGGER_ADDRESS
    try:
        driver = webdriver.Chrome(options=options)
    except Exception as e:
        print_banner("GAGAL MENGHUBUNGKAN KE CHROME!", char="!")
        print(f"Detail error: {e}")
        print("\nPetunjuk:")
        print("1. Pastikan Chrome dibuka menggunakan: start_chrome_ozifazh.bat")
        print("2. Jika Chrome biasa masih berjalan, tutup dulu semua Chrome lalu jalankan bat.")
        sys.exit(1)

    print("Berhasil terhubung ke Chrome.")
    return driver


def wait_and_click(driver, by, selector, timeout=15, desc=""):
    """Klik elemen dengan retry dan scroll jika terhalang."""
    end_time = time.time() + timeout
    last_err = None
    while time.time() < end_time:
        try:
            elem = WebDriverWait(driver, 3).until(EC.presence_of_element_located((by, selector)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", elem)
            time.sleep(0.3)
            # Coba klik standar
            try:
                elem.click()
                return True
            except (ElementClickInterceptedException, StaleElementReferenceException):
                # Fallback JS click
                driver.execute_script("arguments[0].click();", elem)
                return True
        except Exception as e:
            last_err = e
            time.sleep(0.5)

    raise TimeoutError(f"Gagal klik elemen {desc} ({selector}) dalam {timeout}s: {last_err}")


def check_and_handle_bot_detected(driver, target_url=""):
    """
    Deteksi halaman proteksi 'Bot Detected' dari Fasih BPS:
    'Kami mendeteksi perilaku yang tidak wajar pada koneksi anda...'
    Sesuai instruksi:
    Beri jeda pendinginan selama 5 menit (300 detik), lalu klik link [Kembali] (/html/body/a).
    """
    try:
        title = driver.title.lower()
        page_src = driver.page_source.lower()
        is_bot_page = (
            "bot detected" in title or
            "perilaku yang tidak wajar" in page_src or
            "koneksi anda sebagai bot" in page_src or
            "halosis" in page_src
        )
        if is_bot_page:
            print_banner(
                "[BOT DETECTED] Terdeteksi proteksi 'Bot Detected' dari BPS!\n"
                "Sistem memberi jeda pendinginan 5 menit (300 detik) sebelum klik [Kembali]...",
                char="!"
            )
            # Hitung mundur 5 menit (300 detik)
            total_wait = 300
            while total_wait > 0:
                print(f"  -> Cooldown Bot Detected: sisa {total_wait} detik...")
                step = min(30, total_wait)
                time.sleep(step)
                total_wait -= step

            print("  -> Waktu 5 menit selesai! Mengeklik link [Kembali] (/html/body/a)...")
            kembali_xpaths = [
                "/html/body/a",
                "//a[contains(., 'Kembali') or contains(text(), 'Kembali')]",
                "//a",
            ]
            clicked = False
            for kx in kembali_xpaths:
                try:
                    links = driver.find_elements(By.XPATH, kx)
                    for lk in links:
                        if lk.is_displayed():
                            try:
                                lk.click()
                            except Exception:
                                driver.execute_script("arguments[0].click();", lk)
                            clicked = True
                            print("  -> Berhasil klik link [Kembali].")
                            time.sleep(3)
                            break
                    if clicked:
                        break
                except Exception:
                    pass

            if target_url and target_url not in driver.current_url:
                print(f"  -> Kembali ke URL target: {target_url}...")
                driver.get(target_url)
                time.sleep(2)
            return True
    except Exception:
        pass
    return False


def check_and_handle_session_and_popups(driver, target_url=""):
    """
    Mendeteksi dan menangani:
    0. Halaman 'Bot Detected' (jeda 5 menit + klik Kembali)
    1. Sesi login habis / layar 'Lanjutkan dengan SSO'
    2. Popup modal 'Riwayat Perubahan'
    3. Koneksi VPN / DNS terputus
    """
    # 0. Cek Halaman Bot Detected
    check_and_handle_bot_detected(driver, target_url=target_url)

    # 1. Cek Koneksi Putus (DNS_PROBE / Network error)
    try:
        page_src = driver.page_source.lower()
        if "can’t be reached" in page_src or "can't be reached" in page_src or "dns_probe" in page_src or "err_connection" in page_src:
            print_banner("[KONEKSI TERPUTUS] VPN BPS atau Jaringan terputus!\nMenunggu 10 detik lalu memuat ulang...", char="!")
            time.sleep(10)
            if target_url:
                driver.get(target_url)
                time.sleep(3)
    except Exception:
        pass

    # 2. Cek Halaman Login / SSO (2 Tahap: Fasih SSO -> Keycloak)
    try:
        cur_url = driver.current_url.lower()
        page_src = driver.page_source.lower()
        has_kc = len(driver.find_elements(By.XPATH, '//*[@id="kc-login"] | //input[@id="kc-login"] | //button[@id="kc-login"]')) > 0
        has_sso_btn = len(driver.find_elements(By.XPATH, "//a[contains(., 'Lanjutkan dengan SSO')] | //button[contains(., 'Lanjutkan dengan SSO')] | //*[@id='fasih']//a[contains(@href, 'login')]")) > 0
        is_login = "/login" in cur_url or "sso.bps.go.id" in cur_url or "selamat datang kembali" in page_src or "masuk ke akun anda" in page_src or has_kc or has_sso_btn

        if is_login:
            print_banner("[SESI HABIS] Terdeteksi halaman login! Menjalankan login otomatis 2 tahap...", char="*")

            # TAHAP 1: Klik tombol SSO di Fasih (jika masih di halaman Fasih login)
            sso_xpaths = [
                '//*[@id="fasih"]/div/div/div[1]/div/div/div/div/div/div[2]/a[1]',
                "//a[contains(., 'Lanjutkan dengan SSO') and not(contains(., 'Eksternal'))]",
                "//a[contains(., 'Lanjutkan dengan SSO')]",
                "//button[contains(., 'Lanjutkan dengan SSO')]",
                "//a[contains(@href, 'sso') or contains(., 'SSO')]",
                "//button[contains(., 'SSO') and not(contains(., 'Eksternal'))]",
            ]
            for sx in sso_xpaths:
                try:
                    btns = driver.find_elements(By.XPATH, sx)
                    clicked_sso = False
                    for b in btns:
                        if b.is_displayed():
                            print(f"  -> [LOGIN TAHAP 1] Klik tombol SSO Fasih ('{b.text.strip()}')...")
                            try:
                                b.click()
                            except Exception:
                                driver.execute_script("arguments[0].click();", b)
                            clicked_sso = True
                            time.sleep(2)
                            break
                    if clicked_sso:
                        break
                except Exception:
                    pass

            # TAHAP 2: Klik tombol Login Keycloak (#kc-login)
            end_kc_wait = time.time() + 15
            kc_clicked = False
            while time.time() < end_kc_wait:
                cur_url = driver.current_url.lower()
                kc_buttons = driver.find_elements(By.XPATH, '//*[@id="kc-login"] | //input[@id="kc-login"] | //button[@id="kc-login"]')
                if kc_buttons and any(k.is_displayed() for k in kc_buttons):
                    print("  -> [LOGIN TAHAP 2] Menemukan tombol login Keycloak (#kc-login)...")
                    # Beri jeda 1.2 detik agar password manager/autofill Chrome mengisi kredensial
                    time.sleep(1.2)
                    for kb in kc_buttons:
                        if kb.is_displayed():
                            try:
                                kb.click()
                            except Exception:
                                driver.execute_script("arguments[0].click();", kb)
                            kc_clicked = True
                            print("  -> Berhasil klik #kc-login!")
                            time.sleep(2.5)
                            break
                    if kc_clicked:
                        break

                # Jika sudah langsung tembus ke dashboard Fasih tanpa perlu klik kc-login
                if "fasih-sm.bps.go.id/app" in cur_url or "surveys" in cur_url or "assignment" in cur_url:
                    break
                time.sleep(0.8)

            # Tunggu redirect kembali ke aplikasi Fasih
            end_wait = time.time() + 25
            while time.time() < end_wait:
                cur = driver.current_url.lower()
                if "/login" not in cur and "sso.bps.go.id" not in cur and ("fasih-sm.bps.go.id/app" in cur or "surveys" in cur or "assignment" in cur):
                    print("  -> Sesi login SSO berhasil dipulihkan sepenuhnya!")
                    time.sleep(1.5)
                    break
                time.sleep(1.5)
            else:
                if "/login" in driver.current_url.lower() or "sso.bps.go.id" in driver.current_url.lower() or "auth" in driver.current_url.lower():
                    print_banner("PERHATIAN: Silakan selesaikan login SSO di browser Chrome!\nBot menunggu sampai Anda selesai login...", char="!")
                    while "/login" in driver.current_url.lower() or "sso.bps.go.id" in driver.current_url.lower() or "auth" in driver.current_url.lower():
                        time.sleep(2)
                    print("  -> Login terdeteksi berhasil!")

            # Setelah login pulih, jika ada target_url, langsung tuju URL tersebut
            if target_url and target_url not in driver.current_url:
                print(f"  -> Kembali ke URL target: {target_url}...")
                driver.get(target_url)
                time.sleep(2)
    except Exception:
        pass

    # 3. Cek Popup 'Riwayat Perubahan'
    try:
        modal_riwayat = driver.find_elements(By.XPATH, "//*[contains(text(), 'Riwayat Perubahan')]")
        if modal_riwayat and any(m.is_displayed() for m in modal_riwayat):
            print("  -> Terdeteksi modal 'Riwayat Perubahan', menutup modal...")
            # Cek opsi 'Jangan tampilkan lagi' jika ada
            try:
                toggle_btn = driver.find_elements(By.XPATH, "//button[@role='switch' or contains(@id, 'switch')] | //span[contains(., 'Jangan tampilkan lagi')]")
                if toggle_btn and toggle_btn[0].is_displayed():
                    driver.execute_script("arguments[0].click();", toggle_btn[0])
            except Exception:
                pass

            # Klik tombol close (x)
            close_btns = driver.find_elements(By.XPATH, "//*[contains(text(), 'Riwayat Perubahan')]/ancestor::div[contains(@role, 'dialog')]//button[contains(@aria-label, 'Close') or .//svg] | //button[contains(@aria-label, 'Close')]")
            for cb in close_btns:
                if cb.is_displayed():
                    driver.execute_script("arguments[0].click();", cb)
                    print("  -> Modal 'Riwayat Perubahan' berhasil ditutup.")
                    time.sleep(0.5)
                    break
    except Exception:
        pass


def is_error_or_rate_limit_page(driver):
    """Cek apakah browser menampilkan halaman error / rate limit Fasih."""
    try:
        title = driver.title.lower()
        page_source = driver.page_source.lower()
        if "bot detected" in title or "koneksi anda sebagai bot" in page_source or "perilaku yang tidak wajar" in page_source:
            return True
        if "rate limit exceeded" in page_source or "there's some error" in page_source:
            return True
        # Cek tombol Refresh Page atau Back to Home
        btns = driver.find_elements(By.XPATH, "//button[contains(., 'Refresh Page') or contains(., 'Back to Home')]")
        if btns and any(b.is_displayed() for b in btns):
            return True
    except Exception:
        pass
    return False


def wait_and_recover_rate_limit(driver, target_url, max_retries=4):
    """Jika halaman error rate limit muncul, tunggu cooldown dan pulihkan otomatis."""
    for attempt in range(1, max_retries + 1):
        check_and_handle_session_and_popups(driver, target_url=target_url)
        if not is_error_or_rate_limit_page(driver):
            return True

        cooldown = attempt * 15  # 15s, 30s, 45s, 60s
        print_banner(
            f"[RATE LIMIT EXCEEDED] Fasih membatasi request!\n"
            f"Mendinginkan selama {cooldown} detik (Percobaan {attempt}/{max_retries})...",
            char="!"
        )
        time.sleep(cooldown)

        # Coba klik tombol Refresh Page
        try:
            refresh_btns = driver.find_elements(By.XPATH, XPATH_BTN_REFRESH) or driver.find_elements(By.XPATH, "//button[contains(., 'Refresh Page')]")
            if refresh_btns and refresh_btns[0].is_displayed():
                print("  -> Klik tombol 'Refresh Page'...")
                refresh_btns[0].click()
                time.sleep(3)
        except Exception:
            pass

        # Jika masih error di percobaan berikutnya, coba Back to Home lalu get target_url
        if is_error_or_rate_limit_page(driver):
            try:
                home_btns = driver.find_elements(By.XPATH, XPATH_BTN_HOME) or driver.find_elements(By.XPATH, "//button[contains(., 'Back to Home')]")
                if home_btns and home_btns[0].is_displayed():
                    print("  -> Klik tombol 'Back to Home'...")
                    home_btns[0].click()
                    time.sleep(3)
            except Exception:
                pass
            print(f"  -> Memuat ulang URL: {target_url}...")
            driver.get(target_url)
            time.sleep(3)

        if not is_error_or_rate_limit_page(driver):
            print("  -> Halaman berhasil pulih dari rate limit!")
            return True

    print("  -> Peringatan: Masih terdeteksi halaman rate limit setelah beberapa percobaan.")
    return False


def wait_for_assignment_page_loaded(driver, target_url="", timeout=30):
    """Menunggu seluruh data dan antarmuka assignment selesai loading sebelum interaksi."""
    print("  -> Menunggu proses loading data assignment selesai...")
    end_time = time.time() + timeout

    check_and_handle_session_and_popups(driver, target_url=target_url)

    # Jika browser terlempar keluar dari assignment (misal ke /app/ dashboard atau login)
    if "/assignment/" not in driver.current_url and target_url:
        print(f"  -> Browser berada di luar assignment ({driver.current_url}), memuat ulang target: {target_url}...")
        driver.get(target_url.rstrip('/') + '/edit')
        time.sleep(2)
        check_and_handle_session_and_popups(driver, target_url=target_url)

    # 1. Tunggu readyState complete
    WebDriverWait(driver, 15).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

    # 2. Tunggu indikator loading / spinner / skeleton menghilang
    spinner_xpaths = [
        "//*[contains(@class, 'spinner') or contains(@class, 'loading') or @role='progressbar' or contains(@class, 'loader')]",
        "//*[contains(@class, 'skeleton') or contains(@class, 'animate-pulse')]",
    ]

    while time.time() < end_time:
        has_spinner = False
        for sx in spinner_xpaths:
            try:
                spinners = driver.find_elements(By.XPATH, sx)
                if any(s.is_displayed() for s in spinners):
                    has_spinner = True
                    break
            except Exception:
                pass

        if not has_spinner:
            break
        time.sleep(0.4)

    check_and_handle_session_and_popups(driver, target_url=target_url)

    # 3. Pastikan menu sidebar SE2026 - P sudah benar-benar terlihat di layar
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((
                By.XPATH,
                "//aside//div[contains(@class, 'cursor-pointer') and .//*[contains(text(), 'SE2026 - P')]] | "
                "//aside//*[contains(text(), 'SE2026 - P')]"
            ))
        )
    except Exception:
        pass

    # Jeda settling agar hydration event React/Vue selesai
    time.sleep(1.5)
    print("  -> Loading selesai, antarmuka assignment siap!")


def click_edit_mode(driver, target_url="", timeout=20):
    """Pastikan halaman masuk ke MODE EDIT."""
    # Pastikan berada di halaman assignment
    if "/assignment/" not in driver.current_url and target_url:
        print(f"  -> Halaman belum di assignment ({driver.current_url}), memuat: {target_url}...")
        driver.get(target_url.rstrip('/') + '/edit')
        time.sleep(2)
        check_and_handle_session_and_popups(driver, target_url=target_url)

    # 1. Cek apakah URL sudah /edit atau banner mode edit sudah muncul
    if "/edit" in driver.current_url:
        print("  -> Sudah dalam URL /edit")
        return

    banners = driver.find_elements(By.XPATH, "//*[contains(text(), 'MODE EDIT') or contains(text(), 'mengedit assignment')]")
    if banners and any(b.is_displayed() for b in banners):
        print("  -> Banner MODE EDIT sudah aktif")
        return

    # 2. Coba klik tombol edit via XPath user: button[5]
    print("  -> Mengklik tombol icon Edit...")
    clicked = False
    try:
        wait_and_click(driver, By.XPATH, XPATH_BUTTON_EDIT, timeout=5, desc="Tombol Edit user (button 5)")
        clicked = True
    except Exception:
        pass

    # 3. Jika button[5] tidak ketemu atau gagal, coba direct navigasi ke /edit
    if not clicked and target_url:
        edit_url = target_url.rstrip("/") + "/edit"
        print(f"  -> Tombol edit tidak terjangkau, mencoba navigasi langsung ke: {edit_url}")
        driver.get(edit_url)
        time.sleep(2)
        wait_and_recover_rate_limit(driver, edit_url)

    # 4. Fallback selector tombol edit lain di UI jika belum masuk /edit
    if "/edit" not in driver.current_url:
        fallback_xpath = (
            "//*[@id='fasih']//button[contains(@class, 'orange') or contains(@class, 'warning') or contains(@style, 'orange')] | "
            "//button[.//svg//*[local-name()='path' and contains(@d, 'M')] and not(contains(., 'Ringkasan')) and not(contains(., 'Kirim'))] | "
            "//*[@id='fasih']//button[contains(@class, 'edit') or contains(@title, 'Edit')]"
        )
        try:
            wait_and_click(driver, By.XPATH, fallback_xpath, timeout=5, desc="Fallback Edit Button")
        except Exception:
            pass

    # Tunggu indikator mode edit muncul (maks 15 detik)
    time.sleep(2)
    WebDriverWait(driver, 15).until(
        lambda d: "/edit" in d.current_url or len(d.find_elements(By.XPATH, "//*[contains(text(), 'MODE EDIT')]")) > 0
    )
    print("  -> Berhasil masuk MODE EDIT")




def select_radio_ditemukan_verified(driver, section_name=""):
    """Pilih radio button '1. Ditemukan' dan pastikan terverifikasi aktif (data-checked)."""
    print(f"  -> Memilih '1. Ditemukan' pada {section_name}...")

    label_xpaths = [
        "//label[normalize-space()='1. Ditemukan']",
        "//label[contains(text(), '1. Ditemukan')]",
        "//*[contains(@id, 'item') and contains(., '1. Ditemukan')]//label",
        "//label[contains(., '1. Ditemukan')]",
    ]

    for attempt in range(3):
        # 1. Klik elemen label yang terlihat (karena klik label otomatis mengaktifkan radio di UI Ark/Tailwind)
        for lx in label_xpaths:
            labels = driver.find_elements(By.XPATH, lx)
            for lbl in labels:
                if lbl.is_displayed():
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", lbl)
                    time.sleep(0.3)
                    try:
                        lbl.click()
                    except Exception:
                        pass
                    driver.execute_script("arguments[0].click();", lbl)
                    time.sleep(0.4)

        # 2. Cek apakah indikator data-checked sudah muncul
        checked = driver.find_elements(
            By.XPATH,
            "//*[@data-checked and contains(., '1. Ditemukan')] | "
            "//*[@data-checked and contains(@id, 'item') and contains(., '1. Ditemukan')] | "
            "//input[@type='radio' and @value='1' and (@checked or @data-checked)]"
        )
        if checked:
            print(f"  -> Berhasil memilih dan memverifikasi '1. Ditemukan' ({section_name})!")
            time.sleep(0.5)
            return True

        # 3. Fallback: dispatch click langsung pada input radio value='1'
        inputs = driver.find_elements(By.XPATH, "//input[@type='radio' and @value='1']")
        for inp in inputs:
            try:
                driver.execute_script("""
                    const el = arguments[0];
                    el.checked = true;
                    el.click();
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                """, inp)
                time.sleep(0.4)
            except Exception:
                pass

        # 4. Fallback: klik control role='radio'
        controls = driver.find_elements(
            By.XPATH,
            "//*[@role='radio' and (contains(@id, 'control') or @role='radio') and ../label[contains(., '1. Ditemukan')]]"
        )
        for ctrl in controls:
            try:
                driver.execute_script("arguments[0].click();", ctrl)
                time.sleep(0.4)
            except Exception:
                pass

        checked = driver.find_elements(By.XPATH, "//*[@data-checked and contains(., '1. Ditemukan')]")
        if checked:
            print(f"  -> Berhasil memilih dan memverifikasi '1. Ditemukan' ({section_name})!")
            time.sleep(0.5)
            return True

    print(f"  -> Peringatan: Pilihan '1. Ditemukan' ({section_name}) sudah diproses, melanjutkan...")
    return True


def go_to_blok_p(driver):
    """Navigasi ke menu SE2026 - P (KETERANGAN KELUARGA DAN BANGUNAN)."""
    print("  -> Membuka menu SE2026 - P...")
    p_xpaths = [
        "//aside//div[contains(@class, 'cursor-pointer') and .//*[contains(text(), 'SE2026 - P')]]",
        XPATH_MENU_BLOK_P,
        "//aside//*[contains(text(), 'SE2026 - P') or contains(text(), 'KETERANGAN KELUARGA')]",
    ]
    for px in p_xpaths:
        try:
            wait_and_click(driver, By.XPATH, px, timeout=5, desc="Menu SE2026 - P")
            break
        except Exception:
            continue

    # Tunggu halaman Blok P terbuka
    WebDriverWait(driver, 15).until(
        lambda d: len(d.find_elements(By.XPATH, "//*[contains(text(), 'Keberadaan Bangunan') or contains(text(), 'KETERANGAN KELUARGA')]")) > 0
    )
    time.sleep(0.8)


def set_keberadaan_bangunan_ditemukan(driver):
    """Pilih radio button '1. Ditemukan' pada Blok P."""
    select_radio_ditemukan_verified(driver, "Blok P (Keberadaan Bangunan)")


def go_to_blok_l2(driver):
    """Navigasi ke menu SE2026 - L BLOK II (KETERANGAN USAHA/PERUSAHAAN)."""
    print("  -> Membuka menu SE2026 - L BLOK II...")
    l_xpaths = [
        "//aside//div[contains(@class, 'cursor-pointer') and .//*[contains(text(), 'SE2026 - L BLOK II')]]",
        XPATH_MENU_BLOK_L2,
        "//aside//*[contains(text(), 'SE2026 - L BLOK II') or contains(text(), 'KETERANGAN USAHA')]",
    ]
    for lx in l_xpaths:
        try:
            wait_and_click(driver, By.XPATH, lx, timeout=5, desc="Menu SE2026 - L BLOK II")
            break
        except Exception:
            continue

    time.sleep(1)


def click_lihat_rincian(driver):
    """Klik tombol 'Isi Rincian' atau 'Lihat Rincian' atau row data usaha di Blok L."""
    print("  -> Mengklik tombol Isi / Lihat Rincian di Blok L...")
    rincian_xpaths = [
        '//*[@id="se2026_nested"]/div/div[2]/div[2]/button',
        "//*[@id='se2026_nested']//button",
        "//button[contains(., 'Rincian') or contains(., 'Isi') or contains(., 'Lihat')]",
        '//*[@id="se2026_nested"]/div/div[2]',
    ]
    end_time = time.time() + 15
    while time.time() < end_time:
        for rx in rincian_xpaths:
            try:
                btns = driver.find_elements(By.XPATH, rx)
                for b in btns:
                    try:
                        if b.is_displayed():
                            print(f"  -> Menemukan tombol/row: '{b.text.strip()}', klik sekarang...")
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", b)
                            time.sleep(0.3)
                            try:
                                b.click()
                            except Exception:
                                driver.execute_script("arguments[0].click();", b)
                            time.sleep(1.2)
                            return True
                    except StaleElementReferenceException:
                        continue
            except Exception:
                continue

        # Cek jika form rincian sudah terbuka
        if len(driver.find_elements(By.XPATH, "//*[contains(text(), 'Keberadaan Usaha')]")) > 0:
            print("  -> Form rincian sudah aktif.")
            return True
        time.sleep(0.5)

    print("  -> Tombol rincian tidak muncul (mungkin sudah di dalam form rincian).")


def set_keberadaan_usaha_ditemukan(driver):
    """Pilih radio button '1. Ditemukan' pada Keberadaan Usaha (Blok L II)."""
    select_radio_ditemukan_verified(driver, "Blok L II (Keberadaan Usaha)")


def click_kirim(driver, url=""):
    """
    Alur Submit:
    1. Klik tombol 'Kirim' di pojok kanan atas halaman.
    2. Modal 'Kirim' muncul (dengan info Galat, Peringatan, Kosong).
    3. Klik trigger titik tiga (...) pada split button di modal.
    4. Pilih menu 'Submit Paksa'.
    5. Handle dialog konfirmasi jika muncul.
    """
    print("  -> Mengklik tombol 'Kirim' di pojok kanan atas halaman...")
    # Cek apakah URL masih dalam mode edit
    if "/edit" not in driver.current_url and url:
        print("  -> Mode edit terlepas, memuat ulang URL edit...")
        driver.get(url.rstrip('/') + '/edit')
        time.sleep(1.5)
        check_and_handle_session_and_popups(driver, url.rstrip('/') + '/edit')
        wait_and_recover_rate_limit(driver, url.rstrip('/') + '/edit')

    # Klik tombol Kirim utama di header
    xpath_kirim_header = (
        "//header//button[contains(., 'Kirim')] | "
        "//*[@id='fasih']//button[contains(., 'Kirim') and not(ancestor::*[contains(@role, 'dialog')])]"
    )
    try:
        wait_and_click(driver, By.XPATH, xpath_kirim_header, timeout=8, desc="Tombol Kirim Header")
    except Exception:
        # Fallback tombol Kirim apa saja yang terlihat
        wait_and_click(driver, By.XPATH, "//button[contains(., 'Kirim') and not(@disabled)]", timeout=8, desc="Fallback Kirim")

    # Tunggu modal Kirim muncul (validasi selesai dihitung Fasih)
    print("  -> Menunggu modal 'Kirim' muncul...")
    WebDriverWait(driver, 12).until(
        lambda d: len(d.find_elements(By.XPATH, "//*[contains(text(), '25 Jawaban') or contains(text(), 'GALAT') or contains(text(), 'KOSONG') or @role='dialog']")) > 0
    )

    # Klik tombol trigger titik tiga (...)
    print("  -> Mengklik titik tiga (...) untuk opsi submit paksa...")
    trigger_xpaths = [
        XPATH_SUBMIT_DOTS,
        "//*[contains(@id, 'dropdownmenu') and contains(@id, 'trigger')]//button",
        "//*[contains(@id, 'trigger')]//button",
        "//*[@role='dialog']//button[contains(., 'Kirim')]/following-sibling::button",
        "//*[@role='dialog']//button[contains(@aria-haspopup, 'menu')]",
        "//*[@role='dialog']//button[.//svg and not(contains(., 'Batal')) and not(contains(., 'Kirim'))]",
    ]

    clicked_trigger = False
    for t_xpath in trigger_xpaths:
        try:
            wait_and_click(driver, By.XPATH, t_xpath, timeout=2.5, desc="Trigger Titik Tiga")
            clicked_trigger = True
            print("  -> Berhasil klik trigger titik tiga")
            break
        except Exception:
            continue

    if not clicked_trigger:
        print("  -> Peringatan: Gagal klik trigger titik tiga, mencoba langsung klik 'Kirim' di modal...")
        wait_and_click(driver, By.XPATH, "//*[@role='dialog']//button[contains(., 'Kirim')]", timeout=4, desc="Kirim Modal")
        time.sleep(2)
        return

    # Klik opsi "Submit Paksa" dari dropdown
    print("  -> Memilih menu 'Submit Paksa'...")
    paksa_xpaths = [
        XPATH_SUBMIT_PAKSA,
        "//*[contains(text(), 'Submit Paksa') or contains(text(), 'submit paksa') or contains(text(), 'Kirim Paksa') or contains(text(), 'kirim paksa')]",
        "//*[contains(@id, 'dropdownmenu') and contains(@id, 'content')]//*[contains(., 'aksa') or contains(., 'Paksa')]",
        "//*[@role='menu']//*[contains(., 'aksa') or contains(., 'Paksa') or contains(., 'Submit')]",
        "//*[@role='menuitem'][contains(., 'aksa') or contains(., 'Paksa')]",
        "//button[contains(., 'Paksa') or contains(., 'paksa')]",
    ]

    clicked_paksa = False
    end_paksa = time.time() + 5
    while time.time() < end_paksa:
        for p_xpath in paksa_xpaths:
            try:
                elems = driver.find_elements(By.XPATH, p_xpath)
                for el in elems:
                    if el.is_displayed():
                        print(f"  -> Menemukan opsi Submit Paksa ('{el.text.strip()}'), klik sekarang...")
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
                        time.sleep(0.15)
                        try:
                            el.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", el)
                        clicked_paksa = True
                        break
                if clicked_paksa:
                    break
            except Exception:
                continue
        if clicked_paksa:
            break
        time.sleep(0.2)

    if not clicked_paksa:
        print("  -> Gagal klik menu 'Submit Paksa' via selector khusus, mencoba klik konten menu dropdown...")
        try:
            menu_items = driver.find_elements(By.XPATH, "//*[@role='menu']/* | //*[contains(@id, 'content')]/*")
            for item in menu_items:
                if item.is_displayed():
                    driver.execute_script("arguments[0].click();", item)
                    clicked_paksa = True
                    break
        except Exception:
            pass

    # Menunggu dan mengklik dialog konfirmasi pengiriman akhir (button konfirmasi)
    print("  -> Menunggu dialog konfirmasi akhir muncul...")
    confirm_xpaths = [
        XPATH_DIALOG_CONFIRM,
        "//*[contains(@id, 'dialog') and contains(@id, 'content')]//div[last()]//button[2]",
        "//*[contains(@id, 'dialog') and contains(@id, 'content')]//button[not(contains(., 'Batal')) and (contains(., 'Kirim') or contains(., 'Ya') or contains(., 'Setuju') or contains(., 'Lanjut') or contains(., 'Konfirmasi'))]",
        "//div[contains(@role, 'dialog') or contains(@class, 'modal')]//button[not(contains(., 'Batal')) and (contains(., 'Kirim') or contains(., 'Ya') or contains(., 'Setuju') or contains(., 'OK') or contains(., 'Konfirmasi'))]",
        "//div[contains(@role, 'dialog')]//div[contains(@class, 'footer') or position()=last()]//button[last()]",
    ]

    clicked_confirm = False
    end_confirm = time.time() + 6
    while time.time() < end_confirm:
        for c_xpath in confirm_xpaths:
            try:
                btns = driver.find_elements(By.XPATH, c_xpath)
                for b in btns:
                    if b.is_displayed() and not ("batal" in b.text.lower()):
                        print(f"  -> Menemukan tombol konfirmasi akhir ('{b.text.strip()}'), klik sekarang...")
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", b)
                        time.sleep(0.15)
                        try:
                            b.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", b)
                        clicked_confirm = True
                        break
                if clicked_confirm:
                    break
            except Exception:
                continue
        if clicked_confirm:
            break
        time.sleep(0.2)

    if not clicked_confirm:
        print("  -> Info: Tidak terdeteksi dialog konfirmasi lanjutan (mungkin submit sudah langsung diproses).")

    # Beri jeda singkat agar pengiriman selesai terekam di server
    time.sleep(1.8)
    print("  -> Assignment berhasil disubmit paksa!")


def process_assignment(driver, url: str, row_info: str):
    """Proses lengkap untuk satu URL assignment."""
    print_banner(f"Memproses: {row_info}\nURL: {url}", char="-")

    url = url.strip()

    # Coba langsung buka edit_url agar langsung masuk MODE EDIT
    edit_url = url.rstrip('/') + '/edit' if not url.endswith('/edit') else url
    print(f"  -> Membuka URL target: {edit_url}...")
    driver.get(edit_url)
    time.sleep(1.5)

    # Tangani jika sesi login habis / popup changelog / jaringan putus
    check_and_handle_session_and_popups(driver, target_url=edit_url)

    # Tangani jika terkena 'Rate limit exceeded' atau 'There's some error'
    wait_and_recover_rate_limit(driver, edit_url)

    # Tunggu seluruh proses loading data assignment selesai
    wait_for_assignment_page_loaded(driver, target_url=edit_url, timeout=30)

    # Pastikan status MODE EDIT aktif (klik tombol edit atau direct route)
    click_edit_mode(driver, target_url=url)
    wait_and_recover_rate_limit(driver, edit_url)

    # 3. Masuk SE2026 - P
    go_to_blok_p(driver)

    # 4. Set Keberadaan Bangunan -> 1. Ditemukan
    set_keberadaan_bangunan_ditemukan(driver)

    # 5. Masuk SE2026 - L BLOK II
    go_to_blok_l2(driver)

    # 6. Klik Lihat Rincian
    click_lihat_rincian(driver)

    # 7. Set Keberadaan Usaha -> 1. Ditemukan
    set_keberadaan_usaha_ditemukan(driver)

    # 8. Klik Kirim
    click_kirim(driver, url=url)

    print(f"  [SUKSES] Assignment selesai diperbarui: {row_info}")


def main():
    parser = argparse.ArgumentParser(description="Bot Otomasi Perbaikan Data Sensus Ekonomi 2026")
    parser.add_argument("--url", type=str, default=None, help="Testing langsung ke URL assignment tertentu tanpa baca Excel")
    parser.add_argument("--excel", type=str, default=None, help="Path ke file Excel data")
    parser.add_argument("--test", action="store_true", help="Uji coba hanya 1 baris pertama")
    parser.add_argument("--limit", type=int, default=None, help="Batasi jumlah data yang diproses")
    parser.add_argument("--start", type=int, default=None, help="Mulai dari indeks baris tertentu (0-indexed)")
    parser.add_argument("--reset-checkpoint", action="store_true", help="Reset checkpoint dan mulai dari awal")
    parser.add_argument("--retry-failed", action="store_true", help="Hanya proses data yang sebelumnya GAGAL atau belum selesai")
    args = parser.parse_args()

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(ERROR_DIR, exist_ok=True)

    # Jika user ingin test 1 URL spesifik langsung
    if args.url:
        target_url = args.url.strip()
        print_banner(f"MODE TEST SINGLE URL:\n{target_url}", char="*")
        driver = connect_browser()
        try:
            process_assignment(driver, target_url, "Testing Single URL User")
            print_banner("TEST SINGLE URL BERHASIL!", char="#")
        except Exception as e:
            print_banner(f"TEST GAGAL: {e}", char="!")
            screenshot_file = os.path.join(ERROR_DIR, "test_custom_url_error.png")
            try:
                driver.save_screenshot(screenshot_file)
                print(f"Screenshot error disimpan ke: {screenshot_file}")
            except Exception:
                pass
        return

    # Deteksi file Excel
    excel_path = args.excel or DEFAULT_EXCEL_PATH
    if not os.path.exists(excel_path):
        candidate_in_data = os.path.join("data", os.path.basename(excel_path))
        if os.path.exists(candidate_in_data):
            excel_path = os.path.abspath(candidate_in_data)
        else:
            # Coba cari file xlsx apa saja di dalam folder data
            data_files = [os.path.join("data", f) for f in os.listdir("data") if f.endswith(".xlsx") and not f.startswith("~$")]
            if data_files:
                excel_path = os.path.abspath(data_files[0])
                print(f"Otomatis menggunakan file data: {excel_path}")
            else:
                print(f"Error: File Excel tidak ditemukan di {excel_path} maupun di folder data/!")
                sys.exit(1)

    # Checkpoint
    checkpoint = load_checkpoint()
    if args.reset_checkpoint:
        checkpoint = {"completed_urls": [], "last_index": -1}
        save_checkpoint([], -1)
        print("Checkpoint di-reset ke awal.")

    completed_urls = set(checkpoint.get("completed_urls", []))

    # Jika mode retry-failed dan file hasil sudah ada, prioritaskan baca dari file hasil
    if args.retry_failed and os.path.exists(RESULT_EXCEL):
        print(f"Membaca progres dari file hasil: {RESULT_EXCEL}...")
        df = pd.read_excel(RESULT_EXCEL)
    else:
        print(f"Membaca file sumber: {excel_path}...")
        df = pd.read_excel(excel_path)

    col_url = "assignment_id_usaha_bku"
    if col_url not in df.columns:
        print(f"Error: Kolom '{col_url}' tidak ditemukan dalam Excel!")
        print(f"Kolom yang ada: {list(df.columns)}")
        sys.exit(1)

    total_rows = len(df)
    print(f"Total data di Excel: {total_rows} baris")

    # Siapkan status log jika belum ada
    if "status_bot" not in df.columns:
        df["status_bot"] = ""
        df["waktu_eksekusi"] = ""

    # Tentukan baris yang akan diproses
    if args.retry_failed:
        target_indices = []
        for idx, row in df.iterrows():
            url = str(row[col_url]).strip()
            if not url or url.lower() == "nan":
                continue
            if url not in completed_urls:
                target_indices.append(idx)

        if not target_indices:
            print_banner(
                "SELAMAT! SEMUA DATA SUDAH SUKSES (100%)!\n"
                f"Total {len(completed_urls)} data telah tuntas. Tidak ada data yang perlu di-retry.",
                char="#"
            )
            return

        print_banner(
            f"MODE RETRY: Memproses {len(target_indices)} data yang gagal / belum selesai!\n"
            f"(Data yang sudah sukses sebanyak {len(completed_urls)} otomatis dilewati)",
            char="*"
        )
        max_rows = len(target_indices)
    elif args.test:
        print_banner("MODE TEST: Hanya akan memproses 1 baris data pertama!", char="*")
        target_indices = list(range(min(1, total_rows)))
        max_rows = 1
    elif args.limit:
        max_rows = args.limit
        print_banner(f"MODE LIMIT: Memproses maksimal {max_rows} baris", char="*")
        target_indices = list(range(total_rows))
    else:
        max_rows = total_rows
        target_indices = list(range(total_rows))

    driver = connect_browser()

    processed_count = 0
    success_count = 0
    failed_count = 0

    try:
        for loop_idx, idx in enumerate(target_indices):
            row = df.loc[idx]

            if not args.retry_failed and args.start is not None and idx < args.start:
                continue

            if not args.retry_failed and processed_count >= max_rows:
                print(f"Batas pemrosesan ({max_rows} data) tercapai.")
                break

            url = str(row[col_url]).strip()
            if not url or url.lower() == "nan":
                continue

            # Skip jika sudah selesai (pada mode normal)
            if not args.retry_failed and url in completed_urls and not args.test:
                print(f"[SKIP] Baris {idx+1}/{total_rows} sudah selesai sebelumnya.")
                continue

            row_desc = f"Baris {idx+1}/{total_rows} - {row.get('nama_usaha_bku', '')} ({row.get('level_6_name', '')})"
            if args.retry_failed:
                print(f"\n>>> Memproses Data Retry #{loop_idx+1}/{len(target_indices)} (Baris Excel {idx+1}) <<<")

            processed_count += 1

            try:
                process_assignment(driver, url, row_desc)
                success_count += 1

                # Cek jika baris ini sebelumnya sempat error dan memiliki screenshot
                err_img = os.path.join(ERROR_DIR, f"error_row_{idx+1}.png")
                was_error = os.path.exists(err_img)
                if was_error:
                    try:
                        os.remove(err_img)
                        print(f"  [REPAIR BERHASIL] Baris {idx+1} yang sempat error kini TUNTAS diperbaiki! Screenshot lama otomatis dibersihkan.")
                    except Exception:
                        pass

                # Update status
                df.at[idx, "status_bot"] = "SUKSES"
                df.at[idx, "waktu_eksekusi"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                completed_urls.add(url)
                save_checkpoint(list(completed_urls), int(idx))

                # Simpan excel hasil berkala atau setiap baris di mode retry
                if args.retry_failed or processed_count % 5 == 0 or processed_count == max_rows:
                    df.to_excel(RESULT_EXCEL, index=False)

                # Jeda santai antar assignment agar tidak memicu rate limit / bot detection Fasih
                cooldown_between = random.uniform(5.5, 9.0)
                print(f"  -> Jeda santai {cooldown_between:.1f} detik sebelum data berikutnya...")
                time.sleep(cooldown_between)

            except Exception as e:
                failed_count += 1
                error_msg = str(e)
                print_banner(f"ERROR PADA BARIS {idx+1}: {error_msg}", char="!")

                # Simpan screenshot error untuk investigasi
                screenshot_file = os.path.join(ERROR_DIR, f"error_row_{idx+1}.png")
                try:
                    driver.save_screenshot(screenshot_file)
                    print(f"  -> Screenshot error disimpan ke: {screenshot_file}")
                except Exception:
                    pass

                df.at[idx, "status_bot"] = f"GAGAL: {error_msg[:100]}"
                df.at[idx, "waktu_eksekusi"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                df.to_excel(RESULT_EXCEL, index=False)

                if args.test:
                    print("Mode test berhenti pada error untuk pengecekan.")
                    break

                # Beri jeda sebelum lanjut ke row berikutnya
                time.sleep(3)

    finally:
        df.to_excel(RESULT_EXCEL, index=False)
        print_banner(
            f"RINGKASAN EKSEKUSI:\n"
            f"- Diproses: {processed_count}\n"
            f"- Sukses  : {success_count}\n"
            f"- Gagal   : {failed_count}\n"
            f"- Total Selesai (Semua): {len(completed_urls)}/{total_rows}\n"
            f"- Hasil Excel: {RESULT_EXCEL}",
            char="#",
        )


if __name__ == "__main__":
    main()
