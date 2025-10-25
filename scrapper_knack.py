from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
import time
import gspread
from google.oauth2.service_account import Credentials
import os
import pytz
import tempfile

# Jakarta Timezone
JAKARTA_TZ = pytz.timezone('Asia/Jakarta')

# === CONFIG ===
EMAIL = os.getenv("KNACK_EMAIL", "brian.saputra@javamifi.com")
PASSWORD = os.getenv("KNACK_PASSWORD", "Tikuskantor12345#")
LOGIN_URL = "https://jvm.knack.com/app#login"
TARGET_URL = "https://jvm.knack.com/app#main-form2"
GOOGLE_CREDS_PATH = os.getenv("GOOGLE_CREDS_PATH", "workersm.json")


# === LOGGING ===
def write_log(msg):
    log_time = datetime.now(JAKARTA_TZ).strftime('%Y-%m-%d %H:%M:%S')
    print(f"[{log_time}] [LOG] {msg}")
    try:
        log_dir = os.path.dirname(os.path.abspath(__file__))
        log_path = os.path.join(log_dir, "log.txt")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{log_time}] [LOG] {msg}\n")
    except Exception:
        pass


# === DRIVER SETUP ===
def start_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-setuid-sandbox")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--remote-debugging-port=0")
    # Prevent user data directory conflicts
    user_data_dir = tempfile.mkdtemp()
    options.add_argument(f"--user-data-dir={user_data_dir}")
    return webdriver.Chrome(options=options)


# === LOGIN ===
def login_knack(driver, url, email, password):
    driver.get(url)
    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located(
                (By.XPATH, '//input[@type="email"]'))).send_keys(email)
        driver.find_element(By.XPATH,
                            '//input[@type="password"]').send_keys(password)
        write_log("📧 Email dan password diketik")

        login_xpaths = [
            '//button[contains(text(), "Login")]', '//input[@type="submit"]',
            '//button[@type="submit"]'
        ]
        for xpath in login_xpaths:
            try:
                driver.find_element(By.XPATH, xpath).click()
                write_log(f"✅ Klik login berhasil: {xpath}")
                break
            except Exception:
                continue

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((
                By.XPATH,
                '//div[contains(@id, "view_") and contains(@class, "kn-view")]'
            )))
        write_log("✅ Login sukses")
        time.sleep(5)
        return True
    except Exception as e:
        write_log(f"❌ Gagal login otomatis: {type(e).__name__} - {e}")
        return False


# === SEARCH & KLIK DETAIL ORDER (TAB BARU) ===
def search_and_open_detail(driver, url, order_id, imei):
    driver.get(url)
    try:
        # Cari input search dan ketik IMEI (bukan order_id)
        search_input = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((
                By.XPATH,
                '//input[@name="keyword"][@placeholder="Search by keyword"]')))
        search_input.clear()
        search_input.send_keys(str(imei).strip())
        write_log(f"🔍 IMEI diketik: {imei}")

        search_button = driver.find_element(
            By.XPATH, '//*[@id="view_1137"]/div[2]/div[2]/form/p/a')
        search_button.click()
        write_log("🔎 Tombol Search diklik")

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CLASS_NAME, "kn-table")))
        time.sleep(3)

        rows = driver.find_elements(By.CSS_SELECTOR, ".kn-table tbody tr")
        found = False
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            data = [col.text.strip() for col in cols]
            write_log(f"➡️ {data}")

            # Tetap cocokin pakai Order ID
            if any(
                    str(order_id).strip() == str(cell).strip()
                    for cell in data):
                write_log(f"🎯 Match ditemukan: {order_id}")
                found = True
                try:
                    eye_icon = row.find_element(
                        By.XPATH,
                        './/a[contains(@href, "view-order-details12")]//i[contains(@class, "fa-eye")]'
                    )
                    ActionChains(driver).key_down(
                        Keys.CONTROL).click(eye_icon).key_up(
                            Keys.CONTROL).perform()
                    write_log("👁️ Tombol mata diklik (tab baru)")

                    time.sleep(2)
                    tabs = driver.window_handles
                    driver.switch_to.window(tabs[-1])
                    write_log("🪟 Berpindah ke tab detail order")

                    # --- TAMBAHAN: TUNGGU HALAMAN DETAIL ORDER ---
                    WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located((
                            By.XPATH,
                            '//div[contains(@id, "view_") and contains(@class, "kn-view")]'
                        )))
                    time.sleep(2)  # Tambah delay agar semua field muncul

                    current_url = driver.execute_script(
                        "return window.location.href")
                    write_log(f"🌐 URL aktif: {current_url}")
                    uuid = current_url.split("/")[-2] if current_url.endswith(
                        "/") else current_url.split("/")[-1]
                    write_log(f"🆔 UUID detail order: {uuid}")
                    return uuid
                except Exception as e:
                    write_log(f"❌ Gagal klik detail: {type(e).__name__} - {e}")
                    return None
        if not found:
            write_log("⚠️ Order ID tidak ditemukan")
        return None
    except Exception as e:
        write_log(f"❌ Gagal cari Order ID: {type(e).__name__} - {e}")
        return None


# --- FIXED SCRAPE NOMOR RESI ---
def scrape_nomor_resi(driver, target_resi):
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located(
                (By.CLASS_NAME, "kn-table-wrapper")))
        time.sleep(2)
        found_resi = []
        tables = driver.find_elements(By.CLASS_NAME, "kn-table-wrapper")
        for table in tables:
            rows = table.find_elements(By.CSS_SELECTOR, "tbody tr")
            for row in rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                for col in cols:
                    text = col.text
                    clean_text = text.strip()
                    clean_target = str(target_resi).strip()
                    write_log(f"📦 Nomor Resi ditemukan: {clean_text}")
                    write_log(
                        f"🧪 Bandingkan: {repr(clean_text)} vs {repr(clean_target)}"
                    )
                    if clean_text == clean_target:
                        write_log(
                            f"✅ Nomor Resi cocok dengan target: {target_resi}")
                        return True
                    found_resi.append(clean_text)
        write_log(f"❌ Nomor Resi tidak cocok. Target: {target_resi}")
        return False
    except Exception as e:
        write_log(f"❌ Gagal scrape resi: {type(e).__name__} - {e}")
        return False


# === KLIK AWB DI BARIS YANG COCOK (TAB AKTIF) ===
def click_awb_icon_by_resi(driver, target_resi):
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located(
                (By.CLASS_NAME, "kn-table-wrapper")))
        time.sleep(2)
        tables = driver.find_elements(By.CLASS_NAME, "kn-table-wrapper")
        for table in tables:
            rows = table.find_elements(By.CSS_SELECTOR, "tbody tr")
            for row in rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                if any(str(target_resi) in col.text for col in cols):
                    write_log(f"🎯 Baris AWB cocok dengan resi: {target_resi}")
                    icon = row.find_element(
                        By.XPATH, './/i[contains(@class, "fa-eye")]')
                    icon.click()
                    write_log("👁️ Ikon mata AWB diklik (tab aktif)")
                    WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located((
                            By.XPATH,
                            '//div[contains(@id, "view_") and contains(@class, "kn-view")]'
                        )))
                    write_log("📄 Halaman detail AWB berhasil dimuat")
                    return True
        write_log(
            f"❌ Tidak ada baris AWB yang cocok dengan resi: {target_resi}")
        return False
    except Exception as e:
        write_log(f"❌ Gagal klik AWB: {type(e).__name__} - {e}")
        return False


# === SELECT MODEM BY STATUS ===
def select_modem_by_status(driver, status_modem, rental_id):
    """
    Gunakan chzn xpath spesifik:
      READY  -> //*[@id="view_2650_field_1136_chzn"]
      BROKEN -> //*[@id="view_2650_field_1613_chzn"]
    Ketik IMEI di input chzn, tekan ENTER untuk trigger, tunggu hasil, lalu dispatch change ke <select>.
    """
    try:
        write_log(
            f"🔁 select_modem_by_status start | status={status_modem} | imei={rental_id}"
        )

        # tentuin chzn xpath & field id
        if status_modem.strip().upper() == "READY":
            chzn_xpath = '//*[@id="view_2650_field_1136_chzn"]'
            field_id = 1136
        else:
            chzn_xpath = '//*[@id="view_2650_field_1613_chzn"]'
            field_id = 1613

        # ambil container chzn
        try:
            chzn = WebDriverWait(driver, 6).until(
                EC.presence_of_element_located((By.XPATH, chzn_xpath)))
        except Exception:
            write_log(f"❌ chzn container gak ketemu: {chzn_xpath}")
            return False

        # ambil input di dalam chzn
        try:
            chzn_input = chzn.find_element(
                By.XPATH,
                ".//li[contains(@class,'search-field')]//input | .//input[@type='text' or @type='search']"
            )
        except Exception:
            write_log("❌ Input di chzn gak ketemu")
            return False

        # clear default placeholder & fokus
        try:
            driver.execute_script("arguments[0].value = '';", chzn_input)
            driver.execute_script(
                "if(arguments[0].classList) arguments[0].classList.remove('default');",
                chzn_input)
        except Exception:
            pass
        try:
            chzn_input.click()
        except Exception:
            try:
                driver.execute_script("arguments[0].focus();", chzn_input)
            except Exception:
                pass
        time.sleep(0.12)

        # ketik IMEI pelan
        for ch in str(rental_id):
            try:
                chzn_input.send_keys(ch)
            except Exception:
                driver.execute_script("arguments[0].value += arguments[1];",
                                      chzn_input, ch)
            time.sleep(0.04)
        write_log(f"⌨️ IMEI diketik di chzn: {rental_id}")

        # TEKAN ENTER untuk trigger (permintaan lo)
        try:
            chzn_input.send_keys(Keys.ENTER)
            write_log("⏎ ENTER dikirim untuk trigger pilihan")
        except Exception:
            try:
                driver.execute_script(
                    "arguments[0].dispatchEvent(new KeyboardEvent('keydown', {'key':'Enter'}));",
                    chzn_input)
                write_log("⏎ ENTER dispatched via JS fallback")
            except Exception:
                write_log("⚠️ Gagal kirim ENTER ke chzn input")

        time.sleep(0.2)  # beri waktu proses trigger

        # pastiin dropdown visible (bila perlu)
        try:
            driver.execute_script(
                "arguments[0].querySelector('.chzn-drop').style.left = '0px';",
                chzn)
        except Exception:
            pass
        time.sleep(0.12)

        # tunggu dan klik active-result yang sesuai (jika muncul)
        item_xpath = f"({chzn_xpath}//li[contains(@class,'active-result') and normalize-space() = '{rental_id}'] | //ul[contains(@class,'chzn-results')]//li[contains(@class,'active-result') and normalize-space() = '{rental_id}'])"
        try:
            WebDriverWait(driver, 6).until(
                EC.visibility_of_element_located((By.XPATH, item_xpath)))
            item = driver.find_element(By.XPATH, item_xpath)
            driver.execute_script("arguments[0].scrollIntoView(true);", item)
            try:
                item.click()
            except Exception:
                driver.execute_script("arguments[0].click();", item)
            write_log(f"✅ Pilihan chzn diklik untuk IMEI={rental_id}")
        except Exception:
            # item tidak muncul — lanjutkan ke sinkronisasi tanpa log error
            pass

        # dispatch change ke <select> sebagai sinkronisasi
        try:
            sel = driver.find_element(
                By.XPATH,
                f"//select[contains(@id,'field_{field_id}') or contains(@name,'field_{field_id}')]"
            )
            try:
                opt = sel.find_element(
                    By.XPATH, f".//option[normalize-space() = '{rental_id}']")
                val = opt.get_attribute("value")
                if val:
                    driver.execute_script(
                        "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                        sel, val)
                    write_log(f"🔁 Dispatch change ke <select> value={val}")
                else:
                    driver.execute_script(
                        "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                        sel)
                    write_log(
                        "⚠️ Dispatch change fallback (option tidak punya value)"
                    )
            except Exception:
                # fallback: trigger change tanpa set value
                driver.execute_script(
                    "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                    sel)
                write_log("⚠️ Dispatch change fallback ke <select>")
        except Exception:
            write_log(
                "⚠️ <select> sinkronisasi gagal (tapi pilihan chzn mungkin sudah diklik)"
            )

        time.sleep(0.3)
        return True

    except Exception as e:
        write_log(
            f"❌ Error di select_modem_by_status: {type(e).__name__} - {e}")
        return False


# === ISI FORM RETURN SESUAI STATUS MODEM ===
def isi_form_return(driver, status_modem, rental_id, tgl_terima):
    success = True
    try:
        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(3)
        write_log(f"🗓️ Tanggal Terima dari sheet: {tgl_terima}")

        # Tanggal Terima
        try:
            tanggal_input = driver.find_element(
                By.XPATH, '//input[contains(@id, "field_1134")]')
            tanggal_input.clear()
            tanggal_input.send_keys(tgl_terima)
            write_log("✅ Tanggal Terima diisi")
            time.sleep(3)
            try:
                driver.execute_script("arguments[0].blur();", tanggal_input)
            except Exception:
                pass
            try:
                driver.find_element(By.TAG_NAME, "body").click()
            except Exception:
                pass
            try:
                tanggal_input.send_keys(Keys.ESCAPE)
            except Exception:
                pass
            time.sleep(0.5)
        except Exception as e:
            write_log(f"❌ Gagal isi Tanggal Terima: {type(e).__name__}")
            success = False

        # Diterima Oleh
        try:
            diterima_input = driver.find_element(By.XPATH,
                                                 '//input[@id="field_2025"]')
            diterima_input.clear()
            diterima_input.send_keys("Brian Saputra")
            write_log("✅ Diterima Oleh diisi")
        except Exception as e:
            write_log(f"❌ Gagal isi Diterima Oleh: {type(e).__name__}")
            success = False

        # Pilih modem
        try:
            write_log(
                "🔎 Pilih modem langsung berdasarkan status (pakai chzn xpath)..."
            )
            ok = select_modem_by_status(driver, status_modem, rental_id)
            if ok:
                write_log(f"✅ Modem berhasil dipilih untuk IMEI {rental_id}")
            else:
                write_log(f"❌ Gagal memilih modem untuk IMEI {rental_id}")
                success = False
        except Exception as e:
            write_log(f"❌ Error saat memilih modem: {type(e).__name__} - {e}")
            success = False

        # Klik Submit
        try:
            submit_btn = driver.find_element(
                By.XPATH, '//*[@id="view_2650"]/form/div/button')
            driver.execute_script("arguments[0].scrollIntoView(true);",
                                  submit_btn)
            try:
                submit_btn.click()
            except Exception:
                driver.execute_script("arguments[0].click();", submit_btn)
            write_log("📤 Tombol Submit diklik")
        except Exception as e:
            write_log(f"❌ Gagal klik Submit: {type(e).__name__} - {e}")
            success = False

    except Exception as e:
        write_log(f"❌ Gagal isi form return: {type(e).__name__} - {e}")
        success = False

    return success


# === SCRAPE RENTAL STOCKS ===
def scrape_rental_stocks(driver, field_id=1136, timeout=8):
    """
    Scrape data dari kolom Rental Stocks (connection picker).
    - field_id: angka id field (contoh 1136 di screenshot)
    Returns dict:
      {
        "select_id": "<select element id>",
        "options": [ {"value": "...", "text": "..."}, ... ],
        "chosen_texts": [ "1222212004054", ... ],   # yang tampil di chzn-chosen / active tokens
        "selected_values": [ "option_value", ... ]  # value dari <select> yang terpilih (jika ada)
      }
    """
    from selenium.common.exceptions import TimeoutException
    result = {
        "select_id": None,
        "options": [],
        "chosen_texts": [],
        "selected_values": [],
    }
    try:
        # cari <select> yang berkaitan dengan field_id (misal id contains "field_1136")
        sel = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((
                By.XPATH,
                f"//select[contains(@id, 'field_{field_id}') or contains(@name, 'field_{field_id}')]"
            )))
        select_id = sel.get_attribute("id") or sel.get_attribute("name")
        result["select_id"] = select_id

        # ambil semua opsi dari <option>
        option_elems = sel.find_elements(By.TAG_NAME, "option")
        for opt in option_elems:
            val = opt.get_attribute("value") or ""
            txt = (opt.text or "").strip()
            result["options"].append({"value": val, "text": txt})
            if opt.get_attribute("selected") or opt.is_selected():
                if val:
                    result["selected_values"].append(val)

        # ambil teks yang tampil di chzn / chosen container (token yang terlihat)
        try:
            # beberapa Knack pakai id pattern: view_XXXX_field_1136_chzn atau div dengan class chzn-choices
            chzn = None
            try:
                chzn = driver.find_element(
                    By.XPATH,
                    f"//div[contains(@id, 'field_{field_id}') and contains(@class,'chzn-container')]"
                )
            except Exception:
                # fallback cari container berdasarkan select id (if present)
                if select_id:
                    chzn = driver.find_element(
                        By.XPATH,
                        f"//div[contains(@id, '{select_id}') and contains(@class,'chzn-container')]"
                    )
            if chzn:
                # token list di child li / span tergantung implementasi
                texts = []
                # li tokens
                lis = chzn.find_elements(By.XPATH, ".//li")
                for li in lis:
                    t = (li.text or "").strip()
                    if t:
                        texts.append(t)
                # span tokens (alternatif)
                spans = chzn.find_elements(By.XPATH, ".//span")
                for sp in spans:
                    t = (sp.text or "").strip()
                    if t and t not in texts:
                        texts.append(t)
                # input placeholder text (nilai yang tercantum di input)
                inputs = chzn.find_elements(
                    By.XPATH, ".//input[@type='text' or @type='search']")
                for ip in inputs:
                    v = (ip.get_attribute("value") or "").strip()
                    if v and v not in texts:
                        texts.append(v)
                result["chosen_texts"] = texts
        except Exception:
            # ignore, tetap kembalikan apa yang ada
            pass

        return result
    except TimeoutException:
        write_log(f"❌ Timeout: tidak menemukan select untuk field_{field_id}")
        return result
    except Exception as e:
        write_log(f"❌ Error scrape_rental_stocks: {type(e).__name__} - {e}")
        return result


# contoh pemakaian (di mana aja dalam flow lo):
# data = scrape_rental_stocks(driver, field_id=1136)
# write_log(f"📥 Rental Stocks select id: {data['select_id']}")
# write_log(f"📦 Options found: {len(data['options'])}")
# for o in data['options'][:10]:
#     write_log(f" - {o['value']} : {o['text']}")
# write_log(f"🔹 Chosen tokens: {data['chosen_texts']}")
# write_log(f"🔹 Selected values: {data['selected_values']}")


def choose_rental_stock_by_imei(driver, field_id, imei, timeout=8):
    """
    Coba pilih opsi Rental Stocks sesuai IMEI.
    1) cek <select> options, kalau ada option.text == imei -> set value via JS + dispatch change
    2) jika gak, coba buka chzn container dan klik li yang matching imei
    Return True kalau berhasil, False kalau gagal.
    """
    from selenium.common.exceptions import TimeoutException
    try:
        sel = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((
                By.XPATH,
                f"//select[contains(@id,'field_{field_id}') or contains(@name,'field_{field_id}')]"
            )))
    except Exception:
        write_log("❌ choose: tidak menemukan <select> untuk field_id")
        return False

    # 1) coba match option text/content
    options = sel.find_elements(By.TAG_NAME, "option")
    for opt in options:
        txt = (opt.get_attribute("textContent") or "").strip()
        if txt == str(imei).strip():
            val = opt.get_attribute("value")
            if not val:
                continue
            # set value via JS & trigger change
            driver.execute_script(
                """
                var s = arguments[0];
                var v = arguments[1];
                s.value = v;
                s.dispatchEvent(new Event('change', { bubbles: true }));
            """, sel, val)
            write_log(
                f"✅ Set <select> value by JS -> value={val} for IMEI={imei}")
            time.sleep(0.5)
            return True

    # 2) fallback: cari chzn container / dropdown dan klik item yang ada teks IMEI
    select_id = sel.get_attribute("id") or sel.get_attribute(
        "name") or f"field_{field_id}"
    # beberapa pattern: ..._chzn atau view_XXXX_field_1136_chzn
    chzn_xpath = f"//div[contains(@id, '{select_id}') and contains(@class,'chzn-container')] | //div[contains(@id, 'field_{field_id}') and contains(@class,'chzn-container')]"
    try:
        chzn = driver.find_element(By.XPATH, chzn_xpath)
        # buka dropdown (klik input/search)
        try:
            chzn_clickable = chzn.find_element(
                By.XPATH,
                ".//a | .//div[contains(@class,'chzn-single')] | .//input")
            driver.execute_script("arguments[0].click();", chzn_clickable)
        except Exception:
            try:
                chzn.click()
            except Exception:
                pass
        time.sleep(0.5)
        # cari item li yang mengandung IMEI
        items = driver.find_elements(
            By.XPATH,
            f"{chzn_xpath}//li[contains(., '{imei}')] | //ul[contains(@class,'chzn-results')]//li[contains(., '{imei}')]"
        )
        for it in items:
            try:
                driver.execute_script("arguments[0].scrollIntoView(true);", it)
                it.click()
                write_log(f"✅ Klik pilihan di chzn untuk IMEI={imei}")
                time.sleep(0.5)
                return True
            except Exception:
                continue
    except Exception:
        pass

    write_log("❌ choose: tidak menemukan opsi yang cocok untuk IMEI")
    return False


def get_chzn_input(driver, field_id=1136, timeout=6):
    """
    Cari container chzn untuk field_id dan return dict:
     { "chzn": element, "input": input_element, "visible_options": [texts...] }
    """
    try:
        xpath = f"//div[contains(@id, 'field_{field_id}') and contains(@class, 'chzn-container')]"
        chzn = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath)))
        # input tempat ngetik ada di li.search-field > input
        try:
            chzn_input = chzn.find_element(
                By.XPATH, ".//li[contains(@class,'search-field')]//input")
        except Exception:
            chzn_input = chzn.find_element(
                By.XPATH, ".//input[@type='text' or @type='search']")

        # ambil opsi dropdown (bila sudah ada di DOM)
        opts = []
        try:
            items = chzn.find_elements(
                By.XPATH,
                ".//div[contains(@class,'chzn-drop')]//li[contains(@class,'active-result') or contains(@class,'result')]"
            )
            for it in items:
                t = (it.text or "").strip()
                if t:
                    opts.append(t)
        except Exception:
            pass

        return {"chzn": chzn, "input": chzn_input, "visible_options": opts}
    except Exception:
        return None


def type_into_chzn_and_wait(driver, chzn_input, imei, wait_timeout=6):
    """
    Ketik IMEI ke chzn_input lalu tunggu muncul active-result.
    Return list teks hasil dropdown yang muncul.
    """
    chzn_input.clear()
    for ch in str(imei):
        chzn_input.send_keys(ch)
        time.sleep(0.05)
    write_log(f"⌨️ IMEI diketik di chzn input: {imei}")

    # tunggu active-result muncul di seluruh page (lebih robust)
    try:
        WebDriverWait(driver, wait_timeout).until(
            EC.visibility_of_element_located(
                (By.XPATH,
                 "//li[contains(@class,'active-result') and normalize-space()]"
                 )))
        elems = driver.find_elements(
            By.XPATH,
            "//li[contains(@class,'active-result') and normalize-space()]")
        texts = [e.text.strip() for e in elems if e.text.strip()]
        return texts
    except Exception:
        return []


def choose_rental_stock_by_imei_chzn(driver, field_id, imei, timeout=8):
    """
    Robust pilih opsi dari chzn (chosen) berdasarkan IMEI:
    - clear default text di input chzn
    - klik/fokus supaya dropdown muncul
    - ketik IMEI, tunggu active-result, klik item
    - dispatch change ke <select> sebagai backup
    """
    try:
        sel = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((
                By.XPATH,
                f"//select[contains(@id,'field_{field_id}') or contains(@name,'field_{field_id}')]"
            )))
    except Exception:
        write_log("❌ choose_chzn: <select> tidak ditemukan")
        return False

    # cari chzn container & input
    try:
        chzn = driver.find_element(
            By.XPATH,
            f"//div[contains(@id,'field_{field_id}') and contains(@class,'chzn-container')]"
        )
    except Exception:
        # fallback cari berdasarkan select id
        select_id = sel.get_attribute("id") or sel.get_attribute(
            "name") or f"field_{field_id}"
        try:
            chzn = driver.find_element(
                By.XPATH,
                f"//div[contains(@id,'{select_id}') and contains(@class,'chzn-container')]"
            )
        except Exception:
            write_log("❌ choose_chzn: chzn container gak ketemu")
            return False

    try:
        chzn_input = chzn.find_element(
            By.XPATH, ".//li[contains(@class,'search-field')]//input")
    except Exception:
        try:
            chzn_input = chzn.find_element(
                By.XPATH, ".//input[@type='text' or @type='search']")
        except Exception:
            write_log("❌ choose_chzn: input dalam chzn gak ketemu")
            return False

    # clear default value & remove placeholder class supaya input benar-benar kosong
    try:
        driver.execute_script("arguments[0].value = '';", chzn_input)
        driver.execute_script(
            "if(arguments[0].classList) arguments[0].classList.remove('default');",
            chzn_input)
    except Exception:
        pass

    # fokus / klik input supaya chzn-drop bisa terbuka
    try:
        chzn_input.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].focus();", chzn_input)
        except Exception:
            pass
    time.sleep(0.15)

    # ketik IMEI pelan
    for ch in str(imei):
        try:
            chzn_input.send_keys(ch)
        except Exception:
            driver.execute_script("arguments[0].value += arguments[1];",
                                  chzn_input, ch)
        time.sleep(0.04)
    write_log(f"⌨️ IMEI diketik di chzn: {imei}")

    # pastikan dropdown ditampilkan (beberapa chzn render offscreen)
    try:
        driver.execute_script(
            """
            var c = arguments[0];
            var d = c.querySelector('.chzn-drop');
            if(d) d.style.left = '0px';
        """, chzn)
    except Exception:
        pass
    time.sleep(0.15)

    # tunggu active-result yang cocok muncul
    try:
        item_xpath = f"//div[contains(@id,'field_{field_id}') and contains(@class,'chzn-container')]//li[contains(@class,'active-result') and normalize-space() = '{imei}'] | //ul[contains(@class,'chzn-results')]//li[contains(@class,'active-result') and normalize-space() = '{imei}']"
        WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((By.XPATH, item_xpath)))
        item = driver.find_element(By.XPATH, item_xpath)
        driver.execute_script("arguments[0].scrollIntoView(true);", item)
        try:
            item.click()
        except Exception:
            driver.execute_script("arguments[0].click();", item)
        write_log(f"✅ Klik pilihan chzn untuk IMEI={imei}")
    except Exception:
        write_log("❌ Pilihan IMEI gak muncul di dropdown chzn")
        return False

    # dispatch change ke <select> sebagai backup (sinkronisasi)
    try:
        # cari option yang textnya IMEI
        try:
            opt = sel.find_element(By.XPATH,
                                   f".//option[normalize-space() = '{imei}']")
            val = opt.get_attribute("value")
            if val:
                driver.execute_script(
                    "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
                    sel, val)
                write_log(f"🔁 Dispatch change ke <select>, value={val}")
        except Exception:
            # kalau option text gak ada, coba ambil first option yg ada (fallback)
            opts = sel.find_elements(By.TAG_NAME, "option")
            if opts:
                v = opts[0].get_attribute("value")
                if v:
                    driver.execute_script(
                        "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
                        sel, v)
                    write_log("⚠️ Dispatch change fallback ke <select>")
    except Exception:
        pass

    time.sleep(0.3)
    return True


def create_awb_if_resi_missing(driver, uuid, tanggal_kirim, resi_target, imei,
                               status_modem, ekspedisi):
    try:
        awb_xpath = f'//a[@href="https://jvm.knack.com/app#main-form2/view-order-details12/{uuid}/return/{uuid}"]/span'
        awb_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, awb_xpath)))
        awb_link.click()
        write_log(f"📝 AWB link diklik: {awb_xpath}")
        time.sleep(2)

        tanggal_kirim_knack = convert_tanggal_sheet_to_knack(tanggal_kirim)
        tanggal_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="view_1899-field_888"]')))
        tanggal_input.clear()
        tanggal_input.send_keys(tanggal_kirim_knack)
        write_log(f"✅ Tanggal Pengembalian diisi: {tanggal_kirim_knack}")
        time.sleep(2)
        try:
            driver.execute_script("arguments[0].blur();", tanggal_input)
        except Exception:
            pass
        try:
            driver.find_element(By.TAG_NAME, "body").click()
        except Exception:
            pass
        try:
            tanggal_input.send_keys(Keys.ESCAPE)
        except Exception:
            pass
        time.sleep(1)

        # Ekspedisi (ambil dari sheet)
        ekspedisi_select = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((
                By.XPATH,
                '//select[contains(@id,"field_890") and contains(@name,"field_890")]'
            )))
        found_ekspedisi = False
        for option in ekspedisi_select.find_elements(By.TAG_NAME, "option"):
            if ekspedisi.strip().lower() in option.text.strip().lower():
                ekspedisi_select.send_keys(option.text)
                write_log(f"✅ Ekspedisi dipilih: {option.text}")
                found_ekspedisi = True
                break
        if not found_ekspedisi:
            write_log(
                f"❌ Ekspedisi '{ekspedisi}' tidak ditemukan di dropdown, pilih default"
            )
            ekspedisi_select.send_keys(
                option.text)  # fallback: pilih opsi terakhir

        # Nomor Resi (ambil dari sheet)
        resi_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((
                By.XPATH,
                '//input[contains(@id,"field_892") and contains(@name,"field_892")]'
            )))
        resi_input.clear()
        resi_input.send_keys(resi_target)
        write_log(f"✅ Nomor Resi diisi: {resi_target}")

        # Submit
        submit_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//button[contains(text(),"Submit")]')))
        driver.execute_script("arguments[0].scrollIntoView(true);", submit_btn)
        submit_btn.click()
        write_log("📤 Tombol Submit AWB diklik")
        time.sleep(5)  # jeda agar data resi baru sempat masuk ke detail order

    except Exception as e:
        write_log(f"❌ Error saat membuat AWB: {type(e).__name__} - {e}")
        return False
    return True


def update_sheet(order_id, imei):
    # Path ke file credentials JSON
    creds = Credentials.from_service_account_file(
        GOOGLE_CREDS_PATH,
        scopes=[
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ])
    gc = gspread.authorize(creds)
    sh = gc.open(
        "Modem Simcard IN & Out #Give to Product")  # Nama file Google Sheet
    ws = sh.worksheet("Form Return")  # Nama tab/sheet

    # Ambil semua data
    data = ws.get_all_records()
    for row in data:
        if str(row.get("ORDER ID",
                       "")).strip() == str(order_id).strip() and str(
                           row.get("IMEI", "")).strip() == str(imei).strip():
            # Ambil semua kolom sesuai header
            logs = row.get("Logs", "")
            timestamp = row.get("TimeStamp", "")
            pn = row.get("PN", "")
            status = row.get("STATUS", "")
            ekspedisi = row.get("EKSPEDISI", "")
            resi = row.get("NO. RESI", "")
            tgl_kirim = row.get("TGL KIRIM", "")
            tgl_terima = row.get("TGL TERIMA", "")

            # Cari baris kosong berikutnya
            next_row = len(ws.get_all_values()) + 1

            # Tulis data ke baris berikutnya
            ws.update(f"A{next_row}:J{next_row}", [[
                logs, timestamp, imei, pn, status, order_id, ekspedisi, resi,
                tgl_kirim, tgl_terima
            ]])
            return True
    # Jika tidak ketemu, log error
    write_log(
        f"❌ Data order_id={order_id} dan imei={imei} tidak ditemukan di Google Sheet"
    )
    return False


# Contoh pemakaian:
# update_sheet(TARGET_ORDER_ID, RENTAL_ID)


def get_ekspedisi_from_sheet(order_id, sheet_name="Form Return"):
    # Path ke file credentials.json
    creds = Credentials.from_service_account_file(
        GOOGLE_CREDS_PATH,
        scopes=[
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ])
    gc = gspread.authorize(creds)
    sh = gc.open(
        "Modem Simcard IN & Out #Give to Product")  # Nama file Google Sheet
    ws = sh.worksheet("Form Return")

    # Ambil semua data
    data = ws.get_all_records()
    for row in data:
        if str(row.get("ORDER ID", "")).strip() == str(order_id).strip():
            ekspedisi = row.get("EKSPEDISI", "")
            return ekspedisi
    return ""  # Jika tidak ketemu


# Contoh pemakaian:
# ekspedisi = get_ekspedisi_from_sheet(TARGET_ORDER_ID)
# print(ekspedisi)


def get_order_data_from_sheet(order_id, sheet_name="Form Return"):
    creds = Credentials.from_service_account_file(
        GOOGLE_CREDS_PATH,
        scopes=[
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ])
    gc = gspread.authorize(creds)
    sh = gc.open("Modem Simcard IN & Out #Give to Product")
    ws = sh.worksheet("Form Return")
    data = ws.get_all_records()
    for row in data:
        if str(row.get("ORDER ID", "")).strip() == str(order_id).strip():
            return row
    return None


# Contoh pemakaian:
# order_data = get_order_data_from_sheet(TARGET_ORDER_ID)
# if order_data:
#     write_log(f"📦 Data order ditemukan: {order_data}")
# else:
#     write_log("❌ Data order tidak ditemukan")


def convert_tanggal_sheet_to_knack(tanggal_sheet):
    """
    Convert '15 Oktober 2025' -> '15/10/2025'
    """
    bulan_map = {
        "Januari": "01",
        "Februari": "02",
        "Maret": "03",
        "April": "04",
        "Mei": "05",
        "Juni": "06",
        "Juli": "07",
        "Agustus": "08",
        "September": "09",
        "Oktober": "10",
        "November": "11",
        "Desember": "12"
    }
    try:
        parts = tanggal_sheet.strip().split()
        if len(parts) == 3:
            hari, bulan_str, tahun = parts
            bulan = bulan_map.get(bulan_str, "01")
            return f"{hari.zfill(2)}/{bulan}/{tahun}"
    except Exception:
        pass
    return tanggal_sheet  # fallback: return as is


def update_log_sheet(ws, imei, log_msg):
    """
    Update kolom Logs dan TimeStamp pada baris yang IMEI-nya sama.
    """
    import pytz  # Tambahkan ini
    jakarta_tz = pytz.timezone('Asia/Jakarta')  # Tambahkan ini
    now = datetime.now(jakarta_tz).strftime("%d/%m/%Y %H:%M:%S")  # Ubah ini
    values = ws.get_all_values()
    for i, row in enumerate(values, start=1):
        if len(row) > 2 and str(row[2]).strip() == str(imei).strip():
            ws.update(f"A{i}:B{i}", [[log_msg, now]])
            write_log(f"DEBUG: update_log_sheet selesai (baris {i})")
            return True
    write_log(f"❌ Gagal update log: IMEI {imei} tidak ditemukan di sheet")
    return False


# Di dalam create_awb_if_resi_missing:
# tanggal_kirim_knack = convert_tanggal_sheet_to_knack(tanggal_kirim)
# tanggal_input.clear()
# tanggal_input.send_keys(tanggal_kirim_knack)
# write_log(f"✅ Tanggal Pengembalian diisi: {tanggal_kirim_knack}")


def update_worker_status(ws, status):
    """
    Update status worker di kolom R baris 1.
    """
    ws.update('R1', [[status]])
    write_log(f"📝 Worker status di sheet diupdate: {status}")


def update_worker_heartbeat(ws):
    """
    Update kolom R1 dengan timestamp Asia/Jakarta setiap polling.
    """
    now = datetime.now(JAKARTA_TZ).strftime('%Y-%m-%d %H:%M:%S')
    ws.update('R1', [[now]])


if __name__ == "__main__":
    write_log("🚀 Bot mulai jalan...")
    driver = start_driver()
    try:
        if not login_knack(driver, LOGIN_URL, EMAIL, PASSWORD):
            write_log("❌ Tidak bisa login, abort.")
            driver.quit()
            exit(1)
        write_log("✅ Login selesai, mulai polling Google Sheet...")

        failed_imeis = set()

        while True:
            try:
                creds = Credentials.from_service_account_file(
                    GOOGLE_CREDS_PATH,
                    scopes=[
                        'https://www.googleapis.com/auth/spreadsheets',
                        'https://www.googleapis.com/auth/drive'
                    ])
                gc = gspread.authorize(creds)
                sh = gc.open("Modem Simcard IN & Out #Give to Product")
                ws = sh.worksheet("Form Return")
                update_worker_status(ws, "Running 🟢")
                update_worker_heartbeat(ws)  # <-- Tambahkan ini
                data = ws.get_all_records()
                write_log(f"🔄 Mulai polling, total order: {len(data)}")
            except Exception as e:
                write_log(
                    f"❌ Error akses Google Sheet: {type(e).__name__} - {e}")
                time.sleep(10)
                continue
            order_processed = False

            for idx, order_data in enumerate(data, start=2):
                imei = str(order_data.get("IMEI", "")).strip()
                logs = str(order_data.get("Logs", "")).strip()
                timestamp = order_data.get("TimeStamp", "")
                order_id = order_data.get("ORDER ID", "")
                if not imei:
                    continue  # skip IMEI kosong

                # Hanya skip jika Logs mengandung kata "berhasil"
                if "berhasil" in logs.lower():
                    continue  # sudah sukses, skip

                # Kalau belum berhasil (Logs kosong atau error), proses ulang
                write_log(
                    f"🔎 PROSES baris {idx}: imei='{imei}', order_id='{order_id}'"
                )
                TARGET_ORDER_ID = order_data["ORDER ID"]
                TARGET_RESI = order_data["NO. RESI"]
                MODEM_STATUS = order_data["STATUS"]
                RENTAL_ID = order_data["IMEI"]
                EKSPEDISI = order_data["EKSPEDISI"]
                TGL_KIRIM = order_data["TGL KIRIM"]
                TGL_TERIMA = order_data["TGL TERIMA"]

                try:
                    write_log("📄 Mulai scrape order detail")
                    # HAPUS BARIS INI:
                    # write_log(f"🔍 Order ID diketik: {TARGET_ORDER_ID}")

                    uuid = search_and_open_detail(driver, TARGET_URL,
                                                  TARGET_ORDER_ID, RENTAL_ID)
                    write_log(
                        f"📄 Selesai search_and_open_detail, UUID: {uuid}")

                    if uuid:
                        write_log(f"🎯 Match ditemukan: {TARGET_ORDER_ID}")

                        found = scrape_nomor_resi(driver, TARGET_RESI)
                        write_log("✅ Selesai scrape_nomor_resi")

                        if found:
                            write_log(
                                f"✅ Nomor Resi cocok dengan target: {TARGET_RESI}"
                            )
                        else:
                            write_log(
                                f"❌ Nomor Resi tidak cocok. Target: {TARGET_RESI}"
                            )

                        if not found:
                            write_log("📄 Mulai create_awb_if_resi_missing")
                            create_awb_if_resi_missing(driver, uuid, TGL_KIRIM,
                                                       TARGET_RESI, RENTAL_ID,
                                                       MODEM_STATUS, EKSPEDISI)
                            write_log("📤 Tombol Submit diklik (AWB)")
                            found = scrape_nomor_resi(driver, TARGET_RESI)
                            write_log(
                                "✅ Selesai create_awb_if_resi_missing dan scrape_nomor_resi"
                            )
                        if found:
                            write_log(
                                "📄 Mulai click_awb_icon_by_resi dan isi_form_return"
                            )
                            click_awb_icon_by_resi(driver, TARGET_RESI)
                            if all([MODEM_STATUS, RENTAL_ID, TGL_TERIMA]):
                                tgl_terima_knack = convert_tanggal_sheet_to_knack(
                                    TGL_TERIMA)
                                form_ok = isi_form_return(
                                    driver, MODEM_STATUS, RENTAL_ID,
                                    tgl_terima_knack)
                                if form_ok:
                                    update_log_sheet(
                                        ws, RENTAL_ID,
                                        f"✅ Return IMEI {RENTAL_ID} telah berhasil"
                                    )
                                    write_log(
                                        f"✅ Order {TARGET_ORDER_ID} selesai, kolom Logs & TimeStamp direset"
                                    )
                                else:
                                    update_log_sheet(
                                        ws, RENTAL_ID,
                                        f"❌ Gagal isi form return untuk IMEI {RENTAL_ID} di Order ID {TARGET_ORDER_ID}"
                                    )
                                    write_log(
                                        f"❌ Gagal isi form return untuk IMEI {RENTAL_ID} di Order ID {TARGET_ORDER_ID}"
                                    )
                                order_processed = True

                                # Tutup tab detail order & kembali ke main form
                                try:
                                    if len(driver.window_handles) > 1:
                                        driver.close()
                                        driver.switch_to.window(
                                            driver.window_handles[0])
                                        write_log(
                                            "🗂️ Tab detail order ditutup, kembali ke tab utama"
                                        )
                                    driver.get(TARGET_URL)
                                    write_log(
                                        "🏠 Kembali ke halaman Orders Main Form"
                                    )
                                    time.sleep(2)
                                except Exception as e:
                                    write_log(
                                        f"❌ Gagal kembali ke halaman utama: {type(e).__name__} - {e}"
                                    )
                    else:
                        write_log(
                            f"❌ UUID tidak ditemukan untuk order {TARGET_ORDER_ID}"
                        )
                        update_log_sheet(
                            ws, RENTAL_ID,
                            f"❌ UUID tidak ditemukan untuk order {TARGET_ORDER_ID}"
                        )
                except Exception as e:
                    write_log(
                        f"❌ Error di order {TARGET_ORDER_ID}: {type(e).__name__} - {e}"
                    )
                    update_log_sheet(
                        ws, RENTAL_ID,
                        f"❌ Error di order {TARGET_ORDER_ID}: {type(e).__name__} - {e}"
                    )

            if not order_processed:
                write_log("⏳ Tidak ada order baru, menunggu 10 detik...")
                time.sleep(10)  # polling interval

    except KeyboardInterrupt:
        try:
            creds = Credentials.from_service_account_file(
                GOOGLE_CREDS_PATH,
                scopes=[
                    'https://www.googleapis.com/auth/spreadsheets',
                    'https://www.googleapis.com/auth/drive'
                ])
            gc = gspread.authorize(creds)
            sh = gc.open("Modem Simcard IN & Out #Give to Product")
            ws = sh.worksheet("Form Return")
            update_worker_status(ws, "Stopped 🟥")  # Tambahkan ini
        except Exception:
            pass
        write_log("🛑 Bot dihentikan manual (Ctrl+C)")
    finally:
        input("⏳ Tekan Enter untuk menutup browser setelah selesai...")
        driver.quit()
