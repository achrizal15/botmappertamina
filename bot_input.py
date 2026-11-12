import asyncio
import json
import os
import random
import re
import sys
import time
import urllib.request
from datetime import datetime
from pathlib import Path
import openpyxl
from openpyxl.styles import PatternFill
from playwright.async_api import async_playwright

MONTH_NAMES = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]

REGIONAL_CACHE = {
    "3510": "Banyuwangi",
    "1671": "Palembang"
}

def get_tempat_lahir(nik):
    prov_kab = nik[:4]
    if prov_kab in REGIONAL_CACHE:
        return REGIONAL_CACHE[prov_kab]
    try:
        url = f"https://regional-indonesia.vercel.app/api/wilayah?kode={prov_kab}"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=3) as resp:
            data = json.loads(resp.read().decode())
            name = data.get("nama") or data.get("kabupaten_kota")
            if name:
                clean_name = name.replace("KABUPATEN ", "").replace("KOTA ", "").title()
                REGIONAL_CACHE[prov_kab] = clean_name
                return clean_name
    except Exception:
        pass
    return "Banyuwangi"

def parse_nik_dob(nik):
    raw_day = int(nik[6:8])
    month = int(nik[8:10])
    raw_year = int(nik[10:12])
    
    day = raw_day - 40 if raw_day > 40 else raw_day
    curr_yr = datetime.now().year
    year = 2000 + raw_year
    if year > (curr_yr - 17):
        year = 1900 + raw_year
        
    return str(day), month, str(year)

CONFIG_PATH = Path("config.json")
DEFAULT_CONFIG = {
    "phone": "085234104446",
    "pin": "344311",
    "excel_file": "template_nik.xlsx",
    "headless": False,
    "delay_min": 4.0,
    "delay_max": 6.0
}

def load_config():
    if CONFIG_PATH.exists():
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
            return {**DEFAULT_CONFIG, **cfg}
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(DEFAULT_CONFIG, f, indent=2)
    return DEFAULT_CONFIG

def parse_date(val):
    if not val:
        return datetime.min
    if isinstance(val, datetime):
        return val
    val_str = str(val).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"):
        try:
            return datetime.strptime(val_str, fmt)
        except ValueError:
            pass
    return datetime.min

def safe_save_workbook(wb, path, retries=5, delay=2.0):
    path = Path(path)
    for attempt in range(1, retries + 1):
        try:
            wb.save(path)
            return True
        except PermissionError:
            print(f"\n[!] PERINGATAN: File '{path.name}' sedang terbuka di Microsoft Excel / aplikasi lain!")
            print(f"[!] Silakan TUTUP file Excel tersebut agar bot bisa menyimpan progres. Mencoba ulang ({attempt}/{retries})...")
            time.sleep(delay)
        except Exception as e:
            print(f"[-] Terjadi kesalahan saat menyimpan Excel: {e}")
            break

    backup_path = path.parent / f"{path.stem}_backup{path.suffix}"
    try:
        wb.save(backup_path)
        print(f"[+] Data berhasil diselamatkan ke file backup: '{backup_path.name}'")
        return True
    except Exception as e:
        print(f"[-] Gagal menyimpan ke backup: {e}")
        return False

async def wait_for_dashboard_ready(page, delay=4.0):
    """
    Memberi jeda wajib setelah login agar API stok tabung Pertamina selesai dimuat.
    Mencegah bug 'Stok Tabung Kosong' palsu.
    """
    print(f"[*] Menunggu sinkronisasi stok dan sistem MAP...")
    try:
        await page.wait_for_selector('text="Tabung"', timeout=12000)
    except Exception:
        pass

    await asyncio.sleep(delay)

    stock_text = await page.evaluate('''() => {
        const el = Array.from(document.querySelectorAll('*')).find(e => 
            e.textContent?.includes('Tabung') && !e.textContent?.includes('Kosong') && e.children.length === 0
        );
        return el ? el.textContent.trim() : "";
    }''')
    if stock_text:
        print(f"[+] Sisa Stok Pangkalan Terdeteksi: {stock_text}")

async def login(page, phone, pin):
    print(f"[*] Membuka portal MAP Pertamina...")
    await page.goto("https://subsiditepatlpg.mypertamina.id/merchant-login", wait_until="domcontentloaded")
    await asyncio.sleep(2)
    
    if "/merchant/app" in page.url:
        print("[+] Sesi login aktif, langsung ke dashboard.")
        await wait_for_dashboard_ready(page, delay=3.5)
        return True

    print(f"[*] Mengisi No. Ponsel: {phone}")
    phone_input = page.locator('input[placeholder*="Nomor Ponsel"], input[placeholder*="Ponsel"]').first
    await phone_input.wait_for(state="visible", timeout=15000)
    await phone_input.fill(phone)

    print(f"[*] Mengisi PIN...")
    pin_input = page.locator('input[placeholder*="PIN"], input[type="password"]').first
    await pin_input.fill(pin)
    await asyncio.sleep(0.5)

    masuk_btn = page.locator('button:has-text("MASUK")').first
    await masuk_btn.click()
    print("[*] Tombol MASUK diklik. Menunggu login...")

    try:
        await page.wait_for_url("**/merchant/app**", timeout=25000)
        print("[+] LOGIN BERHASIL! Masuk ke dashboard MAP.")
        await wait_for_dashboard_ready(page, delay=4.5)
        return True
    except Exception as e:
        print(f"[-] Gagal masuk ke dashboard: {e}")
        return False

async def close_any_modal(page):
    """Menutup modal yang mungkin masih terbuka di layar."""
    try:
        dialog = page.locator('[role="dialog"]').first
        if await dialog.is_visible():
            d_text = await dialog.text_content()
            if d_text and "MASUKKAN NIK KTP" not in d_text and "Pernyataan Persetujuan" not in d_text and "Centang" not in d_text and "Data Pelanggan belum lengkap" not in d_text and "Lengkapi Data Pelanggan" not in d_text and "Pastikan semua data" not in d_text:
                tutup_btn = dialog.locator('button:has-text("TUTUP"), .mantine-Modal-close').first
                if await tutup_btn.is_visible():
                    await tutup_btn.click()
                    await asyncio.sleep(0.5)
    except Exception:
        pass

async def bypass_puzzle_captcha(page):
    """
    Menyelesaikan captcha puzzle secara otomatis (100% tanpa sentuh mouse manual).
    """
    print("[*] Menunggu modal captcha...")
    modal_opened = False
    for _ in range(30):
        await asyncio.sleep(0.2)
        if await page.locator('text="Cocokan Gambar untuk Proses Keamanan Penjualan"').is_visible():
            modal_opened = True
            break
        if "struk" in page.url or await page.locator('text="LUNAS"').is_visible():
            print("[+] Transaksi langsung sukses tanpa captcha!")
            return True

    if not modal_opened:
        print("[!] Modal captcha tidak muncul, lanjut...")
        return True

    print("[+] MODAL CAPTCHA TERDETEKSI!")
    print("[*] Mengeksekusi bypass captcha otomatis...")

    for attempt in range(25):
        await asyncio.sleep(0.2)
        
        result = await page.evaluate('''() => {
            const elements = [
                document.querySelector('.rc-slider-captcha-jigsaw-bg'),
                document.querySelector('.rc-slider-captcha-control-button'),
                document.querySelector('.rc-slider-captcha'),
                document.querySelector('[role="dialog"]')
            ].filter(Boolean);

            function getFiber(dom) {
                const key = Object.keys(dom || {}).find(k => k.startsWith('__reactFiber$'));
                return key ? dom[key] : null;
            }

            for (const el of elements) {
                let curr = getFiber(el);
                while (curr) {
                    if (curr.memoizedProps) {
                        if (typeof curr.memoizedProps.onSuccessCaptcha === 'function') {
                            curr.memoizedProps.onSuccessCaptcha();
                            return { triggered: true, fn: 'onSuccessCaptcha' };
                        }
                        if (typeof curr.memoizedProps.onSuccess === 'function') {
                            curr.memoizedProps.onSuccess();
                            return { triggered: true, fn: 'onSuccess' };
                        }
                    }
                    curr = curr.return;
                }
            }
            return { triggered: false };
        }''')

        if result.get("triggered"):
            print(f"[+] CAPTCHA BERHASIL DITEMBUS via trigger {result.get('fn')}!")
            return True

        slider_btn = page.locator('.rc-slider-captcha-control-button').first
        if await slider_btn.is_visible():
            expected_x = await page.evaluate('''() => {
                const bg = document.querySelector('.rc-slider-captcha-jigsaw-bg');
                if (!bg) return null;
                function getFiber(dom) {
                    const key = Object.keys(dom || {}).find(k => k.startsWith('__reactFiber$'));
                    return key ? dom[key] : null;
                }
                let curr = getFiber(bg);
                while (curr) {
                    let s = curr.memoizedState;
                    while (s) {
                        if (s.memoizedState && typeof s.memoizedState === 'object' && 'expectedX' in s.memoizedState) {
                            return s.memoizedState.expectedX;
                        }
                        s = s.next;
                    }
                    curr = curr.return;
                }
                return null;
            }''')

            if expected_x:
                print(f"[*] Target koordinat puzzle: {expected_x} px. Menggeser slider...")
                box = await slider_btn.bounding_box()
                if box:
                    start_x = box["x"] + box["width"] / 2
                    start_y = box["y"] + box["height"] / 2
                    await page.mouse.move(start_x, start_y)
                    await page.mouse.down()
                    steps = 15
                    for i in range(1, steps + 1):
                        curr_x = start_x + (expected_x * i / steps)
                        await page.mouse.move(curr_x, start_y)
                        await asyncio.sleep(0.01)
                    await page.mouse.up()
                    print("[+] Slider selesai digeser otomatis!")
                    return True

    return False

async def handle_update_data_customer(page, nik):
    """
    Mengisi form 'Lengkapi Data Pelanggan di bawah ini' di /app/update-data-customer:
    1. Tempat Lahir: nama kabupaten/kota berdasarkan NIK
    2. Tanggal Lahir: tanggal, bulan, tahun dipecah dari NIK
    3. Submit dan konfirmasi hingga masuk ke transaksi
    """
    print(f"[*] Memproses pengisian form 'Lengkapi Data Pelanggan' untuk NIK: {nik}...")
    tempat_lahir = get_tempat_lahir(nik)
    day_str, month_num, year_str = parse_nik_dob(nik)
    month_name = MONTH_NAMES[month_num - 1]
    print(f"[*] Data parsed dari NIK -> Tempat: {tempat_lahir}, Tgl: {day_str}, Bln: {month_name} ({month_num}), Thn: {year_str}")

    try:
        # 1. Isi Tempat Lahir
        pob_input = page.locator('input[name="placeOfBirth"], input[placeholder*="tempat lahir"]').first
        await pob_input.wait_for(state="visible", timeout=8000)
        await pob_input.fill(tempat_lahir)
        await asyncio.sleep(0.4)

        # 2. Isi Tanggal Lahir (Mantine Select)
        # Day (Tgl)
        day_input = page.locator('[data-testid="daySelect"] input, input[placeholder="Tgl"], input[name="dayOB"]').first
        if await day_input.is_visible():
            await day_input.click()
            await asyncio.sleep(0.3)
            day_opt = page.locator(f'[role="option"]:has-text("{day_str}")').first
            if await day_opt.is_visible():
                await day_opt.click()
            else:
                await day_input.fill(day_str)
                await day_input.press("Enter")
        await asyncio.sleep(0.3)

        # Month (Bln)
        month_input = page.locator('[data-testid="monthSelect"] input, input[placeholder="Bln"], input[name="monthOB"]').first
        if await month_input.is_visible():
            await month_input.click()
            await asyncio.sleep(0.3)
            month_opt = page.locator(f'[role="option"]:has-text("{month_name}")').first
            if await month_opt.is_visible():
                await month_opt.click()
            else:
                await month_input.fill(month_name)
                await month_input.press("Enter")
        await asyncio.sleep(0.3)

        # Year (Thn)
        year_input = page.locator('[data-testid="yearSelect"] input, input[placeholder="Thn"], input[name="yearOB"]').first
        if await year_input.is_visible():
            await year_input.click()
            await asyncio.sleep(0.3)
            year_opt = page.locator(f'[role="option"]:has-text("{year_str}")').first
            if await year_opt.is_visible():
                await year_opt.click()
            else:
                await year_input.fill(year_str)
                await year_input.press("Enter")
        await asyncio.sleep(0.5)

        # 3. Klik tombol SELANJUTNYA di form update
        submit_btn = page.locator('button:has-text("SELANJUTNYA"), button[data-testid="btnSubmitUpdate"]').first
        await submit_btn.click()
        print("[*] Tombol 'SELANJUTNYA' pada Update Data Pelanggan diklik.")
        await asyncio.sleep(1.2)

        # 4. Konfirmasi modal ("Pastikan semua data sudah benar") -> Klik YA, Perbarui DATA PELANGGAN
        confirm_btn = page.locator('button:has-text("YA, Perbarui DATA PELANGGAN"), button:has-text("Perbarui DATA PELANGGAN")').first
        try:
            await confirm_btn.wait_for(state="visible", timeout=6000)
            await confirm_btn.click()
            print("[*] Konfirmasi 'YA, Perbarui DATA PELANGGAN' diklik.")
            await asyncio.sleep(1.5)
        except Exception:
            pass

        # 5. Success modal ("Data Pelanggan berhasil diperbarui") -> Klik "LANJUTKAN KE TRANSAKSI"
        lanjut_btn = page.locator('button:has-text("LANJUTKAN KE TRANSAKSI")').first
        try:
            await lanjut_btn.wait_for(state="visible", timeout=8000)
            await lanjut_btn.click()
            print("[*] Tombol 'LANJUTKAN KE TRANSAKSI' diklik.")
            await asyncio.sleep(2.0)
        except Exception:
            pass

        # 6. Cek jika muncul popup "Segera Lengkapi NIB" (untuk Usaha Mikro)
        nib_modal = page.locator('text="Segera Lengkapi NIB"').first
        if await nib_modal.is_visible():
            print("[!] Muncul popup 'Segera Lengkapi NIB'.")
            lanjut_jual = page.locator('button:has-text("NANTI SAJA, LANJUT PENJUALAN")').first
            if await lanjut_jual.is_visible():
                await lanjut_jual.click()
                print("[*] Tombol 'NANTI SAJA, LANJUT PENJUALAN' diklik.")
                await asyncio.sleep(2.0)
            else:
                tutup_nib = page.locator('button:has-text("TUTUP")').first
                if await tutup_nib.is_visible():
                    await tutup_nib.click()
                    print("[*] Tombol 'TUTUP' pada Segera Lengkapi NIB diklik.")
                    await asyncio.sleep(1.0)
                await page.goto("https://subsiditepatlpg.mypertamina.id/merchant/app", wait_until="domcontentloaded")
                await asyncio.sleep(1.5)
                return "RETRY_FROM_APP"

        if "/merchant/app/sale" in page.url:
            return True
        return True

    except Exception as e:
        print(f"[-] Gagal mengisi form update data: {e}")
    return False

async def open_catat_penjualan_modal(page):
    """
    Membuka modal Catat Penjualan dengan penanganan jika muncul popup 'Stok Tabung Kosong' palsu.
    """
    for retry in range(3):
        await close_any_modal(page)
        
        print("[*] Mengklik 'Catat Penjualan'...")
        await page.evaluate('''() => {
            const el = Array.from(document.querySelectorAll('*')).find(e => 
                e.textContent?.trim() === "Catat Penjualan" && e.children.length === 0
            );
            if (el) el.click();
        }''')
        await asyncio.sleep(1.0)

        stok_kosong_modal = page.locator('text="Stok Tabung Kosong"').first
        if await stok_kosong_modal.is_visible():
            print("[!] Muncul peringatan 'Stok Tabung Kosong' (loading awal belum sinkron).")
            print("[*] Menutup popup dan memberi jeda 3 detik agar stok selesai termuat...")
            tutup_btn = page.locator('button:has-text("TUTUP")').first
            if await tutup_btn.is_visible():
                await tutup_btn.click()
            await asyncio.sleep(3.0)
            continue

        nik_input = page.locator('input[placeholder*="16 digit NIK"]').first
        if await nik_input.is_visible():
            return nik_input

    nik_input = page.locator('input[placeholder*="16 digit NIK"]').first
    await nik_input.wait_for(state="visible", timeout=6000)
    return nik_input

async def extract_customer_name(page):
    """
    Mengekstrak nama pelanggan dari berbagai kemungkinan posisi di layar:
    1. Modal 'Pelanggan Terdaftar'
    2. Halaman Penjualan (/sale)
    3. Halaman Struk (/sale/struk)
    """
    return await page.evaluate('''() => {
        // Cek 1: Modal Pelanggan Terdaftar
        const dialog = document.querySelector('[role="dialog"]');
        if (dialog) {
            const allEls = Array.from(dialog.querySelectorAll('*'));
            for (let i = 0; i < allEls.length; i++) {
                if (allEls[i].textContent?.trim() === "Nama Pelanggan") {
                    // Elemen berikutnya biasanya berisi nama
                    for (let j = i + 1; j < Math.min(i + 5, allEls.length); j++) {
                        const txt = allEls[j].textContent?.trim();
                        if (txt && txt !== "Nama Pelanggan" && !txt.includes("NIK") && allEls[j].children.length === 0) {
                            return txt;
                        }
                    }
                }
            }
        }

        // Cek 2: Halaman Struk Lunas
        const strukEls = Array.from(document.querySelectorAll('*'));
        for (let i = 0; i < strukEls.length; i++) {
            if (strukEls[i].textContent?.trim() === "Nama Pelanggan") {
                for (let j = i + 1; j < Math.min(i + 5, strukEls.length); j++) {
                    const txt = strukEls[j].textContent?.trim();
                    if (txt && txt !== "Nama Pelanggan" && !txt.includes("Nomor KK") && strukEls[j].children.length === 0) {
                        return txt;
                    }
                }
            }
        }

        // Cek 3: Halaman Penjualan di atas badge kategori
        const catBadge = Array.from(document.querySelectorAll('*')).find(e => 
            (e.textContent?.trim() === "Rumah Tangga" || e.textContent?.trim() === "Usaha Mikro" || e.textContent?.trim() === "Pengecer") &&
            e.children.length === 0
        );
        if (catBadge && catBadge.parentElement) {
            const siblings = Array.from(catBadge.parentElement.children);
            for (const s of siblings) {
                const t = s.textContent?.trim();
                if (t && t !== "Rumah Tangga" && t !== "Usaha Mikro" && t !== "Pengecer" && !t.includes("Informasi Pelanggan")) {
                    return t;
                }
            }
            // Atau naik 1 level ke container
            const p = catBadge.parentElement.parentElement;
            if (p) {
                const pTexts = Array.from(p.querySelectorAll('div, span, p'))
                    .map(x => x.textContent.trim())
                    .filter(x => x && x !== "Rumah Tangga" && x !== "Usaha Mikro" && x !== "Pengecer" && x !== "Informasi Pelanggan" && x !== "Ganti Pelanggan" && x !== "Lihat Rincian Pelanggan");
                if (pTexts.length > 0) return pTexts[0];
            }
        }

        return "";
    }''')

async def process_single_nik(page, nik, cfg, captured_api_name):
    print(f"\n==================================================")
    print(f"[*] Input NIK: {nik}")
    print(f"==================================================")

    # Reset nama yang tertangkap dari API
    captured_api_name["name"] = ""
    await close_any_modal(page)

    if "/merchant/app/sale" in page.url or "/merchant/app" not in page.url:
        await page.goto("https://subsiditepatlpg.mypertamina.id/merchant/app", wait_until="domcontentloaded")
        await asyncio.sleep(1.5)

    # 1. Buka modal Catat Penjualan
    nik_input = await open_catat_penjualan_modal(page)

    # 2. Masukkan NIK dengan pengetikan alami (human-like)
    await nik_input.click()
    await nik_input.fill("")
    await nik_input.press_sequentially(nik, delay=random.randint(30, 55))
    await asyncio.sleep(random.uniform(0.5, 0.8))
    # 3. Klik "LANJUTKAN PENJUALAN"
    lanjut_btn = page.locator('button:has-text("LANJUTKAN PENJUALAN")').first
    await lanjut_btn.click()
    print("[*] NIK dikirim, menunggu respon sistem...")

    status_result = "UNKNOWN"
    customer_name = ""
    keterangan = ""

    await asyncio.sleep(1.0)

    # 4. Evaluasi respon NIK (Polling hingga ~8 detik)
    for _ in range(40):
        await asyncio.sleep(0.2)

        # PRIORITAS 1: Jika sudah berhasil pindah ke halaman Penjualan /merchant/app/sale
        if "/merchant/app/sale" in page.url:
            status_result = "TERDAFTAR"
            print(f"[+] HASIL: NIK TERDAFTAR! Masuk ke halaman Penjualan.")
            break

        # PRIORITAS 2: Cek popup pilihan jenis pelanggan ("Pelanggan Terdaftar" - Rumah Tangga / Usaha Mikro / Pengecer)
        customer_type_modal = page.locator('[role="dialog"]:has-text("pilihan jenis pelanggan"), [role="dialog"]:has-text("Pelanggan Terdaftar")').first
        if await customer_type_modal.is_visible():
            print("[*] Terdeteksi popup pilihan jenis pelanggan ('Pelanggan Terdaftar').")
            
            # Ambil nama pelanggan jika terlihat di modal ini
            modal_name = await extract_customer_name(page)
            if modal_name:
                customer_name = modal_name
                print(f"[+] Nama Pelanggan terdeteksi di modal: {customer_name}")

            print("[*] Memilih 'Rumah Tangga' sesuai instruksi...")
            
            # Pilih opsi Rumah Tangga
            await page.evaluate('''() => {
                const dialog = document.querySelector('[role="dialog"]');
                if (!dialog) return false;
                
                const radios = Array.from(dialog.querySelectorAll('input[type="radio"]'));
                if (radios.length > 0) {
                    radios[0].click();
                    return true;
                }
                
                const els = Array.from(dialog.querySelectorAll('*'));
                const rtEl = els.find(e => e.textContent?.trim() === "Rumah Tangga" && e.children.length === 0);
                if (rtEl) {
                    (rtEl.closest('[role="radio"]') || rtEl.closest('label') || rtEl.parentElement || rtEl).click();
                    return true;
                }
                return false;
            }''')
            
            await asyncio.sleep(0.6)
            
            lanjut_trx_btn = customer_type_modal.locator('button:has-text("LANJUTKAN PENJUALAN")').first
            if await lanjut_trx_btn.is_visible():
                await lanjut_trx_btn.click()
                print("[*] Tombol 'LANJUTKAN PENJUALAN' diklik setelah memilih Rumah Tangga.")
                await asyncio.sleep(1.2)
                if "/merchant/app/sale" in page.url:
                    status_result = "TERDAFTAR"
                    print(f"[+] HASIL: NIK TERDAFTAR! Masuk ke halaman Penjualan.")
                    break

        # PRIORITAS 3: Cek popup "Data Pelanggan belum lengkap" -> KLIK "UPDATE DATA PELANGGAN"
        incomplete_modal = page.locator('text="Data Pelanggan belum lengkap"').first
        if await incomplete_modal.is_visible():
            print("[*] Terdeteksi popup 'Data Pelanggan belum lengkap'. Mengklik 'UPDATE DATA PELANGGAN'...")
            update_btn = page.locator('button:has-text("UPDATE DATA PELANGGAN"), [data-testid="btnRegister"]').first
            if await update_btn.is_visible():
                await update_btn.click()
                await asyncio.sleep(1.5)
                if "update-data-customer" in page.url or await page.locator('text="Lengkapi Data Pelanggan di bawah ini"').is_visible():
                    updated = await handle_update_data_customer(page, nik)
                    if updated == "RETRY_FROM_APP":
                        print("[*] Mengulang input NIK dari halaman utama setelah data NIB ditutup...")
                        nik_input = await open_catat_penjualan_modal(page)
                        await nik_input.fill(nik)
                        await asyncio.sleep(0.5)
                        lanjut_btn = page.locator('button:has-text("LANJUTKAN PENJUALAN")').first
                        if await lanjut_btn.is_visible():
                            await lanjut_btn.click()
                        await asyncio.sleep(1.2)
                        continue
                    if updated and "/merchant/app/sale" in page.url:
                        status_result = "TERDAFTAR"
                        print(f"[+] HASIL: Data Pelanggan berhasil diperbarui! Masuk ke halaman Penjualan.")
                        break

        # PRIORITAS 3B: Cek jika langsung berada di halaman Update Data Pelanggan
        if "update-data-customer" in page.url or await page.locator('text="Lengkapi Data Pelanggan di bawah ini"').is_visible():
            updated = await handle_update_data_customer(page, nik)
            if updated == "RETRY_FROM_APP":
                print("[*] Mengulang input NIK dari halaman utama setelah data NIB ditutup...")
                nik_input = await open_catat_penjualan_modal(page)
                await nik_input.fill(nik)
                await asyncio.sleep(0.5)
                lanjut_btn = page.locator('button:has-text("LANJUTKAN PENJUALAN")').first
                if await lanjut_btn.is_visible():
                    await lanjut_btn.click()
                await asyncio.sleep(1.2)
                continue
            if updated and "/merchant/app/sale" in page.url:
                status_result = "TERDAFTAR"
                print(f"[+] HASIL: Data Pelanggan berhasil diperbarui! Masuk ke halaman Penjualan.")
                break

        # PRIORITAS 3C: Cek jika popup "Segera Lengkapi NIB" muncul
        nib_modal = page.locator('text="Segera Lengkapi NIB"').first
        if await nib_modal.is_visible():
            print("[!] Terdeteksi popup 'Segera Lengkapi NIB'.")
            lanjut_jual = page.locator('button:has-text("NANTI SAJA, LANJUT PENJUALAN")').first
            if await lanjut_jual.is_visible():
                await lanjut_jual.click()
                print("[*] Tombol 'NANTI SAJA, LANJUT PENJUALAN' diklik.")
                await asyncio.sleep(1.5)
            else:
                tutup_nib = page.locator('button:has-text("TUTUP")').first
                if await tutup_nib.is_visible():
                    await tutup_nib.click()
                    print("[*] Tombol 'TUTUP' pada Segera Lengkapi NIB diklik.")
                    await asyncio.sleep(1.0)
                await page.goto("https://subsiditepatlpg.mypertamina.id/merchant/app", wait_until="domcontentloaded")
                await asyncio.sleep(1.5)
                nik_input = await open_catat_penjualan_modal(page)
                await nik_input.fill(nik)
                await asyncio.sleep(0.5)
                lanjut_btn = page.locator('button:has-text("LANJUTKAN PENJUALAN")').first
                if await lanjut_btn.is_visible():
                    await lanjut_btn.click()
                await asyncio.sleep(1.2)
                continue

        # PRIORITAS 4: Cek Halaman / Modal "Pernyataan Persetujuan" -> Centang & KLIK "SELANJUTNYA"
        is_agreement_page = await page.evaluate('''() => {
            const body = document.body.innerText || "";
            return body.includes("Pernyataan Persetujuan") || 
                   body.includes("Centang") && body.includes("pernyataan di bawah ini") || 
                   window.location.pathname.includes("term-condition");
        }''')
        if is_agreement_page:
            print("[*] Terdeteksi halaman/modal 'Pernyataan Persetujuan'. Memastikan centang dan klik 'SELANJUTNYA'...")
            await page.evaluate('''() => {
                const checkboxes = Array.from(document.querySelectorAll('input[type="checkbox"]'));
                for (const cb of checkboxes) {
                    if (!cb.checked) cb.click();
                }
                const customChecks = Array.from(document.querySelectorAll('[role="checkbox"]'));
                for (const cc of customChecks) {
                    if (cc.getAttribute('aria-checked') !== 'true') cc.click();
                }
            }''')
            await asyncio.sleep(0.5)
            
            selanjutnya_btn = page.locator('button:has-text("SELANJUTNYA")').first
            if await selanjutnya_btn.is_visible():
                await selanjutnya_btn.click()
                print("[*] Tombol 'SELANJUTNYA' diklik.")
                await asyncio.sleep(1.5)
                if "/merchant/app/sale" in page.url:
                    status_result = "TERDAFTAR"
                    print(f"[+] HASIL: NIK TERDAFTAR! Masuk ke halaman Penjualan setelah Pernyataan Persetujuan.")
                    break
        # PRIORITAS 4: Cek Dialog Error Nyata
        dialog_info = await page.evaluate('''() => {
            const dialog = document.querySelector('[role="dialog"]');
            if (!dialog) return null;
            const text = dialog.innerText || "";
            if (text.includes("MASUKKAN NIK KTP") || text.includes("pilihan jenis pelanggan") || text.includes("Pernyataan Persetujuan") || text.includes("Centang") || text.includes("Data Pelanggan belum lengkap") || text.includes("Lengkapi Data Pelanggan") || text.includes("Pastikan semua data") || text.includes("berhasil diperbarui")) return null;

            const hasTutup = !!Array.from(dialog.querySelectorAll('button')).find(b => b.textContent?.includes('TUTUP'));
            return { text, hasTutup };
        }''')

        if dialog_info and dialog_info["hasTutup"]:
            d_text = dialog_info["text"]
            
            if "Pelanggan Tidak Terdaftar" in d_text:
                status_result = "BELUM TERDAFTAR"
                keterangan = "Pelanggan Tidak Terdaftar"
                print(f"[-] HASIL: {status_result} (Pelanggan Tidak Terdaftar)")
            
            elif "Tidak Dapat Melanjutkan" in d_text or "Terlalu banyak melakukan permintaan" in d_text:
                status_result = "BELUM TERDAFTAR"
                keterangan = "Limit Pendaftaran Harian"
                print(f"[-] HASIL: {status_result} ({keterangan})")

            elif "Tidak Padan" in d_text:
                status_result = "NIK BERMASALAH"
                keterangan = "NIK Tidak Padan"
                print(f"[-] HASIL: {status_result} ({keterangan})")
            
            else:
                first_line = d_text.split('\\n')[0].strip()
                status_result = "NIK BERMASALAH"
                keterangan = first_line[:40]
                print(f"[-] HASIL: {status_result} ({keterangan})")

            await page.evaluate('''() => {
                const dialog = document.querySelector('[role="dialog"]');
                if (dialog) {
                    const btn = Array.from(dialog.querySelectorAll('button')).find(b => b.textContent?.includes('TUTUP'));
                    if (btn) btn.click();
                }
            }''')
            await asyncio.sleep(0.5)
            if "Terlalu banyak melakukan permintaan" in d_text or "Tidak Dapat Melanjutkan" in d_text:
                print("[!] Menunggu jeda pendinginan (cooldown) 6 detik...")
                await asyncio.sleep(6.0)
            break

        alert_error = page.locator('.mantine-Alert-root, [role="alert"]').first
        if await alert_error.is_visible():
            msg = await alert_error.text_content()
            if msg and ("gagal" in msg.lower() or "salah" in msg.lower()):
                status_result = "NIK BERMASALAH"
                keterangan = msg.strip()
                print(f"[-] HASIL: {status_result} ({keterangan})")
                break

    # Jika NIK Terdaftar, cek kuota dan proses transaksi sampai tuntas
    if status_result == "TERDAFTAR":
        # Ambil nama pelanggan: prioritaskan hasil API, lalu DOM
        if captured_api_name["name"]:
            customer_name = captured_api_name["name"]
        elif not customer_name:
            dom_name = await extract_customer_name(page)
            if dom_name:
                customer_name = dom_name

        if customer_name:
            print(f"[+] Nama Pelanggan: {customer_name}")

        # ----------------------------------------------------
        # CEK APAKAH KUOTA HABIS / SUDAH LIMIT (Full)
        # ----------------------------------------------------
        await asyncio.sleep(0.8)
        is_limit = await page.evaluate('''() => {
            const bodyText = document.body.innerText || "";
            return bodyText.includes("melebihi batas kewajaran") || 
                   bodyText.includes("melebihi ketentuan") ||
                   bodyText.includes("Tidak dapat transaksi karena telah melebihi");
        }''')

        if is_limit:
            status_result = "TERDAFTAR"
            keterangan = "Full"
            print(f"[!] PERHATIAN: KONSUMEN LIMIT / KUOTA HABIS ({keterangan})")
            print(f"[*] Melanjutkan ke NIK berikutnya secara otomatis...")
            await page.goto("https://subsiditepatlpg.mypertamina.id/merchant/app", wait_until="domcontentloaded")
            return status_result, keterangan, customer_name, None

        # Jika kuota masih ada, klik CEK PESANAN
        print("[*] Mengklik 'CEK PESANAN'...")
        cek_btn = page.locator('button:has-text("CEK PESANAN")').first
        try:
            await cek_btn.wait_for(state="visible", timeout=6000)
            await cek_btn.click()
        except Exception:
            is_limit_late = await page.evaluate('''() => {
                const bodyText = document.body.innerText || "";
                return bodyText.includes("melebihi batas kewajaran") || 
                       bodyText.includes("melebihi ketentuan");
            }''')
            if is_limit_late:
                status_result = "TERDAFTAR"
                keterangan = "Full"
                print(f"[!] KONSUMEN LIMIT (Full): {customer_name}")
                await page.goto("https://subsiditepatlpg.mypertamina.id/merchant/app", wait_until="domcontentloaded")
                return status_result, keterangan, customer_name, None
            else:
                status_result = "TERDAFTAR"
                keterangan = "Tombol Cek Pesanan tidak ditemukan"
                print(f"[!] {keterangan}")
                await page.goto("https://subsiditepatlpg.mypertamina.id/merchant/app", wait_until="domcontentloaded")
                return status_result, keterangan, customer_name, None

        # Klik PROSES PENJUALAN
        await asyncio.sleep(0.8)
        print("[*] Mengklik 'PROSES PENJUALAN'...")
        proses_btn = page.locator('button:has-text("PROSES PENJUALAN")').first
        await proses_btn.wait_for(state="visible", timeout=8000)
        await proses_btn.click()

        # BYPASS CAPTCHA
        await bypass_puzzle_captcha(page)

        # Tunggu sampai halaman LUNAS / Struk muncul
        timestamp_now = None
        try:
            print("[*] Menunggu konfirmasi LUNAS...")
            await page.wait_for_selector('text="LUNAS"', timeout=15000)
            status_result = "SUKSES"
            keterangan = "LUNAS (1 Tabung LPG 3 Kg)"
            timestamp_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # Cek sekali lagi jika nama belum terisi, ambil dari struk
            if not customer_name:
                struk_name = await extract_customer_name(page)
                if struk_name:
                    customer_name = struk_name

            print(f"[+] ==========================================")
            print(f"[+] TRANSAKSI LUNAS BERHASIL! ({customer_name}) - {timestamp_now}")
            print(f"[+] ==========================================")

            kembali_btn = page.locator('text="KEMBALI KE HALAMAN UTAMA"').first
            if await kembali_btn.is_visible():
                await kembali_btn.click()
            else:
                await page.goto("https://subsiditepatlpg.mypertamina.id/merchant/app")
        except Exception as e:
            print(f"[!] Timeout menunggu konfirmasi Lunas: {e}")
            status_result = "GAGAL VERIFIKASI"
            keterangan = "Timeout konfirmasi lunas"

        return status_result, keterangan, customer_name, timestamp_now

    elif status_result == "UNKNOWN":
        status_result = "NIK BERMASALAH"
        keterangan = "Tidak ada respon dari server"
        print(f"[-] HASIL: {status_result}")

    await close_any_modal(page)
    if "/merchant/app/sale" in page.url or "struk" in page.url or "term-condition" in page.url:
        await page.goto("https://subsiditepatlpg.mypertamina.id/merchant/app", wait_until="domcontentloaded")

    return status_result, keterangan, customer_name, None

async def main():
    cfg = load_config()
    excel_path = Path(cfg["excel_file"])
    if not excel_path.exists():
        print(f"[!] File Excel '{excel_path}' belum ada. Membuat template baru otomatis...")
        wb_new = openpyxl.Workbook()
        ws_new = wb_new.active
        headers = ["NIK", "STATUS", "KETERANGAN", "NAMA PELANGGAN", "TERAKHIR DIGUNAKAN"]
        for idx_h, h in enumerate(headers, 1):
            cell = ws_new.cell(row=1, column=idx_h, value=h)
            cell.font = openpyxl.styles.Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            cell.fill = openpyxl.styles.PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
        wb_new.save(excel_path)
        print(f"[+] File '{excel_path}' berhasil dibuat dengan format header yang benar!")
        print(f"[!] Silakan masukkan daftar NIK ke Kolom A pada file '{excel_path}', lalu jalankan kembali bot.")
        sys.exit(0)

    print(f"[*] Membaca file: {excel_path}")
    
    try:
        with open(excel_path, "a+b"):
            pass
    except PermissionError:
        print("\n" + "="*50)
        print(f"[!] PERHATIAN: File '{excel_path.name}' sedang terbuka di Microsoft Excel!")
        print(f"[!] Harap TUTUP file Excel terlebih dahulu agar bot dapat menulis hasil transaksi.")
        print("="*50 + "\n")
        input("Tekan Enter setelah Anda menutup file Excel...")

    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    col_nik = 1
    col_status = 2
    col_ket = 3
    col_nama = 4
    col_terakhir = 5

    for c in range(1, 15):
        val = str(ws.cell(row=1, column=c).value or "").strip().upper()
        if "NIK" in val:
            col_nik = c
        elif "STATUS" in val:
            col_status = c
        elif "KET" in val:
            col_ket = c
        elif "NAMA" in val:
            col_nama = c
        elif "TERAKHIR" in val or "LAST" in val:
            col_terakhir = c

    if not ws.cell(row=1, column=col_terakhir).value or "TERAKHIR" not in str(ws.cell(row=1, column=col_terakhir).value).upper():
        col_terakhir = max(col_nik, col_status, col_ket, col_nama) + 1
        ws.cell(row=1, column=col_terakhir, value="TERAKHIR DIGUNAKAN")

    fill_sukses = PatternFill(start_color="D1E7DD", end_color="D1E7DD", fill_type="solid")
    fill_gagal = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
    fill_limit = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")

    rows_data = []
    max_row = ws.max_row
    for r in range(2, max_row + 1):
        nik_val = ws.cell(row=r, column=col_nik).value
        if not nik_val:
            continue
        nik_str = str(nik_val).strip().replace(".0", "").replace(" ", "").replace("'", "")
        if len(nik_str) < 15:
            continue
        
        stat_val = str(ws.cell(row=r, column=col_status).value or "").strip()
        ket_val = str(ws.cell(row=r, column=col_ket).value or "").strip()
        last_used_val = ws.cell(row=r, column=col_terakhir).value
        parsed_dt = parse_date(last_used_val)
        
        rows_data.append({
            "row_idx": r,
            "nik": nik_str,
            "status": stat_val,
            "ket": ket_val,
            "last_used_raw": last_used_val,
            "last_used_dt": parsed_dt
        })

    # Urutkan antrean: belum pernah dipakai duluan, lalu tanggal terlama
    rows_data.sort(key=lambda item: item["last_used_dt"])

    never_used_count = sum(1 for x in rows_data if x["last_used_dt"] == datetime.min)
    used_count = len(rows_data) - never_used_count
    print(f"\n[+] Total NIK valid: {len(rows_data)}")
    print(f"[+] Antrean: {never_used_count} NIK belum pernah dipakai (prioritas terdepan), {used_count} NIK diurutkan dari yang terlama.")

    async with async_playwright() as p:
        user_data_dir = Path("user_profile_pertamina")
        context = await p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            headless=cfg["headless"],
            viewport={"width": 1280, "height": 850},
            args=["--disable-blink-features=AutomationControlled"]
        )
        page = context.pages[0] if context.pages else await context.new_page()

        # Listener response API verify-nik dipasang 1x saja agar tidak menumpuk
        captured_api_name = {"name": ""}
        async def on_response(response):
            try:
                if "verify-nik" in response.url:
                    data = await response.json()
                    if data and isinstance(data, dict):
                        d = data.get("data") or {}
                        if isinstance(d, dict) and d.get("name"):
                            captured_api_name["name"] = d["name"].strip()
            except Exception:
                pass
        page.on("response", on_response)

        logged_in = await login(page, cfg["phone"], cfg["pin"])
        if not logged_in:
            print("[-] Gagal login. Script dihentikan.")
            await context.close()
            return

        for idx, item in enumerate(rows_data, 1):
            r = item["row_idx"]
            nik_str = item["nik"]
            current_status = str(ws.cell(row=r, column=col_status).value or "").strip().upper()
            current_ket = str(ws.cell(row=r, column=col_ket).value or "").strip().upper()
            print(f"\n[*] Memproses Antrean #{idx}: NIK {nik_str} (Baris Excel {r}) | Terakhir digunakan: {item['last_used_raw'] or 'Belum Pernah'}")

            status_res, ket_res, name_res, timestamp_res = await process_single_nik(page, nik_str, cfg, captured_api_name)

            c_stat = ws.cell(row=r, column=col_status, value=status_res)
            c_ket = ws.cell(row=r, column=col_ket, value=ket_res)
            
            # Pastikan nama pelanggan selalu ditulis ke Kolom D
            if name_res:
                ws.cell(row=r, column=col_nama, value=name_res)
                
            if timestamp_res:
                ws.cell(row=r, column=col_terakhir, value=timestamp_res)

            if "SUKSES" in status_res:
                c_stat.fill = fill_sukses
            elif "Full" in ket_res:
                c_stat.fill = fill_limit
                c_ket.fill = fill_limit
            else:
                c_stat.fill = fill_gagal

            safe_save_workbook(wb, excel_path)
            print(f"[+] Baris {r} tersimpan: {status_res} | {ket_res} | Nama: {name_res or '-'} | {timestamp_res or '-'}")

            delay = random.uniform(cfg["delay_min"], cfg["delay_max"])
            print(f"[*] Jeda {delay:.1f} detik...")
            await asyncio.sleep(delay)

        print("\n==================================================")
        print("[+] Semua antrean NIK telah selesai diproses!")
        print("==================================================")
        await context.close()

if __name__ == "__main__":
    asyncio.run(main())
