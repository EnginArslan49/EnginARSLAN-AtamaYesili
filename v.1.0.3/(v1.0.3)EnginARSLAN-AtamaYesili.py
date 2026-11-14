# -*- coding: utf-8 -*-
"""
Engin ARSLAN ile Atanma Yeşili - v1.0.3
Kariyer Kapısı İlan Tarama Aracı - Paralel Sürüm
"""

import os
import sys
import time
import random
import threading
import re
import tkinter as tk
from tkinter import scrolledtext, messagebox
from datetime import datetime

from selenium.webdriver import Chrome
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import (
    StaleElementReferenceException,
    ElementClickInterceptedException,
    WebDriverException
)

# ------------------------------------------------------------
# SABİTLER
# ------------------------------------------------------------
SURUM = "v1.0.3"
PROGRAM_ADI = "Engin ARSLAN ile Atanma Yeşili"
HEDEF_URL = "https://kariyerkapisi.gov.tr/isealim"

MIN_BEKLEME = 1.2
MAX_BEKLEME = 2.2
ILAN_YUKLEME_BEKLEME = 25
PARALEL_TARAYICI_SAYISI = 2

IFADE_BP = "bilgisayar programcılığı"
IFADE_ONLISANS_VARYANTLARI = [
    "herhangi bir önlisans",
    "herhangi bir ön lisans",
    "herhangi bir önlisans bölüm",
    "herhangi bir ön lisans bölüm",
    "önlisans bölümlerinden",
    "ön lisans bölümlerinden",
    "önlisans programlarından",
    "ön lisans programlarından",
    "önlisans mezunu",
    "ön lisans mezunu",
]

# ---------- Tema Renkleri ----------
TEMA_LIGHT = {
    "arka": "#f0f4f0",
    "baslik_metin": "#2e7d32",
    "alt_metin": "#555555",
    "label_frame_fg": "#2e7d32",
    "label_frame_bg": "#f0f4f0",
    "log_bg": "#ffffff",  # Light temada log arka planı beyaz
    "log_fg": "#000000",  # Light temada yazı siyah
    "buton_yesil": "#4caf50",
    "buton_kirmizi": "#f44336",
    "buton_tema": "#455a64",
    "buton_ayar": "#546e7a",
    "buton_devam": "#2e7d32",
    "ilan_ad": "#1565c0",
    "basari": "#2e7d32",
    "uyari": "#d84315",
    "hata": "#c62828",
    "muted": "#888888",
    "info": "#000000",
    "bot1_renk": "#1565c0",
    "bot2_renk": "#c2185b",
    "kart_bg": "#ffffff",
    "kart_border": "#c8e6c9",
    "secim_bg": "#ffffff",
    "adim_aktif": "#2e7d32",
    "adim_pasif": "#bdbdbd",
}

TEMA_DARK = {
    "arka": "#121212",
    "baslik_metin": "#7cd492",
    "alt_metin": "#aaaaaa",
    "label_frame_fg": "#7cd492",
    "label_frame_bg": "#121212",
    "log_bg": "#1e1e1e",  # Dark temada log arka planı koyu gri
    "log_fg": "#ffffff",  # Dark temada yazı beyaz
    "buton_yesil": "#43a047",
    "buton_kirmizi": "#e53935",
    "buton_tema": "#90a4ae",
    "buton_ayar": "#78909c",
    "buton_devam": "#43a047",
    "ilan_ad": "#64b5f6",
    "basari": "#7cd492",
    "uyari": "#ffb74d",
    "hata": "#ef5350",
    "muted": "#aaaaaa",
    "info": "#ffffff",
    "bot1_renk": "#64b5f6",
    "bot2_renk": "#f06292",
    "kart_bg": "#1e1e1e",
    "kart_border": "#2e7d32",
    "secim_bg": "#1e1e1e",
    "adim_aktif": "#7cd492",
    "adim_pasif": "#555555",
}


# ------------------------------------------------------------
# YARDIMCI
# ------------------------------------------------------------
def rastgele_bekle(min_s=MIN_BEKLEME, max_s=MAX_BEKLEME):
    time.sleep(random.uniform(min_s, max_s))


def masaustu_yolu():
    ev = os.path.expanduser("~")
    for isim in ["Masaüstü", "Desktop"]:
        yol = os.path.join(ev, isim)
        if os.path.exists(yol):
            return yol
    return ev


def normalize(metin):
    if not metin:
        return ""
    metin = metin.replace("\xa0", " ").replace("&nbsp;", " ")
    return re.sub(r"\s+", " ", metin.lower()).strip()


# ------------------------------------------------------------
# EXCEL YAZICI
# ------------------------------------------------------------
def excel_yaz(dosya_yolu, bp_liste, onlisans_liste, secili_kriterler):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment

        wb = Workbook()
        sheet_eklendi = False

        if secili_kriterler.get("bp", False):
            ws1 = wb.active if not sheet_eklendi else wb.create_sheet()
            ws1.title = "Bilg. Programcılığı"
            ws1.append(["#", "İlan Adı", "Alt Başlık", "Link"])
            for c in ws1[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="2E7D32")
                c.alignment = Alignment(horizontal="center")
            if bp_liste:
                for i, ilan in enumerate(bp_liste, 1):
                    ws1.append([i, ilan["baslik"], ilan["alt_baslik"], ilan["url"]])
            else:
                ws1.append(["-", "Bu kriterde uygun ilan bulunamadı", "", ""])
            ws1.column_dimensions["A"].width = 5
            ws1.column_dimensions["B"].width = 50
            ws1.column_dimensions["C"].width = 40
            ws1.column_dimensions["D"].width = 60
            sheet_eklendi = True

        if secili_kriterler.get("onlisans", False):
            ws2 = wb.active if not sheet_eklendi else wb.create_sheet()
            ws2.title = "Herhangi Önlisans"
            ws2.append(["#", "İlan Adı", "Alt Başlık", "Link"])
            for c in ws2[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="C2185B")
                c.alignment = Alignment(horizontal="center")
            if onlisans_liste:
                for i, ilan in enumerate(onlisans_liste, 1):
                    ws2.append([i, ilan["baslik"], ilan["alt_baslik"], ilan["url"]])
            else:
                ws2.append(["-", "Bu kriterde uygun ilan bulunamadı", "", ""])
            ws2.column_dimensions["A"].width = 5
            ws2.column_dimensions["B"].width = 50
            ws2.column_dimensions["C"].width = 40
            ws2.column_dimensions["D"].width = 60

        wb.save(dosya_yolu)
        return True, ""

    except ImportError:
        return False, "openpyxl kütüphanesi kurulu değil. 'pip install openpyxl' komutunu çalıştırın."
    except Exception as e:
        return False, str(e)


# ------------------------------------------------------------
# BOT
# ------------------------------------------------------------
class KariyerKapisiBot:

    def __init__(self, bot_id, log_cb, sayaç_cb=None, ilan_cb=None, secilen_kriterler=None):
        self.bot_id = bot_id
        self.driver = None
        self.log = log_cb
        self.sayaç_cb = sayaç_cb or (lambda *a: None)
        self.ilan_cb = ilan_cb or (lambda *a: None)
        self.durduruldu = False
        self.secilen_kriterler = secilen_kriterler or {"bp": True, "onlisans": True}

        self.uygun_ilanlar_bp = []
        self.uygun_ilanlar_onlisans = []
        self.incelenen_sayisi = 0

    def tarayici_baslat(self, sol_taraf=True):
        self.log("🌐 Chrome başlatılıyor...", "info")
        options = Options()
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-popup-blocking")

        ekran_w = self._ekran_genisligi()
        yarim = ekran_w // 2
        if sol_taraf:
            options.add_argument("--window-position=0,0")
            options.add_argument(f"--window-size={yarim},1080")
        else:
            options.add_argument(f"--window-position={yarim},0")
            options.add_argument(f"--window-size={yarim},1080")

        self.driver = webdriver.Chrome(options=options)
        self.driver.execute_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        self.log("✅ Chrome hazır.", "success")
        return self.driver

    @staticmethod
    def _ekran_genisligi():
        try:
            import ctypes
            return ctypes.windll.user32.GetSystemMetrics(0)
        except Exception:
            return 1920

    def taramayi_baslat(self, baslangic_idx, bitis_idx, toplam):
        self.log("🔗 Siteye gidiliyor...", "info")
        self.driver.get(HEDEF_URL)
        self.log("⏳ İlanların yüklenmesi bekleniyor...", "info")
        if not self._ilanlarin_yuklenmesini_bekle():
            self.log("⚠️ İlanlar yüklenemedi.", "warning")
            return
        rastgele_bekle(2, 3)
        self.log(f"📋 Toplam {toplam} ilan var. "
                 f"Bu bot {baslangic_idx + 1}-{min(bitis_idx, toplam)} arasını tarayacak.", "info")

        for idx in range(baslangic_idx, min(bitis_idx, toplam)):
            if self.durduruldu:
                break
            try:
                guncel_butonlar = self._basvur_butonlarini_bul()
                if idx >= len(guncel_butonlar):
                    break
                buton = guncel_butonlar[idx]
                ilan_adi = self._kart_ilan_adini_bul(buton, idx + 1)
                self.ilan_cb(f"Bot {self.bot_id}: {ilan_adi}")
                self.log(f"🔎 İlan {idx + 1}/{toplam}: {ilan_adi[:70]}", "info")
                self._detaya_git_ve_tara(buton, ilan_adi, idx + 1)
            except StaleElementReferenceException:
                continue
            except Exception as e:
                self.log(f"⚠️ İlan {idx + 1} hatası: {str(e)}", "warning")
                continue
            rastgele_bekle()
        self.log("🏁 Bot taraması bitti.", "success")

    def _ilanlarin_yuklenmesini_bekle(self):
        bitis = time.time() + ILAN_YUKLEME_BEKLEME
        while time.time() < bitis:
            if self.durduruldu:
                return False
            try:
                pnl = self.driver.find_element(By.ID, "pnlIlanlar")
                butonlar = pnl.find_elements(
                    By.XPATH,
                    ".//a[contains(translate(., 'BAŞVUR', 'başvur'), 'başvur')] | "
                    ".//button[contains(translate(., 'BAŞVUR', 'başvur'), 'başvur')]"
                )
                if butonlar:
                    return True
            except Exception:
                pass
            time.sleep(0.5)
        return False

    def _basvur_butonlarini_bul(self):
        try:
            pnl = self.driver.find_element(By.ID, "pnlIlanlar")
            return pnl.find_elements(
                By.XPATH,
                ".//a[contains(translate(., 'BAŞVUR', 'başvur'), 'başvur')] | "
                ".//button[contains(translate(., 'BAŞVUR', 'başvur'), 'başvur')]"
            )
        except Exception:
            return []

    def _kart_ilan_adini_bul(self, buton, sira):
        try:
            kart = None
            for xp in [
                "./ancestor::div[contains(@class,'row')][1]",
                "./ancestor::div[contains(@class,'card')][1]",
                "./ancestor::tr[1]",
                "./ancestor::li[1]",
                "./ancestor::div[position()<=3][1]",
            ]:
                try:
                    kart = buton.find_element(By.XPATH, xp)
                    if kart:
                        break
                except Exception:
                    continue
            if not kart:
                return f"İlan {sira}"
            metin = kart.text.strip()
            if metin:
                satirlar = [s.strip() for s in metin.split("\n") if s.strip()]
                atlanacak = {"başvur", "başvuru", "detay", "detaylar", "i̇ncele", "incele"}
                adaylar = [s for s in satirlar if normalize(s) not in atlanacak and len(s) > 3]
                if adaylar:
                    return " - ".join(adaylar[:2])[:150]
            for xp in [".//h1", ".//h2", ".//h3", ".//h4", ".//h5", ".//h6", ".//b", ".//strong"]:
                try:
                    el = kart.find_element(By.XPATH, xp)
                    t = el.text.strip()
                    if t and len(t) > 3:
                        return t[:150]
                except Exception:
                    continue
        except Exception:
            pass
        return f"İlan {sira}"

    def _detaya_git_ve_tara(self, buton, ilan_adi, global_sira):
        ana_pencere = self.driver.current_window_handle
        yeni_pencere = False
        try:
            try:
                buton.click()
            except ElementClickInterceptedException:
                self.driver.execute_script("arguments[0].click();", buton)
            rastgele_bekle(1.5, 2.5)
            if len(self.driver.window_handles) > 1:
                self.driver.switch_to.window(self.driver.window_handles[-1])
                yeni_pencere = True
            rastgele_bekle(2, 3)

            bp_bulunan = []
            onlisans_bulunan = []
            self._sayfadaki_accordionlari_tara(ilan_adi, bp_bulunan, onlisans_bulunan)

            self.incelenen_sayisi += 1
            self.sayaç_cb(self.bot_id, self.incelenen_sayisi,
                          len(self.uygun_ilanlar_bp),
                          len(self.uygun_ilanlar_onlisans))

            detay_url = self.driver.current_url
            if self.secilen_kriterler.get("bp", False):
                for alt_baslik, _ in bp_bulunan:
                    self._listeye_ekle(self.uygun_ilanlar_bp, ilan_adi, detay_url, alt_baslik)
                    self.log(f"   ✅ BİLG. PROGRAMCILIĞI: {alt_baslik[:60]}", "success")
            if self.secilen_kriterler.get("onlisans", False):
                for alt_baslik, _ in onlisans_bulunan:
                    self._listeye_ekle(self.uygun_ilanlar_onlisans, ilan_adi, detay_url, alt_baslik)
                    self.log(f"   ✅ ÖNLİSANS: {alt_baslik[:60]}", "success")
            self.sayaç_cb(self.bot_id, self.incelenen_sayisi,
                          len(self.uygun_ilanlar_bp),
                          len(self.uygun_ilanlar_onlisans))
            if not bp_bulunan and not onlisans_bulunan:
                self.log("   ❌ Uygun değil.", "muted")
        except Exception as e:
            self.log(f"⚠️ Detay hatası: {str(e)}", "warning")
        finally:
            try:
                if yeni_pencere and len(self.driver.window_handles) > 1:
                    self.driver.close()
                    self.driver.switch_to.window(ana_pencere)
                else:
                    self.driver.switch_to.window(ana_pencere)
                    self.driver.back()
                rastgele_bekle(1.5, 2.5)
            except Exception:
                try:
                    self.driver.switch_to.window(ana_pencere)
                except Exception:
                    pass

    def _sayfadaki_accordionlari_tara(self, ana_ilan_adi, bp_liste, onlisans_liste):
        try:
            butonlar = self.driver.find_elements(
                By.XPATH,
                "//button[@data-bs-toggle='collapse' or @data-toggle='collapse']"
            )
        except Exception:
            butonlar = []

        if not butonlar:
            try:
                metin = self.driver.find_element(By.TAG_NAME, "body").text
                n = normalize(metin)
                bp = IFADE_BP in n
                onl = any(v in n for v in IFADE_ONLISANS_VARYANTLARI)
                if bp:
                    bp_liste.append((ana_ilan_adi, self.driver.current_url))
                if onl:
                    onlisans_liste.append((ana_ilan_adi, self.driver.current_url))
            except Exception:
                pass
            return

        for idx, buton in enumerate(butonlar):
            if self.durduruldu:
                break
            try:
                alt_baslik = self._accordion_basligini_oku(buton) or f"Alt {idx + 1}"
                self.ilan_cb(f"Bot {self.bot_id}: {ana_ilan_adi}  →  {alt_baslik}")
                try:
                    self.driver.execute_script(
                        "arguments[0].scrollIntoView({block:'center', behavior:'smooth'});",
                        buton
                    )
                    time.sleep(0.4)
                except Exception:
                    pass
                self._accordion_ac(buton)
                time.sleep(0.35)
                icerik = self._accordion_icerigini_oku(buton)
                icerik_norm = normalize(icerik)
                bp = IFADE_BP in icerik_norm
                onl = any(v in icerik_norm for v in IFADE_ONLISANS_VARYANTLARI)
                if bp:
                    bp_liste.append((alt_baslik, self.driver.current_url))
                    self.log(f"   ✅ Bilg. Prog.: {alt_baslik[:50]}", "success")
                if onl:
                    onlisans_liste.append((alt_baslik, self.driver.current_url))
                    self.log(f"   ✅ Önlisans: {alt_baslik[:50]}", "success")
                if not bp and not onl:
                    self.log(f"   ❌ {alt_baslik[:50]}", "muted")
                rastgele_bekle(0.3, 0.7)
            except Exception:
                continue

    def _accordion_basligini_oku(self, buton):
        try:
            for xp in [".//h6", ".//h5", ".//h4", ".//h3", ".//b", ".//strong"]:
                try:
                    el = buton.find_element(By.XPATH, xp)
                    t = re.sub(r"\s+", " ", el.text).strip()
                    if t:
                        return t
                except Exception:
                    continue
        except Exception:
            pass
        try:
            t = re.sub(r"\s+", " ", buton.text).strip()
            if t:
                return t
        except Exception:
            pass
        return ""

    def _accordion_ac(self, buton):
        try:
            hedef_id = (buton.get_attribute("data-bs-target") or
                        buton.get_attribute("data-target") or "")
            if hedef_id:
                hedef_id = hedef_id.lstrip("#.")
                try:
                    icerik_div = self.driver.find_element(By.ID, hedef_id)
                    if "show" in (icerik_div.get_attribute("class") or "").split():
                        return True
                except Exception:
                    pass
            try:
                buton.click()
            except ElementClickInterceptedException:
                self.driver.execute_script("arguments[0].click();", buton)
            except Exception:
                self.driver.execute_script("arguments[0].click();", buton)
            time.sleep(0.35)
            return True
        except Exception:
            return False

    def _accordion_icerigini_oku(self, buton):
        parcalar = []
        try:
            hedef_id = (buton.get_attribute("data-bs-target") or
                        buton.get_attribute("data-target") or "")
            if hedef_id:
                hedef_id = hedef_id.lstrip("#.")
                icerik_div = self.driver.find_element(By.ID, hedef_id)
                parcalar.append(icerik_div.text or "")
                try:
                    parcalar.append(icerik_div.get_attribute("textContent") or "")
                except Exception:
                    pass
        except Exception:
            pass
        try:
            parent = buton.find_element(By.XPATH, "./..")
            parcalar.append(parent.text or "")
        except Exception:
            pass
        return "\n".join(p for p in parcalar if p)

    def _listeye_ekle(self, liste, baslik, url, alt_baslik):
        anahtar = (normalize(baslik), url, normalize(alt_baslik))
        for mevcut in liste:
            if (normalize(mevcut["baslik"]), mevcut["url"],
                normalize(mevcut["alt_baslik"])) == anahtar:
                return
        liste.append({"baslik": baslik, "url": url, "alt_baslik": alt_baslik})

    def durdur(self):
        self.durduruldu = True

    def tarayiciyi_kapat(self):
        try:
            if self.driver:
                self.driver.quit()
                self.driver = None
        except Exception:
            pass


# ------------------------------------------------------------
# ARAYÜZ
# ------------------------------------------------------------
class BotArayuz:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(f"{PROGRAM_ADI}  ({SURUM})")
        self.root.geometry("1300x900")
        self.root.resizable(False, False)
        self.tema = TEMA_LIGHT
        self.tema_adi = "light"

        # Kapatma protokolünü (X butonu) kendi fonksiyonumuza yönlendir
        self.root.protocol("WM_DELETE_WINDOW", self._guvenli_kapat)

        self.root.attributes("-topmost", True)
        self._pencereyi_ortala(1300, 900)

        self.ana_container = tk.Frame(self.root)
        self.ana_container.pack(fill="both", expand=True)

        self.kriter_bp_var = tk.BooleanVar(value=True)
        self.kriter_onlisans_var = tk.BooleanVar(value=True)
        self.cikti_secim = tk.StringVar(value="excel")

        self.botlar = []
        self.threadler = []
        self.tarama_aktif = False
        self.hazir_bot_sayisi = 0

        self.ekran = None
        self._ekran_kurulum_olustur()
        self.tema_uygula()

    def _guvenli_kapat(self):
        """X butonuna basıldığında çalışır: açık Chrome'ları kapatıp uygulamayı sonlandırır."""
        if self.tarama_aktif:
            if not messagebox.askyesno("Çıkış", "Tarama devam ediyor. Kapatmak istediğinize emin misiniz?"):
                return
        # Tüm botları durdur ve Chrome'ları kapat
        for bot in self.botlar:
            try:
                bot.durdur()
                bot.tarayiciyi_kapat()
            except Exception:
                pass
        # Uygulamayı kapat
        self.root.destroy()

    def _pencereyi_ortala(self, g, y):
        self.root.update_idletasks()
        eg, ey = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x, y2 = (eg - g) // 2, (ey - y) // 2
        self.root.geometry(f"{g}x{y}+{x}+{y2}")

    def _ekran_gecis_animasyonu(self, yeni_ekran_olusturucu):
        eski = self.ekran
        yeni_container = tk.Frame(self.ana_container)

        if eski is not None:
            eski.pack_forget()

        toplam_adim = 12
        gecikme = 12

        def adim(i):
            if i > toplam_adim:
                if eski is not None:
                    try:
                        eski.destroy()
                    except Exception:
                        pass
                yeni_ekran_olusturucu(yeni_container)
                self.ekran = yeni_container
                self.tema_uygula()
                return

            oran = i / toplam_adim

            if eski is not None:
                try:
                    eski.place(
                        x=int(-1300 * oran), y=0,
                        width=1300, height=900
                    )
                    eski.place_configure(anchor="nw")
                except Exception:
                    pass

            try:
                yeni_container.place(
                    x=int(1300 * (1 - oran)), y=0,
                    width=1300, height=900
                )
            except Exception:
                pass

            self.root.after(gecikme, lambda: adim(i + 1))

        yeni_container.place(x=1300, y=0, width=1300, height=900)
        self.root.after(10, lambda: adim(1))

    # --------------------------------------------------------
    # KURULUM EKRANI
    # --------------------------------------------------------
    def _ekran_kurulum_olustur(self, container=None):
        if container is None:
            container = tk.Frame(self.ana_container)
            container.pack(fill="both", expand=True)
            self.ekran = container
        else:
            container.pack(fill="both", expand=True)
            self.ekran = container

        self.kurulum_ust = tk.Frame(container)
        self.kurulum_ust.pack(fill="x", padx=30, pady=(25, 10))

        self.lbl_adim1 = tk.Label(
            self.kurulum_ust, text="  1  Kurulum  ",
            font=("Segoe UI", 11, "bold")
        )
        self.lbl_adim1.pack(side="left")

        self.lbl_ok = tk.Label(
            self.kurulum_ust, text="  ────────  ",
            font=("Segoe UI", 11)
        )
        self.lbl_ok.pack(side="left")

        self.lbl_adim2 = tk.Label(
            self.kurulum_ust, text="  2  Tarama  ",
            font=("Segoe UI", 11)
        )
        self.lbl_adim2.pack(side="left")

        # Tema Butonu (Kurulum Ekranı)
        self.btn_tema_k = tk.Button(
            self.kurulum_ust,
            text="🌙 Dark" if self.tema_adi == "light" else "☀️ Light",
            font=("Segoe UI", 10, "bold"),
            relief="flat", cursor="hand2",
            padx=15, pady=5, command=self._tema_degistir
        )
        self.btn_tema_k.pack(side="right")

        self.kurulum_baslik_frame = tk.Frame(container)
        self.kurulum_baslik_frame.pack(fill="x", padx=30, pady=(20, 5))

        self.lbl_kurulum_baslik = tk.Label(
            self.kurulum_baslik_frame,
            text="⚙️  Kurulum",
            font=("Segoe UI", 22, "bold")
        )
        self.lbl_kurulum_baslik.pack()

        self.lbl_kurulum_alt = tk.Label(
            self.kurulum_baslik_frame,
            text="Arama kriterlerini ve çıktı tercihini belirleyin",
            font=("Segoe UI", 11)
        )
        self.lbl_kurulum_alt.pack(pady=(4, 0))

        self.kurulum_icerik = tk.Frame(container)
        self.kurulum_icerik.pack(fill="both", expand=True, padx=60, pady=20)

        self.kart_kriter = tk.Frame(self.kurulum_icerik, bd=1, relief="solid")
        self.kart_kriter.pack(fill="x", pady=(0, 15))

        self.lbl_kart_kriter = tk.Label(
            self.kart_kriter,
            text="🔍  Arama Kriterleri",
            font=("Segoe UI", 13, "bold"),
            anchor="w"
        )
        self.lbl_kart_kriter.pack(fill="x", padx=20, pady=(15, 8))

        self.lbl_kart_kriter_alt = tk.Label(
            self.kart_kriter,
            text="En az bir kriter seçmelisiniz",
            font=("Segoe UI", 9, "italic"),
            anchor="w"
        )
        self.lbl_kart_kriter_alt.pack(fill="x", padx=20, pady=(0, 10))

        self.chk_bp = tk.Checkbutton(
            self.kart_kriter,
            text="💻  Bilgisayar Programcılığı mezunu şartı",
            variable=self.kriter_bp_var,
            font=("Segoe UI", 11),
            anchor="w", cursor="hand2"
        )
        self.chk_bp.pack(fill="x", padx=30, pady=5)

        self.chk_onlisans = tk.Checkbutton(
            self.kart_kriter,
            text="🎓  Herhangi bir önlisans bölümünden mezun olmak şartı",
            variable=self.kriter_onlisans_var,
            font=("Segoe UI", 11),
            anchor="w", cursor="hand2"
        )
        self.chk_onlisans.pack(fill="x", padx=30, pady=(5, 15))

        self.kart_cikti = tk.Frame(self.kurulum_icerik, bd=1, relief="solid")
        self.kart_cikti.pack(fill="x", pady=(0, 15))

        self.lbl_kart_cikti = tk.Label(
            self.kart_cikti,
            text="📄  Çıktı Tercihi",
            font=("Segoe UI", 13, "bold"),
            anchor="w"
        )
        self.lbl_kart_cikti.pack(fill="x", padx=20, pady=(15, 8))

        self.lbl_kart_cikti_alt = tk.Label(
            self.kart_cikti,
            text="Sonuçların hangi formatta kaydedileceğini seçin",
            font=("Segoe UI", 9, "italic"),
            anchor="w"
        )
        self.lbl_kart_cikti_alt.pack(fill="x", padx=20, pady=(0, 10))

        self.rd_excel = tk.Radiobutton(
            self.kart_cikti, text="📊  Excel (.xlsx)  —  önerilen",
            variable=self.cikti_secim, value="excel",
            font=("Segoe UI", 11), anchor="w", cursor="hand2"
        )
        self.rd_excel.pack(fill="x", padx=30, pady=5)

        self.rd_txt = tk.Radiobutton(
            self.kart_cikti, text="📃  TXT (.txt)",
            variable=self.cikti_secim, value="txt",
            font=("Segoe UI", 11), anchor="w", cursor="hand2"
        )
        self.rd_txt.pack(fill="x", padx=30, pady=(5, 15))

        self.btn_devam = tk.Button(
            container, text="✅  Devam  →",
            font=("Segoe UI", 13, "bold"),
            relief="flat", cursor="hand2",
            padx=40, pady=14,
            command=self._devam_tiklandi
        )
        self.btn_devam.pack(pady=(5, 20))

        self.lbl_surum_k = tk.Label(
            container, text=f"Sürüm: {SURUM}",
            font=("Segoe UI", 8)
        )
        self.lbl_surum_k.pack(side="bottom", pady=3)

    def _devam_tiklandi(self):
        if not self.kriter_bp_var.get() and not self.kriter_onlisans_var.get():
            messagebox.showwarning(
                "Kriter Seçilmedi",
                "Lütfen en az bir arama kriteri seçin."
            )
            return
        self._ekran_gecis_animasyonu(self._ekran_tarama_olustur)

    # --------------------------------------------------------
    # TARAMA EKRANI
    # --------------------------------------------------------
    def _ekran_tarama_olustur(self, container):
        self.tarama_container = container

        self.tarama_ust = tk.Frame(container)
        self.tarama_ust.pack(fill="x", padx=20, pady=(12, 4))

        self.lbl_tarama_baslik = tk.Label(
            self.tarama_ust, text=f"{PROGRAM_ADI}",
            font=("Segoe UI", 16, "bold")
        )
        self.lbl_tarama_baslik.pack(side="left")

        self.btn_tema_t = tk.Button(
            self.tarama_ust,
            text="🌙 Dark" if self.tema_adi == "light" else "☀️ Light",
            font=("Segoe UI", 10, "bold"),
            relief="flat", cursor="hand2",
            padx=12, pady=5, command=self._tema_degistir
        )
        self.btn_tema_t.pack(side="right", padx=(5, 0))

        self.btn_ayarlar = tk.Button(
            self.tarama_ust, text="⚙️  Ayarlar",
            font=("Segoe UI", 10, "bold"),
            relief="flat", cursor="hand2",
            padx=12, pady=5, command=self._ayarlara_don
        )
        self.btn_ayarlar.pack(side="right", padx=(5, 0))

        self.lbl_kriter_ozet = tk.Label(
            container, text=self._kriter_ozet_metni(),
            font=("Segoe UI", 9, "italic")
        )
        self.lbl_kriter_ozet.pack(pady=(0, 4))

        self.durum_frame = tk.LabelFrame(
            container, text="📊 Genel Durum",
            font=("Segoe UI", 10, "bold"),
            padx=12, pady=6
        )
        self.durum_frame.pack(fill="x", padx=20, pady=4)

        self.lbl_incelenen = tk.Label(self.durum_frame, text="İncelenen: 0",
                                      font=("Segoe UI", 10))
        self.lbl_incelenen.pack(side="left", padx=10)

        self.lbl_bp = None
        self.lbl_onlisans = None

        if self.kriter_bp_var.get():
            self.lbl_bp = tk.Label(self.durum_frame, text="Bilg. Prog.: 0",
                                   font=("Segoe UI", 10, "bold"))
            self.lbl_bp.pack(side="left", padx=10)

        if self.kriter_onlisans_var.get():
            self.lbl_onlisans = tk.Label(self.durum_frame, text="Önlisans: 0",
                                         font=("Segoe UI", 10, "bold"))
            self.lbl_onlisans.pack(side="left", padx=10)

        self.lbl_bot_durum = tk.Label(self.durum_frame, text="Botlar: hazır",
                                      font=("Segoe UI", 9, "italic"))
        self.lbl_bot_durum.pack(side="right", padx=10)

        self.ilan_frame = tk.LabelFrame(
            container, text="🔎 Şu An Kontrol Edilen İlanlar",
            font=("Segoe UI", 10, "bold"),
            padx=12, pady=6
        )
        self.ilan_frame.pack(fill="x", padx=20, pady=4)

        b1_frame = tk.Frame(self.ilan_frame)
        b1_frame.pack(fill="x", pady=1)
        self.lbl_bot1_et = tk.Label(b1_frame, text="Bot 1:",
                                    font=("Segoe UI", 9, "bold"))
        self.lbl_bot1_et.pack(side="left", padx=(0, 6))
        self.lbl_ilan_bot1 = tk.Label(
            b1_frame, text="—", font=("Segoe UI", 9, "bold"),
            wraplength=1100, justify="left", anchor="w"
        )
        self.lbl_ilan_bot1.pack(side="left", fill="x", expand=True)

        b2_frame = tk.Frame(self.ilan_frame)
        b2_frame.pack(fill="x", pady=1)
        self.lbl_bot2_et = tk.Label(b2_frame, text="Bot 2:",
                                    font=("Segoe UI", 9, "bold"))
        self.lbl_bot2_et.pack(side="left", padx=(0, 6))
        self.lbl_ilan_bot2 = tk.Label(
            b2_frame, text="—", font=("Segoe UI", 9, "bold"),
            wraplength=1100, justify="left", anchor="w"
        )
        self.lbl_ilan_bot2.pack(side="left", fill="x", expand=True)

        self.log_container = tk.Frame(container)
        self.log_container.pack(fill="both", expand=True, padx=20, pady=4)

        self.log1_frame = tk.LabelFrame(
            self.log_container, text="📝 İşlem Günlüğü — Bot 1",
            font=("Segoe UI", 10, "bold"), padx=5, pady=5
        )
        self.log1_frame.pack(side="left", fill="both", expand=True, padx=(0, 4))

        self.log1 = scrolledtext.ScrolledText(
            self.log1_frame, font=("Consolas", 8),
            wrap="word", state="disabled"
        )
        self.log1.pack(fill="both", expand=True)

        self.log2_frame = tk.LabelFrame(
            self.log_container, text="📝 İşlem Günlüğü — Bot 2",
            font=("Segoe UI", 10, "bold"), padx=5, pady=5
        )
        self.log2_frame.pack(side="right", fill="both", expand=True, padx=(4, 0))

        self.log2 = scrolledtext.ScrolledText(
            self.log2_frame, font=("Consolas", 8),
            wrap="word", state="disabled"
        )
        self.log2.pack(fill="both", expand=True)

        for tag in ("info", "success", "warning", "error", "muted"):
            self.log1.tag_config(tag)
            self.log2.tag_config(tag)

        self.btnf = tk.Frame(container)
        self.btnf.pack(fill="x", padx=20, pady=(4, 8))

        self.btn_baslat = tk.Button(
            self.btnf, text="▶ Taramayı Başlat",
            font=("Segoe UI", 11, "bold"),
            relief="flat", cursor="hand2",
            padx=30, pady=10, command=self._baslat
        )
        self.btn_baslat.pack(side="left", padx=6)

        self.btn_durdur = tk.Button(
            self.btnf, text="⏹ Durdur",
            font=("Segoe UI", 11, "bold"),
            relief="flat", cursor="hand2",
            padx=30, pady=10, state="disabled", command=self._durdur
        )
        self.btn_durdur.pack(side="left", padx=6)

        self.lbl_surum_t = tk.Label(
            container, text=f"Sürüm: {SURUM}",
            font=("Segoe UI", 8)
        )
        self.lbl_surum_t.pack(side="bottom", pady=3)

    def _kriter_ozet_metni(self):
        parcaciklar = []
        if self.kriter_bp_var.get():
            parcaciklar.append("💻 Bilgisayar Programcılığı")
        if self.kriter_onlisans_var.get():
            parcaciklar.append("🎓 Herhangi Önlisans")
        return "Aranan: " + " | ".join(parcaciklar)

    def _ayarlara_don(self):
        if self.tarama_aktif:
            if not messagebox.askyesno(
                    "Tarama Devam Ediyor",
                    "Tarama devam ediyor. Ayarlara dönmek istediğinize emin misiniz?\n"
                    "Tarama durdurulacak."
            ):
                return
            self._durdur()
        self._ekran_gecis_animasyonu(self._ekran_kurulum_olustur)

    # --------------------------------------------------------
    # TEMA UYGULAMA
    # --------------------------------------------------------
    def tema_uygula(self):
        t = self.tema
        self.root.configure(bg=t["arka"])
        self.ana_container.configure(bg=t["arka"])

        if hasattr(self, "kurulum_ust") and self.kurulum_ust.winfo_exists():
            self._tema_kurulum_uygula(t)
        if hasattr(self, "tarama_ust") and self.tarama_ust.winfo_exists():
            self._tema_tarama_uygula(t)

    def _tema_kurulum_uygula(self, t):
        for w in (self.ekran, self.kurulum_ust,
                  self.kurulum_baslik_frame, self.kurulum_icerik,
                  self.kart_kriter, self.kart_cikti):
            try:
                w.configure(bg=t["arka"])
            except Exception:
                pass

        try:
            self.lbl_adim1.configure(bg=t["arka"], fg=t["adim_aktif"])
            self.lbl_ok.configure(bg=t["arka"], fg=t["adim_pasif"])
            self.lbl_adim2.configure(bg=t["arka"], fg=t["adim_pasif"])
        except Exception:
            pass

        try:
            self.lbl_kurulum_baslik.configure(bg=t["arka"], fg=t["baslik_metin"])
            self.lbl_kurulum_alt.configure(bg=t["arka"], fg=t["alt_metin"])
        except Exception:
            pass

        for kart in (self.kart_kriter, self.kart_cikti):
            try:
                kart.configure(bg=t["kart_bg"], bd=1, relief="solid",
                               highlightbackground=t["kart_border"],
                               highlightthickness=1)
            except Exception:
                pass

        try:
            self.lbl_kart_kriter.configure(bg=t["kart_bg"], fg=t["info"])
            self.lbl_kart_kriter_alt.configure(bg=t["kart_bg"], fg=t["muted"])
            self.lbl_kart_cikti.configure(bg=t["kart_bg"], fg=t["info"])
            self.lbl_kart_cikti_alt.configure(bg=t["kart_bg"], fg=t["muted"])
        except Exception:
            pass

        try:
            self.chk_bp.configure(
                bg=t["kart_bg"], fg=t["info"],
                activebackground=t["kart_bg"],
                activeforeground=t["info"],
                selectcolor=t["kart_bg"]
            )
            self.chk_onlisans.configure(
                bg=t["kart_bg"], fg=t["info"],
                activebackground=t["kart_bg"],
                activeforeground=t["info"],
                selectcolor=t["kart_bg"]
            )
        except Exception:
            pass

        try:
            self.rd_excel.configure(
                bg=t["kart_bg"], fg=t["info"],
                activebackground=t["kart_bg"],
                activeforeground=t["info"],
                selectcolor=t["kart_bg"]
            )
            self.rd_txt.configure(
                bg=t["kart_bg"], fg=t["info"],
                activebackground=t["kart_bg"],
                activeforeground=t["info"],
                selectcolor=t["kart_bg"]
            )
        except Exception:
            pass

        try:
            self.btn_devam.configure(bg=t["buton_devam"], fg="white",
                                     activebackground=t["buton_devam"],
                                     activeforeground="white")
            self.btn_tema_k.configure(bg=t["buton_tema"], fg="white",
                                      activebackground=t["buton_tema"])
            self.lbl_surum_k.configure(bg=t["arka"], fg=t["alt_metin"])
        except Exception:
            pass

    def _tema_tarama_uygula(self, t):
        try:
            self.tarama_container.configure(bg=t["arka"])
            self.tarama_ust.configure(bg=t["arka"])
            self.lbl_tarama_baslik.configure(bg=t["arka"], fg=t["baslik_metin"])
            self.lbl_kriter_ozet.configure(bg=t["arka"], fg=t["alt_metin"])

            self.btn_tema_t.configure(bg=t["buton_tema"], fg="white")
            self.btn_ayarlar.configure(bg=t["buton_ayar"], fg="white")

            for lf in (self.durum_frame, self.ilan_frame,
                       self.log1_frame, self.log2_frame):
                lf.configure(bg=t["arka"], fg=t["label_frame_fg"])

            for lbl in (self.lbl_incelenen, self.lbl_bot_durum,
                        self.lbl_bot1_et, self.lbl_bot2_et):
                lbl.configure(bg=t["arka"], fg=t["info"])

            if self.lbl_bp is not None:
                self.lbl_bp.configure(bg=t["arka"], fg=t["basari"])
            if self.lbl_onlisans is not None:
                self.lbl_onlisans.configure(bg=t["arka"], fg=t["basari"])

            self.lbl_ilan_bot1.configure(bg=t["arka"], fg=t["bot1_renk"])
            self.lbl_ilan_bot2.configure(bg=t["arka"], fg=t["bot2_renk"])
            self.lbl_bot1_et.configure(fg=t["bot1_renk"])
            self.lbl_bot2_et.configure(fg=t["bot2_renk"])

            for b_frame in self.ilan_frame.winfo_children():
                b_frame.configure(bg=t["arka"])

            self.log_container.configure(bg=t["arka"])

            # LOG RENKLERİ BURADA AYARLANIYOR
            for log in (self.log1, self.log2):
                log.configure(bg=t["log_bg"], fg=t["log_fg"],
                              insertbackground=t["log_fg"])
                log.tag_config("info", foreground=t["info"])
                log.tag_config("success", foreground=t["basari"])
                log.tag_config("warning", foreground=t["uyari"])
                log.tag_config("error", foreground=t["hata"])
                log.tag_config("muted", foreground=t["muted"])

            self.btnf.configure(bg=t["arka"])
            self.btn_baslat.configure(bg=t["buton_yesil"], fg="white", activebackground=t["buton_yesil"])
            self.btn_durdur.configure(bg=t["buton_kirmizi"], fg="white", activebackground=t["buton_kirmizi"])
            self.lbl_surum_t.configure(bg=t["arka"], fg=t["alt_metin"])
        except Exception:
            pass

    def _tema_degistir(self):
        # Tema değiştir
        if self.tema_adi == "light":
            self.tema = TEMA_DARK
            self.tema_adi = "dark"
        else:
            self.tema = TEMA_LIGHT
            self.tema_adi = "light"

        # Buton yazısını GÜNCELLE (Zıt temayı göster)
        yeni_buton_metni = "🌙 Dark" if self.tema_adi == "light" else "☀️ Light"

        if hasattr(self, "btn_tema_k") and self.btn_tema_k.winfo_exists():
            self.btn_tema_k.configure(text=yeni_buton_metni)
        if hasattr(self, "btn_tema_t") and self.btn_tema_t.winfo_exists():
            self.btn_tema_t.configure(text=yeni_buton_metni)

        self.tema_uygula()

    # --------------------------------------------------------
    def _log_bot1(self, mesaj, tip="info"):
        self._log_yaz(self.log1, mesaj, tip)

    def _log_bot2(self, mesaj, tip="info"):
        self._log_yaz(self.log2, mesaj, tip)

    def _log_yaz(self, log, mesaj, tip):
        log.configure(state="normal")
        log.insert("end", mesaj + "\n", tip)
        log.see("end")
        log.configure(state="disabled")

    def _sayaç_anlik(self, bot_id, incelenen_bot, bp_bot, onlisans_bot):
        try:
            self.root.after(0, self._sayaç_anlik_ui,
                            bot_id, incelenen_bot, bp_bot, onlisans_bot)
        except Exception:
            pass

    def _sayaç_anlik_ui(self, bot_id, incelenen_bot, bp_bot, onlisans_bot):
        if not self.botlar:
            return
        toplam_inc = 0
        toplam_bp = 0
        toplam_onl = 0
        for b in self.botlar:
            toplam_inc += b.incelenen_sayisi
            toplam_bp += len(b.uygun_ilanlar_bp)
            toplam_onl += len(b.uygun_ilanlar_onlisans)

        self.lbl_incelenen.configure(text=f"İncelenen: {toplam_inc}")
        if self.lbl_bp is not None:
            self.lbl_bp.configure(text=f"Bilg. Prog.: {toplam_bp}")
        if self.lbl_onlisans is not None:
            self.lbl_onlisans.configure(text=f"Önlisans: {toplam_onl}")

    def _ilan_guncelle(self, metin):
        try:
            if metin.startswith("Bot 1:"):
                ui_metin = metin[len("Bot 1:"):].strip()
                self.root.after(0, self.lbl_ilan_bot1.configure, {"text": ui_metin})
            elif metin.startswith("Bot 2:"):
                ui_metin = metin[len("Bot 2:"):].strip()
                self.root.after(0, self.lbl_ilan_bot2.configure, {"text": ui_metin})
        except Exception:
            pass

    def _bot_durum_guncelle(self, metin):
        try:
            self.root.after(0, self.lbl_bot_durum.configure, {"text": metin})
        except Exception:
            pass

    def _baslat(self):
        self.btn_baslat.configure(state="disabled")
        self.btn_durdur.configure(state="normal")
        self.tarama_aktif = True

        for log in (self.log1, self.log2):
            log.configure(state="normal")
            log.delete("1.0", "end")
            log.configure(state="disabled")
        self.lbl_ilan_bot1.configure(text="—")
        self.lbl_ilan_bot2.configure(text="—")
        self.lbl_incelenen.configure(text="İncelenen: 0")
        if self.lbl_bp is not None:
            self.lbl_bp.configure(text="Bilg. Prog.: 0")
        if self.lbl_onlisans is not None:
            self.lbl_onlisans.configure(text="Önlisans: 0")

        self.botlar = []
        self.threadler = []
        self.hazir_bot_sayisi = 0
        self._bot_durum_guncelle("Botlar: başlatılıyor")

        secili = {
            "bp": self.kriter_bp_var.get(),
            "onlisans": self.kriter_onlisans_var.get(),
        }

        for i in range(PARALEL_TARAYICI_SAYISI):
            log_cb = self._log_bot1 if i == 0 else self._log_bot2
            bot = KariyerKapisiBot(
                bot_id=i + 1,
                log_cb=log_cb,
                sayaç_cb=self._sayaç_anlik,
                ilan_cb=self._ilan_guncelle,
                secilen_kriterler=secili,
            )
            self.botlar.append(bot)
            th = threading.Thread(
                target=self._botu_calistir,
                args=(bot, i, PARALEL_TARAYICI_SAYISI),
                daemon=True
            )
            self.threadler.append(th)
            th.start()

        threading.Thread(target=self._bitis_bekle_ve_birlestir, daemon=True).start()

    def _botu_calistir(self, bot, bot_index, toplam_bot):
        try:
            sol = (bot_index == 0)
            bot.tarayici_baslat(sol_taraf=sol)
            bot.driver.get(HEDEF_URL)
            if not bot._ilanlarin_yuklenmesini_bekle():
                bot.log("⚠️ İlanlar yüklenemedi.", "warning")
                return
            toplam = len(bot._basvur_butonlarini_bul())
            if toplam == 0:
                bot.log("⚠️ Hiç ilan yok.", "warning")
                return
            pay = toplam // toplam_bot
            kalan = toplam % toplam_bot
            baslangic = bot_index * pay + min(bot_index, kalan)
            bitis = baslangic + pay + (1 if bot_index < kalan else 0)
            bot.log(f"📋 Toplam {toplam} ilan, bu bot {baslangic + 1}-{bitis} arasını tarayacak.", "info")

            self.hazir_bot_sayisi += 1
            if self.hazir_bot_sayisi < len(self.botlar):
                self._bot_durum_guncelle(
                    f"Botlar: {self.hazir_bot_sayisi}/{len(self.botlar)} hazır"
                )
            else:
                self._bot_durum_guncelle(
                    f"Botlar: {len(self.botlar)}/{len(self.botlar)} çalışıyor"
                )
            bot.taramayi_baslat(baslangic, bitis, toplam)
        except Exception as e:
            bot.log(f"❌ Bot hatası: {str(e)}", "error")
        finally:
            bot.tarayiciyi_kapat()
            bot.log("🔒 Tarayıcı kapatıldı.", "info")

    def _bitis_bekle_ve_birlestir(self):
        for th in self.threadler:
            th.join()
        self.root.after(0, self._birlestir_ve_goster)

    def _birlestir_ve_goster(self):
        tum_bp = []
        tum_onlisans = []
        toplam_incelenen = 0
        secili = {
            "bp": self.kriter_bp_var.get(),
            "onlisans": self.kriter_onlisans_var.get(),
        }

        for bot in self.botlar:
            toplam_incelenen += bot.incelenen_sayisi
            for ilan in bot.uygun_ilanlar_bp:
                if not self._listede_var_mi(tum_bp, ilan):
                    tum_bp.append(ilan)
            for ilan in bot.uygun_ilanlar_onlisans:
                if not self._listede_var_mi(tum_onlisans, ilan):
                    tum_onlisans.append(ilan)

        self.lbl_incelenen.configure(text=f"İncelenen: {toplam_incelenen}")
        if self.lbl_bp is not None:
            self.lbl_bp.configure(text=f"Bilg. Prog.: {len(tum_bp)}")
        if self.lbl_onlisans is not None:
            self.lbl_onlisans.configure(text=f"Önlisans: {len(tum_onlisans)}")
        self._bot_durum_guncelle("Botlar: tamamlandı")

        self._log_bot1("=" * 55, "info")
        self._log_bot1("📊 TARAMA TAMAMLANDI", "success")
        self._log_bot1(f"✅ Bilg. Prog.: {len(tum_bp)} | Önlisans: {len(tum_onlisans)}", "info")
        self._log_bot2("=" * 55, "info")
        self._log_bot2("📊 TARAMA TAMAMLANDI", "success")
        self._log_bot2(f"✅ Bilg. Prog.: {len(tum_bp)} | Önlisans: {len(tum_onlisans)}", "info")

        toplam = len(tum_bp) + len(tum_onlisans)
        cikti_tipi = self.cikti_secim.get()

        if toplam == 0:
            self._log_bot1("ℹ️ Uygun ilan bulunamadı, dosya oluşturulmadı.", "warning")
            self._log_bot2("ℹ️ Uygun ilan bulunamadı, dosya oluşturulmadı.", "warning")
        else:
            try:
                masaustu = masaustu_yolu()
                zaman = datetime.now().strftime("%Y%m%d_%H%M%S")

                if cikti_tipi == "excel":
                    dosya_adi = f"uygun_ilanlar_{zaman}.xlsx"
                    dosya_yolu = os.path.join(masaustu, dosya_adi)
                    basarili, hata = excel_yaz(dosya_yolu, tum_bp, tum_onlisans, secili)
                    if basarili:
                        self._log_bot1(f"💾 Excel kaydedildi: {dosya_adi}", "success")
                        self._log_bot2(f"💾 Excel kaydedildi: {dosya_adi}", "success")
                    else:
                        self._log_bot1(f"❌ Excel kaydedilemedi: {hata}", "error")
                        self._log_bot2(f"❌ Excel kaydedilemedi: {hata}", "error")
                else:
                    dosya_adi = f"uygun_ilanlar_{zaman}.txt"
                    dosya_yolu = os.path.join(masaustu, dosya_adi)
                    with open(dosya_yolu, "w", encoding="utf-8") as f:
                        f.write("=" * 60 + "\n")
                        f.write(f"{PROGRAM_ADI}\n")
                        f.write(f"Tarih: {datetime.now().strftime('%d.%m.%Y %H:%M')}\n")
                        f.write("=" * 60 + "\n\n")

                        if secili["bp"]:
                            f.write("=" * 60 + "\n")
                            f.write("[BÖLÜM 1] BİLGİSAYAR PROGRAMCILIĞI MEZUNU İLANLAR\n")
                            f.write("=" * 60 + "\n\n")
                            if tum_bp:
                                for i, ilan in enumerate(tum_bp, 1):
                                    f.write(f"{i}. {ilan['baslik']}\n")
                                    f.write(f"   Alt Başlık: {ilan['alt_baslik']}\n")
                                    f.write(f"   Link: {ilan['url']}\n\n")
                            else:
                                f.write("(Bu kriterde uygun ilan bulunamadı)\n\n")

                        if secili["onlisans"]:
                            f.write("\n" + "=" * 60 + "\n")
                            f.write("[BÖLÜM] HERHANGİ BİR ÖNLİSANS BÖLÜMÜNDEN MEZUN OLMAK ŞARTLI İLANLAR\n")
                            f.write("=" * 60 + "\n\n")
                            if tum_onlisans:
                                for i, ilan in enumerate(tum_onlisans, 1):
                                    f.write(f"{i}. {ilan['baslik']}\n")
                                    f.write(f"   Alt Başlık: {ilan['alt_baslik']}\n")
                                    f.write(f"   Link: {ilan['url']}\n\n")
                            else:
                                f.write("(Bu kriterde uygun ilan bulunamadı)\n\n")

                    self._log_bot1(f"💾 TXT kaydedildi: {dosya_adi}", "success")
                    self._log_bot2(f"💾 TXT kaydedildi: {dosya_adi}", "success")

            except Exception as e:
                self._log_bot1(f"❌ Dosya kaydedilemedi: {str(e)}", "error")
                self._log_bot2(f"❌ Dosya kaydedilemedi: {str(e)}", "error")

        self.tarama_aktif = False
        self.btn_baslat.configure(state="normal")
        self.btn_durdur.configure(state="disabled")

    def _listede_var_mi(self, liste, ilan):
        anahtar = (normalize(ilan["baslik"]), ilan["url"],
                   normalize(ilan.get("alt_baslik", "")))
        for mevcut in liste:
            if (normalize(mevcut["baslik"]), mevcut["url"],
                normalize(mevcut.get("alt_baslik", ""))) == anahtar:
                return True
        return False

    def _durdur(self):
        for bot in self.botlar:
            bot.durdur()
        self.tarama_aktif = False
        self.btn_baslat.configure(state="normal")
        self.btn_durdur.configure(state="disabled")
        self._bot_durum_guncelle("Botlar: durduruldu")

    def calistir(self):
        self.root.mainloop()


# ------------------------------------------------------------
if __name__ == "__main__":
    arayuz = BotArayuz()
    arayuz.calistir()