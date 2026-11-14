# -*- coding: utf-8 -*-
"""
Engin ARSLAN ile Atama Yeşili
Kariyer Kapısı İlan Tarama Aracı - Paralel Sürüm
"""

import os
import sys
import time
import random
import threading
import re
import tkinter as tk
from tkinter import scrolledtext
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
SURUM = "v1.0.2"
PROGRAM_ADI = "Engin ARSLAN ile Atama Yeşili"
HEDEF_URL = "https://kariyerkapisi.gov.tr/isealim"

ARANAN_IFADE_1 = "bilgisayar programcılığı"
ARANAN_IFADE_2_VARYANTLARI = [
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

MIN_BEKLEME = 1.2
MAX_BEKLEME = 2.2
ILAN_YUKLEME_BEKLEME = 25
PARALEL_TARAYICI_SAYISI = 2

# ---------- Tema Renkleri ----------
TEMA_LIGHT = {
    "arka":           "#f0f4f0",
    "baslik_metin":   "#2e7d32",
    "alt_metin":      "#555555",
    "label_frame_fg": "#2e7d32",
    "log_bg":         "#1e1e1e",
    "log_fg":         "#d4d4d4",
    "buton_yesil":    "#4caf50",
    "buton_kirmizi":  "#f44336",
    "buton_tema":     "#455a64",
    "ilan_ad":        "#1565c0",
    "basari":         "#2e7d32",
    "uyari":          "#d84315",
    "hata":           "#c62828",
    "muted":          "#888888",
    "info":           "#d4d4d4",
    "bot1_renk":      "#1565c0",
    "bot2_renk":      "#c2185b",
}

TEMA_DARK = {
    "arka":           "#1f1f1f",
    "baslik_metin":   "#7cd492",
    "alt_metin":      "#aaaaaa",
    "label_frame_fg": "#7cd492",
    "log_bg":         "#0d0d0d",
    "log_fg":         "#e8e8e8",
    "buton_yesil":    "#43a047",
    "buton_kirmizi":  "#e53935",
    "buton_tema":     "#90a4ae",
    "ilan_ad":        "#64b5f6",
    "basari":         "#7cd492",
    "uyari":          "#ffb74d",
    "hata":           "#ef5350",
    "muted":          "#aaaaaa",
    "info":           "#e8e8e8",
    "bot1_renk":      "#64b5f6",
    "bot2_renk":      "#f06292",
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
# BOT
# ------------------------------------------------------------
class KariyerKapisiBot:

    def __init__(self, bot_id, log_cb, sayaç_cb=None, ilan_cb=None):
        self.bot_id = bot_id
        self.driver = None
        self.log = log_cb
        self.sayaç_cb = sayaç_cb or (lambda *a: None)
        self.ilan_cb = ilan_cb or (lambda *a: None)
        self.durduruldu = False

        self.uygun_ilanlar_bp = []
        self.uygun_ilanlar_onlisans = []
        self.incelenen_sayisi = 0

    # --------------------------------------------------------
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

    # --------------------------------------------------------
    def taramayi_baslat(self, baslangic_idx, bitis_idx, toplam):
        self.log("🔗 Siteye gidiliyor...", "info")
        self.driver.get(HEDEF_URL)

        self.log("⏳ İlanların yüklenmesi bekleniyor...", "info")
        if not self._ilanlarin_yuklenmesini_bekle():
            self.log("⚠️ İlanlar yüklenemedi.", "warning")
            return

        rastgele_bekle(2, 3)

        self.log(f"📋 Toplam {toplam} ilan var. "
                 f"Bu bot {baslangic_idx+1}-{min(bitis_idx, toplam)} arasını tarayacak.", "info")

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
                self.log(f"🔎 İlan {idx+1}/{toplam}: {ilan_adi[:70]}", "info")

                self._detaya_git_ve_tara(buton, ilan_adi, idx + 1)
            except StaleElementReferenceException:
                continue
            except Exception as e:
                self.log(f"⚠️ İlan {idx+1} hatası: {str(e)}", "warning")
                continue

            rastgele_bekle()

        self.log("🏁 Bot taraması bitti.", "success")

    # --------------------------------------------------------
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

    # --------------------------------------------------------
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

    # --------------------------------------------------------
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

    # --------------------------------------------------------
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

            for alt_baslik, _ in bp_bulunan:
                self._listeye_ekle(self.uygun_ilanlar_bp, ilan_adi, detay_url, alt_baslik)
                self.log(f"   ✅ BİLG. PROGRAMCILIĞI: {alt_baslik[:60]}", "success")

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

    # --------------------------------------------------------
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
                bp = ARANAN_IFADE_1 in n
                onl = any(v in n for v in ARANAN_IFADE_2_VARYANTLARI)
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
                alt_baslik = self._accordion_basligini_oku(buton) or f"Alt {idx+1}"
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

                bp = ARANAN_IFADE_1 in icerik_norm
                onl = any(v in icerik_norm for v in ARANAN_IFADE_2_VARYANTLARI)

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

    # --------------------------------------------------------
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
        self.root.title(PROGRAM_ADI)
        self.root.geometry("1300x820")
        self.root.resizable(False, False)
        self.tema = TEMA_LIGHT
        self.tema_adi = "light"

        self._pencereyi_ortala(1300, 820)
        self._arayuz_olustur()
        self.tema_uygula()

        self.botlar = []
        self.threadler = []
        self.tarama_aktif = False
        self.hazir_bot_sayisi = 0

    def _pencereyi_ortala(self, g, y):
        self.root.update_idletasks()
        eg, ey = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x, y2 = (eg - g) // 2, (ey - y) // 2
        self.root.geometry(f"{g}x{y}+{x}+{y2}")

    # --------------------------------------------------------
    def _arayuz_olustur(self):
        # --- Başlık ---
        self.baslik_frame = tk.Frame(self.root)
        self.baslik_frame.pack(fill="x", padx=20, pady=(15, 5))

        self.lbl_baslik = tk.Label(
            self.baslik_frame, text=f"{PROGRAM_ADI}  ({SURUM})",
            font=("Segoe UI", 17, "bold")
        )
        self.lbl_baslik.pack(side="left")

        self.btn_tema = tk.Button(
            self.baslik_frame, text="🌙 Dark",
            font=("Segoe UI", 10, "bold"),
            relief="flat", cursor="hand2",
            padx=15, pady=5, command=self._tema_degistir
        )
        self.btn_tema.pack(side="right")

        self.lbl_alt = tk.Label(
            self.root, text="Kariyer Kapısı İlan Tarama Aracı (Paralel Mod)",
            font=("Segoe UI", 10)
        )
        self.lbl_alt.pack(pady=(0, 5))

        # --- Durum Paneli ---
        self.durum_frame = tk.LabelFrame(
            self.root, text="📊 Genel Durum",
            font=("Segoe UI", 10, "bold"),
            padx=12, pady=8
        )
        self.durum_frame.pack(fill="x", padx=20, pady=4)

        self.lbl_incelenen = tk.Label(self.durum_frame, text="İncelenen: 0",
                                      font=("Segoe UI", 10))
        self.lbl_incelenen.pack(side="left", padx=10)

        self.lbl_bp = tk.Label(self.durum_frame, text="Bilg. Prog.: 0",
                               font=("Segoe UI", 10, "bold"))
        self.lbl_bp.pack(side="left", padx=10)

        self.lbl_onlisans = tk.Label(self.durum_frame, text="Önlisans: 0",
                                     font=("Segoe UI", 10, "bold"))
        self.lbl_onlisans.pack(side="left", padx=10)

        self.lbl_bot_durum = tk.Label(self.durum_frame, text="Botlar: hazır",
                                      font=("Segoe UI", 9, "italic"))
        self.lbl_bot_durum.pack(side="right", padx=10)

        # --- Anlık İlan Bilgisi ---
        self.ilan_frame = tk.LabelFrame(
            self.root, text="🔎 Şu An Kontrol Edilen İlanlar",
            font=("Segoe UI", 10, "bold"),
            padx=12, pady=8
        )
        self.ilan_frame.pack(fill="x", padx=20, pady=4)

        b1_frame = tk.Frame(self.ilan_frame)
        b1_frame.pack(fill="x", pady=2)
        self.lbl_bot1_et = tk.Label(b1_frame, text="Bot 1:",
                                     font=("Segoe UI", 9, "bold"))
        self.lbl_bot1_et.pack(side="left", padx=(0, 6))
        self.lbl_ilan_bot1 = tk.Label(
            b1_frame, text="—", font=("Segoe UI", 10, "bold"),
            wraplength=1100, justify="left", anchor="w"
        )
        self.lbl_ilan_bot1.pack(side="left", fill="x", expand=True)

        b2_frame = tk.Frame(self.ilan_frame)
        b2_frame.pack(fill="x", pady=2)
        self.lbl_bot2_et = tk.Label(b2_frame, text="Bot 2:",
                                     font=("Segoe UI", 9, "bold"))
        self.lbl_bot2_et.pack(side="left", padx=(0, 6))
        self.lbl_ilan_bot2 = tk.Label(
            b2_frame, text="—", font=("Segoe UI", 10, "bold"),
            wraplength=1100, justify="left", anchor="w"
        )
        self.lbl_ilan_bot2.pack(side="left", fill="x", expand=True)

        # --- İki Ayrı Log Alanı ---
        self.log_container = tk.Frame(self.root)
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

        # --- Butonlar ---
        self.btnf = tk.Frame(self.root)
        self.btnf.pack(fill="x", padx=20, pady=(4, 10))

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

        # Sürüm (sadece burada gösteriliyor)
        self.lbl_surum = tk.Label(
            self.root, text=f"Sürüm: {SURUM}",
            font=("Segoe UI", 8)
        )
        self.lbl_surum.pack(side="bottom", pady=3)

    # --------------------------------------------------------
    def tema_uygula(self):
        t = self.tema
        self.root.configure(bg=t["arka"])
        self.baslik_frame.configure(bg=t["arka"])
        self.lbl_baslik.configure(bg=t["arka"], fg=t["baslik_metin"])
        self.lbl_alt.configure(bg=t["arka"], fg=t["alt_metin"])

        for lf in (self.durum_frame, self.ilan_frame,
                   self.log1_frame, self.log2_frame):
            lf.configure(bg=t["arka"], fg=t["label_frame_fg"])

        for lbl in (self.lbl_incelenen, self.lbl_bot_durum,
                    self.lbl_bot1_et, self.lbl_bot2_et):
            lbl.configure(bg=t["arka"], fg=t["info"])
        self.lbl_bp.configure(bg=t["arka"], fg=t["basari"])
        self.lbl_onlisans.configure(bg=t["arka"], fg=t["basari"])

        self.lbl_ilan_bot1.configure(bg=t["arka"], fg=t["bot1_renk"])
        self.lbl_ilan_bot2.configure(bg=t["arka"], fg=t["bot2_renk"])
        self.lbl_bot1_et.configure(fg=t["bot1_renk"])
        self.lbl_bot2_et.configure(fg=t["bot2_renk"])

        for log in (self.log1, self.log2):
            log.configure(bg=t["log_bg"], fg=t["log_fg"],
                          insertbackground=t["log_fg"])
            log.tag_config("info", foreground=t["info"])
            log.tag_config("success", foreground=t["basari"])
            log.tag_config("warning", foreground=t["uyari"])
            log.tag_config("error", foreground=t["hata"])
            log.tag_config("muted", foreground=t["muted"])

        self.btnf.configure(bg=t["arka"])

        for child in self.ilan_frame.winfo_children():
            try:
                child.configure(bg=t["arka"])
                for g in child.winfo_children():
                    try:
                        g.configure(bg=t["arka"])
                    except Exception:
                        pass
            except Exception:
                pass

        self.btn_baslat.configure(bg=t["buton_yesil"], fg="white")
        self.btn_durdur.configure(bg=t["buton_kirmizi"], fg="white")
        self.btn_tema.configure(bg=t["buton_tema"], fg="white")
        self.lbl_surum.configure(bg=t["arka"], fg=t["alt_metin"])

    def _tema_degistir(self):
        if self.tema_adi == "light":
            self.tema = TEMA_DARK
            self.tema_adi = "dark"
            self.btn_tema.configure(text="☀️ Light")
        else:
            self.tema = TEMA_LIGHT
            self.tema_adi = "light"
            self.btn_tema.configure(text="🌙 Dark")
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

    # --------------------------------------------------------
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
        self.lbl_bp.configure(text=f"Bilg. Prog.: {toplam_bp}")
        self.lbl_onlisans.configure(text=f"Önlisans: {toplam_onl}")

    # --------------------------------------------------------
    def _ilan_guncelle(self, metin):
        try:
            if metin.startswith("Bot 1:"):
                ui_metin = metin[len("Bot 1:"):].strip()
                self.root.after(0, self.lbl_ilan_bot1.configure, {"text": ui_metin})
            elif metin.startswith("Bot 2:"):
                ui_metin = metin[len("Bot 2:"):].strip()
                self.root.after(0, self.lbl_ilan_bot2.configure, {"text": ui_metin})
            else:
                self.root.after(0, self.lbl_ilan_bot1.configure, {"text": metin})
                self.root.after(0, self.lbl_ilan_bot2.configure, {"text": metin})
        except Exception:
            pass

    def _bot_durum_guncelle(self, metin):
        try:
            self.root.after(0, self.lbl_bot_durum.configure, {"text": metin})
        except Exception:
            pass

    # --------------------------------------------------------
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
        self.lbl_bp.configure(text="Bilg. Prog.: 0")
        self.lbl_onlisans.configure(text="Önlisans: 0")

        self.botlar = []
        self.threadler = []
        self.hazir_bot_sayisi = 0
        self._bot_durum_guncelle("Botlar: başlatılıyor")

        for i in range(PARALEL_TARAYICI_SAYISI):
            log_cb = self._log_bot1 if i == 0 else self._log_bot2
            bot = KariyerKapisiBot(
                bot_id=i + 1,
                log_cb=log_cb,
                sayaç_cb=self._sayaç_anlik,
                ilan_cb=self._ilan_guncelle,
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

            bot.log(f"📋 Toplam {toplam} ilan, bu bot {baslangic+1}-{bitis} arasını tarayacak.", "info")

            # --- Bot hazır → durum metnini güncelle ---
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

        for bot in self.botlar:
            toplam_incelenen += bot.incelenen_sayisi
            for ilan in bot.uygun_ilanlar_bp:
                if not self._listede_var_mi(tum_bp, ilan):
                    tum_bp.append(ilan)
            for ilan in bot.uygun_ilanlar_onlisans:
                if not self._listede_var_mi(tum_onlisans, ilan):
                    tum_onlisans.append(ilan)

        self.lbl_incelenen.configure(text=f"İncelenen: {toplam_incelenen}")
        self.lbl_bp.configure(text=f"Bilg. Prog.: {len(tum_bp)}")
        self.lbl_onlisans.configure(text=f"Önlisans: {len(tum_onlisans)}")
        self._bot_durum_guncelle("Botlar: tamamlandı")

        self._log_bot1("=" * 55, "info")
        self._log_bot1("📊 TARAMA TAMAMLANDI", "success")
        self._log_bot1(f"✅ Bilg. Prog.: {len(tum_bp)} | Önlisans: {len(tum_onlisans)}", "info")
        self._log_bot2("=" * 55, "info")
        self._log_bot2("📊 TARAMA TAMAMLANDI", "success")
        self._log_bot2(f"✅ Bilg. Prog.: {len(tum_bp)} | Önlisans: {len(tum_onlisans)}", "info")

        toplam = len(tum_bp) + len(tum_onlisans)

        if toplam == 0:
            self._log_bot1("ℹ️ Uygun ilan bulunamadı, TXT oluşturulmadı.", "warning")
            self._log_bot2("ℹ️ Uygun ilan bulunamadı, TXT oluşturulmadı.", "warning")
        else:
            try:
                masaustu = masaustu_yolu()
                zaman = datetime.now().strftime("%Y%m%d_%H%M%S")
                dosya_adi = f"uygun_ilanlar_{zaman}.txt"
                dosya_yolu = os.path.join(masaustu, dosya_adi)

                with open(dosya_yolu, "w", encoding="utf-8") as f:
                    f.write("=" * 60 + "\n")
                    f.write(f"{PROGRAM_ADI}\n")
                    f.write(f"Tarih: {datetime.now().strftime('%d.%m.%Y %H:%M')}\n")
                    f.write("=" * 60 + "\n\n")

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

                    f.write("\n" + "=" * 60 + "\n")
                    f.write("[BÖLÜM 2] HERHANGİ BİR ÖNLİSANS BÖLÜMÜNDEN MEZUN OLMAK ŞARTLI İLANLAR\n")
                    f.write("=" * 60 + "\n\n")
                    if tum_onlisans:
                        for i, ilan in enumerate(tum_onlisans, 1):
                            f.write(f"{i}. {ilan['baslik']}\n")
                            f.write(f"   Alt Başlık: {ilan['alt_baslik']}\n")
                            f.write(f"   Link: {ilan['url']}\n\n")
                    else:
                        f.write("(Bu kriterde uygun ilan bulunamadı)\n\n")

                self._log_bot1(f"💾 Kaydedildi: {dosya_adi}", "success")
                self._log_bot2(f"💾 Kaydedildi: {dosya_adi}", "success")
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