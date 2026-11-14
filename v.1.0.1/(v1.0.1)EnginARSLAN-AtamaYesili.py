# -*- coding: utf-8 -*-
"""
Engin ARSLAN ile Atanma Yeşili
Kariyer Kapısı İlan Tarama Aracı - v6 (Doğru akış)

AKIŞ:
1. Ana sayfaya git (kariyerkapisi.gov.tr/isealim)
2. İlanları yükle (#pnlIlanlar dolsun)
3. Her ilan kartındaki "Başvur" butonuna TIKLA
4. Açılan sayfada accordion'ları (TKN 01, ST 01...) TEK TEK AÇ
5. Her accordion içeriğinde "Bilgisayar Programcılığı" ara
6. Uygunsa kaydet, değilse ana sayfaya dön
7. Sonraki ilana geç
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
PROGRAM_ADI = "Engin ARSLAN ile Atanma Yeşili"
HEDEF_URL = "https://kariyerkapisi.gov.tr/isealim"
ARANAN_IFADE = "bilgisayar programcılığı"
MIN_BEKLEME = 1.2
MAX_BEKLEME = 2.2
ILAN_YUKLEME_BEKLEME = 25

# Renkler
RENK_ARKA_PLAN = "#f0f4f0"
RENK_BASLIK = "#2e7d32"
RENK_LOG_ARKA = "#1e1e1e"
RENK_LOG_YAZI = "#d4d4d4"
RENK_BASARI = "#4ec9b0"
RENK_UYARI = "#ce9178"
RENK_HATA = "#f44747"
RENK_BUTON_YESIL = "#4caf50"
RENK_BUTON_KIRMIZI = "#f44336"
RENK_ILAN_ADI = "#1565c0"


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

    def __init__(self, log_cb, durum_cb=None, ilan_cb=None):
        self.driver = None
        self.log = log_cb
        self.durum_guncelle = durum_cb or (lambda *a: None)
        self.ilan_adi_guncelle = ilan_cb or (lambda *a: None)
        self.durduruldu = False
        self.uygun_ilanlar = []
        self.incelenen_sayisi = 0
        self.islenen_kartlar = set()  # Aynı kartı iki kez işlememek için

    # --------------------------------------------------------
    def tarayici_baslat(self):
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
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-popup-blocking")
        options.add_experimental_option("detach", True)

        self.driver = webdriver.Chrome(options=options)
        self.driver.execute_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        self.log("✅ Chrome hazır.", "success")
        return self.driver

    # --------------------------------------------------------
    def taramayi_baslat(self):
        self.log(f"🔗 Siteye gidiliyor: {HEDEF_URL}", "info")
        self.driver.get(HEDEF_URL)

        self.log("⏳ İlanların yüklenmesi bekleniyor...", "info")
        if self._ilanlarin_yuklenmesini_bekle():
            self.log("✅ İlanlar yüklendi.", "success")
        else:
            self.log("⚠️ İlanlar yüklenemedi veya sayfa boş.", "warning")

        rastgele_bekle(2, 3)
        self.tarama_dongusu()

    # --------------------------------------------------------
    def _ilanlarin_yuklenmesini_bekle(self):
        """#pnlIlanlar içine 'Başvur' butonu gelene kadar bekler."""
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
    def tarama_dongusu(self):
        self.durduruldu = False
        self.uygun_ilanlar = []
        self.incelenen_sayisi = 0

        try:
            self.durum_guncelle("-", self.incelenen_sayisi, len(self.uygun_ilanlar))

            # #pnlIlanlar içindeki "Başvur" butonlarını bul
            basvur_butonlari = self._basvur_butonlarini_bul()

            if not basvur_butonlari:
                self.log("ℹ️ Başvur butonu bulunamadı.", "warning")
                return

            self.log(f"📋 {len(basvur_butonlari)} ilan bulundu.", "info")

            # Her başvur butonu için işlem yap
            for idx in range(len(basvur_butonlari)):
                if self.durduruldu:
                    break

                try:
                    # Butonları her seferinde YENİDEN bul (DOM değişiyor)
                    guncel_butonlar = self._basvur_butonlarini_bul()
                    if idx >= len(guncel_butonlar):
                        break
                    buton = guncel_butonlar[idx]

                    # İlan adını al (karttan)
                    ilan_adi = self._kart_ilan_adini_bul(buton, idx + 1)
                    self.ilan_adi_guncelle(ilan_adi)
                    self.log(f"🔎 İlan {idx+1}/{len(basvur_butonlari)}: {ilan_adi[:70]}", "info")

                    # Butona tıkla ve detay sayfasına git
                    self._detaya_git_ve_tara(buton, ilan_adi)

                except StaleElementReferenceException:
                    self.log(f"⚠️ İlan {idx+1} eski, atlanıyor.", "warning")
                    continue
                except Exception as e:
                    self.log(f"⚠️ İlan {idx+1} hatası: {str(e)}", "warning")
                    continue

                rastgele_bekle()

        except Exception as e:
            self.log(f"❌ Tarama hatası: {str(e)}", "error")
        finally:
            self._ozet_goster()

    # --------------------------------------------------------
    def _basvur_butonlarini_bul(self):
        """
        #pnlIlanlar içindeki tüm 'Başvur' butonlarını/linklerini bulur.
        Her seferinde DOM'dan yeniden okur (dinamik).
        """
        try:
            pnl = self.driver.find_element(By.ID, "pnlIlanlar")
            butonlar = pnl.find_elements(
                By.XPATH,
                ".//a[contains(translate(., 'BAŞVUR', 'başvur'), 'başvur')] | "
                ".//button[contains(translate(., 'BAŞVUR', 'başvur'), 'başvur')] | "
                ".//a[contains(@class, 'btn')][contains(translate(., 'BAŞVUR', 'başvur'), 'başvur')]"
            )
            return butonlar
        except Exception:
            return []

    # --------------------------------------------------------
    def _kart_ilan_adini_bul(self, buton, sira):
        """
        Bir 'Başvur' butonunun bağlı olduğu karttaki ilan adını bulur.
        HTML'de üst row'da kurum adı ve ilan adı yazar.
        """
        try:
            # Butonun en yakın kart/row kapsayıcısını bul
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

            # Karttaki metni al
            metin = kart.text.strip()
            if metin:
                satirlar = [s.strip() for s in metin.split("\n") if s.strip()]
                # "Başvur", "Detay" gibi genel metinleri atla
                atlanacak = {"başvur", "başvuru", "detay", "detaylar", "i̇ncele", "incele"}
                adaylar = []
                for s in satirlar:
                    if normalize(s) not in atlanacak and len(s) > 3:
                        adaylar.append(s)
                if adaylar:
                    # İlk 2 satırı birleştir (kurum + ilan adı olabilir)
                    return " - ".join(adaylar[:2])[:150]

            # Alternatif: h1-h6 veya bold içindeki metin
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
    def _detaya_git_ve_tara(self, buton, ilan_adi):
        """
        'Başvur' butonuna tıklar, açılan sayfadaki accordion'ları tarar,
        sonra ana sayfaya döner.
        """
        ana_pencere = self.driver.current_window_handle
        yeni_pencere = False

        try:
            # Butona tıkla
            try:
                buton.click()
            except ElementClickInterceptedException:
                self.driver.execute_script("arguments[0].click();", buton)

            rastgele_bekle(2, 3)

            # Yeni pencere açıldıysa ona geç
            if len(self.driver.window_handles) > 1:
                self.driver.switch_to.window(self.driver.window_handles[-1])
                yeni_pencere = True

            # Sayfanın yüklenmesini bekle
            rastgele_bekle(3, 4)

            # Bu sayfadaki accordion'ları tara
            bulundu = self._sayfadaki_accordionlari_tara(ilan_adi)

            self.incelenen_sayisi += 1
            self.durum_guncelle("-", self.incelenen_sayisi, len(self.uygun_ilanlar))

            if bulundu:
                self.uygun_ilanlar.append({
                    "baslik": ilan_adi,
                    "url": self.driver.current_url
                })
                self.durum_guncelle("-", self.incelenen_sayisi, len(self.uygun_ilanlar))
                self.log(f"✅ UYGUN: {ilan_adi}", "success")
            else:
                self.log(f"❌ Uygun değil: {ilan_adi[:80]}", "muted")

        except Exception as e:
            self.log(f"⚠️ Detay hatası: {str(e)}", "warning")
        finally:
            # Ana pencereye dön
            try:
                if yeni_pencere and len(self.driver.window_handles) > 1:
                    self.driver.close()
                    self.driver.switch_to.window(ana_pencere)
                else:
                    self.driver.switch_to.window(ana_pencere)
                    self.driver.back()
                rastgele_bekle(2, 3)
            except Exception:
                try:
                    self.driver.switch_to.window(ana_pencere)
                except Exception:
                    pass

    # --------------------------------------------------------
    def _sayfadaki_accordionlari_tara(self, ana_ilan_adi):
        """
        Detay sayfasındaki TÜM accordion'ları TEK TEK açar ve içeriğini okur.
        Eğer herhangi bir accordion içinde ARANAN_IFADE geçiyorsa True döner.
        """
        bulundu = False

        # Accordion butonlarını bul (data-bs-toggle="collapse" veya data-toggle="collapse")
        try:
            butonlar = self.driver.find_elements(
                By.XPATH,
                "//button[@data-bs-toggle='collapse' or @data-toggle='collapse']"
            )
        except Exception:
            butonlar = []

        if not butonlar:
            # Accordion yoksa tüm sayfa metnine bak
            try:
                metin = self.driver.find_element(By.TAG_NAME, "body").text
                if ARANAN_IFADE in normalize(metin):
                    return True
            except Exception:
                pass
            return False

        self.log(f"   🔽 {len(butonlar)} açılır alan bulundu.", "info")

        for idx, buton in enumerate(butonlar):
            if self.durduruldu:
                break

            try:
                # Accordion başlığını oku (ilan alt kırılımı adı)
                alt_baslik = self._accordion_basligini_oku(buton)
                if not alt_baslik:
                    alt_baslik = f"Alt Alan {idx+1}"

                # Arayüzde alt başlıkla birlikte göster
                goster = f"{ana_ilan_adi}  →  {alt_baslik}"
                self.ilan_adi_guncelle(goster[:150])

                # Sayfayı o alana kaydır (görsel takip)
                try:
                    self.driver.execute_script(
                        "arguments[0].scrollIntoView({block:'center', behavior:'smooth'});",
                        buton
                    )
                    time.sleep(0.7)
                except Exception:
                    pass

                # Accordion açık mı kontrol et, kapalıysa aç
                self._accordion_ac(buton)
                time.sleep(0.5)

                # İçeriği oku
                icerik = self._accordion_icerigini_oku(buton)
                icerik_norm = normalize(icerik)

                if ARANAN_IFADE in icerik_norm:
                    bulundu = True
                    self.log(f"   ✅ UYGUN (alt): {alt_baslik[:60]}", "success")
                else:
                    self.log(f"   ❌ Uygun değil (alt): {alt_baslik[:60]}", "muted")

                rastgele_bekle(0.4, 0.9)

            except StaleElementReferenceException:
                continue
            except Exception as e:
                self.log(f"   ⚠️ Alt alan hatası: {str(e)}", "warning")
                continue

        return bulundu

    # --------------------------------------------------------
    def _accordion_basligini_oku(self, buton):
        """Accordion başlığını (h6, h5 vs.) okur."""
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

    # --------------------------------------------------------
    def _accordion_ac(self, buton):
        """
        Accordion kapalıysa açar. Açıksa (class'ta 'show' varsa) tıklamaz.
        """
        try:
            hedef_id = (buton.get_attribute("data-bs-target") or
                        buton.get_attribute("data-target") or "")
            if hedef_id:
                hedef_id = hedef_id.lstrip("#.")
                try:
                    icerik_div = self.driver.find_element(By.ID, hedef_id)
                    class_attr = (icerik_div.get_attribute("class") or "").split()
                    if "show" in class_attr:
                        return True  # Zaten açık
                except Exception:
                    pass

            try:
                buton.click()
            except ElementClickInterceptedException:
                self.driver.execute_script("arguments[0].click();", buton)
            except Exception:
                self.driver.execute_script("arguments[0].click();", buton)

            time.sleep(0.5)
            return True
        except Exception:
            return False

    # --------------------------------------------------------
    def _accordion_icerigini_oku(self, buton):
        """data-bs-target ile hedef div'in içeriğini okur."""
        parcalar = []

        # 1) data-bs-target ile hedef div
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

        # 2) Ebeveyn div'in tamamı
        try:
            parent = buton.find_element(By.XPATH, "./..")
            parcalar.append(parent.text or "")
        except Exception:
            pass

        return "\n".join(p for p in parcalar if p)

    # --------------------------------------------------------
    def _ozet_goster(self):
        self.log("=" * 55, "info")
        self.log(f"📊 TARAMA TAMAMLANDI", "success")
        self.log(f"📄 Toplam incelenen: {self.incelenen_sayisi}", "info")
        self.log(f"✅ Uygun bulunan: {len(self.uygun_ilanlar)}", "success")

        if not self.uygun_ilanlar:
            self.log("ℹ️ Uygun ilan bulunamadı.", "warning")
            self.log("=" * 55, "info")
            return

        try:
            masaustu = masaustu_yolu()
            zaman = datetime.now().strftime("%Y%m%d_%H%M%S")
            dosya_adi = f"uygun_ilanlar_{zaman}.txt"
            dosya_yolu = os.path.join(masaustu, dosya_adi)

            with open(dosya_yolu, "w", encoding="utf-8") as f:
                f.write(f"{PROGRAM_ADI}\n")
                f.write(f"Tarih: {datetime.now().strftime('%d.%m.%Y %H:%M')}\n")
                f.write(f"Aranan İfade: {ARANAN_IFADE.title()}\n")
                f.write("=" * 60 + "\n\n")
                for i, ilan in enumerate(self.uygun_ilanlar, 1):
                    f.write(f"{i}. {ilan['baslik']}\n")
                    f.write(f"   Link: {ilan['url']}\n\n")

            self.log(f"💾 Sonuçlar kaydedildi: {dosya_adi}", "success")
        except Exception as e:
            self.log(f"❌ Dosya kaydedilemedi: {str(e)}", "error")

        self.log("=" * 55, "info")

    # --------------------------------------------------------
    def durdur(self):
        self.durduruldu = True
        self.log("⏹ Tarama durduruluyor...", "warning")

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
        self.root.geometry("820x700")
        self.root.resizable(False, False)
        self.root.configure(bg=RENK_ARKA_PLAN)
        self._pencereyi_ortala(820, 700)
        self._arayuz_olustur()
        self.bot = None
        self.tarama_thread = None

    def _pencereyi_ortala(self, g, y):
        self.root.update_idletasks()
        eg, ey = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x, y2 = (eg - g) // 2, (ey - y) // 2
        self.root.geometry(f"{g}x{y}+{x}+{y2}")

    def _arayuz_olustur(self):
        bf = tk.Frame(self.root, bg=RENK_ARKA_PLAN)
        bf.pack(fill="x", padx=25, pady=(20, 10))
        tk.Label(bf, text=PROGRAM_ADI, font=("Segoe UI", 18, "bold"),
                 fg=RENK_BASLIK, bg=RENK_ARKA_PLAN).pack()
        tk.Label(bf, text="Kariyer Kapısı İlan Tarama Aracı",
                 font=("Segoe UI", 11), fg="#555555",
                 bg=RENK_ARKA_PLAN).pack(pady=(5, 0))

        df = tk.LabelFrame(self.root, text="📊 Tarama Durumu",
                           font=("Segoe UI", 10, "bold"), bg=RENK_ARKA_PLAN,
                           fg=RENK_BASLIK, padx=15, pady=10)
        df.pack(fill="x", padx=25, pady=5)

        self.lbl_incelenen = tk.Label(df, text="İncelenen: 0",
                                      font=("Segoe UI", 10), bg=RENK_ARKA_PLAN)
        self.lbl_incelenen.pack(side="left", padx=15)

        self.lbl_uygun = tk.Label(df, text="Uygun: 0",
                                  font=("Segoe UI", 10, "bold"),
                                  fg=RENK_BASARI, bg=RENK_ARKA_PLAN)
        self.lbl_uygun.pack(side="left", padx=15)

        ilan_frame = tk.LabelFrame(self.root, text="🔎 Şu An Kontrol Edilen İlan",
                                   font=("Segoe UI", 10, "bold"),
                                   bg=RENK_ARKA_PLAN, fg=RENK_ILAN_ADI,
                                   padx=15, pady=10)
        ilan_frame.pack(fill="x", padx=25, pady=5)

        self.lbl_ilan_adi = tk.Label(
            ilan_frame, text="—",
            font=("Segoe UI", 10, "bold"),
            fg=RENK_ILAN_ADI, bg=RENK_ARKA_PLAN,
            wraplength=740, justify="left", anchor="w"
        )
        self.lbl_ilan_adi.pack(fill="x")

        lf = tk.LabelFrame(self.root, text="📝 İşlem Günlüğü",
                           font=("Segoe UI", 10, "bold"), bg=RENK_ARKA_PLAN,
                           fg=RENK_BASLIK, padx=5, pady=5)
        lf.pack(fill="both", expand=True, padx=25, pady=5)

        self.log_alani = scrolledtext.ScrolledText(
            lf, font=("Consolas", 9), bg=RENK_LOG_ARKA, fg=RENK_LOG_YAZI,
            insertbackground=RENK_LOG_YAZI, wrap="word", state="disabled"
        )
        self.log_alani.pack(fill="both", expand=True)

        self.log_alani.tag_config("info", foreground=RENK_LOG_YAZI)
        self.log_alani.tag_config("success", foreground=RENK_BASARI)
        self.log_alani.tag_config("warning", foreground=RENK_UYARI)
        self.log_alani.tag_config("error", foreground=RENK_HATA)
        self.log_alani.tag_config("muted", foreground="#808080")

        btnf = tk.Frame(self.root, bg=RENK_ARKA_PLAN)
        btnf.pack(fill="x", padx=25, pady=(5, 20))

        self.btn_baslat = tk.Button(
            btnf, text="▶ Taramayı Başlat", font=("Segoe UI", 11, "bold"),
            bg=RENK_BUTON_YESIL, fg="white", relief="flat", cursor="hand2",
            padx=30, pady=10, command=self._baslat
        )
        self.btn_baslat.pack(side="left", padx=8)

        self.btn_durdur = tk.Button(
            btnf, text="⏹ Durdur", font=("Segoe UI", 11, "bold"),
            bg=RENK_BUTON_KIRMIZI, fg="white", relief="flat", cursor="hand2",
            padx=30, pady=10, state="disabled", command=self._durdur
        )
        self.btn_durdur.pack(side="left", padx=8)

    def _log_mesaj(self, mesaj, tip="info"):
        self.log_alani.configure(state="normal")
        self.log_alani.insert("end", mesaj + "\n", tip)
        self.log_alani.see("end")
        self.log_alani.configure(state="disabled")

    def _durum_guncelle(self, sayfa=None, incelenen=None, uygun=None):
        if incelenen is not None:
            self.lbl_incelenen.configure(text=f"İncelenen: {incelenen}")
        if uygun is not None:
            self.lbl_uygun.configure(text=f"Uygun: {uygun}")

    def _ilan_adi_guncelle(self, metin):
        try:
            self.lbl_ilan_adi.configure(text=metin if metin else "—")
        except Exception:
            pass

    def _baslat(self):
        self.btn_baslat.configure(state="disabled")
        self.btn_durdur.configure(state="normal")
        self._log_mesaj("=" * 55, "info")
        self._log_mesaj("🚀 Bot başlatılıyor...", "info")

        self.bot = KariyerKapisiBot(
            self._log_mesaj, self._durum_guncelle, self._ilan_adi_guncelle
        )
        self.tarama_thread = threading.Thread(target=self._botu_calistir, daemon=True)
        self.tarama_thread.start()

    def _botu_calistir(self):
        try:
            self.bot.tarayici_baslat()
            self.bot.taramayi_baslat()
        except Exception as e:
            self.root.after(0, lambda: self._log_mesaj(f"❌ Hata: {str(e)}", "error"))
        finally:
            self.root.after(0, self._tarama_bitti)

    def _durdur(self):
        if self.bot:
            self.bot.durdur()
            self.bot.tarayiciyi_kapat()
        self._tarama_bitti()
        self._log_mesaj("⏹ Tarama durduruldu.", "warning")

    def _tarama_bitti(self):
        self.btn_baslat.configure(state="normal")
        self.btn_durdur.configure(state="disabled")

    def calistir(self):
        self.root.mainloop()


# ------------------------------------------------------------
if __name__ == "__main__":
    arayuz = BotArayuz()
    arayuz.calistir()