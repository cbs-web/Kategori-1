import tkinter as tk
from tkinter import ttk, messagebox
import ttkbootstrap as tb
import logging
import sys
import tkintermapview
import os
import tempfile
from copy import deepcopy
from ac_yn_islemleri import AcYnIslemleri
from arayuz_yardimcilari import ArayuzYardimcilari
from raporlama_islemleri import RaporlamaIslemleri
from ekler_islemleri import EklerIslemleri
from harita_islemleri import HaritaIslemleri
from harita_renkleri import JEOLOJI_HARITA_RENK_ACIKLAMASI
from jeofizik_islemleri import JeofizikIslemleri
from jeoloji_kutuphanesi_islemleri import JeolojiKutuphanesiIslemleri
from jeoloji_pafta_islemleri import JeolojiPaftaIslemleri
from genel_jeoloji_islemleri import GenelJeolojiIslemleri
from is_takibi import IsTakibiIslemleri
from kayit import KayitYoneticisi
from laboratuvar_islemleri import LaboratuvarIslemleri
from on_deger import bos_is_akisi_verisi, bos_on_deger_verisi, bos_tdth_verisi
from on_deger_islemleri import OnDegerIslemleri
from proje_durumu_islemleri import ProjeDurumuIslemleri
from taahhutname_islemleri import TaahhutnameIslemleri
from tasima_islemleri import TasimaIslemleri
from temel_bilgiler_islemleri import TemelBilgilerIslemleri


UYGULAMA_ADI = "K-1"


def kullanici_veri_klasoru_yolu_bul():
    """Yazılabilir, kullanıcıya özel uygulama veri klasörünü döndürür."""
    adaylar = []
    for ortam_adi in ("LOCALAPPDATA", "APPDATA"):
        kok = os.environ.get(ortam_adi)
        if kok:
            adaylar.append(os.path.join(kok, UYGULAMA_ADI))
    adaylar.append(os.path.join(os.path.expanduser("~"), f".{UYGULAMA_ADI.lower()}"))
    adaylar.append(os.path.join(tempfile.gettempdir(), UYGULAMA_ADI))

    denenmis = set()
    for aday in adaylar:
        aday = os.path.abspath(aday)
        anahtar = os.path.normcase(aday)
        if anahtar in denenmis:
            continue
        denenmis.add(anahtar)
        try:
            os.makedirs(aday, exist_ok=True)
            fd, deneme_yolu = tempfile.mkstemp(prefix=".yazma_deneme_", dir=aday)
            os.close(fd)
            os.remove(deneme_yolu)
            return aday
        except OSError:
            continue
    return None


KULLANICI_VERI_KLASORU = kullanici_veri_klasoru_yolu_bul()

def log_dosyasi_yolu_bul():
    if not KULLANICI_VERI_KLASORU:
        return None
    log_yolu = os.path.join(KULLANICI_VERI_KLASORU, "hata_kaydi.log")
    try:
        with open(log_yolu, "a", encoding="utf-8"):
            pass
        return log_yolu
    except OSError:
        return None

LOG_DOSYASI = log_dosyasi_yolu_bul()
LOG_AYARLARI = {
    "level": logging.INFO,
    "format": "%(asctime)s [%(levelname)s] %(message)s",
}
if LOG_DOSYASI:
    LOG_AYARLARI.update({"filename": LOG_DOSYASI, "encoding": "utf-8"})
else:
    LOG_AYARLARI["stream"] = sys.stderr
logging.basicConfig(**LOG_AYARLARI)
logger = logging.getLogger("ZeminRaporPro")

DESTEK_KLASORLERI = ("rapor", "taahhutname", "excel", "kml", "gorseller", "word", "jeoloji")
DESTEK_DOSYALARI = [
    ("rapor", "TASLAK.docx"),
    ("taahhutname", "örnek taahütname.docx"),
    ("word", "AÇ Log.docx"),
    ("word", "Bina Bilgileri.docx"),
    ("word", "Jeofizik Koordinat.docx"),
    ("word", "Tasimagücütablo.docx"),
    ("excel", "Jeofizik Parametre.xlsx"),
    ("excel", "LAB.xlsx"),
    ("excel", "LAB_1.xlsx"),
    ("gorseller", "MJH.jpg"),
    ("kml", "tkgm-parsel-sorgu-sonuc-119-ada-7-parsel.kml"),
]

# --- ANA UYGULAMA ARAYÜZÜ ---
class RaporProApp:
    def __init__(self, root):
        self.root = root
        self.root.title("K-1 - Zemin Etüt Programı")
        self.root.geometry("1200x800")
        self.arayuz_stillerini_hazirla()
        logger.info("Program başlatıldı. Log dosyası: %s", LOG_DOSYASI or "stderr")
        self.ornek_dosya_klasoru = self.destek_dosyalari_hazirla()
        
        self.guncel_dosya_yolu = None
        self.son_projeler = self.son_projeleri_yukle()
        self.veri_alanlari = {} 
        self.bina_alanlari = {}
        self.ac_yn_sekme_bilgileri = {}
        self.ek_kategorileri = ["EVRAKLAR", "LOG", "LABORATUVAR", "JEOFİZİK", "FOTOĞRAFLAR", "TDTH"]
        self.ekler = {kategori: [] for kategori in self.ek_kategorileri}
        self.ek_treeviewler = {}
        self.ek_durum_etiketleri = {}
        self.on_deger_verisi = bos_on_deger_verisi()
        self.tdth_verisi = bos_tdth_verisi()
        self.is_akisi_verisi = bos_is_akisi_verisi()
        self.proje_salt_okunur = False
        self.animasyonlar_aktif = True
        # Etkileşim sırasında tüm proje ağacını yeniden serileştirmek yerine
        # ucuz bir kirli bayrağı tutulur. Tam karşılaştırma yalnızca dosya
        # değiştirme/kapatma gibi doğrulama gereken sınır noktalarında yapılır.
        self._proje_kirli = False
        self.taahhut_varsayilanlari = {
            "JEOFIZIK_MUH_AD": "",
            "JEOFIZIK_MUH_SICIL": "",
            "JEOFIZIK_MUH_ADRES": "",
            "JEOFIZIK_MUH_TELEFON": "",
            "JEOLOJI_MUH_AD": "",
            "JEOLOJI_MUH_SICIL": "",
            "JEOLOJI_MUH_ADRES": "",
            "JEOLOJI_MUH_TELEFON": "",
        }
        self.taahhut_bilgileri = dict(self.taahhut_varsayilanlari)
        self.sablon_yolu = self.varsayilan_rapor_sablonu()
        self.jeoloji_sablon_yolu = ""
        self.jeoloji_kutuphanesi_bolumu_aktif = False
        self.jeoloji_kutuphanesi_kayit_id = None
        self.jeoloji_kutuphanesi_bolum_yolu = ""
        self.jeoloji_kutuphanesi_bolum_hash = ""
        self.jeoloji_kutuphanesi_uygulanan_genel = ""
        self.jeoloji_kutuphanesi_uygulanan_inceleme = ""
        self.taahhut_word_sablon_yolu = self.varsayilan_taahhut_word_sablonu()
        self.son_qk = "-" 
        self.son_qt = "-" 
        
        self.img_mjh = None
        self.img_jeofizik_lok = None
        self.img_jeoloji_lok = None
        self.img_yerbulduru = None
        self.img_parsel_haritasi = None
        self.img_genel_jeoloji = None
        self.img_pga_haritasi = None
        self.parsel_haritasi_kaynak_url = ""
        self.parsel_haritasi_geometri_hash = ""
        self.parsel_haritasi_ada = ""
        self.parsel_haritasi_parsel = ""

        self.harita_nokta_sayaclari = {"AÇ": 1, "YN": 1, "SS": 1}
        self.harita_isaretleri = {} 
        self.ss_ilk_nokta = None
        self.temp_ss_marker = None
        self.kml_polygon_obj = None
        self.yuklu_kml_yolu = ""
        self.yuklu_kml_points = []
        self.jeoloji_pafta_sonucu = {}
        self.genel_jeoloji_verisi = {}
        self.jeoloji_kutuphane_polygonlari = []
        self.jeoloji_kutuphane_secili_polygon = None
        self._jeoloji_harita_after_id = None
        self._jeoloji_harita_son_gorunum = None
        
        # Formasyon Listesi (İstediğiniz zaman buraya yeni eklemeler yapabilirsiniz)
        self.formasyonlar = [
            "Seçiniz...",
            "Ayvacık Volkaniti (Tmay)",
            "Ezine Volkaniti (Tme)",
            "Hüseyinfakı Volkaniti (Tmhü)",
            "Arıklı İgnimbiriti (Tmar)",
            "İlyasbaşı Formasyonu (Tmi)",
            "Çamkabalak İgnimbiriti (Tmç)",
            "Çetmi Melanjı (Kç)",
            "Üst Oligosen-Alt Miyosen Granitoyitleri (Tg)",
            "Şahinli Formasyonu (Teşa)",
            "Bayramiç Formasyonu (Tplb)",
            "Çanakkale Formasyonu (Tmçk)",
            "Kirazlı Üyesi (Tmki)",
            "Çamrakdere Üyesi (Tmçd)",
            "Alçıtepe Üyesi (Tmal)",
            "Tüf Üyesi (Tmçt)"
        ]
        
        # Seçilen formasyon için şablon metinler
        self.formasyon_metinleri = {
            "Seçiniz...": "",
            "Ayvacık Volkaniti (Tmay)": '"Tmay" simgesiyle gösterilen "Ayvacık Volkaniti" adı ile anılan birimler',
            "Ezine Volkaniti (Tme)": '"Tme" simgesiyle gösterilen "Ezine Volkaniti" adı ile anılan',
            "Hüseyinfakı Volkaniti (Tmhü)": '"Tmhü" simgesiyle gösterilen "Hüseyinfakı Volkaniti" adı ile anılan',
            "Arıklı İgnimbiriti (Tmar)": '"Tmar" simgesiyle gösterilen "Arıklı İgnimbiriti" adı ile anılan',
            "İlyasbaşı Formasyonu (Tmi)": '"Tmi" simgesiyle gösterilen "İlyasbaşı Formasyonu" adı ile anılan',
            "Çamkabalak İgnimbiriti (Tmç)": '"Tmç" simgesiyle gösterilen "Çamkabalak İgnimbiriti" adı ile anılan',
            "Çetmi Melanjı (Kç)": '"Kç" simgesiyle gösterilen "Çetmi Melanjı" adı ile anılan',
            "Üst Oligosen-Alt Miyosen Granitoyitleri (Tg)": '"Tg" simgesiyle gösterilen "Üst Oligosen-Alt Miyosen Granitoyitleri" adı ile anılan',
            "Şahinli Formasyonu (Teşa)": '"Teşa" simgesiyle gösterilen "Şahinli Formasyonu" adı ile anılan',
            "Bayramiç Formasyonu (Tplb)": '"Tplb" simgesiyle gösterilen "Bayramiç Formasyonu" adı ile anılan',
            "Çanakkale Formasyonu (Tmçk)": '"Tmçk" simgesiyle gösterilen "Çanakkale Formasyonu" adı ile anılan',
            "Kirazlı Üyesi (Tmki)": '"Tmki" simgesiyle gösterilen "Kirazlı Üyesi" adı ile anılan',
            "Çamrakdere Üyesi (Tmçd)": '"Tmçd" simgesiyle gösterilen "Çamrakdere Üyesi" adı ile anılan',
            "Alçıtepe Üyesi (Tmal)": '"Tmal" simgesiyle gösterilen "Alçıtepe Üyesi" adı ile anılan',
            "Tüf Üyesi (Tmçt)": '"Tmçt" simgesiyle gösterilen "Tüf Üyesi" adı ile anılan'
        }

        self.ikonlari_olustur() 
        self.menu_olustur()
        self.sekmeleri_olustur()
        self.durum_cubugu_olustur()
        self.durum_mesaji_yaz("Hazır", "Yeni proje")
        
        ilk_veri = self.verileri_topla()
        self.varsayilan_proje_verisi = deepcopy(ilk_veri)
        self.son_kayit_verisi = deepcopy(ilk_veri)
        self._proje_kirli = False
        self.proje_durum_seridi_guncelle(kaydedilmedi=False)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def hata_kaydet(self, baslik, hata=None):
        if hata is None:
            logger.exception(baslik)
        else:
            logger.error("%s: %s", baslik, hata, exc_info=True)

    def uygulama_klasoru_bul(self):
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))

    def kullanici_veri_klasoru_bul(self):
        if not KULLANICI_VERI_KLASORU:
            raise OSError("K-1 için yazılabilir kullanıcı veri klasörü bulunamadı.")
        return KULLANICI_VERI_KLASORU

    def sablon_kok_adaylari(self):
        adaylar = []
        if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
            adaylar.append(os.path.join(sys._MEIPASS, "ornek_sablonlar"))
        adaylar.append(os.path.join(self.uygulama_klasoru_bul(), "ornek_sablonlar"))
        if KULLANICI_VERI_KLASORU:
            adaylar.append(os.path.join(KULLANICI_VERI_KLASORU, "ornek_sablonlar"))

        sonuc = []
        gorulen = set()
        for aday in adaylar:
            tam_yol = os.path.abspath(aday)
            anahtar = os.path.normcase(tam_yol)
            if anahtar not in gorulen:
                gorulen.add(anahtar)
                sonuc.append(tam_yol)
        return sonuc

    def sablon_kok_klasoru(self):
        for aday in self.sablon_kok_adaylari():
            if os.path.isdir(aday):
                return aday
        if KULLANICI_VERI_KLASORU:
            return os.path.join(KULLANICI_VERI_KLASORU, "ornek_sablonlar")
        return os.path.join(self.uygulama_klasoru_bul(), "ornek_sablonlar")

    def sablon_alt_klasoru(self, kategori):
        return os.path.join(self.sablon_kok_klasoru(), kategori)

    def kaynak_dosya_adaylari(self, kategori, dosya_adi):
        adaylar = []
        for sablon_kok in self.sablon_kok_adaylari():
            adaylar.append(os.path.join(sablon_kok, kategori, dosya_adi))
            adaylar.append(os.path.join(sablon_kok, dosya_adi))
        adaylar.append(os.path.join(self.uygulama_klasoru_bul(), dosya_adi))
        if kategori in ("rapor", "taahhutname"):
            adaylar.append(os.path.join(os.path.expanduser("~"), "Desktop", "K1 Dosya", dosya_adi))
        return adaylar

    def destek_dosyasi_bul(self, kategori, dosya_adi):
        for aday in self.kaynak_dosya_adaylari(kategori, dosya_adi):
            if os.path.exists(aday):
                return aday
        return ""

    def destek_dosyalari_hazirla(self):
        hedef_klasor = self.sablon_kok_klasoru()
        try:
            if not os.path.isdir(hedef_klasor) and KULLANICI_VERI_KLASORU:
                os.makedirs(hedef_klasor, exist_ok=True)
                for kategori in DESTEK_KLASORLERI:
                    os.makedirs(os.path.join(hedef_klasor, kategori), exist_ok=True)
            eksik = []
            for kategori, dosya_adi in DESTEK_DOSYALARI:
                if not self.destek_dosyasi_bul(kategori, dosya_adi):
                    eksik.append(os.path.join(kategori, dosya_adi))
            if eksik:
                logger.warning("Eksik örnek destek dosyaları: %s", ", ".join(eksik))
        except Exception as e:
            self.hata_kaydet("Örnek destek dosyaları hazırlanamadı", e)
        return hedef_klasor

    def on_closing(self):
        try:
            if self.degisiklik_gecisine_izin_ver("Çıkış"):
                self.root.destroy()
        except Exception as e:
            self.hata_kaydet("Program kapatılırken hata oluştu", e)
            messagebox.showerror(
                "Kapatma Hatası",
                "Proje durumu güvenli biçimde denetlenemedi. Veri kaybını önlemek için program açık tutuldu.\n\n"
                f"{e}",
            )

    def formasyon_degisti(self, event=None):
        return self.temel_bilgiler_islemleri().formasyon_degisti(event)

    def formasyon_bilgilerini_hazirla(self):
        return self.temel_bilgiler_islemleri().formasyon_bilgilerini_hazirla()

    def arayuz_yardimcilari(self):
        return ArayuzYardimcilari(self)

    def arayuz_stillerini_hazirla(self):
        return self.arayuz_yardimcilari().genel_stilleri_hazirla()

    def durum_cubugu_olustur(self):
        return self.arayuz_yardimcilari().durum_cubugu_olustur()

    def durum_mesaji_yaz(self, mesaj, dosya=None):
        return self.arayuz_yardimcilari().durum_mesaji_yaz(mesaj, dosya)

    def animasyonlu_pencere(self, parent=None, **kwargs):
        return self.arayuz_yardimcilari().animasyonlu_pencere(parent, **kwargs)

    def ikonlari_olustur(self):
        return self.arayuz_yardimcilari().ikonlari_olustur()

    def menu_olustur(self):
        menu_cubugu = tk.Menu(self.root)
        self.menu_cubugu = menu_cubugu
        self.root.config(menu=menu_cubugu)
        dosya_menusu = tk.Menu(menu_cubugu, tearoff=0)
        self.dosya_menusu = dosya_menusu
        menu_cubugu.add_cascade(label="Dosya", menu=dosya_menusu)
        dosya_menusu.add_command(label="Yeni", command=self.yeni_dosya)
        dosya_menusu.add_command(label="Aç", command=self.dosya_ac)
        self.son_projeler_menusu = tk.Menu(dosya_menusu, tearoff=0)
        dosya_menusu.add_cascade(label="Son Açılan Projeler", menu=self.son_projeler_menusu)
        dosya_menusu.add_command(label="Kaydet", command=self.kaydet)
        dosya_menusu.add_command(label="Farklı Kaydet", command=self.farkli_kaydet)
        dosya_menusu.add_separator()
        dosya_menusu.add_command(label="Eski Kayıt Karakterlerini Onar", command=self.eski_json_karakterlerini_onar)
        dosya_menusu.add_separator()
        dosya_menusu.add_command(label="Çıkış", command=self.on_closing)

        araclar_menusu = tk.Menu(menu_cubugu, tearoff=0)
        self.araclar_menusu = araclar_menusu
        menu_cubugu.add_cascade(label="Araçlar", menu=araclar_menusu)
        araclar_menusu.add_command(
            label="Çanakkale Jeoloji Kütüphanesi",
            command=self.jeoloji_kutuphanesi_penceresi,
        )
        araclar_menusu.add_command(
            label="1/100.000 Jeoloji Paftaları",
            command=self.jeoloji_pafta_kutuphanesi_penceresi,
        )
        araclar_menusu.add_command(label="İş Takibi", command=self.is_takibi_penceresi)
        araclar_menusu.add_separator()
        self.animasyonlar_aktif_var = tk.BooleanVar(
            master=self.root,
            value=bool(getattr(self, "animasyonlar_aktif", True)),
        )
        araclar_menusu.add_checkbutton(
            label="Pencere Animasyonları",
            variable=self.animasyonlar_aktif_var,
        )
        self.son_projeler_menusunu_guncelle()

    def sekmeleri_olustur(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(side="top", expand=True, fill='both', padx=10, pady=(8, 0))
        self.tablo_stillerini_hazirla()
        
        # --- SEKMELERİ İSTEDİĞİNİZ YENİ SIRAYA GÖRE ÇAĞIRIYORUZ ---
        self.proje_ozet_sekmesi_olustur()
        self.sekme7_harita()
        self.sekme1_proje()
        self.sekme2_arazi()
        self.sekme3_bina()
        self.sekme5_jeofizik()
        self.sekme9_lab()
        self.sekme4_cukur()
        self.sekme6_tasima()
        self.sekme10_ekler()
        self.sekme8_rapor()

        # --- SEKME İSİMLERİNİ VE NUMARALARINI OTOMATİK GÜNCELLİYORUZ ---
        yeni_isimler = [
            "0. Özet",
            "1. Haritalar",
            "2. Proje Bilgileri",
            "3. Arazi Bilgileri",
            "4. Bina Bilgileri",
            "5. Jeofizik",
            "6. Laboratuvar",
            "7. Araştırma Çukuru / Yüzey Numunesi",
            "8. Taşıma Gücü",
            "9. Ekler",
            "10. Raporlama"
        ]
        
        for i, isim in enumerate(yeni_isimler):
            self.notebook.tab(i, text=isim)

    def tablo_stillerini_hazirla(self):
        return self.arayuz_yardimcilari().tablo_stillerini_hazirla()

    def hucre_duzenle(self, event, tree, set_row=None, set_col=None):
        return self.arayuz_yardimcilari().hucre_duzenle(event, tree, set_row, set_col)

    def ac_yn_islemleri(self):
        return AcYnIslemleri(self)

    def temel_bilgiler_islemleri(self):
        return TemelBilgilerIslemleri(self)

    def jeoloji_kutuphanesi_islemleri(self):
        return JeolojiKutuphanesiIslemleri(self)

    def jeoloji_kutuphanesi_penceresi(self):
        return self.jeoloji_kutuphanesi_islemleri().jeoloji_kutuphanesi_penceresi()

    def jeoloji_harita_katmanini_degistir(self):
        return self.jeoloji_kutuphanesi_islemleri().jeoloji_harita_katmanini_degistir()

    def jeoloji_harita_katmanini_yenile(self, zorla=True):
        return self.jeoloji_kutuphanesi_islemleri().jeoloji_harita_katmanini_yenile(zorla=zorla)

    def jeoloji_pafta_islemleri(self):
        return JeolojiPaftaIslemleri(self)

    def jeoloji_pafta_kutuphanesi_penceresi(self):
        return self.jeoloji_pafta_islemleri().jeoloji_pafta_kutuphanesi_penceresi()

    def formasyonu_jeoloji_haritasindan_bul(self):
        return self.jeoloji_pafta_islemleri().formasyonu_jeoloji_haritasindan_bul()

    def jeoloji_pafta_durumunu_guncelle(self):
        return self.jeoloji_pafta_islemleri().jeoloji_pafta_durumunu_guncelle()

    def genel_jeoloji_islemleri(self):
        return GenelJeolojiIslemleri(self)

    def genel_jeoloji_haritasi_hazirla(self):
        return self.genel_jeoloji_islemleri().genel_jeoloji_haritasi_hazirla()

    def genel_jeoloji_durumunu_guncelle(self):
        return self.genel_jeoloji_islemleri().genel_jeoloji_durumunu_guncelle()

    def sekme1_proje(self):
        return self.temel_bilgiler_islemleri().sekme1_proje()

    def sekme2_arazi(self):
        return self.temel_bilgiler_islemleri().sekme2_arazi()

    def sekme3_bina(self):
        return self.temel_bilgiler_islemleri().sekme3_bina()

    def varsayilan_alan_degerlerini_yerlestir(self, yalniz_bos=True):
        return self.temel_bilgiler_islemleri().varsayilan_alan_degerlerini_yerlestir(yalniz_bos)

    def sekme4_cukur(self):
        return self.ac_yn_islemleri().sekme4_cukur()

    def cukur_sekmesi_ekle(self, isim, enlem="", boylam="", tarih=""):
        return self.ac_yn_islemleri().cukur_sekmesi_ekle(isim, enlem, boylam, tarih)

    def ac_yn_sekme_bilgisi(self, sekme):
        return self.ac_yn_islemleri().ac_yn_sekme_bilgisi(sekme)

    def ac_yn_sekme_kayitlari(self):
        return self.ac_yn_islemleri().ac_yn_sekme_kayitlari()

    def ac_yn_satirlari(self, kayit):
        return self.ac_yn_islemleri().ac_yn_satirlari(kayit)

    def ac_yn_kaydi_verisini_oku(self, kayit):
        return self.ac_yn_islemleri().ac_yn_kaydi_verisini_oku(kayit)

    def ac_yn_sekmelerini_temizle(self):
        return self.ac_yn_islemleri().ac_yn_sekmelerini_temizle()

    def jeofizik_islemleri(self):
        return JeofizikIslemleri(self)

    def sekme5_jeofizik(self):
        return self.jeofizik_islemleri().sekme5_jeofizik()

    def excel_yukle(self):
        return self.jeofizik_islemleri().excel_yukle()

    def jeofizik_excel_yolu_al(self):
        return self.jeofizik_islemleri().jeofizik_excel_yolu_al()

    def jeofizik_excel_yolu_ayarla(self, yol):
        return self.jeofizik_islemleri().jeofizik_excel_yolu_ayarla(yol)

    def jeofizik_koordinatlari_al(self):
        return self.jeofizik_islemleri().jeofizik_koordinatlari_al()

    def jeofizik_koordinatlarini_temizle(self):
        return self.jeofizik_islemleri().jeofizik_koordinatlarini_temizle()

    def jeofizik_koordinat_ekle(self, calisma_no, enlem, boylam):
        return self.jeofizik_islemleri().jeofizik_koordinat_ekle(calisma_no, enlem, boylam)

    def jeofizik_koordinatlari_yerlestir(self, satirlar):
        return self.jeofizik_islemleri().jeofizik_koordinatlari_yerlestir(satirlar)

    def jeofon_dizilim_bilgileri_al(self):
        return self.jeofizik_islemleri().jeofon_dizilim_bilgileri_al()

    def jeofon_dizilim_bilgileri_yerlestir(self, bilgiler):
        return self.jeofizik_islemleri().jeofon_dizilim_bilgileri_yerlestir(bilgiler)

    def tasima_islemleri(self):
        return TasimaIslemleri(self)

    def sekme6_tasima(self):
        return self.tasima_islemleri().sekme6_tasima()

    def tasima_ekran_guncelle(self):
        return self.tasima_islemleri().tasima_ekran_guncelle()

    def tasima_dayanim_23_arayuz_guncelle(self):
        return self.tasima_islemleri().tasima_dayanim_23_arayuz_guncelle()

    def tasima_dayanim_23_degistir(self):
        return self.tasima_islemleri().tasima_dayanim_23_degistir()

    def sayisal_tasima_girdisi_oku(self, kod, etiket):
        return self.tasima_islemleri().sayisal_tasima_girdisi_oku(kod, etiket)

    def tasima_giris_hatasi_goster(self, hatalar):
        return self.tasima_islemleri().tasima_giris_hatasi_goster(hatalar)

    def kaya_tasima_girdilerini_oku(self):
        return self.tasima_islemleri().kaya_tasima_girdilerini_oku()

    def zemin_tasima_girdilerini_oku(self):
        return self.tasima_islemleri().zemin_tasima_girdilerini_oku()

    def nihai_qt_oku(self):
        return self.tasima_islemleri().nihai_qt_oku()

    def hesaplanmis_qk_oku(self):
        return self.tasima_islemleri().hesaplanmis_qk_oku()

    def tasima_hesapla(self):
        return self.tasima_islemleri().tasima_hesapla()

    def tasima_metni_olustur(self):
        return self.tasima_islemleri().tasima_metni_olustur()

    def tasima_raporu_guncel_mi(self):
        return self.tasima_islemleri().tasima_raporu_guncel_mi()

    def sekme7_harita(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="7. Haritalar")
        ust_frame = ttk.Frame(frame)
        ust_frame.pack(fill='x', pady=5)
        ttk.Button(ust_frame, text="KML Yükle", command=self.kml_haritaya_yukle).pack(side="left", padx=5)
        self.btn_tkgm_kml = ttk.Button(
            ust_frame,
            text="TKGM'den Al",
            command=self.tkgm_kml_al,
            bootstyle="success-outline",
        )
        self.btn_tkgm_kml.pack(side="left", padx=2)
        ttk.Label(ust_frame, text="İşaretleme Aracı:").pack(side="left", padx=(15, 5))
        self.aktif_harita_araci = tk.StringVar(value="Yok")
        ttk.Radiobutton(ust_frame, text="İzleme", variable=self.aktif_harita_araci, value="Yok", command=self.harita_araci_degisti).pack(side="left", padx=2)
        ttk.Radiobutton(ust_frame, text="M (Merkez)", variable=self.aktif_harita_araci, value="M", command=self.harita_araci_degisti).pack(side="left", padx=2)
        ttk.Radiobutton(ust_frame, text="AÇ", variable=self.aktif_harita_araci, value="AÇ", command=self.harita_araci_degisti).pack(side="left", padx=2)
        ttk.Radiobutton(ust_frame, text="YN", variable=self.aktif_harita_araci, value="YN", command=self.harita_araci_degisti).pack(side="left", padx=2)
        ttk.Radiobutton(ust_frame, text="SS", variable=self.aktif_harita_araci, value="SS", command=self.harita_araci_degisti).pack(side="left", padx=2)
        
        btn_haritalar_f = ttk.Frame(ust_frame)
        btn_haritalar_f.pack(side="left", padx=15)
        ttk.Button(btn_haritalar_f, text="Yakın Haritaları Hazırla", command=self.haritalari_hazirla, style="Secondary.TButton").pack(side="left", padx=2)
        ttk.Button(btn_haritalar_f, text="Yerbulduru Hazırla", command=self.yerbulduru_hazirla, style="Secondary.TButton").pack(side="left", padx=2)
        ttk.Button(btn_haritalar_f, text="Parsel Haritası Hazırla", command=self.parsel_haritasi_hazirla, style="Secondary.TButton").pack(side="left", padx=2)
        
        ttk.Button(ust_frame, text="İşaretleri Tablolara Aktar", command=self.harita_verilerini_senkronize_et, style="Primary.TButton").pack(side="right", padx=5)

        kutuphane_frame = ttk.Frame(frame)
        kutuphane_frame.pack(fill="x", padx=5, pady=(0, 5))
        self.jeoloji_kutuphane_harita_var = tk.BooleanVar(value=True)
        self.jeoloji_kutuphane_taslak_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            kutuphane_frame,
            text="Jeoloji kütüphanesi parsellerini göster",
            variable=self.jeoloji_kutuphane_harita_var,
            command=self.jeoloji_harita_katmanini_degistir,
        ).pack(side="left")
        ttk.Checkbutton(
            kutuphane_frame,
            text="Taslakları da göster",
            variable=self.jeoloji_kutuphane_taslak_var,
            command=self.jeoloji_harita_katmanini_degistir,
        ).pack(side="left", padx=(12, 0))
        ttk.Button(
            kutuphane_frame,
            text="Katmanı Yenile",
            command=self.jeoloji_harita_katmanini_yenile,
        ).pack(side="left", padx=(12, 0))
        self.jeoloji_harita_durum = tk.StringVar(value=JEOLOJI_HARITA_RENK_ACIKLAMASI)
        ttk.Label(kutuphane_frame, textvariable=self.jeoloji_harita_durum).pack(side="right")

        pafta_frame = ttk.Frame(frame)
        pafta_frame.pack(fill="x", padx=5, pady=(0, 5))
        ttk.Button(
            pafta_frame,
            text="1/100.000 Pafta Kütüphanesi",
            command=self.jeoloji_pafta_kutuphanesi_penceresi,
        ).pack(side="left")
        ttk.Button(
            pafta_frame,
            text="Formasyonu Haritadan Bul",
            command=self.formasyonu_jeoloji_haritasindan_bul,
        ).pack(side="left", padx=(8, 0))
        self.jeoloji_pafta_durum = tk.StringVar(
            value="1/100.000 formasyon tanıması için parsel KML'si yükleyin"
        )
        ttk.Label(pafta_frame, textvariable=self.jeoloji_pafta_durum).pack(side="right")

        genel_jeoloji_frame = ttk.Frame(frame)
        genel_jeoloji_frame.pack(fill="x", padx=5, pady=(0, 5))
        ttk.Button(
            genel_jeoloji_frame,
            text="Genel Jeoloji Haritası ve 2.1 Hazırla",
            command=self.genel_jeoloji_haritasi_hazirla,
            bootstyle="primary-outline",
        ).pack(side="left")
        self.genel_jeoloji_durum = tk.StringVar(
            value="Genel jeoloji için parsel KML'si yükleyin"
        )
        ttk.Label(genel_jeoloji_frame, textvariable=self.genel_jeoloji_durum).pack(side="right")

        self.map_widget = tkintermapview.TkinterMapView(frame, width=800, height=400, corner_radius=0)
        self.map_widget.set_tile_server("https://mt0.google.com/vt/lyrs=s&hl=en&x={x}&y={y}&z={z}&s=Ga", max_zoom=22)
        self.map_widget.set_position(39.524, 26.120) 
        self.map_widget.set_zoom(15)
        self.map_widget.pack(fill="both", expand=True)
        self.map_widget.add_left_click_map_command(self.harita_sol_tik)
        self.map_widget.add_right_click_menu_command(label="En Yakın Noktayı Sil", command=self.harita_sag_tik, pass_coords=True)
        self.root.after(900, lambda: self.jeoloji_harita_katmanini_yenile(zorla=True))

    def harita_islemleri(self):
        return HaritaIslemleri(self)

    def harita_araci_degisti(self):
        return self.harita_islemleri().harita_araci_degisti()

    def sil_ve_yeniden_ciz(self, ac_yn_goster=True, ss_goster=True, m_goster=True):
        return self.harita_islemleri().sil_ve_yeniden_ciz(ac_yn_goster, ss_goster, m_goster)

    def harita_sol_tik(self, coords):
        return self.harita_islemleri().harita_sol_tik(coords)

    def harita_sag_tik(self, coords):
        return self.harita_islemleri().harita_sag_tik(coords)

    def harita_verilerini_senkronize_et(self):
        return self.harita_islemleri().harita_verilerini_senkronize_et()

    def stripe_tree(self, tree):
        return self.arayuz_yardimcilari().stripe_tree(tree)

    def tree_secili_satirlari_tasi(self, tree, yon):
        return self.arayuz_yardimcilari().tree_secili_satirlari_tasi(tree, yon)

    def cizim_uretici(self):
        # Harita dışa aktarım bağımlılıkları yalnızca ilk gerçek kullanımda yüklenir.
        from cizimler import CizimUretici

        return CizimUretici(self)

    def harita_kirpma_miktarlari(self, genislik, yukseklik):
        return self.cizim_uretici().harita_kirpma_miktarlari(genislik, yukseklik)

    def harita_yakalama_widgeti(self):
        return self.cizim_uretici().harita_yakalama_widgeti()

    def harita_goruntusu_yakala(self):
        return self.cizim_uretici().harita_goruntusu_yakala()

    def resim_cek_ve_isaj_ekle(self, path, ac_yn_goster, ss_goster, baslik="MÜHENDİSLİK JEOLOJİSİ HARİTASI", a4_format=True):
        return self.cizim_uretici().resim_cek_ve_isaj_ekle(path, ac_yn_goster, ss_goster, baslik, a4_format)

    def haritalari_hazirla(self):
        return self.cizim_uretici().haritalari_hazirla()

    def yerbulduru_hazirla(self):
        return self.cizim_uretici().yerbulduru_hazirla()

    def parsel_haritasi_hazirla(self):
        return self.cizim_uretici().parsel_haritasi_hazirla()

    def tum_loglari_ciz(self):
        return self.cizim_uretici().tum_loglari_ciz()

    def tekil_log_ciz(self, kayit, isim, kayit_yolu):
        return self.cizim_uretici().tekil_log_ciz(kayit, isim, kayit_yolu)

    def kesit_ciz_olustur(self, ac_sekmeleri, kayit_yolu):
        return self.cizim_uretici().kesit_ciz_olustur(ac_sekmeleri, kayit_yolu)
    def kml_haritaya_yukle(self):
        return self.harita_islemleri().kml_haritaya_yukle()
    def tkgm_kml_al(self):
        return self.harita_islemleri().tkgm_kml_al()
    def ciz_tarama_deseni(self, draw, zemin_tipi, x1, y1, x2, y2):
        return self.cizim_uretici().ciz_tarama_deseni(draw, zemin_tipi, x1, y1, x2, y2)
    def rapor_uretici(self):
        # pandas/python-docx ağırlıklı rapor motoru açılış yolunu yavaşlatmasın.
        from rapor import RaporUretici

        return RaporUretici(self)

    def docx_paragraflarini_dolas(self, doc):
        return self.rapor_uretici().docx_paragraflarini_dolas(doc)

    def docx_xml_metin_dugumlerini_dolas(self, doc):
        return self.rapor_uretici().docx_xml_metin_dugumlerini_dolas(doc)

    def desteklenen_sablon_etiketleri(self):
        return self.rapor_uretici().desteklenen_sablon_etiketleri()

    def sablon_etiketlerini_oku(self, dosya_yolu):
        return self.rapor_uretici().sablon_etiketlerini_oku(dosya_yolu)

    def raporlama_islemleri(self):
        return RaporlamaIslemleri(self)

    def rapor_on_kontrol_uretici(self):
        from rapor_on_kontrol import RaporOnKontrol

        return RaporOnKontrol(self)

    def rapor_on_kontrol(self, devam_sor=True):
        return self.rapor_on_kontrol_uretici().calistir(devam_sor)

    def rapor_on_kontrol_sonuclari(self):
        return self.rapor_on_kontrol_uretici().sonuclari_hazirla()

    def sablon_kontrol_et(self):
        return self.raporlama_islemleri().sablon_kontrol_et()

    def rapor_etiket_verilerini_hazirla(self):
        return self.rapor_uretici().rapor_etiket_verilerini_hazirla()

    def ekler_islemleri(self):
        return EklerIslemleri(self)

    def sekme10_ekler(self):
        return self.ekler_islemleri().sekme10_ekler()

    def taahhutname_paneli_ekle(self, parent):
        return self.taahhutname_islemleri().taahhutname_paneli_ekle(parent)


    def sablon_dosyasi_bul(self, kategori, dosya_adlari=None, ad_parcasi=None):
        arama_klasorleri = []
        for kok in self.sablon_kok_adaylari():
            arama_klasorleri.extend([os.path.join(kok, kategori), kok])
        arama_klasorleri.append(self.uygulama_klasoru_bul())
        if kategori in ("rapor", "taahhutname"):
            arama_klasorleri.append(os.path.join(os.path.expanduser("~"), "Desktop", "K1 Dosya"))

        for klasor in arama_klasorleri:
            if not os.path.isdir(klasor):
                continue
            if dosya_adlari:
                for dosya_adi in dosya_adlari:
                    aday = os.path.join(klasor, dosya_adi)
                    if os.path.exists(aday):
                        return aday
            if ad_parcasi:
                try:
                    for dosya_adi in os.listdir(klasor):
                        kucuk = dosya_adi.lower()
                        if kucuk.endswith(".docx") and ad_parcasi.lower() in kucuk:
                            return os.path.join(klasor, dosya_adi)
                except OSError:
                    continue
        return ""

    def varsayilan_rapor_sablonu(self):
        return self.sablon_dosyasi_bul("rapor", ["TASLAK.docx"])

    def taahhutname_islemleri(self):
        return TaahhutnameIslemleri(self)

    def varsayilan_taahhut_word_sablonu(self):
        return self.taahhutname_islemleri().varsayilan_taahhut_word_sablonu()

    def taahhut_word_sablonu_sec(self):
        return self.taahhutname_islemleri().taahhut_word_sablonu_sec()

    def taahhut_bilgilerini_duzenle(self):
        return self.taahhutname_islemleri().taahhut_bilgilerini_duzenle()

    def proje_deger(self, kod, varsayilan=""):
        return self.temel_bilgiler_islemleri().proje_deger(kod, varsayilan)

    def ek_dosyayi_listeye_ekle(self, kategori, baslik, yol):
        return self.ekler_islemleri().ek_dosyayi_listeye_ekle(kategori, baslik, yol)

    def bina_deger(self, etiket, varsayilan=""):
        return self.temel_bilgiler_islemleri().bina_deger(etiket, varsayilan)

    def taahhutnameleri_olustur(self):
        return self.taahhutname_islemleri().taahhutnameleri_olustur()

    def ek_dosya_turu(self, yol):
        return self.ekler_islemleri().ek_dosya_turu(yol)

    def ek_kategori_durumunu_hazirla(self, kategori):
        return self.ekler_islemleri().ek_kategori_durumunu_hazirla(kategori)

    def ek_dosya_ekle(self, kategori):
        return self.ekler_islemleri().ek_dosya_ekle(kategori)

    def ek_secili_index(self, kategori):
        return self.ekler_islemleri().ek_secili_index(kategori)

    def ek_secili_sil(self, kategori):
        return self.ekler_islemleri().ek_secili_sil(kategori)

    def ek_secili_tasi(self, kategori, yon):
        return self.ekler_islemleri().ek_secili_tasi(kategori, yon)

    def ek_baslik_duzenle(self, kategori):
        return self.ekler_islemleri().ek_baslik_duzenle(kategori)

    def ek_listeleri_guncelle(self, kategori=None, secili_index=None):
        return self.ekler_islemleri().ek_listeleri_guncelle(kategori, secili_index)

    def ek_denetim_mesaji(self, denetim):
        return self.ekler_islemleri().ek_denetim_mesaji(denetim)

    def ek_kontrol_ozeti_goster(self):
        return self.ekler_islemleri().ek_kontrol_ozeti_goster()

    def ekler_pdf_olustur(self):
        return self.ekler_islemleri().ekler_pdf_olustur()

    def ekler_verisini_topla(self):
        return self.ekler_islemleri().ekler_verisini_topla()

    def ekler_verisini_yerlestir(self, veriler):
        return self.ekler_islemleri().ekler_verisini_yerlestir(veriler)

    def sekme8_rapor(self):
        return self.raporlama_islemleri().sekme8_rapor()


    def laboratuvar_islemleri(self):
        return LaboratuvarIslemleri(self)

    def sekme9_lab(self):
        return self.laboratuvar_islemleri().sekme9_lab()

    def lab_excel_yukle(self):
        return self.laboratuvar_islemleri().lab_excel_yukle()

    def lab_ac_satirlari_al(self):
        return self.laboratuvar_islemleri().lab_ac_satirlari_al()

    def lab_yn_satirlari_al(self):
        return self.laboratuvar_islemleri().lab_yn_satirlari_al()

    def lab_ac_satirlari_yerlestir(self, satirlar):
        return self.laboratuvar_islemleri().lab_ac_satirlari_yerlestir(satirlar)

    def lab_yn_satirlari_yerlestir(self, satirlar):
        return self.laboratuvar_islemleri().lab_yn_satirlari_yerlestir(satirlar)

    def lab_ac_numaralari_al(self):
        return self.laboratuvar_islemleri().lab_ac_numaralari_al()

    def lab_yn_numaralari_al(self):
        return self.laboratuvar_islemleri().lab_yn_numaralari_al()

    def lab_ac_bos_satir_ekle(self, no):
        return self.laboratuvar_islemleri().lab_ac_bos_satir_ekle(no)

    def lab_yn_bos_satir_ekle(self, no):
        return self.laboratuvar_islemleri().lab_yn_bos_satir_ekle(no)

    def sablon_sec(self):
        return self.raporlama_islemleri().sablon_sec()

    def jeoloji_sablon_etiket_metni(self):
        return self.raporlama_islemleri().jeoloji_sablon_etiket_metni()

    def jeoloji_sablonu_sec(self):
        return self.raporlama_islemleri().jeoloji_sablonu_sec()

    def jeoloji_sablonu_otomatik_kullan(self):
        return self.raporlama_islemleri().jeoloji_sablonu_otomatik_kullan()

    def rapor_paragraf_bul(self, doc, tag):
        return self.rapor_uretici().rapor_paragraf_bul(doc, tag)

    def rapor_tabloyu_ortala(self, tbl):
        return self.rapor_uretici().rapor_tabloyu_ortala(tbl)

    def rapor_tablo_stili_uygula(self, tablo, header_rows=1, label_columns=0, font_size=None, banded_rows=True):
        return self.rapor_uretici().rapor_tablo_stili_uygula(
            tablo,
            header_rows=header_rows,
            label_columns=label_columns,
            font_size=font_size,
            banded_rows=banded_rows,
        )

    def rapor_metin_degistir(self, doc, tag, value):
        return self.rapor_uretici().rapor_metin_degistir(doc, tag, value)

    def rapor_xml_metin_degistir(self, doc, tag, value):
        return self.rapor_uretici().rapor_xml_metin_degistir(doc, tag, value)

    def rapor_deger_formatla(self, v):
        return self.rapor_uretici().rapor_deger_formatla(v)

    def rapor_formatli_metin_ekle(self, p_obj, metin_satiri):
        return self.rapor_uretici().rapor_formatli_metin_ekle(p_obj, metin_satiri)

    def rapor_run_stilini_kopyala(self, kaynak, hedef):
        return self.rapor_uretici().rapor_run_stilini_kopyala(kaynak, hedef)

    def rapor_cok_satirli_etiket_degistir(self, doc, tag, value):
        return self.rapor_uretici().rapor_cok_satirli_etiket_degistir(doc, tag, value)

    def rapor_xml_sonrasina_ekle(self, anchor, yeni_eleman):
        return self.rapor_uretici().rapor_xml_sonrasina_ekle(anchor, yeni_eleman)

    def rapor_statik_etiketleri_degistir(self, doc):
        return self.rapor_uretici().rapor_statik_etiketleri_degistir(doc)

    def rapor_bina_tablosu_ekle(self, doc):
        return self.rapor_uretici().rapor_bina_tablosu_ekle(doc)

    def rapor_resimleri_ekle(self, doc):
        return self.rapor_uretici().rapor_resimleri_ekle(doc)

    def rapor_jeofizik_koordinat_tablosu_ekle(self, doc):
        return self.rapor_uretici().rapor_jeofizik_koordinat_tablosu_ekle(doc)

    def rapor_jeoloji_koordinat_tablosu_ekle(self, doc):
        return self.rapor_uretici().rapor_jeoloji_koordinat_tablosu_ekle(doc)

    def rapor_koordinat_tablolarini_ekle(self, doc):
        return self.rapor_uretici().rapor_koordinat_tablolarini_ekle(doc)

    def rapor_jeofizik_excel_yolu(self):
        return self.rapor_uretici().rapor_jeofizik_excel_yolu()

    def rapor_jeofizik_parametrelerini_oku(self):
        return self.rapor_uretici().rapor_jeofizik_parametrelerini_oku()

    def rapor_tablo_basligini_kalin_yap(self, tablo):
        return self.rapor_uretici().rapor_tablo_basligini_kalin_yap(tablo)

    def rapor_jeofizik_parametre_tablosu_ekle(self, doc, param_ss_list):
        return self.rapor_uretici().rapor_jeofizik_parametre_tablosu_ekle(doc, param_ss_list)

    def rapor_masw_liste_hazirla(self, param_ss_list):
        return self.rapor_uretici().rapor_masw_liste_hazirla(param_ss_list)

    def rapor_masw_tablosu_ekle(self, doc, param_ss_list):
        return self.rapor_uretici().rapor_masw_tablosu_ekle(doc, param_ss_list)

    def rapor_vp_liste_hazirla(self, param_ss_list):
        return self.rapor_uretici().rapor_vp_liste_hazirla(param_ss_list)

    def rapor_vp_tablosu_ekle(self, doc, param_ss_list):
        return self.rapor_uretici().rapor_vp_tablosu_ekle(doc, param_ss_list)

    def rapor_jeofizik_tablolarini_ekle(self, doc):
        return self.rapor_uretici().rapor_jeofizik_tablolarini_ekle(doc)

    def rapor_tasima_float_al(self, anahtar, varsayilan):
        return self.rapor_uretici().rapor_tasima_float_al(anahtar, varsayilan)

    def rapor_tasima_gucu_tablosu_olustur(self, doc):
        return self.rapor_uretici().rapor_tasima_gucu_tablosu_olustur(doc)

    def rapor_tasima_gucu_ekle(self, doc):
        return self.rapor_uretici().rapor_tasima_gucu_ekle(doc)

    def rapor_kesit_ekle(self, doc):
        return self.rapor_uretici().rapor_kesit_ekle(doc)

    def rapor_ac_loglarini_ekle(self, doc):
        return self.rapor_uretici().rapor_ac_loglarini_ekle(doc)

    def rapor_lab_tablosu_olustur(self, doc, tree, kolon_map):
        return self.rapor_uretici().rapor_lab_tablosu_olustur(doc, tree, kolon_map)

    def rapor_lab_tablolarini_ekle(self, doc):
        return self.rapor_uretici().rapor_lab_tablolarini_ekle(doc)

    def rapor_icerigini_doldur(self, doc):
        return self.rapor_uretici().rapor_icerigini_doldur(doc)

    def rapor_olustur(self):
        return self.raporlama_islemleri().rapor_olustur()

    def nihai_rapor_pdf_olustur(self):
        return self.raporlama_islemleri().nihai_rapor_pdf_olustur()

    def on_deger_islemleri(self):
        return OnDegerIslemleri(self)

    def proje_durumu_islemleri(self):
        return ProjeDurumuIslemleri(self)

    def proje_durum_seridi_olustur(self):
        return self.proje_durumu_islemleri().proje_durum_seridi_olustur()

    def proje_ozet_sekmesi_olustur(self):
        return self.proje_durumu_islemleri().proje_ozet_sekmesi_olustur()

    def proje_durum_seridi_guncelle(self, kaydedilmedi=None):
        return self.proje_durumu_islemleri().proje_durum_seridi_guncelle(kaydedilmedi)

    def proje_durumu_yenilemeyi_planla(self, event=None):
        return self.proje_durumu_islemleri().proje_durumu_yenilemeyi_planla(event)

    def is_takibi_islemleri(self):
        return IsTakibiIslemleri(self)

    def is_takibi_penceresi(self):
        return self.is_takibi_islemleri().is_takibi_penceresi()

    def is_takibi_kaydi_guncelle(self, proje_yolu=None, veriler=None):
        return self.is_takibi_islemleri().is_takibi_kaydi_guncelle(proje_yolu, veriler)

    def on_deger_paneli_olustur(self, parent):
        return self.on_deger_islemleri().on_deger_paneli_olustur(parent)

    def on_deger_verisini_topla(self):
        return self.on_deger_islemleri().on_deger_verisini_topla()

    def on_deger_verisini_yerlestir(self, on_deger, tdth, is_akisi):
        return self.on_deger_islemleri().on_deger_verisini_yerlestir(on_deger, tdth, is_akisi)

    def on_deger_ekranini_guncelle(self):
        return self.on_deger_islemleri().on_deger_ekranini_guncelle()

    def is_asamasini_belirle(self):
        return self.on_deger_islemleri().is_asamasini_belirle()

    def bitmis_proje_acilis_secimi(self):
        return self.on_deger_islemleri().bitmis_proje_acilis_secimi()

    def eski_proje_asama_secimi(self):
        return self.on_deger_islemleri().eski_proje_asama_secimi()

    def proje_salt_okunur_ayarla(self, aktif):
        return self.on_deger_islemleri().proje_salt_okunur_ayarla(aktif)

    def degisiklik_izni_kontrol_et(self, eylem="Bu işlem"):
        return self.on_deger_islemleri().degisiklik_izni_kontrol_et(eylem)

    def nihai_zemin_sinifi_degisti(self, event=None):
        return self.on_deger_islemleri().nihai_zemin_sinifi_degisti(event)


    def kayit_yoneticisi(self):
        return KayitYoneticisi(self)

    def son_projeler_dosya_yolu(self):
        return self.kayit_yoneticisi().son_projeler_dosya_yolu()

    def son_projeleri_yukle(self):
        return self.kayit_yoneticisi().son_projeleri_yukle()

    def son_projeleri_kaydet(self):
        return self.kayit_yoneticisi().son_projeleri_kaydet()

    def son_proje_ekle(self, yol):
        return self.kayit_yoneticisi().son_proje_ekle(yol)

    def son_projeleri_temizle(self):
        return self.kayit_yoneticisi().son_projeleri_temizle()

    def son_projeler_menusunu_guncelle(self):
        return self.kayit_yoneticisi().son_projeler_menusunu_guncelle()

    def son_proje_ac(self, yol):
        return self.kayit_yoneticisi().son_proje_ac(yol)

    def yedek_klasoru(self):
        return self.kayit_yoneticisi().yedek_klasoru()

    def otomatik_yedek_olustur(self, kaynak_yolu, veriler=None):
        return self.kayit_yoneticisi().otomatik_yedek_olustur(kaynak_yolu, veriler)

    def atomik_dosya_kopyala(self, kaynak_yolu, hedef_yolu):
        return self.kayit_yoneticisi().atomik_dosya_kopyala(kaynak_yolu, hedef_yolu)

    def proje_dosyasini_ac(self, yol):
        return self.kayit_yoneticisi().proje_dosyasini_ac(yol)

    def verileri_topla(self):
        return self.kayit_yoneticisi().verileri_topla()

    def verileri_yerlestir(self, veriler):
        return self.kayit_yoneticisi().verileri_yerlestir(veriler)

    def kaydedilmemis_degisiklik_var_mi(self):
        return self.kayit_yoneticisi().kaydedilmemis_degisiklik_var_mi()

    def degisiklik_gecisine_izin_ver(self, eylem):
        return self.kayit_yoneticisi().degisiklik_gecisine_izin_ver(eylem)

    def mojibake_puani(self, metin):
        return self.kayit_yoneticisi().mojibake_puani(metin)

    def mojibake_metin_onar(self, metin):
        return self.kayit_yoneticisi().mojibake_metin_onar(metin)

    def json_verisini_karakter_onar(self, veri):
        return self.kayit_yoneticisi().json_verisini_karakter_onar(veri)

    def json_onarim_yedek_yolu(self, yol):
        return self.kayit_yoneticisi().json_onarim_yedek_yolu(yol)

    def eski_json_karakterlerini_onar(self):
        return self.kayit_yoneticisi().eski_json_karakterlerini_onar()

    def yeni_dosya(self):
        return self.kayit_yoneticisi().yeni_dosya()

    def dosya_ac(self):
        return self.kayit_yoneticisi().dosya_ac()

    def kaydet(self):
        return self.kayit_yoneticisi().kaydet()

    def farkli_kaydet(self):
        return self.kayit_yoneticisi().farkli_kaydet()


if __name__ == "__main__":
    root = tb.Window(themename="cosmo")
    app = RaporProApp(root)
    root.mainloop()
