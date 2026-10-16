# -*- coding: utf-8 -*-
r"""
Alış Faturası Kayıt ve Takip  -  Sürüm 2

Sadece Python standart kütüphanesi kullanır (tkinter, sqlite3, csv, ...).
Ek paket kurulumu GEREKMEZ.

Çalıştırma:   python fatura_takip.py      (konsolsuz açmak için dosya adını .pyw yapabilirsiniz)

Ayarlar: Programla aynı klasördeki config.ini dosyası.
  [ayarlar]
  veritabani = fatura.db                     ; ağ için: \\sunucu\paylasim\FaturaTakip\fatura.db
  tarama_klasoru =                           ; PDF "Gözat" penceresinin ilk açılacağı klasör (isteğe bağlı)
  tema = Açık                                ; Açık / Koyu / Mavi (menüden de değiştirilir)
  kurum_anahtar =                            ; kendi kurumunuzun adından kelime(ler), virgülle (PDF'te alıcıyı firma sanmasın)
  kurum_vkn =                                ; kendi kurumunuzun vergi numarası (isteğe bağlı)

Tutarlar veritabanında KURUŞ (tam sayı) olarak tutulur; böylece yuvarlama hatası olmaz.
"""

import configparser
import csv
import getpass
import os
import re
import sqlite3
import base64
import shutil
import subprocess
import sys
import tempfile
import threading
import zlib
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

# İsteğe bağlı: PyMuPDF kuruluysa PDF metni onunla okunur (daha güvenilir). Yoksa program
# yerleşik okuyucuyla aynen çalışır. Ağ klasöründeki "kutuphaneler" dizini de aranır.
_KUTUPHANE_KLASORU = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0] if sys.argv and sys.argv[0] else __file__)),
                                  "kutuphaneler")
if os.path.isdir(_KUTUPHANE_KLASORU) and _KUTUPHANE_KLASORU not in sys.path:
    sys.path.append(_KUTUPHANE_KLASORU)
try:
    import openpyxl  # isteğe bağlı: biçimlendirilmiş .xlsx dışa aktarma; yoksa CSV kullanılır
    import openpyxl.styles
    import openpyxl.worksheet.properties
except Exception:
    openpyxl = None
try:
    import sv_ttk  # modern (Windows 11 tarzı) arayüz teması; yoksa klasik tema kullanılır
except Exception:
    sv_ttk = None
try:
    import pymupdf as fitz  # PyMuPDF (yeni ad)
except Exception:
    try:
        import fitz  # PyMuPDF (eski ad)
    except Exception:
        fitz = None

# ----------------------------------------------------------------------------
# 1) AYARLAR (config.ini)
# ----------------------------------------------------------------------------

UYGULAMA_ADI = "Alış Faturası Kayıt ve Takip"
UYGULAMA_KLASORU = os.path.dirname(os.path.abspath(sys.argv[0] if sys.argv and sys.argv[0] else __file__))
CONFIG_YOLU = os.path.join(UYGULAMA_KLASORU, "config.ini")


def ayarlari_oku(config_yolu=CONFIG_YOLU):
    """config.ini'yi okur; yoksa varsayılan değerlerle oluşturur."""
    cfg = configparser.ConfigParser(interpolation=None)
    varsayilan = {"veritabani": "fatura.db", "tarama_klasoru": "", "tema": "Açık",
                  "kurum_anahtar": "", "kurum_vkn": ""}
    if os.path.exists(config_yolu):
        cfg.read(config_yolu, encoding="utf-8")
    if not cfg.has_section("ayarlar"):
        cfg.add_section("ayarlar")
    degisti = False
    for anahtar, deger in varsayilan.items():
        if not cfg.has_option("ayarlar", anahtar):
            cfg.set("ayarlar", anahtar, deger)
            degisti = True
    if degisti or not os.path.exists(config_yolu):
        try:
            with open(config_yolu, "w", encoding="utf-8") as f:
                cfg.write(f)
        except OSError:
            pass  # yazılamıyorsa varsayılanlarla devam
    veritabani = cfg.get("ayarlar", "veritabani").strip() or "fatura.db"
    if not os.path.isabs(veritabani) and not veritabani.startswith("\\\\"):
        veritabani = os.path.join(os.path.dirname(config_yolu), veritabani)
    return {
        "veritabani": veritabani,
        "tarama_klasoru": cfg.get("ayarlar", "tarama_klasoru").strip(),
        "tema": cfg.get("ayarlar", "tema").strip() or "Açık",
        "kurum_anahtar": cfg.get("ayarlar", "kurum_anahtar").strip(),
        "kurum_vkn": re.sub(r"\D", "", cfg.get("ayarlar", "kurum_vkn")),
    }


def ayar_yaz(config_yolu, anahtar, deger):
    """config.ini içindeki tek bir ayarı günceller (yazılamazsa sessizce geçer)."""
    try:
        cfg = configparser.ConfigParser(interpolation=None)
        cfg.read(config_yolu, encoding="utf-8")
        if not cfg.has_section("ayarlar"):
            cfg.add_section("ayarlar")
        cfg.set("ayarlar", anahtar, deger)
        with open(config_yolu, "w", encoding="utf-8") as f:
            cfg.write(f)
    except (OSError, configparser.Error):
        pass


# ----------------------------------------------------------------------------
# 2) YARDIMCI FONKSİYONLAR (tutar / tarih biçimleri)
# ----------------------------------------------------------------------------

def tutar_ayristir(metin):
    """'1.234,56', '1234,56', '1234.56', '1.234' , '1.234,56 TL' -> kuruş (int). Boşsa None."""
    if metin is None:
        return None
    s = str(metin).upper().replace("TL", "").replace("₺", "").replace(" ", "").strip()
    if s == "":
        return None
    if "," in s and "." in s:
        # Hangisi sondaysa ondalık ayracıdır
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    elif "." in s:
        # Yalnız nokta: 1.234 veya 1.234.567 ise binlik, aksi halde ondalık
        if re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
            s = s.replace(".", "")
    try:
        d = Decimal(s)
    except InvalidOperation:
        raise ValueError("Geçersiz tutar: %s" % metin)
    if d < 0:
        raise ValueError("Tutar negatif olamaz.")
    return int((d * 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def tutar_yaz(kurus, tl=False):
    """kuruş (int) -> '1.234,56' (tl=True ise sonuna ' TL')."""
    if kurus is None:
        return ""
    negatif = kurus < 0
    kurus = abs(int(kurus))
    lira, kr = divmod(kurus, 100)
    s = "{:,}".format(lira).replace(",", ".") + "," + "%02d" % kr
    if negatif:
        s = "-" + s
    return s + " TL" if tl else s


def tarih_ayristir(metin):
    """'GG.AA.YYYY' (veya GG/AA/YYYY, GG-AA-YYYY) -> 'YYYY-MM-DD'."""
    s = (metin or "").strip().replace("/", ".").replace("-", ".")
    try:
        return datetime.strptime(s, "%d.%m.%Y").strftime("%Y-%m-%d")
    except ValueError:
        raise ValueError("Tarih GG.AA.YYYY biçiminde olmalı (örnek: 05.03.2026).")


def tarih_yaz(iso):
    """'YYYY-MM-DD' -> 'GG.AA.YYYY'."""
    if not iso:
        return ""
    try:
        return datetime.strptime(iso[:10], "%Y-%m-%d").strftime("%d.%m.%Y")
    except ValueError:
        return iso


def tr_kucuk(s):
    """Türkçe harfleri de hesaba katan küçük harfe çevirme (arama/mükerrer kontrolü için)."""
    return (s or "").replace("İ", "i").replace("I", "ı").lower().strip()


def yol_norm(yol):
    return os.path.normcase(os.path.normpath(yol.strip())) if yol and yol.strip() else ""


# ----------------------------------------------------------------------------
# 3) VERİTABANI KATMANI
# ----------------------------------------------------------------------------

class VeritabaniHatasi(Exception):
    pass


class MukerrerKayit(VeritabaniHatasi):
    pass


ALANLAR = ["fatura_tarihi", "fatura_no", "firma_adi", "yapilan_is",
           "tutar_haric", "tutar_dahil", "pdf_yolu"]


class FaturaDB:
    def __init__(self, yol):
        self.yol = yol
        klasor = os.path.dirname(yol)
        if klasor and not os.path.isdir(klasor):
            raise VeritabaniHatasi("Veritabanı klasörüne ulaşılamıyor:\n%s\n\n"
                                   "Ağ bağlantısını ve config.ini içindeki yolu kontrol edin." % klasor)
        self._olustur()

    def _baglan(self):
        try:
            con = sqlite3.connect(self.yol, timeout=15)
            con.row_factory = sqlite3.Row
            return con
        except sqlite3.Error as e:
            raise VeritabaniHatasi("Veritabanı açılamadı:\n%s\n\n%s" % (self.yol, e))

    def _calistir(self, fonk):
        """Kilitlenme gibi hataları anlaşılır Türkçe mesaja çevirir."""
        con = self._baglan()
        try:
            sonuc = fonk(con)
            con.commit()
            return sonuc
        except sqlite3.OperationalError as e:
            if "locked" in str(e).lower():
                raise VeritabaniHatasi("Veritabanı şu an başka bir kullanıcı tarafından kullanılıyor.\n"
                                       "Birkaç saniye sonra tekrar deneyin.")
            raise VeritabaniHatasi("Veritabanı hatası: %s" % e)
        except sqlite3.Error as e:
            raise VeritabaniHatasi("Veritabanı hatası: %s" % e)
        finally:
            con.close()

    def _olustur(self):
        def f(con):
            con.execute("""
                CREATE TABLE IF NOT EXISTS faturalar (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    fatura_tarihi TEXT NOT NULL,
                    fatura_no TEXT NOT NULL,
                    firma_adi TEXT NOT NULL,
                    yapilan_is TEXT,
                    tutar_haric INTEGER,
                    tutar_dahil INTEGER,
                    pdf_yolu TEXT,
                    kaydeden TEXT,
                    kayit_tarihi TEXT,
                    son_duzenleyen TEXT,
                    duzenleme_tarihi TEXT,
                    firma_anahtar TEXT,
                    pdf_anahtar TEXT
                )""")
            con.execute("CREATE INDEX IF NOT EXISTS ix_tarih ON faturalar(fatura_tarihi)")
            # Sürüm 1'den gelen veritabanını bozmadan yeni sütun/tablo ekle
            sutunlar = [r[1] for r in con.execute("PRAGMA table_info(faturalar)")]
            if "vkn" not in sutunlar:
                con.execute("ALTER TABLE faturalar ADD COLUMN vkn TEXT")
            con.execute("CREATE TABLE IF NOT EXISTS firma_vkn (vkn TEXT PRIMARY KEY, firma_adi TEXT NOT NULL)")
            con.execute("""
                CREATE TABLE IF NOT EXISTS ekler (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    fatura_id INTEGER NOT NULL,
                    tur TEXT,
                    dosya_yolu TEXT NOT NULL,
                    aciklama TEXT,
                    ekleyen TEXT,
                    ekleme_tarihi TEXT
                )""")
            con.execute("CREATE INDEX IF NOT EXISTS ix_ek_fatura ON ekler(fatura_id)")
        self._calistir(f)

    # -- kontroller ----------------------------------------------------------
    @staticmethod
    def _anahtar_firma(veri):
        return tr_kucuk(veri["firma_adi"]) + "|" + tr_kucuk(veri["fatura_no"])

    def _mukerrer_kontrol(self, con, veri, kendi_id=None):
        cur = con.execute("SELECT id FROM faturalar WHERE firma_anahtar=? AND id IS NOT ?",
                          (self._anahtar_firma(veri), kendi_id))
        if cur.fetchone():
            raise MukerrerKayit("Bu firmaya ait aynı fatura numarası zaten kayıtlı:\n%s  /  %s"
                                % (veri["firma_adi"], veri["fatura_no"]))
        pa = yol_norm(veri.get("pdf_yolu"))
        if pa:
            cur = con.execute("SELECT id, fatura_no, firma_adi FROM faturalar WHERE pdf_anahtar=? AND id IS NOT ?",
                              (pa, kendi_id))
            r = cur.fetchone()
            if r:
                raise MukerrerKayit("Bu PDF dosyası zaten başka bir faturaya bağlı:\n%s  /  %s"
                                    % (r["firma_adi"], r["fatura_no"]))

    # -- işlemler ------------------------------------------------------------
    def ekle(self, veri, kullanici):
        simdi = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        def f(con):
            self._mukerrer_kontrol(con, veri)
            cur = con.execute("""
                INSERT INTO faturalar (fatura_tarihi, fatura_no, firma_adi, yapilan_is, tutar_haric,
                    tutar_dahil, pdf_yolu, kaydeden, kayit_tarihi, firma_anahtar, pdf_anahtar, vkn)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                (veri["fatura_tarihi"], veri["fatura_no"], veri["firma_adi"], veri["yapilan_is"],
                 veri["tutar_haric"], veri["tutar_dahil"], veri["pdf_yolu"], kullanici, simdi,
                 self._anahtar_firma(veri), yol_norm(veri.get("pdf_yolu")), veri.get("vkn") or None))
            self._vkn_isle(con, veri)
            return cur.lastrowid
        return self._calistir(f)

    def guncelle(self, fid, veri, kullanici):
        simdi = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        def f(con):
            self._mukerrer_kontrol(con, veri, kendi_id=fid)
            con.execute("""
                UPDATE faturalar SET fatura_tarihi=?, fatura_no=?, firma_adi=?, yapilan_is=?,
                    tutar_haric=?, tutar_dahil=?, pdf_yolu=?, son_duzenleyen=?, duzenleme_tarihi=?,
                    firma_anahtar=?, pdf_anahtar=?, vkn=COALESCE(NULLIF(?, ''), vkn) WHERE id=?""",
                (veri["fatura_tarihi"], veri["fatura_no"], veri["firma_adi"], veri["yapilan_is"],
                 veri["tutar_haric"], veri["tutar_dahil"], veri["pdf_yolu"], kullanici, simdi,
                 self._anahtar_firma(veri), yol_norm(veri.get("pdf_yolu")), veri.get("vkn") or "", fid))
            self._vkn_isle(con, veri)
        self._calistir(f)

    def sil(self, fid):
        def f(con):
            con.execute("DELETE FROM ekler WHERE fatura_id=?", (fid,))
            con.execute("DELETE FROM faturalar WHERE id=?", (fid,))
        self._calistir(f)

    # -- ek evraklar (talep, yazı, irsaliye ...): dosya kopyalanmaz, yalnızca konumu tutulur --
    def ekler(self, fid):
        def f(con):
            return [dict(r) for r in con.execute("SELECT * FROM ekler WHERE fatura_id=? ORDER BY id", (fid,))]
        return self._calistir(f)

    def ek_ekle(self, fid, tur, yol, aciklama, kullanici):
        simdi = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        def f(con):
            if con.execute("SELECT 1 FROM ekler WHERE fatura_id=? AND dosya_yolu=?", (fid, yol)).fetchone():
                raise MukerrerKayit("Bu dosya bu faturaya zaten eklenmiş:\n%s" % yol)
            con.execute("INSERT INTO ekler (fatura_id, tur, dosya_yolu, aciklama, ekleyen, ekleme_tarihi) "
                        "VALUES (?,?,?,?,?,?)", (fid, tur, yol, aciklama, kullanici, simdi))
        self._calistir(f)

    def ek_sil(self, ek_id):
        self._calistir(lambda con: con.execute("DELETE FROM ekler WHERE id=?", (ek_id,)))

    def getir(self, fid):
        def f(con):
            r = con.execute("SELECT * FROM faturalar WHERE id=?", (fid,)).fetchone()
            return dict(r) if r else None
        return self._calistir(f)

    def firmalar(self):
        def f(con):
            return [r[0] for r in con.execute(
                "SELECT DISTINCT firma_adi FROM faturalar ORDER BY firma_adi COLLATE NOCASE")]
        return sorted(self._calistir(f), key=tr_kucuk)

    @staticmethod
    def _vkn_isle(con, veri):
        """VKN okunmuşsa 'vkn -> firma adı' eşlemesini kaydeder (sonraki faturalarda ad otomatik gelsin)."""
        vkn = (veri.get("vkn") or "").strip()
        if vkn and (veri.get("firma_adi") or "").strip():
            con.execute("INSERT OR REPLACE INTO firma_vkn (vkn, firma_adi) VALUES (?, ?)",
                        (vkn, veri["firma_adi"].strip()))

    def vkn_firma(self, vkn):
        """Daha önce kaydedilmiş VKN için firma adını döndürür (yoksa None)."""
        if not vkn:
            return None

        def f(con):
            r = con.execute("SELECT firma_adi FROM firma_vkn WHERE vkn=?", (vkn,)).fetchone()
            return r[0] if r else None
        return self._calistir(f)

    def hepsi(self):
        def f(con):
            return [dict(r) for r in con.execute(
                "SELECT f.*, (SELECT COUNT(*) FROM ekler e WHERE e.fatura_id = f.id) AS ek_sayisi "
                "FROM faturalar f ORDER BY f.fatura_tarihi DESC, f.id DESC")]
        return self._calistir(f)

    def liste(self, arama="", firma="", bas=None, bit=None, yil=None, ay=None):
        """Filtrelenmiş kayıtları döndürür. bas/bit: 'YYYY-MM-DD' veya None."""
        return filtrele(self.hepsi(), arama, firma, bas, bit, yil, ay)


def toplamlar(satirlar):
    adet = len(satirlar)
    haric = sum(r["tutar_haric"] or 0 for r in satirlar)
    dahil = sum(r["tutar_dahil"] or 0 for r in satirlar)
    return adet, haric, dahil


AY_ADLARI = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
             "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]


def filtrele(satirlar, arama="", firma="", bas=None, bit=None, yil=None, ay=None):
    """Kayıt listesini arama, firma, tarih aralığı, yıl ve aya göre süzer (yil/ay: int veya None)."""
    sonuc = []
    a = tr_kucuk(arama)
    for r in satirlar:
        t = r["fatura_tarihi"] or ""
        if firma and r["firma_adi"] != firma:
            continue
        if bas and t < bas:
            continue
        if bit and t > bit:
            continue
        if yil and t[:4] != "%04d" % yil:
            continue
        if ay and t[5:7] != "%02d" % ay:
            continue
        if a:
            havuz = tr_kucuk(" ".join([r["fatura_no"] or "", r["firma_adi"] or "", r["yapilan_is"] or ""]))
            if a not in havuz:
                continue
        sonuc.append(r)
    return sonuc


def aylik_ozet(satirlar, yil):
    """Seçilen yılın 12 ayı için adet / KDV hariç / KDV dahil toplamları + son satırda TOPLAM."""
    sonuc = []
    for k in range(1, 13):
        adet, haric, dahil = toplamlar(filtrele(satirlar, yil=yil, ay=k))
        sonuc.append({"ay": AY_ADLARI[k - 1], "adet": adet, "haric": haric, "dahil": dahil})
    sonuc.append({"ay": "TOPLAM", "adet": sum(x["adet"] for x in sonuc),
                  "haric": sum(x["haric"] for x in sonuc), "dahil": sum(x["dahil"] for x in sonuc)})
    return sonuc


def _xlsx_stilleri():
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    ince = Side(style="thin", color="BFC7D1")
    return {
        "kenar": Border(left=ince, right=ince, top=ince, bottom=ince),
        "baslik_dolgu": PatternFill("solid", fgColor="1F4E79"),
        "baslik_yazi": Font(name="Calibri", bold=True, color="FFFFFF", size=11),
        "zebra": PatternFill("solid", fgColor="EEF3F9"),
        "toplam_dolgu": PatternFill("solid", fgColor="D9E2EF"),
        "toplam_yazi": Font(name="Calibri", bold=True, size=11),
        "normal": Font(name="Calibri", size=11),
        "baslik_yazisi": Font(name="Calibri", bold=True, size=14, color="1F4E79"),
        "orta": Alignment(horizontal="center", vertical="center"),
        "sol": Alignment(horizontal="left", vertical="center", wrap_text=True),
        "sag": Alignment(horizontal="right", vertical="center"),
        "baslik_orta": Alignment(horizontal="center", vertical="center", wrap_text=True),
    }


def _xlsx_sayfa_adi(metin):
    return re.sub(r"[\\/*?:\[\]]", "-", metin)[:31] or "Faturalar"


def xlsx_yaz(yol, satirlar, baslik="Alış Faturaları"):
    """Faturaları biçimlendirilmiş bir Excel (.xlsx) tablosu olarak yazar (openpyxl gerekir)."""
    if openpyxl is None:
        raise RuntimeError("XLSX için 'openpyxl' kütüphanesi kurulu değil.")
    from openpyxl.utils import get_column_letter
    st = _xlsx_stilleri()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = _xlsx_sayfa_adi(baslik)
    sutunlar = ["Fatura Tarihi", "Fatura No", "Firma Adı", "Yapılan İş", "Tutar (KDV Hariç)", "KDV Tutarı",
                "Tutar (KDV Dahil)", "Ek Evrak", "PDF Konumu"]
    genislik_en_az = [14, 20, 34, 40, 18, 16, 18, 10, 40]
    son = len(sutunlar)

    ws["A1"] = "%s   (%d fatura, düzenlenme: %s)" % (baslik, len(satirlar), date.today().strftime("%d.%m.%Y"))
    ws["A1"].font = st["baslik_yazisi"]
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=son)
    ws.row_dimensions[1].height = 26

    bas = 3
    for k, ad in enumerate(sutunlar, 1):
        h = ws.cell(row=bas, column=k, value=ad)
        h.font, h.fill, h.border, h.alignment = st["baslik_yazi"], st["baslik_dolgu"], st["kenar"], st["baslik_orta"]
    ws.row_dimensions[bas].height = 30

    sirali = sorted(satirlar, key=lambda r: (r["fatura_tarihi"] or "", r["id"]))
    for i, r in enumerate(sirali):
        sat = bas + 1 + i
        try:
            tarih = datetime.strptime((r["fatura_tarihi"] or "")[:10], "%Y-%m-%d").date()
        except ValueError:
            tarih = r["fatura_tarihi"]
        haric = r["tutar_haric"] / 100 if r["tutar_haric"] is not None else None
        dahil = r["tutar_dahil"] / 100 if r["tutar_dahil"] is not None else None
        kdv = (r["tutar_dahil"] - r["tutar_haric"]) / 100 if haric is not None and dahil is not None else None
        degerler = [tarih, r["fatura_no"], r["firma_adi"], r["yapilan_is"] or "", haric, kdv, dahil,
                    r.get("ek_sayisi") or None, r["pdf_yolu"] or ""]
        for k, d in enumerate(degerler, 1):
            c = ws.cell(row=sat, column=k, value=d)
            c.font, c.border = st["normal"], st["kenar"]
            if i % 2:
                c.fill = st["zebra"]
            if k == 1:
                c.number_format, c.alignment = "DD.MM.YYYY", st["orta"]
            elif k in (5, 6, 7):
                c.number_format, c.alignment = "#,##0.00", st["sag"]
            elif k == 8:
                c.alignment = st["orta"]
            else:
                c.alignment = st["sol"]
        pdf = ws.cell(row=sat, column=9)
        if r["pdf_yolu"]:
            try:
                pdf.hyperlink = r["pdf_yolu"]
                pdf.font = openpyxl.styles.Font(name="Calibri", size=11, color="0563C1", underline="single")
            except Exception:
                pass

    ilk, bit = bas + 1, bas + len(sirali)
    tr = bit + 1
    for k in range(1, son + 1):
        c = ws.cell(row=tr, column=k)
        c.font, c.fill, c.border = st["toplam_yazi"], st["toplam_dolgu"], st["kenar"]
    ws.cell(row=tr, column=3, value="TOPLAM (%d fatura)" % len(sirali)).alignment = st["sag"]
    for k in (5, 6, 7):
        harf = get_column_letter(k)
        c = ws.cell(row=tr, column=k, value="=SUBTOTAL(109,%s%d:%s%d)" % (harf, ilk, harf, bit))
        c.number_format, c.alignment = "#,##0.00", st["sag"]
    ws.cell(row=tr, column=8, value="=SUBTOTAL(109,H%d:H%d)" % (ilk, bit)).alignment = st["orta"]

    for k, ad in enumerate(sutunlar, 1):
        en = max([len(str(ws.cell(row=s, column=k).value or "")) for s in range(ilk, bit + 1)] + [len(ad)])
        sinir = 60 if k in (3, 4, 9) else 24
        ws.column_dimensions[get_column_letter(k)].width = min(max(genislik_en_az[k - 1], en + 3), sinir)

    ws.auto_filter.ref = "A%d:%s%d" % (bas, get_column_letter(son), bit)
    ws.freeze_panes = ws.cell(row=bas + 1, column=1)
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    ws.print_title_rows = "%d:%d" % (bas, bas)
    wb.save(yol)


def ozet_xlsx_yaz(yol, ozet, yil):
    """Aylık özeti biçimlendirilmiş Excel (.xlsx) olarak yazar."""
    if openpyxl is None:
        raise RuntimeError("XLSX için 'openpyxl' kütüphanesi kurulu değil.")
    st = _xlsx_stilleri()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Aylık Özet %s" % yil
    ws["A1"] = "Aylık Fatura Özeti - %s" % yil
    ws["A1"].font = st["baslik_yazisi"]
    ws.merge_cells("A1:E1")
    ws.row_dimensions[1].height = 26
    for k, ad in enumerate(["Ay", "Fatura Adedi", "Tutar (KDV Hariç)", "KDV Tutarı", "Tutar (KDV Dahil)"], 1):
        h = ws.cell(row=3, column=k, value=ad)
        h.font, h.fill, h.border, h.alignment = st["baslik_yazi"], st["baslik_dolgu"], st["kenar"], st["orta"]
    ws.row_dimensions[3].height = 24
    for i, x in enumerate(ozet):
        sat = 4 + i
        toplam = x["ay"] == "TOPLAM"
        degerler = [x["ay"], x["adet"], x["haric"] / 100, (x["dahil"] - x["haric"]) / 100, x["dahil"] / 100]
        for k, d in enumerate(degerler, 1):
            c = ws.cell(row=sat, column=k, value=d)
            c.border = st["kenar"]
            c.font = st["toplam_yazi"] if toplam else st["normal"]
            if toplam:
                c.fill = st["toplam_dolgu"]
            elif i % 2:
                c.fill = st["zebra"]
            if k >= 3:
                c.number_format, c.alignment = "#,##0.00", st["sag"]
            elif k == 2:
                c.alignment = st["orta"]
    for harf, gen in zip("ABCDE", (16, 16, 22, 20, 22)):
        ws.column_dimensions[harf].width = gen
    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False
    wb.save(yol)


def ozet_csv_yaz(yol, ozet):
    with open(yol, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Ay", "Fatura Adedi", "KDV Hariç Toplam", "KDV Tutarı", "KDV Dahil Toplam"])
        for x in ozet:
            w.writerow([x["ay"], x["adet"], tutar_yaz(x["haric"]), tutar_yaz(x["dahil"] - x["haric"]),
                        tutar_yaz(x["dahil"])])


def csv_yaz(yol, satirlar):
    """Excel'de Türkçe karakter bozulmadan açılır: utf-8-sig + noktalı virgül."""
    with open(yol, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Fatura Tarihi", "Fatura No", "Firma Adı", "Yapılan İş",
                    "Fatura Tutarı (KDV Hariç)", "Fatura Tutarı (KDV Dahil)", "KDV Tutarı", "PDF Konumu"])
        for r in satirlar:
            kdv = None
            if r["tutar_haric"] is not None and r["tutar_dahil"] is not None:
                kdv = r["tutar_dahil"] - r["tutar_haric"]
            w.writerow([tarih_yaz(r["fatura_tarihi"]), r["fatura_no"], r["firma_adi"], r["yapilan_is"] or "",
                        tutar_yaz(r["tutar_haric"]), tutar_yaz(r["tutar_dahil"]), tutar_yaz(kdv),
                        r["pdf_yolu"] or ""])


# ----------------------------------------------------------------------------
# 4) WINDOWS YARDIMCILARI (PDF aç / klasörde göster)
# ----------------------------------------------------------------------------

def pdf_ac(yol):
    if not yol or not os.path.exists(yol):
        raise FileNotFoundError("Dosya bulunamadı:\n%s\n\nTaşınmış ya da silinmiş olabilir; "
                                "ağ bağlantısını da kontrol edin." % (yol or "(yol girilmemiş)"))
    if hasattr(os, "startfile"):
        os.startfile(yol)  # Windows
    else:
        subprocess.Popen(["xdg-open", yol])


def klasorde_goster(yol):
    if not yol or not os.path.exists(yol):
        raise FileNotFoundError("Dosya bulunamadı:\n%s" % (yol or "(yol girilmemiş)"))
    if os.name == "nt":
        subprocess.Popen(["explorer", "/select,", os.path.normpath(yol)])
    else:
        subprocess.Popen(["xdg-open", os.path.dirname(yol)])


# ----------------------------------------------------------------------------
# 5) PDF'DEN FATURA BİLGİSİ OKUMA  (yalnızca standart kütüphane)
# ----------------------------------------------------------------------------
# Yöntem 1: PDF'in içindeki yazı katmanı (e-fatura PDF'leri ve tarayıcıda "aranabilir PDF"
#           seçeneğiyle taranmış PDF'ler). Hızlıdır ve en güvenilir yoldur.
# Yöntem 2: Windows'un yerleşik OCR'ı (PowerShell üzerinden). Yazı katmanı olmayan,
#           düz resim olarak taranmış PDF'ler için yedek yöntemdir.
# Okunan metinden fatura alanları düzenli ifadelerle (regex) ayıklanır. Sonuç her zaman
# kullanıcı tarafından kontrol edilir; hiçbir şey onaysız kaydedilmez.


class _Ad(bytes):
    """İçerik akışındaki /Ad (name) işlenenlerini string işlenenlerden ayırmak için."""


_ACIK_DIZI = object()
_KIMLIK = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


def _carp(a, b):
    return (a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3],
            a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3],
            a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5])


def _akis_coz(sozluk, ham):
    """Akışın filtrelerini (Flate, ASCII85, ASCIIHex) çözer. Desteklenmeyen filtrede None döner."""
    m = re.search(rb"/Filter\s*(\[[^\]]*\]|/[A-Za-z0-9]+)", sozluk)
    if not m:
        return ham
    veri = ham
    for ad in re.findall(rb"/([A-Za-z0-9]+)", m.group(1)):
        try:
            if ad in (b"FlateDecode", b"Fl"):
                veri = zlib.decompressobj().decompress(veri)
            elif ad in (b"ASCII85Decode", b"A85"):
                v = veri.strip()
                if v.startswith(b"<~"):
                    v = v[2:]
                if v.endswith(b"~>"):
                    v = v[:-2]
                veri = base64.a85decode(v)
            elif ad in (b"ASCIIHexDecode", b"AHx"):
                hx = re.sub(rb"[^0-9A-Fa-f]", b"", veri.split(b">")[0])
                if len(hx) % 2:
                    hx += b"0"
                veri = bytes.fromhex(hx.decode())
            else:
                return None
        except Exception:
            return None
    return veri


def _govde_ayir(govde):
    """'<< sözlük >> stream ... endstream' gövdesini (sözlük, akış) olarak ayırır."""
    m = re.search(rb"(?<!end)stream(?:\r\n|\n|\r)", govde)
    if not m:
        return govde, None
    sozluk = govde[:m.start()]
    ham = govde[m.end():]
    uz = re.search(rb"/Length\s+(\d+)\b(?!\s+\d+\s+R)", sozluk)
    if uz and int(uz.group(1)) <= len(ham):
        ham = ham[:int(uz.group(1))]
    else:
        k = ham.rfind(b"endstream")
        if k >= 0:
            ham = ham[:k]
    return sozluk, ham


def _pdf_nesneleri(veri):
    """PDF'teki nesneleri {no: (sözlük_baytları, akış_baytları|None)} olarak çıkarır (ObjStm dahil)."""
    nesneler = {}
    for m in re.finditer(rb"(?<![0-9])(\d{1,7})\s+\d{1,5}\s+obj\b", veri):
        bit = veri.find(b"endobj", m.end())
        govde = veri[m.end(): bit if bit >= 0 else len(veri)]
        nesneler[int(m.group(1))] = _govde_ayir(govde)
    for no, (sozluk, akis) in list(nesneler.items()):
        if akis is None or not re.search(rb"/Type\s*/ObjStm", sozluk):
            continue
        try:
            d = _akis_coz(sozluk, akis)
            n = int(re.search(rb"/N\s+(\d+)", sozluk).group(1))
            first = int(re.search(rb"/First\s+(\d+)", sozluk).group(1))
            sayilar = re.findall(rb"\d+", d[:first])
            cift = [(int(sayilar[2 * i]), int(sayilar[2 * i + 1])) for i in range(min(n, len(sayilar) // 2))]
            for i, (onum, off) in enumerate(cift):
                son = cift[i + 1][1] if i + 1 < len(cift) else len(d) - first
                nesneler[onum] = (d[first + off: first + son], None)
        except Exception:
            continue
    return nesneler


def _deger_al(s, anahtar):
    """Sözlük baytlarında /Anahtar değerini döndürür (sözlük, dizi, referans veya düz değer)."""
    m = re.search(rb"/" + anahtar + rb"(?![A-Za-z0-9])\s*", s)
    if not m:
        return None
    i = m.end()
    if s[i:i + 2] == b"<<":
        derin, j = 0, i
        while j < len(s):
            if s[j:j + 2] == b"<<":
                derin += 1
                j += 2
            elif s[j:j + 2] == b">>":
                derin -= 1
                j += 2
                if derin == 0:
                    return s[i:j]
            else:
                j += 1
        return s[i:]
    if s[i:i + 1] == b"[":
        k = s.find(b"]", i)
        return s[i:k + 1] if k >= 0 else s[i:]
    r = re.match(rb"\d+\s+\d+\s+R", s[i:])
    if r:
        return r.group(0)
    t = re.match(rb"[^\s/<>\[\]()]+", s[i:])
    return t.group(0) if t else None


def _ref_no(deger):
    m = re.match(rb"\s*(\d+)\s+\d+\s+R", deger or b"")
    return int(m.group(1)) if m else None


def _cozumle(nesneler, deger):
    """Değer bir referanssa o nesnenin sözlüğünü, değilse değerin kendisini verir."""
    no = _ref_no(deger)
    if no is not None:
        return nesneler.get(no, (b"", None))[0]
    return deger or b""


def _hex_utf16(hx):
    hx = hx.decode() if isinstance(hx, bytes) else hx
    if len(hx) % 2:
        hx += "0"
    try:
        return bytes.fromhex(hx).decode("utf-16-be", "ignore")
    except ValueError:
        return ""


def _cmap_oku(cm):
    """ToUnicode CMap'ini okur -> (kod_genişliği_bayt, {kod: metin})."""
    harita = {}
    genislik = 0
    cs = re.search(rb"begincodespacerange(.*?)endcodespacerange", cm, re.S)
    if cs:
        h = re.findall(rb"<([0-9A-Fa-f]+)>", cs.group(1))
        if h:
            genislik = len(h[0]) // 2
    for blok in re.findall(rb"beginbfchar(.*?)endbfchar", cm, re.S):
        for a, b in re.findall(rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]*)>", blok):
            harita[int(a, 16)] = _hex_utf16(b)
    for blok in re.findall(rb"beginbfrange(.*?)endbfrange", cm, re.S):
        for m in re.finditer(rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(\[[^\]]*\]|<[0-9A-Fa-f]*>)", blok):
            lo, hi = int(m.group(1), 16), int(m.group(2), 16)
            if hi < lo or hi - lo > 65535:
                continue
            hedef = m.group(3)
            if hedef.startswith(b"["):
                for k, hx in enumerate(re.findall(rb"<([0-9A-Fa-f]*)>", hedef)):
                    if lo + k <= hi:
                        harita[lo + k] = _hex_utf16(hx)
            else:
                hx = hedef[1:-1].decode()
                if not hx:
                    continue
                taban, uz = int(hx, 16), len(hx)
                for k in range(hi - lo + 1):
                    harita[lo + k] = _hex_utf16(format(taban + k, "0%dx" % uz))
    if not genislik:
        genislik = 2 if any(k > 255 for k in harita) else 1
    return genislik, harita


_VARSAYILAN_FONT = {"genislik": 1, "harita": None, "kodlama": "cp1252", "okunamaz": False}


def _font_bilgisi(nesneler, no, onbellek):
    if no is None:
        return _VARSAYILAN_FONT
    if no in onbellek:
        return onbellek[no]
    bilgi = dict(_VARSAYILAN_FONT)
    try:
        s = nesneler[no][0]
        tu = _deger_al(s, b"ToUnicode")
        tno = _ref_no(tu) if tu else None
        if tno in nesneler and nesneler[tno][1] is not None:
            cm = _akis_coz(nesneler[tno][0], nesneler[tno][1])
            if cm:
                bilgi["genislik"], bilgi["harita"] = _cmap_oku(cm)
        if bilgi["harita"] is None:
            if re.search(rb"/Subtype\s*/Type0", s):
                bilgi["okunamaz"] = True
            enc = _deger_al(s, b"Encoding") or b""
            if b"MacRoman" in enc:
                bilgi["kodlama"] = "mac_roman"
    except Exception:
        pass
    onbellek[no] = bilgi
    return bilgi


def _metni_coz(ham, bilgi):
    if bilgi["harita"] is not None:
        g = bilgi["genislik"] or 1
        cikti = []
        for i in range(0, len(ham) - g + 1, g):
            kod = int.from_bytes(ham[i:i + g], "big")
            cikti.append(bilgi["harita"].get(kod, " " if (kod == 32 and g == 1) else ""))
        return "".join(cikti)
    if bilgi["okunamaz"]:
        return ""
    return ham.decode(bilgi["kodlama"], "ignore")


_ADI_RE = re.compile(rb"/[^\s/<>\[\]()%]*")
_TOKEN_RE = re.compile(rb"[^\s/<>\[\]()%]+")
_SAYI_RE = re.compile(rb"[+-]?(?:\d+\.?\d*|\.\d+)")
_KACIS = {ord("n"): 10, ord("r"): 13, ord("t"): 9, ord("b"): 8, ord("f"): 12}


def _icerik_parcalari(icerik, font_bul):
    """İçerik akışındaki metin parçalarını (y, x, bitiş_x, metin, boyut) olarak döndürür."""
    parcalar = []
    n = len(icerik)
    i = 0
    ope = []
    ctm, ctm_yigin = _KIMLIK, []
    tm = tlm = _KIMLIK
    lead, boyut, bilgi = 0.0, 10.0, _VARSAYILAN_FONT

    def parca_ekle(metin):
        nonlocal tm
        if not metin:
            return
        efb = (boyut * (tm[0] ** 2 + tm[1] ** 2) ** 0.5 * (ctm[0] ** 2 + ctm[1] ** 2) ** 0.5) or boyut
        x = tm[4] * ctm[0] + tm[5] * ctm[2] + ctm[4]
        y = tm[4] * ctm[1] + tm[5] * ctm[3] + ctm[5]
        parcalar.append((y, x, x + len(metin) * 0.5 * efb, metin, efb))
        tm = (tm[0], tm[1], tm[2], tm[3],
              tm[4] + len(metin) * 0.5 * boyut * tm[0], tm[5] + len(metin) * 0.5 * boyut * tm[1])

    while i < n:
        c = icerik[i:i + 1]
        if c in (b" ", b"\t", b"\r", b"\n", b"\x00", b"\x0c"):
            i += 1
            continue
        if c == b"%":
            j = i
            while j < n and icerik[j] not in (10, 13):
                j += 1
            i = j
            continue
        if c == b"(":
            derin, j, buf = 1, i + 1, bytearray()
            while j < n and derin:
                ch = icerik[j]
                if ch == 0x5C:
                    j += 1
                    if j >= n:
                        break
                    e = icerik[j]
                    if e in _KACIS:
                        buf.append(_KACIS[e])
                        j += 1
                    elif 48 <= e <= 55:
                        okt = bytes([e])
                        j += 1
                        while len(okt) < 3 and j < n and 48 <= icerik[j] <= 55:
                            okt += bytes([icerik[j]])
                            j += 1
                        buf.append(int(okt, 8) & 255)
                    elif e in (10, 13):
                        j += 1
                        if e == 13 and j < n and icerik[j] == 10:
                            j += 1
                    else:
                        buf.append(e)
                        j += 1
                elif ch == 0x28:
                    derin += 1
                    buf.append(ch)
                    j += 1
                elif ch == 0x29:
                    derin -= 1
                    if derin:
                        buf.append(ch)
                    j += 1
                else:
                    buf.append(ch)
                    j += 1
            ope.append(bytes(buf))
            i = j
            continue
        if c == b"<":
            if icerik[i:i + 2] == b"<<":
                ope.append(None)
                i += 2
                continue
            j = icerik.find(b">", i)
            if j < 0:
                break
            hx = re.sub(rb"[^0-9A-Fa-f]", b"", icerik[i + 1:j])
            if len(hx) % 2:
                hx += b"0"
            ope.append(bytes.fromhex(hx.decode()))
            i = j + 1
            continue
        if c == b">":
            i += 2 if icerik[i:i + 2] == b">>" else 1
            continue
        if c == b"[":
            ope.append(_ACIK_DIZI)
            i += 1
            continue
        if c == b"]":
            k = len(ope) - 1
            while k >= 0 and ope[k] is not _ACIK_DIZI:
                k -= 1
            dizi = ope[k + 1:] if k >= 0 else []
            del ope[max(k, 0):]
            ope.append(dizi)
            i += 1
            continue
        if c == b"/":
            m = _ADI_RE.match(icerik, i)
            ope.append(_Ad(m.group(0)[1:]))
            i = m.end()
            continue
        m = _TOKEN_RE.match(icerik, i)
        if not m:
            i += 1
            continue
        tok = m.group(0)
        i = m.end()
        if _SAYI_RE.fullmatch(tok):
            ope.append(float(tok))
            continue
        # ---- operatör ----
        if tok == b"BI":                       # satır içi resim: ikili veriyi atla
            k = re.compile(rb"\sEI(?=\s|$)").search(icerik, i)
            i = k.end() if k else n
            ope = []
            continue
        sayilar = [x for x in ope if isinstance(x, float)]
        if tok == b"q":
            ctm_yigin.append(ctm)
        elif tok == b"Q":
            if ctm_yigin:
                ctm = ctm_yigin.pop()
        elif tok == b"cm" and len(sayilar) >= 6:
            ctm = _carp(tuple(sayilar[-6:]), ctm)
        elif tok == b"BT":
            tm = tlm = _KIMLIK
        elif tok == b"Tf" and len(ope) >= 2 and isinstance(ope[-2], _Ad):
            bilgi = font_bul(bytes(ope[-2]))
            if isinstance(ope[-1], float) and ope[-1]:
                boyut = abs(ope[-1])
        elif tok in (b"Td", b"TD") and len(sayilar) >= 2:
            tx, ty = sayilar[-2], sayilar[-1]
            if tok == b"TD":
                lead = -ty
            tlm = _carp((1.0, 0.0, 0.0, 1.0, tx, ty), tlm)
            tm = tlm
        elif tok == b"Tm" and len(sayilar) >= 6:
            tlm = tm = tuple(sayilar[-6:])
        elif tok == b"TL" and sayilar:
            lead = sayilar[-1]
        elif tok in (b"T*", b"'", b'"'):
            tlm = _carp((1.0, 0.0, 0.0, 1.0, 0.0, -lead), tlm)
            tm = tlm
            if tok != b"T*" and ope and isinstance(ope[-1], bytes):
                parca_ekle(_metni_coz(ope[-1], bilgi))
        elif tok == b"Tj" and ope and isinstance(ope[-1], bytes):
            parca_ekle(_metni_coz(ope[-1], bilgi))
        elif tok == b"TJ" and ope and isinstance(ope[-1], list):
            metin = []
            for oge in ope[-1]:
                if isinstance(oge, bytes):
                    metin.append(_metni_coz(oge, bilgi))
                elif isinstance(oge, float) and oge < -250:
                    metin.append(" ")
            parca_ekle("".join(metin))
        ope = []
    return parcalar


def _satirlara_cevir(parcalar):
    """Konumlarına göre parçaları satırlara dizer (yukarıdan aşağıya, soldan sağa)."""
    satirlar = []
    for p in sorted(parcalar, key=lambda p: -p[0]):
        if satirlar and abs(satirlar[-1][0] - p[0]) <= max(2.0, 0.45 * p[4]):
            satirlar[-1][1].append(p)
        else:
            satirlar.append([p[0], [p]])
    cikti = []
    for _, ogeler in satirlar:
        ogeler.sort(key=lambda p: p[1])
        s = ""
        onceki = None
        for p in ogeler:
            if onceki is None:
                s = p[3]
            elif len(onceki[3].strip()) == 1 and len(p[3].strip()) == 1 and p[1] - onceki[2] < 0.3 * p[4]:
                s += p[3]
            else:
                s += " " + p[3]
            onceki = p
        s = re.sub(r"[ \t]+", " ", s).strip()
        if s:
            cikti.append(s)
    return "\n".join(cikti)


def _sayfa_sirasi(nesneler):
    sonuc = []
    kok = None
    for no, (s, _) in nesneler.items():
        if re.search(rb"/Type\s*/Catalog", s):
            p = _deger_al(s, b"Pages")
            kok = _ref_no(p) if p else None
            break

    def gez(no, d=0):
        if no not in nesneler or d > 20:
            return
        s = nesneler[no][0]
        if re.search(rb"/Type\s*/Pages", s):
            for k in re.findall(rb"(\d+)\s+\d+\s+R", _deger_al(s, b"Kids") or b""):
                gez(int(k), d + 1)
        elif re.search(rb"/Type\s*/Page(?![A-Za-z])", s):
            sonuc.append(no)
    if kok is not None:
        gez(kok)
    if not sonuc:
        sonuc = sorted(no for no, (s, _) in nesneler.items() if re.search(rb"/Type\s*/Page(?![A-Za-z])", s))
    return sonuc


def _sayfa_fontlari(nesneler, sozluk, derinlik=0):
    kay = _deger_al(sozluk, b"Resources")
    if kay is None:
        par = _deger_al(sozluk, b"Parent")
        no = _ref_no(par) if par else None
        if no in nesneler and derinlik < 10:
            return _sayfa_fontlari(nesneler, nesneler[no][0], derinlik + 1)
        return {}
    font = _deger_al(_cozumle(nesneler, kay), b"Font")
    if font is None:
        return {}
    font_s = _cozumle(nesneler, font)
    return {ad: int(no) for ad, no in re.findall(rb"/([^\s/<>\[\]()]+)\s+(\d+)\s+\d+\s+R", font_s)}


def pdf_sayfa_metinleri(yol, en_cok_sayfa=60):
    """Yerleşik okuyucuyla her sayfanın yazı katmanını ayrı ayrı okur -> (sayfa_metinleri, toplam_sayfa).
    Yazı katmanı olmayan sayfaların metni boş gelir."""
    with open(yol, "rb") as f:
        veri = f.read()
    if re.search(rb"/Encrypt\b", veri):
        return [], 0
    nesneler = _pdf_nesneleri(veri)
    onbellek = {}
    sayfa_metinleri = []
    sirali = _sayfa_sirasi(nesneler)
    for sayfa_no in sirali[:en_cok_sayfa]:
        s = nesneler[sayfa_no][0]
        icerik_deger = _deger_al(s, b"Contents") or b""
        if icerik_deger.startswith(b"["):
            refler = [int(x) for x in re.findall(rb"(\d+)\s+\d+\s+R", icerik_deger)]
        else:
            r = _ref_no(icerik_deger)
            refler = [r] if r is not None else []
        icerik = b"\n".join(
            (_akis_coz(nesneler[r][0], nesneler[r][1]) or b"")
            for r in refler if r in nesneler and nesneler[r][1] is not None)
        fontlar = _sayfa_fontlari(nesneler, s)
        parcalar = _icerik_parcalari(
            icerik, lambda ad: _font_bilgisi(nesneler, fontlar.get(ad), onbellek))
        sayfa_metinleri.append(_satirlara_cevir(parcalar))
    return sayfa_metinleri, len(sirali)


def pdf_metin_cikar(yol, en_cok_sayfa=5):
    """PDF'in ilk sayfalarındaki yazı katmanı metni (birleşik). Yazı katmanı yoksa boş metin gelir."""
    metinler, _ = pdf_sayfa_metinleri(yol, en_cok_sayfa)
    return "\n\n".join(m for m in metinler if m)


def _metin_yeterli(metin):
    """Çıkan metin gerçekten bir faturaya benziyor mu? (Resim PDF'lerde metin boş ya da anlamsız gelir.)"""
    k = tr_kucuk(metin)
    anahtarlar = ("fatura", "toplam", "tutar", "tarih", "kdv", "vergi", "vkn", "ödenecek", "mal hizmet")
    return len(k) > 80 and sum(1 for a in anahtarlar if a in k) >= 2


# ---- Windows OCR (PowerShell + Windows.Media.Ocr) ----------------------------

_OCR_PS = r'''
$ErrorActionPreference = 'Stop'
try {
Add-Type -AssemblyName System.Runtime.WindowsRuntime
[void][Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]
[void][Windows.Data.Pdf.PdfDocument,Windows.Data.Pdf,ContentType=WindowsRuntime]
[void][Windows.Data.Pdf.PdfPageRenderOptions,Windows.Data.Pdf,ContentType=WindowsRuntime]
[void][Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime]
[void][Windows.Media.Ocr.OcrResult,Windows.Foundation,ContentType=WindowsRuntime]
[void][Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics.Imaging,ContentType=WindowsRuntime]
[void][Windows.Graphics.Imaging.SoftwareBitmap,Windows.Graphics.Imaging,ContentType=WindowsRuntime]
[void][Windows.Storage.Streams.InMemoryRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime]
[void][Windows.Globalization.Language,Windows.Globalization,ContentType=WindowsRuntime]
$m = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 }
$asOp = $m | Where-Object { $_.IsGenericMethod -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1
$asAct = $m | Where-Object { -not $_.IsGenericMethod -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' } | Select-Object -First 1
function Op($o, $t) { $k = $asOp.MakeGenericMethod($t).Invoke($null, @($o)); $k.Wait(-1) | Out-Null; $k.Result }
function Act($o) { $k = $asAct.Invoke($null, @($o)); $k.Wait(-1) | Out-Null }
$dosya = Op ([Windows.Storage.StorageFile]::GetFileFromPathAsync($env:FATURA_PDF)) ([Windows.Storage.StorageFile])
$belge = Op ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($dosya)) ([Windows.Data.Pdf.PdfDocument])
$motor = $null
try { $motor = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage((New-Object Windows.Globalization.Language 'tr')) } catch {}
if (-not $motor) { $motor = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }
if (-not $motor) { throw 'Windows OCR dil paketi bulunamadi' }
$sb = New-Object System.Text.StringBuilder
$liste = @($env:FATURA_SAYFALAR.Split(',') | ForEach-Object { [int]$_ })
foreach ($i in $liste) {
  if ($i -ge [int]$belge.PageCount) { [void]$sb.Append([char]12); continue }
  $sayfa = $belge.GetPage($i)
  $akis = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
  $sec = New-Object Windows.Data.Pdf.PdfPageRenderOptions
  $sec.DestinationWidth = [uint32]2480
  Act ($sayfa.RenderToStreamAsync($akis, $sec))
  $akis.Seek([uint64]0)
  $kod = Op ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($akis)) ([Windows.Graphics.Imaging.BitmapDecoder])
  $bmp = Op ($kod.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
  $sonuc = Op ($motor.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
  foreach ($s in $sonuc.Lines) { [void]$sb.AppendLine($s.Text) }
  [void]$sb.Append([char]12)
  $sayfa.Dispose()
}
[IO.File]::WriteAllText($env:FATURA_OUT, $sb.ToString(), (New-Object System.Text.UTF8Encoding $false))
} catch { [Console]::Error.WriteLine($_.Exception.Message); exit 2 }
'''


def ocr_ile_oku(yol, sayfalar=(0, 1, 2), zaman_asimi=None):
    """Windows'un yerleşik OCR'ı ile PDF'in verilen sayfalarını (0 tabanlı dizinler) okur.
    -> sayfalarla aynı sırada metin listesi. Windows dışında hata verir."""
    if os.name != "nt":
        raise RuntimeError("Windows OCR yalnızca Windows'ta çalışır.")
    sayfalar = [int(i) for i in sayfalar]
    if not sayfalar:
        return []
    zaman_asimi = zaman_asimi or max(150, 30 * len(sayfalar))
    klasor = tempfile.mkdtemp(prefix="fatura_ocr_")
    try:
        yerel = os.path.join(klasor, "belge.pdf")      # ağ yolundan değil, yerel kopyadan okunur
        shutil.copyfile(yol, yerel)
        cikti = os.path.join(klasor, "cikti.txt")
        env = dict(os.environ, FATURA_PDF=yerel, FATURA_OUT=cikti,
                   FATURA_SAYFALAR=",".join(str(i) for i in sayfalar))
        betik = os.path.join(klasor, "ocr.ps1")        # -EncodedCommand yerine dosyadan çalıştırılır (daha güvenilir)
        with open(betik, "w", encoding="utf-8-sig") as bf:
            bf.write(_OCR_PS)
        p = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", betik],
            env=env, capture_output=True, timeout=zaman_asimi,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        if p.returncode != 0 or not os.path.exists(cikti):
            hata = (p.stderr or b"").decode("utf-8", "ignore").strip() or "bilinmeyen hata"
            raise RuntimeError(hata[:300])
        with open(cikti, encoding="utf-8") as f:
            parcalar = [x.strip() for x in f.read().split("\f")]
        return (parcalar + [""] * len(sayfalar))[:len(sayfalar)]
    except subprocess.TimeoutExpired:
        raise RuntimeError("OCR zaman aşımına uğradı.")
    except FileNotFoundError:
        raise RuntimeError("PowerShell bulunamadı.")
    finally:
        shutil.rmtree(klasor, ignore_errors=True)


# ---- Metinden fatura alanlarını ayıklama -----------------------------------------

_TUTAR_RE = r"(\d{1,3}(?:\.\d{3})+,\d{2}|\d+,\d{2})"
_TARIH_RE = r"(\d{1,2}\s*[./-]\s*\d{1,2}\s*[./-]\s*\d{4})"
_ORANLAR = (0, 1, 8, 10, 18, 20)
_I = "[İIiı]"
_FIRMA_SONEK = re.compile(
    r"(?<![A-Za-zÇĞİÖŞÜçğıöşü])A\.?\s?Ş\b|ANON" + _I + r"M|L" + _I + r"M" + _I + r"TED|\bLTD\b|\bŞT" + _I +
    r"\b|SAN\.|T" + _I + r"C\.|KOLL\.|ORTAKLI[ĞG]I|HOLD" + _I + r"NG|\bŞ" + _I + r"RKET" + _I, re.I)
_KARA_LISTE = ("sayın", "fatura", "tarih", "vkn", "tckn", "vergi", "adres", "tel", "mersis", "ticaret sicil",
               "web", "posta", "senaryo", "ettn", "irsaliye", "sipariş", "ödeme", "banka", "iban", "e-arşiv",
               "e-fatura", "no:", "no ", "mah.", "cad.", "sok.")
_BIRIMLER = {"adet", "ad", "kg", "kgm", "mt", "m", "mtr", "lt", "ltr", "paket", "set", "kutu", "saat", "gün",
             "gun", "ay", "c62", "niu", "top", "çift", "rulo", "metre"}


def _tarih_bul(metin):
    for et in (r"Fatura\s*Tarihi", r"D.zenlenme\s*Tarihi", r"D.zenleme\s*Tarihi"):
        for m in re.finditer(et + r"\s*:?\s*" + _TARIH_RE, metin, re.I):
            try:
                iso = tarih_ayristir(re.sub(r"\s+", "", m.group(1)))
                if 2000 <= int(iso[:4]) <= 2100:
                    return iso
            except ValueError:
                pass
    for m in re.finditer(_TARIH_RE, metin):
        try:
            iso = tarih_ayristir(re.sub(r"\s+", "", m.group(1)))
            if 2000 <= int(iso[:4]) <= 2100:
                return iso
        except ValueError:
            pass
    return None


def _fatura_no_bul(metin):
    m = re.search(r"\b([A-Z0-9]{3}20\d{2}\d{9})\b", metin.upper())
    if m:
        return m.group(1)
    m = re.search(r"Fatura\s*(?:No|Numaras.)\s*[:.]?\s*([A-Za-z0-9][A-Za-z0-9\-/_.]{2,24})", metin, re.I)
    if m:
        return m.group(1).strip(".")
    return None


def _tutar_etiket(metin, kaliplar):
    for k in kaliplar:
        for m in re.finditer(k, metin, re.I):
            a = re.search(_TUTAR_RE, metin[m.end(): m.end() + 90])
            if a:
                try:
                    return tutar_ayristir(a.group(1))
                except ValueError:
                    pass
    return None


def _oran_uyumlu(haric, dahil):
    return any(abs(haric * (100 + o) - dahil * 100) <= 100 for o in _ORANLAR)


def _tutar_cifti_bul(metin):
    """Etiket bulunamazsa: metindeki tutarlar arasında 'KDV hariç x (1+oran) = KDV dahil' olan çifti ara."""
    tutarlar = set()
    for m in re.finditer(_TUTAR_RE, metin):
        try:
            tutarlar.add(tutar_ayristir(m.group(1)))
        except ValueError:
            pass
    tutarlar = sorted((t for t in tutarlar if t > 0), reverse=True)
    for d in tutarlar:
        for h in tutarlar:
            if h < d and _oran_uyumlu(h, d):
                return h, d
    return None


def _tutarlari_bul(metin):
    mal = _tutar_etiket(metin, [r"Mal\s*Hizmet\s*Toplam\s*Tutar"])
    isk = _tutar_etiket(metin, [r"Toplam\s*.skonto"])
    kdv = _tutar_etiket(metin, [r"Hesaplanan\s*K\.?D\.?V"])
    dahil = _tutar_etiket(metin, [r"Vergiler\s*Dahil\s*Toplam\s*Tutar", r"Vergiler\s*Dahil",
                                  r"KDV\s*Dahil\s*Toplam", r".denecek\s*Tutar", r"Genel\s*Toplam"])
    haric = None
    if mal is not None:
        haric = mal - (isk or 0)
    else:
        haric = _tutar_etiket(metin, [r"Vergiler\s*Hari.", r"KDV\s*Hari.\s*Toplam", r"Matrah", r"Ara\s*Toplam"])
    if haric is None and dahil is not None and kdv is not None:
        haric = dahil - kdv
    if dahil is None and haric is not None and kdv is not None:
        dahil = haric + kdv
    tutarli = haric is not None and dahil is not None and 0 <= haric <= dahil and _oran_uyumlu(haric, dahil)
    if not tutarli:
        cift = _tutar_cifti_bul(metin)
        if cift:
            haric, dahil = cift
        elif haric is not None and dahil is not None and haric > dahil:
            haric = dahil = None
    return haric, dahil


_VKN_ETIKET = re.compile(
    r"(?:VKN|TCKN|Vergi\s*Kimlik\s*(?:No|Numaras.)|T\.?\s?C\.?\s*Kimlik\s*(?:No|Numaras.)|"
    r"Vergi\s*(?:No|Numaras.)|V\.?\s?K\.?\s?N\.?)\s*[:/.]?\s*(?<!\d)(\d{10}|\d{11})(?!\d)", re.I)


def _vkn_bul(metin, kurum_vkn=""):
    """Satıcının VKN (10 hane) ya da TCKN (11 hane, şahıs işletmesi) numarasını bulur.
    Önce 'SAYIN' (alıcı) bölümünden önceye bakar; iki sütunlu faturalarda satıcı bilgisi
    'Sayın'dan sonra gelebildiği için bulunamazsa tüm metinde ilk uygun numara alınır."""
    sayin = re.search(r"SAYIN", metin.replace("ı", "I").upper())
    bolgeler = [metin[:sayin.start()]] if sayin else []
    bolgeler.append(metin)
    for bolge in bolgeler:
        for m in _VKN_ETIKET.finditer(bolge):
            if m.group(1) != kurum_vkn:
                return m.group(1)
    return None


_SADE = str.maketrans("ıİIşŞğĞüÜöÖçÇ", "iiissgguuoocc")


def _sade(s):
    """Türkçe harfleri ASCII'ye indirger (OCR 'TEİAŞ' yerine 'TEIAS' okusa da eşleşsin)."""
    return (s or "").translate(_SADE).lower().strip()


_BELGE_ONEKI = re.compile(r"^\s*e\s*-\s*(?:fatura|ar[şs]iv(?:\s*fatura)?)\s*", re.I)
_DEVAM_RE = re.compile(r"^(?:T[İI]C|T[İI]CARET|LTD|ŞT[İI]|SAN|SANAY[İI]|A\.?\s?Ş|ANON[İI]M|L[İI]M[İI]TED|"
                       r"[İI]TH|[İI]HR|PAZ|MAK[İI]NE)", re.I)


def _devam_satiri_ekle(ad, bolge, i):
    """Firma adı satır sonunda bağlaçla bitiyorsa ya da sonraki satır unvan devamı gibi duruyorsa (TİC.LTD.ŞTİ) ekler."""
    if i + 1 >= len(bolge):
        return ad
    sonraki = bolge[i + 1].strip()
    bagla_biter = re.search(r"(VE|&|T[İI]C\.?|SAN\.?)\s*$", ad, re.I)
    kisa_metin = not re.search(r"\d", sonraki) and len(sonraki) <= 40
    if kisa_metin and (bagla_biter or not _FIRMA_SONEK.search(ad)) and _DEVAM_RE.match(sonraki):
        return ad + " " + sonraki
    if kisa_metin and bagla_biter and not any(b in tr_kucuk(sonraki) for b in _KARA_LISTE):
        return ad + " " + sonraki
    return ad


def _firma_bul(satirlar, kurum_anahtar=""):
    anahtarlar = [_sade(a) for a in kurum_anahtar.split(",") if a.strip()]
    bolge = []
    for s in satirlar[:45]:
        if "sayın" in tr_kucuk(s):
            break
        bolge.append(s)

    def kurum_mu(s):
        k = _sade(s)
        return any(a in k for a in anahtarlar)

    adaylar = [i for i, s in enumerate(bolge)
               if _FIRMA_SONEK.search(s) and not kurum_mu(s) and len(re.findall(r"[A-Za-zÇĞİÖŞÜçğıöşü]{2,}", s)) >= 2]
    if adaylar:
        i = adaylar[0]
        ad = _BELGE_ONEKI.sub("", bolge[i])
        ad = _devam_satiri_ekle(ad, bolge, i)
        baglac = re.match(r"\s*(VE|T[İI]CARET|SANAY[İI]|T[İI]C\.|SAN\.)\b", ad, re.I)
        if baglac and i > 0 and not kurum_mu(bolge[i - 1]) and not re.search(r"\d", bolge[i - 1]):
            ad = bolge[i - 1] + " " + ad
        return re.sub(r"\s+", " ", ad).strip(" -:,")
    for s in bolge:
        harfler = re.findall(r"[A-Za-zÇĞİÖŞÜçğıöşü]", s)
        if len(harfler) < 6 or kurum_mu(s):
            continue
        if any(b in tr_kucuk(s) for b in _KARA_LISTE) or re.search(r"\d{4,}", s):
            continue
        if sum(1 for h in harfler if h == h.upper()) / len(harfler) > 0.7:
            return re.sub(r"\s+", " ", s).strip(" -:,")
    return None


def _is_bul(satirlar):
    bas = None
    for i, s in enumerate(satirlar):
        k = tr_kucuk(s)
        if re.search(r"mal\s*hizmet|a[çc][ıi]klama", k) and re.search(r"miktar|birim|fiyat", k):
            bas = i
            break
    if bas is None:
        return None
    kalemler = []
    for s in satirlar[bas + 1: bas + 9]:
        k = tr_kucuk(s)
        if re.search(r"toplam|iskonto|hesaplanan|kdv|yaln[ıi]z|not\s*:", k):
            break
        parcalar = re.sub(r"^\s*\d{1,3}\s+", "", s).split()
        cikti = []
        for j, t in enumerate(parcalar):
            tl = tr_kucuk(t)
            if cikti and (re.fullmatch(r"%?\d+[.,]\d+%?|%\d+", t) or tl in _BIRIMLER):
                break
            if cikti and re.fullmatch(r"\d+", t) and j + 1 < len(parcalar) and tr_kucuk(parcalar[j + 1]) in _BIRIMLER:
                break
            cikti.append(t)
        ad = " ".join(cikti).strip()
        if len(ad) >= 3 and re.search(r"[A-Za-zÇĞİÖŞÜçğıöşü]{3}", ad):
            kalemler.append(ad)
        if len(kalemler) == 2:
            break
    return ("; ".join(kalemler))[:120] if kalemler else None


def fatura_metnini_coz(metin, ayar=None):
    """Okunan metinden alanları ayıklar. Bulunamayanlar sonuca girmez. Tutarlar kuruş, tarih ISO biçimindedir."""
    ayar = ayar or {}
    metin = (metin or "").replace("\xa0", " ")
    satirlar = [s.strip() for s in metin.splitlines() if s.strip()]
    sonuc = {}
    tarih = _tarih_bul(metin)
    if tarih:
        sonuc["fatura_tarihi"] = tarih
    no = _fatura_no_bul(metin)
    if no:
        sonuc["fatura_no"] = no
    vkn = _vkn_bul(metin, ayar.get("kurum_vkn", ""))
    if vkn:
        sonuc["vkn"] = vkn
    firma = _firma_bul(satirlar, ayar.get("kurum_anahtar", ""))
    if firma:
        sonuc["firma_adi"] = firma
    haric, dahil = _tutarlari_bul(metin)
    if haric is not None:
        sonuc["tutar_haric"] = haric
    if dahil is not None:
        sonuc["tutar_dahil"] = dahil
    return sonuc


def pdf_sayfa_metinleri_mupdf(yol, en_cok_sayfa=60):
    """PyMuPDF kuruluysa sayfa sayfa metin çıkarır -> (sayfa_metinleri, toplam_sayfa); yoksa ([], 0)."""
    if fitz is None:
        return [], 0
    try:
        with fitz.open(yol) as belge:
            if belge.needs_pass:
                return [], 0
            toplam = len(belge)
            return [belge[i].get_text("text", sort=True).strip() for i in range(min(toplam, en_cok_sayfa))], toplam
    except Exception:
        return [], 0


_FATURA_IPUCLARI = ("fatura no", "fatura numaras", "fatura tarihi", "ettn", "mal hizmet toplam", "vergiler dahil",
                    "ödenecek tutar", "hesaplanan kdv", "e-arşiv", "e-fatura", "senaryo", "özelleştirme",
                    "düzenlenme tarihi", "sayın")
_OCR_SAYFA_SINIRI = 20


def _fatura_puani(metin):
    """Sayfada kaç farklı 'fatura ipucu' var? (Fatura sayfasını talep/yazı sayfalarından ayırmak için.)"""
    k = tr_kucuk(metin)
    return sum(1 for a in _FATURA_IPUCLARI if a in k)


def sayfa_ayristir(metin, toplam):
    """'7', '7-8' veya '2,5' -> 0 tabanlı sayfa dizinleri (boşsa None)."""
    s = (metin or "").strip().replace(" ", "")
    if not s:
        return None
    biçim_hata = ValueError("Fatura sayfası '7', '7-8' ya da '2,5' biçiminde yazılmalı.")
    numaralar = []
    try:
        for parca in s.split(","):
            if "-" in parca:
                a, b = (int(x) for x in parca.split("-", 1))
                numaralar += list(range(min(a, b), max(a, b) + 1))
            else:
                numaralar.append(int(parca))
    except ValueError:
        raise biçim_hata
    if not numaralar or len(numaralar) > 20:
        raise biçim_hata
    if any(n < 1 or (toplam and n > toplam) for n in numaralar):
        raise ValueError("Fatura sayfası 1 ile %d arasında olmalı." % toplam if toplam else "Sayfa numarası geçersiz.")
    return sorted({n - 1 for n in numaralar})


def pdf_bilgi_cikar(yol, ayar=None, sayfa=None):
    """PDF'i okur ve fatura sayfasındaki bilgileri ayıklar.
    sayfa: None/boş = fatura sayfasını otomatik bul; '7', '7-8', '2,5' = yalnızca bu sayfalar okunur.
    Okuma sırası: PyMuPDF (varsa) -> yerleşik yazı katmanı okuyucu -> Windows OCR (yazısı olmayan sayfalar).
    -> {yontem, metin, alanlar, hata, sayfa, sayfa_sayisi, sayfa_bilgisi}"""
    sonuc = {"yontem": None, "metin": "", "alanlar": {}, "hata": None,
             "sayfa": None, "sayfa_sayisi": None, "sayfa_bilgisi": ""}
    sayfalar, toplam = pdf_sayfa_metinleri_mupdf(yol)
    if not any(_metin_yeterli(m) for m in sayfalar):
        try:
            s2, t2 = pdf_sayfa_metinleri(yol)
        except Exception:
            s2, t2 = [], 0
        if any(_metin_yeterli(m) for m in s2) or not sayfalar:
            sayfalar = s2
        toplam = max(toplam, t2)
    sayfalar = list(sayfalar)
    n = max(toplam, len(sayfalar))
    sonuc["sayfa_sayisi"] = n or None

    ocr = {"hata": None, "yapilan": set()}

    def ocr_calistir(dizinler):
        if not dizinler:
            return
        if os.name != "nt":
            ocr["hata"] = "PDF'te okunabilir yazı katmanı yok (resim olarak taranmış). OCR yalnızca Windows'ta çalışır."
            return
        try:
            metinler = ocr_ile_oku(yol, dizinler)
        except Exception as ex:
            ocr["hata"] = "Windows OCR çalışmadı: %s" % ex
            return
        for i, m in zip(dizinler, metinler):
            if len(m.strip()) > 40:
                while len(sayfalar) <= i:
                    sayfalar.append("")
                sayfalar[i] = m
                ocr["yapilan"].add(i)

    def en_iyi_sayfa():
        puanlar = [(_fatura_puani(m) if _metin_yeterli(m) else 0) for m in sayfalar]
        if not puanlar or max(puanlar) < 2:
            return None
        return puanlar.index(max(puanlar))

    istenen = sayfa_ayristir(sayfa, n)
    secilen = None
    if istenen:
        ocr_calistir([i for i in istenen if i >= len(sayfalar) or not _metin_yeterli(sayfalar[i])])
        secilen = istenen
        sonuc["sayfa_bilgisi"] = "Yalnızca %s. sayfa okundu." % ", ".join(str(i + 1) for i in istenen)
    else:
        i = en_iyi_sayfa()
        if i is None:                # yazı katmanı olmayan (taranmış) sayfaları OCR'dan geçirip tekrar dene
            hedef = [k for k in range(min(n or 3, _OCR_SAYFA_SINIRI))
                     if k >= len(sayfalar) or len(sayfalar[k].strip()) < 40]
            ocr_calistir(hedef)
            i = en_iyi_sayfa()
        if i is not None:
            secilen = [i]
            if n > 1:
                sonuc["sayfa_bilgisi"] = ("%d sayfalık belgede fatura %d. sayfada bulundu. Yanlışsa 'Fatura sayfası' "
                                          "kutusuna doğru sayfayı yazıp 'PDF'den Oku'ya basın." % (n, i + 1))

    if secilen:
        metin = "\n\n".join(sayfalar[i] for i in secilen if i < len(sayfalar))
        sonuc["sayfa"] = ", ".join(str(i + 1) for i in secilen)
        kaynak_ocr = any(i in ocr["yapilan"] for i in secilen)
    else:                            # fatura sayfası ayırt edilemedi: eski davranış, ilk sayfaların metni
        ilk = range(min(5, len(sayfalar)))
        metin = "\n\n".join(sayfalar[i] for i in ilk if sayfalar[i].strip())
        kaynak_ocr = any(i in ocr["yapilan"] for i in ilk)
        if n > 1 and metin.strip():
            sonuc["sayfa_bilgisi"] = ("Fatura sayfası otomatik belirlenemedi, ilk sayfalar okundu. Doğru sayfayı "
                                      "'Fatura sayfası' kutusuna yazıp 'PDF'den Oku'ya basın.")
    if metin.strip() and (secilen or _metin_yeterli(metin)):
        sonuc["yontem"] = "ocr" if kaynak_ocr else "metin"
    if sonuc["yontem"] is None or not _metin_yeterli(metin):
        sonuc["hata"] = ocr["hata"] or (None if sonuc["yontem"] else
                                        "OCR çalıştı ama sayfada okunabilir yazı bulamadı (tarama çok silik/eğik olabilir).")
    sonuc["metin"] = metin
    if metin.strip():
        sonuc["alanlar"] = fatura_metnini_coz(metin, ayar)
    return sonuc


# ----------------------------------------------------------------------------
# 6) ARAYÜZ
# ----------------------------------------------------------------------------

TEMALAR = {
    "Açık": dict(bg="#fafafa", fg="#1c1c1c", alan="#ffffff", sec="#005fb8", sec_fg="#ffffff",
                 baslik="#e9e9e9", oto="#fff3c4", cizgi="#d0d0d0", dugme="#f3f3f3", zebra="#f4f7fb",
                 **{"not": "#6b6b6b"}),
    "Koyu": dict(bg="#1c1c1c", fg="#ffffff", alan="#2b2b2b", sec="#57c8ff", sec_fg="#000000",
                 baslik="#323232", oto="#5a4a12", cizgi="#454545", dugme="#2f2f2f", zebra="#242424",
                 **{"not": "#a0a0a0"}),
    "Mavi": dict(bg="#e8f0fa", fg="#12263f", alan="#ffffff", sec="#1f5fae", sec_fg="#ffffff",
                 baslik="#c9dcf2", oto="#fff0b3", cizgi="#a9c3e3", dugme="#d3e3f6", zebra="#f0f6fd",
                 **{"not": "#4a6a8f"}),
}


def arayuzu_calistir():
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog

    ayar = ayarlari_oku()
    kullanici = getpass.getuser()

    kok = tk.Tk()
    kok.title(UYGULAMA_ADI)
    kok.geometry("1240x680")
    kok.minsize(980, 520)

    try:
        db = FaturaDB(ayar["veritabani"])
    except VeritabaniHatasi as e:
        messagebox.showerror(UYGULAMA_ADI, str(e))
        kok.destroy()
        return

    # ---- tema ----------------------------------------------------------------
    stil = ttk.Style(kok)
    tema = {"ad": "Açık", "r": TEMALAR["Açık"]}
    v_tema = tk.StringVar(value=ayar["tema"] if ayar["tema"] in TEMALAR else "Açık")

    def tema_uygula(ad, kaydet=False):
        if ad not in TEMALAR:
            ad = "Açık"
        r = TEMALAR[ad]
        tema["ad"], tema["r"] = ad, r
        if sv_ttk is not None and ad in ("Açık", "Koyu"):
            sv_ttk.set_theme("light" if ad == "Açık" else "dark")
            stil.configure("Treeview", rowheight=30, font=("Segoe UI", 10))
            stil.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
            stil.configure("Oto.TEntry", fieldbackground=r["oto"], foreground=r["fg"])
            stil.configure("Oto.TCombobox", fieldbackground=r["oto"], foreground=r["fg"])
            stil.configure("Toplam.TLabel", font=("Segoe UI", 10, "bold"))
            stil.configure("Not.TLabel", foreground=r["not"])
            stil.map("Toolbutton", foreground=[("selected", r["sec"])])
            kok.configure(bg=r["bg"])
            for w in kok.winfo_children():
                if isinstance(w, tk.Toplevel):
                    w.configure(bg=r["bg"])
            if kaydet:
                ayar_yaz(CONFIG_YOLU, "tema", ad)
            return
        try:
            stil.theme_use("clam")
        except tk.TclError:
            pass
        stil.configure(".", background=r["bg"], foreground=r["fg"], fieldbackground=r["alan"],
                       bordercolor=r["cizgi"], troughcolor=r["bg"], lightcolor=r["bg"], darkcolor=r["bg"])
        stil.configure("TButton", background=r["dugme"], foreground=r["fg"], padding=(8, 4))
        stil.map("TButton", background=[("active", r["baslik"]), ("pressed", r["cizgi"])])
        stil.configure("TEntry", fieldbackground=r["alan"], foreground=r["fg"], insertcolor=r["fg"])
        stil.configure("Oto.TEntry", fieldbackground=r["oto"], foreground=r["fg"], insertcolor=r["fg"])
        stil.configure("TCombobox", fieldbackground=r["alan"], background=r["dugme"], foreground=r["fg"],
                       arrowcolor=r["fg"], insertcolor=r["fg"])
        stil.configure("Oto.TCombobox", fieldbackground=r["oto"], background=r["dugme"], foreground=r["fg"],
                       arrowcolor=r["fg"], insertcolor=r["fg"])
        stil.map("TCombobox", fieldbackground=[("readonly", r["alan"])], foreground=[("readonly", r["fg"])],
                 selectbackground=[("readonly", r["alan"])], selectforeground=[("readonly", r["fg"])])
        stil.configure("Treeview", background=r["alan"], fieldbackground=r["alan"], foreground=r["fg"],
                       rowheight=24)
        stil.map("Treeview", background=[("selected", r["sec"])], foreground=[("selected", r["sec_fg"])])
        stil.configure("Treeview.Heading", background=r["baslik"], foreground=r["fg"], relief="flat",
                       padding=(4, 4))
        stil.map("Treeview.Heading", background=[("active", r["cizgi"])])
        stil.configure("Toolbutton", background=r["dugme"], foreground=r["fg"], padding=(7, 3))
        stil.map("Toolbutton", background=[("selected", r["sec"]), ("active", r["baslik"])],
                 foreground=[("selected", r["sec_fg"])])
        stil.configure("TScrollbar", background=r["dugme"], troughcolor=r["bg"], arrowcolor=r["fg"])
        stil.configure("Toplam.TLabel", font=("Segoe UI", 10, "bold"))
        stil.configure("Not.TLabel", foreground=r["not"])
        kok.configure(bg=r["bg"])
        kok.option_add("*TCombobox*Listbox.background", r["alan"])
        kok.option_add("*TCombobox*Listbox.foreground", r["fg"])
        kok.option_add("*TCombobox*Listbox.selectBackground", r["sec"])
        kok.option_add("*TCombobox*Listbox.selectForeground", r["sec_fg"])
        for w in kok.winfo_children():
            if isinstance(w, tk.Toplevel):
                w.configure(bg=r["bg"])
                for t in w.winfo_children():
                    if isinstance(t, tk.Text):
                        t.configure(bg=r["alan"], fg=r["fg"], insertbackground=r["fg"])
        if kaydet:
            ayar_yaz(CONFIG_YOLU, "tema", ad)

    tema_uygula(v_tema.get())

    def yeni_pencere(baslik):
        p = tk.Toplevel(kok)
        p.title(baslik)
        p.configure(bg=tema["r"]["bg"])
        p.transient(kok)
        return p

    def arka_planda(is_fonk, bitince):
        """Uzun işi (PDF okuma) ayrı iş parçacığında yapar; arayüz donmaz. bitince(değer, hata) ana iş parçacığında çağrılır."""
        kutu = {}

        def calis():
            try:
                kutu["deger"] = is_fonk()
            except Exception as ex:  # noqa
                kutu["hata"] = ex
        t = threading.Thread(target=calis, daemon=True)
        t.start()

        def bekle():
            if t.is_alive():
                kok.after(200, bekle)
                return
            bitince(kutu.get("deger"), kutu.get("hata"))
        kok.after(200, bekle)

    def hata_goster(e):
        messagebox.showerror(UYGULAMA_ADI, str(e), parent=kok)

    # ---- durum ---------------------------------------------------------------
    durum = {"satirlar": [], "gorunen": [], "sirala": ("fatura_tarihi", True)}
    v_arama = tk.StringVar()
    v_firma = tk.StringVar(value="Tümü")
    v_bas = tk.StringVar()
    v_bit = tk.StringVar()
    v_toplam = tk.StringVar()
    v_yil = tk.StringVar()
    v_ay = tk.IntVar(value=0)     # 0 = Tümü, 1..12 = ay

    # ---- üst bar 1: yıl + ay sekmeleri -----------------------------------------
    ust1 = ttk.Frame(kok, padding=(8, 8, 8, 0))
    ust1.pack(fill="x")
    ttk.Label(ust1, text="Yıl:").pack(side="left")
    cb_yil = ttk.Combobox(ust1, textvariable=v_yil, width=7, state="readonly")
    cb_yil.pack(side="left", padx=(4, 10))
    ay_dugmeleri = []
    rb = ttk.Radiobutton(ust1, text="Tümü", value=0, variable=v_ay, style="Toolbutton")
    rb.pack(side="left", padx=1)
    ay_dugmeleri.append(rb)
    for k in range(1, 13):
        rb = ttk.Radiobutton(ust1, text=AY_ADLARI[k - 1], value=k, variable=v_ay, style="Toolbutton")
        rb.pack(side="left", padx=1)
        ay_dugmeleri.append(rb)

    # ---- üst bar 2: arama + filtre -------------------------------------------
    ust = ttk.Frame(kok, padding=(8, 8, 8, 2))
    ust.pack(fill="x")
    ttk.Label(ust, text="Ara:").grid(row=0, column=0, sticky="w")
    e_ara = ttk.Entry(ust, textvariable=v_arama, width=34)
    e_ara.grid(row=0, column=1, padx=(4, 12))
    ttk.Label(ust, text="Firma:").grid(row=0, column=2, sticky="w")
    cb_firma = ttk.Combobox(ust, textvariable=v_firma, width=30, state="readonly")
    cb_firma.grid(row=0, column=3, padx=(4, 12))
    ttk.Label(ust, text="Tarih (GG.AA.YYYY):").grid(row=0, column=4, sticky="w")
    ttk.Entry(ust, textvariable=v_bas, width=11).grid(row=0, column=5, padx=(4, 2))
    ttk.Label(ust, text="-").grid(row=0, column=6)
    ttk.Entry(ust, textvariable=v_bit, width=11).grid(row=0, column=7, padx=(2, 8))
    btn_temizle = ttk.Button(ust, text="Filtreleri Temizle")
    btn_temizle.grid(row=0, column=8, padx=4)

    # ---- tablo ---------------------------------------------------------------
    orta = ttk.Frame(kok, padding=(8, 4))
    orta.pack(fill="both", expand=True)
    sutunlar = [
        ("fatura_tarihi", "Fatura Tarihi", 95, "center"),
        ("fatura_no", "Fatura No", 150, "w"),
        ("firma_adi", "Firma Adı", 340, "w"),
        ("yapilan_is", "Yapılan İş", 280, "w"),
        ("tutar_haric", "Tutar (KDV Hariç)", 130, "e"),
        ("tutar_dahil", "Tutar (KDV Dahil)", 130, "e"),
        ("pdf_var", "PDF", 45, "center"),
        ("ek_sayisi", "Ek", 40, "center"),
    ]
    tree = ttk.Treeview(orta, columns=[s[0] for s in sutunlar], show="headings", selectmode="browse")
    for ad, baslik, genislik, hiza in sutunlar:
        tree.heading(ad, text=baslik, command=lambda a=ad: sirala_tikla(a))
        tree.column(ad, width=genislik, anchor=hiza, stretch=(ad in ("firma_adi", "yapilan_is")))
    sb = ttk.Scrollbar(orta, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")

    # ---- alt bar -------------------------------------------------------------
    alt = ttk.Frame(kok, padding=(8, 4, 8, 8))
    alt.pack(fill="x")
    btn_ekle = ttk.Button(alt, text="Yeni Fatura")
    btn_duzenle = ttk.Button(alt, text="Düzenle")
    btn_sil = ttk.Button(alt, text="Sil")
    btn_pdf = ttk.Button(alt, text="PDF'i Aç")
    btn_evrak = ttk.Button(alt, text="Evraklar...")
    btn_klasor = ttk.Button(alt, text="Klasörde Göster")
    btn_ozet = ttk.Button(alt, text="Aylık Özet")
    btn_csv = ttk.Button(alt, text="Excel'e Aktar")
    for b in (btn_ekle, btn_duzenle, btn_sil, btn_pdf, btn_evrak, btn_klasor, btn_ozet, btn_csv):
        b.pack(side="left", padx=3)
    ttk.Label(alt, textvariable=v_toplam, style="Toplam.TLabel").pack(side="right")

    # ---- liste yenileme ------------------------------------------------------
    def filtre_degerleri():
        bas = bit = None
        try:
            if v_bas.get().strip():
                bas = tarih_ayristir(v_bas.get())
            if v_bit.get().strip():
                bit = tarih_ayristir(v_bit.get())
        except ValueError:
            return None  # yazım tamamlanmamış; filtreyi uygulama
        return bas, bit

    def anahtar(r, sutun):
        if sutun in ("tutar_haric", "tutar_dahil"):
            return r[sutun] if r[sutun] is not None else -1
        if sutun == "pdf_var":
            return 1 if r["pdf_yolu"] else 0
        if sutun == "ek_sayisi":
            return r.get("ek_sayisi") or 0
        if sutun == "fatura_tarihi":
            return r["fatura_tarihi"] or ""
        return tr_kucuk(r[sutun] or "")

    def goster(donem=""):
        tree.delete(*tree.get_children())
        sutun, azalan = durum["sirala"]
        satirlar = sorted(durum["satirlar"], key=lambda r: anahtar(r, sutun), reverse=azalan)
        tree.tag_configure("cift", background=tema["r"]["zebra"])
        for sira, r in enumerate(satirlar):
            tree.insert("", "end", iid=str(r["id"]), tags=("cift",) if sira % 2 else (), values=(
                tarih_yaz(r["fatura_tarihi"]), r["fatura_no"], r["firma_adi"], r["yapilan_is"] or "",
                tutar_yaz(r["tutar_haric"]), tutar_yaz(r["tutar_dahil"]), "✔" if r["pdf_yolu"] else "",
                r.get("ek_sayisi") or ""))
        durum["gorunen"] = satirlar
        adet, haric, dahil = toplamlar(satirlar)
        v_toplam.set("%s%d fatura   |   KDV Hariç: %s   |   KDV Dahil: %s"
                     % (donem + "  -  " if donem else "", adet, tutar_yaz(haric, True), tutar_yaz(dahil, True)))
        for ad, baslik, _, _ in sutunlar:
            ok = (" ▼" if azalan else " ▲") if ad == sutun else ""
            tree.heading(ad, text=baslik + ok)

    def yenile(*_):
        try:
            fv = filtre_degerleri()
            if fv is None:
                return
            tum = db.hepsi()
            # yıl listesi
            yillar = sorted({int(r["fatura_tarihi"][:4]) for r in tum} | {date.today().year}, reverse=True)
            cb_yil["values"] = ["Tümü"] + [str(y) for y in yillar]
            if v_yil.get() not in ["Tümü"] + [str(y) for y in yillar]:
                v_yil.set(str(yillar[0]))
            yil = None if v_yil.get() == "Tümü" else int(v_yil.get())
            # firma listesi
            firmalar = sorted({r["firma_adi"] for r in tum}, key=tr_kucuk)
            cb_firma["values"] = ["Tümü"] + firmalar
            if v_firma.get() not in ["Tümü"] + firmalar:
                v_firma.set("Tümü")
            # ay düğmelerindeki adetler (seçili yıla göre)
            yil_kayitlari = filtrele(tum, yil=yil)
            sayac = {k: 0 for k in range(1, 13)}
            for r in yil_kayitlari:
                sayac[int(r["fatura_tarihi"][5:7])] += 1
            ay_dugmeleri[0].configure(text="Tümü (%d)" % len(yil_kayitlari))
            for k in range(1, 13):
                ay_dugmeleri[k].configure(text="%s (%d)" % (AY_ADLARI[k - 1], sayac[k]) if sayac[k] else AY_ADLARI[k - 1])
            ay = v_ay.get() or None
            firma = "" if v_firma.get() == "Tümü" else v_firma.get()
            durum["satirlar"] = filtrele(tum, v_arama.get(), firma, fv[0], fv[1], yil, ay)
            donem = ((AY_ADLARI[ay - 1] + " ") if ay else "") + (str(yil) if yil else "Tüm yıllar")
            goster(donem)
        except VeritabaniHatasi as e:
            hata_goster(e)

    def sirala_tikla(sutun):
        eski, azalan = durum["sirala"]
        durum["sirala"] = (sutun, (not azalan) if eski == sutun else False)
        yenile()

    def temizle():
        v_arama.set("")
        v_firma.set("Tümü")
        v_bas.set("")
        v_bit.set("")
        v_ay.set(0)
        yenile()

    def secili_id():
        s = tree.selection()
        if not s:
            messagebox.showinfo(UYGULAMA_ADI, "Önce listeden bir fatura seçin.", parent=kok)
            return None
        return int(s[0])

    # ---- metin gösterici (okunan ham metin) ----------------------------------
    def metin_goster(baslik, metin):
        pen = yeni_pencere(baslik)
        pen.geometry("760x560")
        t = tk.Text(pen, wrap="word", bg=tema["r"]["alan"], fg=tema["r"]["fg"], insertbackground=tema["r"]["fg"],
                    font=("Consolas", 10))
        t.insert("1.0", metin or "(boş)")
        t.configure(state="normal")
        sb2 = ttk.Scrollbar(pen, command=t.yview)
        t.configure(yscrollcommand=sb2.set)
        sb2.pack(side="right", fill="y")
        t.pack(fill="both", expand=True, padx=6, pady=6)
        pen.bind("<Escape>", lambda e: pen.destroy())

    def durum_metni(sonuc, bulunan):
        if sonuc["yontem"] == "metin":
            kaynak = "PDF'in yazı katmanından okundu"
        elif sonuc["yontem"] == "ocr":
            kaynak = "Tarama Windows OCR ile okundu (hata payı daha yüksek olabilir)"
        else:
            return "PDF'den okunamadı. %s Bilgileri elle girebilirsiniz." % (sonuc["hata"] or "")
        if bulunan == 0:
            return "%s ama fatura bilgisi ayıklanamadı. Bilgileri elle girin." % kaynak
        return "%s. %d/5 alan bulundu; sarı alanları mutlaka kontrol edin. (Yapılan İş elle girilir.)" % (kaynak, bulunan)

    # ---- fatura formu --------------------------------------------------------
    def form_ac(fid=None):
        mevcut = db.getir(fid) if fid else None
        if fid and not mevcut:
            messagebox.showwarning(UYGULAMA_ADI, "Kayıt bulunamadı (başkası silmiş olabilir).", parent=kok)
            yenile()
            return
        pen = yeni_pencere("Fatura Düzenle" if fid else "Yeni Fatura")
        pen.grab_set()
        pen.resizable(False, False)

        v = {
            "tarih": tk.StringVar(value=tarih_yaz(mevcut["fatura_tarihi"]) if mevcut else date.today().strftime("%d.%m.%Y")),
            "no": tk.StringVar(value=mevcut["fatura_no"] if mevcut else ""),
            "firma": tk.StringVar(value=mevcut["firma_adi"] if mevcut else ""),
            "is": tk.StringVar(value=(mevcut["yapilan_is"] or "") if mevcut else ""),
            "haric": tk.StringVar(value=tutar_yaz(mevcut["tutar_haric"]) if mevcut else ""),
            "dahil": tk.StringVar(value=tutar_yaz(mevcut["tutar_dahil"]) if mevcut else ""),
            "pdf": tk.StringVar(value=(mevcut["pdf_yolu"] or "") if mevcut else ""),
            "vkn": tk.StringVar(value=(mevcut.get("vkn") or "") if mevcut else ""),
            "sayfa": tk.StringVar(value=""),
        }
        v_durum = tk.StringVar(value="")
        son_metin = {"metin": ""}

        cerceve = ttk.Frame(pen, padding=14)
        cerceve.pack()
        etiketler = [("Fatura Tarihi (GG.AA.YYYY)", "tarih"), ("Fatura No", "no"), ("Firma Adı", "firma"),
                     ("Yapılan İş", "is"), ("Tutar (KDV Hariç)", "haric"), ("Tutar (KDV Dahil)", "dahil")]
        girisler = {}

        def normal_stil(k):
            girisler[k].configure(style="TCombobox" if k == "firma" else "TEntry")

        for i, (et, k) in enumerate(etiketler):
            ttk.Label(cerceve, text=et).grid(row=i, column=0, sticky="w", pady=4)
            if k == "firma":
                g = ttk.Combobox(cerceve, textvariable=v[k], width=52, values=db.firmalar())
                g.bind("<<ComboboxSelected>>", lambda e, kk=k: normal_stil(kk))
            else:
                g = ttk.Entry(cerceve, textvariable=v[k], width=55 if k not in ("haric", "dahil") else 20)
            g.bind("<Key>", lambda e, kk=k: normal_stil(kk))
            g.grid(row=i, column=1, sticky="w", padx=(10, 0), pady=4, columnspan=2)
            girisler[k] = g

        kdv_cerceve = ttk.Frame(cerceve)
        kdv_cerceve.grid(row=5, column=1, sticky="e", columnspan=2)
        v_oran = tk.StringVar(value="20")
        ttk.Label(kdv_cerceve, text="KDV %").pack(side="left")
        ttk.Entry(kdv_cerceve, textvariable=v_oran, width=4).pack(side="left", padx=3)

        def kdv_hesapla():
            try:
                h = tutar_ayristir(v["haric"].get())
                if h is None:
                    raise ValueError("Önce KDV hariç tutarı girin.")
                oran = Decimal(v_oran.get().replace(",", ".") or "0")
                d = int((Decimal(h) * (100 + oran) / 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
                v["dahil"].set(tutar_yaz(d))
                normal_stil("dahil")
            except (ValueError, InvalidOperation) as ex:
                messagebox.showwarning(UYGULAMA_ADI, str(ex), parent=pen)
        ttk.Button(kdv_cerceve, text="Dahil'i hesapla", command=kdv_hesapla).pack(side="left")

        ttk.Label(cerceve, text="PDF Dosyası").grid(row=6, column=0, sticky="w", pady=4)
        ttk.Entry(cerceve, textvariable=v["pdf"], width=55).grid(row=6, column=1, sticky="w", padx=(10, 0), pady=4)
        pdf_dugmeleri = ttk.Frame(cerceve)
        pdf_dugmeleri.grid(row=6, column=2, padx=6)

        # ---- PDF'den otomatik doldurma
        def pdf_oku(sessiz=False):
            yol = v["pdf"].get().strip()
            if not yol or not os.path.exists(yol):
                if not sessiz:
                    messagebox.showwarning(UYGULAMA_ADI, "Önce okunacak PDF dosyasını seçin (Gözat).", parent=pen)
                return
            sayfa_spec = v["sayfa"].get().strip()
            v_durum.set("PDF okunuyor... (resim olarak taranmış çok sayfalı belgelerde 30-60 saniye sürebilir)")
            btn_oku.state(["disabled"])

            def bitti(sonuc, hata):
                try:
                    if not pen.winfo_exists():
                        return
                    btn_oku.state(["!disabled"])
                except tk.TclError:
                    return
                if hata or not sonuc:
                    v_durum.set("PDF okunurken hata oluştu: %s. Bilgileri elle girebilirsiniz." % (hata,))
                    return
                son_metin["metin"] = sonuc["metin"]
                if sonuc.get("sayfa") and (sonuc.get("sayfa_sayisi") or 0) > 1:
                    v["sayfa"].set(sonuc["sayfa"])      # bulunan fatura sayfası kutuda görünür (elle değiştirilebilir)
                a = sonuc["alanlar"]
                vkn = a.get("vkn", "")
                if vkn:
                    v["vkn"].set(vkn)
                    bilinen = None
                    try:
                        bilinen = db.vkn_firma(vkn)
                    except VeritabaniHatasi:
                        pass
                    if bilinen:
                        a["firma_adi"] = bilinen     # aynı firma hep aynı yazımla gelsin
                bulunan = 0
                eslesme = [("fatura_tarihi", "tarih", tarih_yaz), ("fatura_no", "no", str),
                           ("firma_adi", "firma", str),
                           ("tutar_haric", "haric", tutar_yaz), ("tutar_dahil", "dahil", tutar_yaz)]
                for alan, k, bicim in eslesme:
                    if alan in a and a[alan] not in (None, ""):
                        v[k].set(bicim(a[alan]))
                        girisler[k].configure(style="Oto.TCombobox" if k == "firma" else "Oto.TEntry")
                        bulunan += 1
                v_durum.set((durum_metni(sonuc, bulunan) + "  " + sonuc.get("sayfa_bilgisi", "")).strip())
            arka_planda(lambda: pdf_bilgi_cikar(yol, ayar, sayfa_spec or None), bitti)

        def gozat():
            ilk = ""
            if v["pdf"].get() and os.path.isdir(os.path.dirname(v["pdf"].get())):
                ilk = os.path.dirname(v["pdf"].get())
            elif ayar["tarama_klasoru"] and os.path.isdir(ayar["tarama_klasoru"]):
                ilk = ayar["tarama_klasoru"]
            yol = filedialog.askopenfilename(parent=pen, title="Fatura PDF'ini seçin", initialdir=ilk or None,
                                             filetypes=[("PDF dosyaları", "*.pdf"), ("Tüm dosyalar", "*.*")])
            if yol:
                v["pdf"].set(os.path.normpath(yol))
                v["sayfa"].set("")                # yeni dosyada sayfa otomatik bulunur
                if not fid:           # yeni kayıtta seçer seçmez otomatik oku
                    pdf_oku(sessiz=True)

        ttk.Button(pdf_dugmeleri, text="Gözat...", command=gozat).pack(side="left")
        btn_oku = ttk.Button(pdf_dugmeleri, text="PDF'den Oku", command=pdf_oku)
        btn_oku.pack(side="left", padx=(4, 0))

        sayfa_cerceve = ttk.Frame(cerceve)
        sayfa_cerceve.grid(row=7, column=1, columnspan=2, sticky="w", padx=(10, 0))
        ttk.Label(sayfa_cerceve, text="(Dosya kopyalanmaz, yalnızca konumu kaydedilir.)",
                  style="Not.TLabel").pack(side="left")
        ttk.Label(sayfa_cerceve, text="     Fatura sayfası:").pack(side="left")
        ttk.Entry(sayfa_cerceve, textvariable=v["sayfa"], width=7).pack(side="left", padx=4)
        ttk.Label(sayfa_cerceve, text="(boş = otomatik bul; örn. 7 ya da 7-8)", style="Not.TLabel").pack(side="left")
        ttk.Label(cerceve, textvariable=v_durum, style="Not.TLabel", wraplength=560, justify="left").grid(
            row=8, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Button(cerceve, text="Okunan metni göster",
                   command=lambda: metin_goster("PDF'den okunan ham metin", son_metin["metin"])).grid(
            row=9, column=0, columnspan=2, sticky="w", pady=(4, 0))

        def kaydet():
            try:
                if not v["firma"].get().strip():
                    raise ValueError("Firma adı boş olamaz.")
                if not v["no"].get().strip():
                    raise ValueError("Fatura numarası boş olamaz.")
                veri = {
                    "fatura_tarihi": tarih_ayristir(v["tarih"].get()),
                    "fatura_no": v["no"].get().strip(),
                    "firma_adi": v["firma"].get().strip(),
                    "yapilan_is": v["is"].get().strip(),
                    "tutar_haric": tutar_ayristir(v["haric"].get()),
                    "tutar_dahil": tutar_ayristir(v["dahil"].get()),
                    "pdf_yolu": v["pdf"].get().strip(),
                    "vkn": v["vkn"].get().strip(),
                }
                if veri["tutar_haric"] is None and veri["tutar_dahil"] is None:
                    raise ValueError("En az bir tutar girin.")
                if (veri["tutar_haric"] is not None and veri["tutar_dahil"] is not None
                        and veri["tutar_dahil"] < veri["tutar_haric"]):
                    raise ValueError("KDV dahil tutar, KDV hariç tutardan küçük olamaz.")
                if veri["pdf_yolu"] and not os.path.exists(veri["pdf_yolu"]):
                    if not messagebox.askyesno(UYGULAMA_ADI, "Seçilen PDF dosyası bulunamıyor.\n"
                                               "Yine de bu yolla kaydedilsin mi?", parent=pen):
                        return
                if fid:
                    db.guncelle(fid, veri, kullanici)
                else:
                    db.ekle(veri, kullanici)
            except ValueError as ex:
                messagebox.showwarning(UYGULAMA_ADI, str(ex), parent=pen)
                return
            except VeritabaniHatasi as ex:
                messagebox.showerror(UYGULAMA_ADI, str(ex), parent=pen)
                return
            pen.destroy()
            yenile()

        dugme = ttk.Frame(cerceve)
        dugme.grid(row=10, column=0, columnspan=3, pady=(14, 0), sticky="e")
        ttk.Button(dugme, text="Kaydet", command=kaydet).pack(side="left", padx=4)
        ttk.Button(dugme, text="Vazgeç", command=pen.destroy).pack(side="left", padx=4)
        if mevcut:
            ttk.Label(cerceve, style="Not.TLabel", text="Kaydeden: %s  (%s)%s" % (
                mevcut["kaydeden"] or "-", mevcut["kayit_tarihi"] or "",
                ("   |   Son düzenleyen: %s (%s)" % (mevcut["son_duzenleyen"], mevcut["duzenleme_tarihi"]))
                if mevcut["son_duzenleyen"] else "")).grid(row=11, column=0, columnspan=3, sticky="w", pady=(10, 0))
        girisler["tarih"].focus_set()
        pen.bind("<Escape>", lambda e: pen.destroy())

    # ---- PDF okuma testi (menüden) ---------------------------------------------
    def okuma_testi():
        yol = filedialog.askopenfilename(parent=kok, title="Okuma testi için bir fatura PDF'i seçin",
                                         initialdir=ayar["tarama_klasoru"] or None,
                                         filetypes=[("PDF dosyaları", "*.pdf")])
        if not yol:
            return
        bilgi = yeni_pencere("PDF okunuyor...")
        ttk.Label(bilgi, text="Okunuyor, lütfen bekleyin...", padding=20).pack()

        def bitti(sonuc, hata):
            try:
                bilgi.destroy()
            except tk.TclError:
                pass
            if hata or not sonuc:
                hata_goster("Okuma sırasında hata: %s" % hata)
                return
            a = sonuc["alanlar"]
            satirlar = ["YÖNTEM: " + {"metin": "PDF yazı katmanı", "ocr": "Windows OCR"}.get(sonuc["yontem"], "okunamadı")]
            if sonuc.get("sayfa_bilgisi"):
                satirlar.append("SAYFA: " + sonuc["sayfa_bilgisi"])
            if sonuc["hata"]:
                satirlar.append("UYARI: " + sonuc["hata"])
            satirlar += ["", "BULUNAN ALANLAR:"]
            for ad, anahtar_, bicim in [("Fatura Tarihi", "fatura_tarihi", tarih_yaz), ("Fatura No", "fatura_no", str),
                                        ("Firma Adı", "firma_adi", str), ("VKN", "vkn", str),
                                        ("Tutar (KDV Hariç)", "tutar_haric", tutar_yaz),
                                        ("Tutar (KDV Dahil)", "tutar_dahil", tutar_yaz)]:
                satirlar.append("  %-18s %s" % (ad, bicim(a[anahtar_]) if anahtar_ in a else "- bulunamadı -"))
            satirlar += ["", "---------- OKUNAN HAM METİN ----------", sonuc["metin"] or "(metin yok)"]
            metin_goster("PDF Okuma Testi", "\n".join(satirlar))
        arka_planda(lambda: pdf_bilgi_cikar(yol, ayar), bitti)

    # ---- aylık özet penceresi ------------------------------------------------
    def ozet_penceresi():
        tum = db.hepsi()
        yillar = sorted({int(r["fatura_tarihi"][:4]) for r in tum} | {date.today().year}, reverse=True)
        pen = yeni_pencere("Aylık Özet")
        v_y = tk.StringVar(value=v_yil.get() if v_yil.get() != "Tümü" else str(yillar[0]))
        ust_ = ttk.Frame(pen, padding=8)
        ust_.pack(fill="x")
        ttk.Label(ust_, text="Yıl:").pack(side="left")
        cb = ttk.Combobox(ust_, textvariable=v_y, values=[str(y) for y in yillar], width=8, state="readonly")
        cb.pack(side="left", padx=6)
        kolonlar = [("ay", "Ay", 110, "w"), ("adet", "Fatura Adedi", 100, "center"),
                    ("haric", "KDV Hariç Toplam", 150, "e"), ("kdv", "KDV Tutarı", 130, "e"),
                    ("dahil", "KDV Dahil Toplam", 150, "e")]
        agac = ttk.Treeview(pen, columns=[k[0] for k in kolonlar], show="headings", height=14)
        for ad, baslik, gen, hiza in kolonlar:
            agac.heading(ad, text=baslik)
            agac.column(ad, width=gen, anchor=hiza)
        agac.tag_configure("toplam", font=("Segoe UI", 10, "bold"))
        agac.pack(fill="both", expand=True, padx=8, pady=4)
        veri_ = {"ozet": []}

        def doldur(*_):
            agac.delete(*agac.get_children())
            try:
                veri_["ozet"] = aylik_ozet(db.hepsi(), int(v_y.get()))
            except VeritabaniHatasi as e:
                hata_goster(e)
                return
            for x in veri_["ozet"]:
                agac.insert("", "end", tags=("toplam",) if x["ay"] == "TOPLAM" else (), values=(
                    x["ay"], x["adet"], tutar_yaz(x["haric"], True), tutar_yaz(x["dahil"] - x["haric"], True),
                    tutar_yaz(x["dahil"], True)))

        def aktar():
            tipler = [("Excel çalışma kitabı", "*.xlsx"), ("CSV (Excel'de açılır)", "*.csv")] if openpyxl else \
                     [("CSV (Excel'de açılır)", "*.csv")]
            yol = filedialog.asksaveasfilename(parent=pen, title="Aylık özeti kaydet", defaultextension=tipler[0][1],
                                               initialfile="Aylik_Ozet_%s" % v_y.get(), filetypes=tipler)
            if not yol:
                return
            try:
                if yol.lower().endswith(".csv") or openpyxl is None:
                    ozet_csv_yaz(yol, veri_["ozet"])
                else:
                    ozet_xlsx_yaz(yol, veri_["ozet"], v_y.get())
                messagebox.showinfo(UYGULAMA_ADI, "Aylık özet kaydedildi:\n%s" % yol, parent=pen)
            except PermissionError:
                messagebox.showerror(UYGULAMA_ADI, "Dosya yazılamadı. Excel'de açık olabilir; kapatıp tekrar deneyin.", parent=pen)
            except OSError as e:
                hata_goster(e)
        alt_ = ttk.Frame(pen, padding=8)
        alt_.pack(fill="x")
        ttk.Button(alt_, text="Excel'e Aktar", command=aktar).pack(side="left")
        ttk.Button(alt_, text="Kapat", command=pen.destroy).pack(side="right")
        cb.bind("<<ComboboxSelected>>", doldur)
        pen.bind("<Escape>", lambda e: pen.destroy())
        doldur()

    # ---- düğme işlevleri -----------------------------------------------------
    def duzenle(*_):
        fid = secili_id()
        if fid:
            form_ac(fid)

    def sil():
        fid = secili_id()
        if not fid:
            return
        k = db.getir(fid)
        if not k:
            yenile()
            return
        if messagebox.askyesno(UYGULAMA_ADI, "Bu fatura kalıcı olarak silinecek:\n\n%s  /  %s\n\nEmin misiniz?\n"
                               "(PDF dosyası silinmez, yalnızca kayıt silinir.)" % (k["firma_adi"], k["fatura_no"]),
                               icon="warning", parent=kok):
            try:
                db.sil(fid)
            except VeritabaniHatasi as e:
                hata_goster(e)
            yenile()

    # ---- fatura evrakları (fatura PDF'i + talep, yazı vb. ekler) ------------------
    EK_TURLERI = ["Talep", "Yazı", "Teklif", "İrsaliye", "Sözleşme", "Fatura Ön İzleme", "Diğer"]

    def evrak_penceresi(fid=None):
        fid = fid or secili_id()
        if not fid:
            return
        fatura = db.getir(fid)
        if not fatura:
            messagebox.showwarning(UYGULAMA_ADI, "Kayıt bulunamadı (başkası silmiş olabilir).", parent=kok)
            yenile()
            return
        pen = yeni_pencere("Evraklar  -  %s  /  %s" % (fatura["firma_adi"], fatura["fatura_no"]))
        pen.geometry("860x420")
        pen.grab_set()
        satirlar = {}                       # iid -> (yol, ek_id veya None)

        ttk.Label(pen, text="%s   |   %s   |   %s" % (fatura["firma_adi"], fatura["fatura_no"],
                                                       tarih_yaz(fatura["fatura_tarihi"])),
                  padding=(10, 8, 10, 2)).pack(anchor="w")
        orta2 = ttk.Frame(pen, padding=(10, 4))
        orta2.pack(fill="both", expand=True)
        kol = [("tur", "Tür", 120), ("dosya", "Dosya", 300), ("aciklama", "Açıklama", 190),
               ("ekleyen", "Ekleyen", 100), ("tarih", "Tarih", 90), ("durum", "Durum", 70)]
        agac = ttk.Treeview(orta2, columns=[k[0] for k in kol], show="headings", selectmode="browse")
        for ad, baslik, gen in kol:
            agac.heading(ad, text=baslik)
            agac.column(ad, width=gen, anchor="w", stretch=(ad in ("dosya", "aciklama")))
        sb3 = ttk.Scrollbar(orta2, orient="vertical", command=agac.yview)
        agac.configure(yscrollcommand=sb3.set)
        agac.pack(side="left", fill="both", expand=True)
        sb3.pack(side="right", fill="y")

        def doldur():
            agac.delete(*agac.get_children())
            satirlar.clear()
            kayitlar = [("f", "Fatura", fatura["pdf_yolu"], "", "", "", None)]
            try:
                for e in db.ekler(fid):
                    kayitlar.append(("e%d" % e["id"], e["tur"] or "", e["dosya_yolu"], e["aciklama"] or "",
                                     e["ekleyen"] or "", tarih_yaz((e["ekleme_tarihi"] or "")[:10]), e["id"]))
            except VeritabaniHatasi as ex:
                hata_goster(ex)
            for iid, tur, yol, acik, ekleyen, tarih, ek_id in kayitlar:
                if iid == "f" and not yol:
                    continue
                satirlar[iid] = (yol, ek_id)
                var = os.path.exists(yol)
                agac.insert("", "end", iid=iid, values=(tur, os.path.basename(yol) or yol, acik, ekleyen, tarih,
                                                        "var" if var else "YOK"))

        def secili():
            s = agac.selection()
            if not s:
                messagebox.showinfo(UYGULAMA_ADI, "Önce listeden bir evrak seçin.", parent=pen)
                return None
            return s[0]

        def ac(*_):
            iid = secili()
            if iid:
                try:
                    pdf_ac(satirlar[iid][0])
                except (OSError, FileNotFoundError) as ex:
                    messagebox.showerror(UYGULAMA_ADI, str(ex), parent=pen)

        def tumunu_ac():
            if not satirlar:
                return
            if len(satirlar) > 5 and not messagebox.askyesno(
                    UYGULAMA_ADI, "%d dosya birden açılacak. Devam edilsin mi?" % len(satirlar), parent=pen):
                return
            eksik = []
            for yol, _ in satirlar.values():
                try:
                    pdf_ac(yol)
                except (OSError, FileNotFoundError):
                    eksik.append(yol)
            if eksik:
                messagebox.showwarning(UYGULAMA_ADI, "Şu dosyalar açılamadı:\n\n" + "\n".join(eksik), parent=pen)

        def klasorde():
            iid = secili()
            if iid:
                try:
                    klasorde_goster(satirlar[iid][0])
                except (OSError, FileNotFoundError) as ex:
                    messagebox.showerror(UYGULAMA_ADI, str(ex), parent=pen)

        v_tur = tk.StringVar(value=EK_TURLERI[0])
        v_acik = tk.StringVar()

        def dosya_ekle():
            klasor = os.path.dirname(fatura["pdf_yolu"]) if fatura["pdf_yolu"] else ""
            ilk = klasor if os.path.isdir(klasor) else (ayar["tarama_klasoru"] or None)
            yollar = filedialog.askopenfilenames(parent=pen, title="Faturaya eklenecek evrakları seçin",
                                                 initialdir=ilk, filetypes=[("Tüm dosyalar", "*.*")])
            if not yollar:
                return
            hatalar = []
            for y in yollar:
                try:
                    db.ek_ekle(fid, v_tur.get().strip() or "Diğer", os.path.normpath(y), v_acik.get().strip(), kullanici)
                except VeritabaniHatasi as ex:
                    hatalar.append(str(ex))
            v_acik.set("")
            doldur()
            yenile()
            if hatalar:
                messagebox.showwarning(UYGULAMA_ADI, "\n\n".join(hatalar), parent=pen)

        def kaldir():
            iid = secili()
            if not iid:
                return
            if satirlar[iid][1] is None:
                messagebox.showinfo(UYGULAMA_ADI, "Fatura PDF'i buradan kaldırılamaz.\n"
                                    "Değiştirmek için faturayı düzenleyin.", parent=pen)
                return
            if messagebox.askyesno(UYGULAMA_ADI, "Bu evrak faturadan çıkarılsın mı?\n(Dosyanın kendisi silinmez.)",
                                   parent=pen):
                try:
                    db.ek_sil(satirlar[iid][1])
                except VeritabaniHatasi as ex:
                    hata_goster(ex)
                doldur()
                yenile()

        ekle_bar = ttk.Frame(pen, padding=(10, 4))
        ekle_bar.pack(fill="x")
        ttk.Label(ekle_bar, text="Tür:").pack(side="left")
        ttk.Combobox(ekle_bar, textvariable=v_tur, values=EK_TURLERI, width=16).pack(side="left", padx=(4, 10))
        ttk.Label(ekle_bar, text="Açıklama:").pack(side="left")
        ttk.Entry(ekle_bar, textvariable=v_acik, width=30).pack(side="left", padx=(4, 10))
        ttk.Button(ekle_bar, text="Dosya Ekle...", command=dosya_ekle).pack(side="left")
        alt2 = ttk.Frame(pen, padding=(10, 4, 10, 10))
        alt2.pack(fill="x")
        for metin, komut in (("Aç", ac), ("Tümünü Aç", tumunu_ac), ("Klasörde Göster", klasorde),
                             ("Evrakı Kaldır", kaldir)):
            ttk.Button(alt2, text=metin, command=komut).pack(side="left", padx=3)
        ttk.Button(alt2, text="Kapat", command=pen.destroy).pack(side="right")
        ttk.Label(pen, text="Dosyalar kopyalanmaz, yalnızca konumları kaydedilir. Bu evraklar PDF okumaya katılmaz.",
                  style="Not.TLabel", padding=(10, 0, 10, 8)).pack(anchor="w")
        agac.bind("<Double-1>", ac)
        agac.bind("<Return>", ac)
        pen.bind("<Escape>", lambda e: pen.destroy())
        doldur()
        if agac.get_children():
            agac.selection_set(agac.get_children()[0])

    def pdf_ac_btn(*_):
        fid = secili_id()
        if not fid:
            return
        k = db.getir(fid)
        try:
            if not k or not k["pdf_yolu"]:
                raise FileNotFoundError("Bu faturaya bağlı bir PDF yok.")
            pdf_ac(k["pdf_yolu"])
        except (OSError, FileNotFoundError) as e:
            hata_goster(e)

    def klasor_btn():
        fid = secili_id()
        if not fid:
            return
        k = db.getir(fid)
        try:
            if not k or not k["pdf_yolu"]:
                raise FileNotFoundError("Bu faturaya bağlı bir PDF yok.")
            klasorde_goster(k["pdf_yolu"])
        except (OSError, FileNotFoundError) as e:
            hata_goster(e)

    def donem_adi():
        ay = v_ay.get()
        yil = v_yil.get()
        return ((AY_ADLARI[ay - 1] + " ") if ay else "") + (yil if yil != "Tümü" else "Tüm Yıllar")

    def csv_btn():
        satirlar = durum.get("gorunen", [])
        if not satirlar:
            messagebox.showinfo(UYGULAMA_ADI, "Aktarılacak kayıt yok.", parent=kok)
            return
        donem = donem_adi()
        dosya_donem = donem.replace(" ", "_")
        tipler = [("Excel çalışma kitabı", "*.xlsx"), ("CSV (Excel'de açılır)", "*.csv")] if openpyxl else \
                 [("CSV (Excel'de açılır)", "*.csv")]
        yol = filedialog.asksaveasfilename(parent=kok, title="Excel'e aktar (%s)" % donem,
                                           defaultextension=tipler[0][1], filetypes=tipler,
                                           initialfile="Faturalar_%s" % dosya_donem)
        if not yol:
            return
        try:
            if yol.lower().endswith(".csv") or openpyxl is None:
                csv_yaz(yol, satirlar)
            else:
                xlsx_yaz(yol, satirlar, "Alış Faturaları - %s" % donem)
            if messagebox.askyesno(UYGULAMA_ADI, "%d kayıt aktarıldı:\n%s\n\nDosya şimdi açılsın mı?" % (len(satirlar), yol),
                                   parent=kok):
                pdf_ac(yol)
        except PermissionError:
            messagebox.showerror(UYGULAMA_ADI, "Dosya yazılamadı. Excel'de açık olabilir; kapatıp tekrar deneyin.", parent=kok)
        except OSError as e:
            hata_goster(e)

    # ---- menü ----------------------------------------------------------------
    menu = tk.Menu(kok)
    m_dosya = tk.Menu(menu, tearoff=0)
    m_dosya.add_command(label="Excel'e Aktar...", command=csv_btn)
    m_dosya.add_command(label="Aylık Özet...", command=ozet_penceresi)
    m_dosya.add_command(label="PDF Okuma Testi...", command=okuma_testi)
    m_dosya.add_separator()
    m_dosya.add_command(label="Çıkış", command=kok.destroy)
    m_gor = tk.Menu(menu, tearoff=0)
    m_tema = tk.Menu(m_gor, tearoff=0)
    for ad in TEMALAR:
        m_tema.add_radiobutton(label=ad, variable=v_tema, value=ad, command=lambda: (tema_uygula(v_tema.get(), True), yenile()))
    m_gor.add_cascade(label="Tema", menu=m_tema)
    menu.add_cascade(label="Dosya", menu=m_dosya)
    menu.add_cascade(label="Görünüm", menu=m_gor)
    kok.configure(menu=menu)

    # ---- bağlantılar ---------------------------------------------------------
    btn_ekle.configure(command=lambda: form_ac())
    btn_duzenle.configure(command=duzenle)
    btn_sil.configure(command=sil)
    btn_pdf.configure(command=pdf_ac_btn)
    btn_klasor.configure(command=klasor_btn)
    btn_evrak.configure(command=lambda: evrak_penceresi())
    btn_ozet.configure(command=ozet_penceresi)
    btn_csv.configure(command=csv_btn)
    btn_temizle.configure(command=temizle)
    tree.bind("<Double-1>", lambda e: evrak_penceresi())
    tree.bind("<Return>", duzenle)
    tree.bind("<Delete>", lambda e: sil())
    v_arama.trace_add("write", yenile)
    v_bas.trace_add("write", yenile)
    v_bit.trace_add("write", yenile)
    v_ay.trace_add("write", yenile)
    cb_firma.bind("<<ComboboxSelected>>", yenile)
    cb_yil.bind("<<ComboboxSelected>>", yenile)

    # açılışta: verisi olan en yeni yılı göster
    try:
        ilk = db.hepsi()
        v_yil.set(str(max(int(r["fatura_tarihi"][:4]) for r in ilk)) if ilk else str(date.today().year))
    except (VeritabaniHatasi, ValueError):
        v_yil.set(str(date.today().year))
    kok.title("%s   -   %s" % (UYGULAMA_ADI, kullanici))
    yenile()
    e_ara.focus_set()
    kok.mainloop()


if __name__ == "__main__":
    arayuzu_calistir()
