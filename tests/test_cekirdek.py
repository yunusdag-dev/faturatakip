# -*- coding: utf-8 -*-
"""Çekirdek mantık ve güvenlik testleri (arayüz gerektirmez; gerçek veri kullanmaz).

Çalıştırma (proje klasöründe):   python -m unittest discover -s tests -v
"""
import os
import sqlite3
import sys
import tempfile
import unittest
import zlib

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fatura_takip as f

_TUR = f._PBKDF2_TUR


def setUpModule():
    f._PBKDF2_TUR = 1000            # testler hızlı olsun; gerçek tur sayısı üretimde 600.000


def tearDownModule():
    f._PBKDF2_TUR = _TUR


def yeni_db():
    klasor = tempfile.mkdtemp()
    return f.FaturaDB(os.path.join(klasor, "t.db")), klasor


def fatura(i, **k):
    v = dict(id=i, fatura_tarihi="2026-10-%02d" % (1 + i % 28), fatura_no="F-%d" % i, firma_adi="ALE YÜK A.Ş.",
             yapilan_is="Vinç", tutar_haric=100000 + i, kategori="Genel Alım", talep_no="4035%03d" % i,
             talep_tarihi="2026-09-29", talep_eden_birim="Aliağa İletim Şebekeleri İşletme ve Bakım Müdürlüğü",
             dt_no="26DT500%04d" % i, pdf_yolu="", ek_sayisi=None)
    v.update(k)
    return v


class TutarVeTarih(unittest.TestCase):
    def test_tutar_bicimleri(self):
        for metin, kurus in (("1.234,56", 123456), ("1234,56", 123456), ("1234.56", 123456), ("1.234", 123400),
                             ("1.234,56 TL", 123456), ("0,5", 50), ("", None)):
            self.assertEqual(f.tutar_ayristir(metin), kurus, metin)
        for kotu in ("abc", "-5"):
            with self.assertRaises(ValueError):
                f.tutar_ayristir(kotu)

    def test_tarih(self):
        self.assertEqual(f.tarih_ayristir("04.06.2026"), "2026-06-04")
        with self.assertRaises(ValueError):
            f.tarih_ayristir("32.13.2026")


class Hesaplar(unittest.TestCase):
    def setUp(self):
        self.db, self.klasor = yeni_db()

    def test_parola_en_az_uzunluk(self):
        with self.assertRaises(f.VeritabaniHatasi):
            self.db.ilk_yonetici_olustur("Y", "1234567")
        self.assertEqual(self.db.kullanici_sayisi(), 0)

    def test_ilk_yonetici_yalniz_bir_kez(self):
        self.db.ilk_yonetici_olustur("Yonetici", "gizli1234")
        with self.assertRaises(f.VeritabaniHatasi):
            self.db.ilk_yonetici_olustur("Baska", "gizli1234")
        self.assertEqual(self.db.kullanici_sayisi(), 1)

    def test_giris_ve_turkce_ad_eslemesi(self):
        self.db.ilk_yonetici_olustur("Yönetici", "gizli1234")
        self.assertTrue(self.db.giris_dene("YÖNETİCİ", "gizli1234")[0])
        self.assertIsNone(self.db.giris_dene("Yönetici", "gizli1235")[0])
        self.assertIsNone(self.db.giris_dene("yok", "gizli1234")[0])

    def test_sifre_duz_metin_saklanmaz(self):
        self.db.ilk_yonetici_olustur("Y", "gizli1234")
        kayit = sqlite3.connect(self.db.yol).execute("SELECT sifre FROM kullanicilar").fetchone()[0]
        self.assertTrue(kayit.startswith("pbkdf2$"))
        self.assertNotIn("gizli1234", kayit)

    def test_eski_hash_giriste_yenilenir(self):
        self.db.ilk_yonetici_olustur("Y", "gizli1234")
        eski = f._PBKDF2_TUR
        f._PBKDF2_TUR = 500                                  # "eski" daha az turlu hash üret
        eski_hash = f.sifre_hashle("gizli1234")
        f._PBKDF2_TUR = eski
        con = sqlite3.connect(self.db.yol); con.execute("UPDATE kullanicilar SET sifre=?", (eski_hash,)); con.commit(); con.close()
        self.assertTrue(self.db.giris_dene("y", "gizli1234")[0])
        yeni = sqlite3.connect(self.db.yol).execute("SELECT sifre FROM kullanicilar").fetchone()[0]
        self.assertEqual(int(yeni.split("$")[1]), eski)      # artık güncel tur sayısıyla
        self.assertTrue(self.db.giris_dene("y", "gizli1234")[0])

    def test_kullanici_adi_sinirlari(self):
        self.db.ilk_yonetici_olustur("Y", "gizli1234")
        for kotu in ("", "a" * 41, "ad\x00soyad", "ad\tsoyad"):
            with self.assertRaises(f.VeritabaniHatasi):
                self.db.kullanici_ekle(kotu, "gecici1234")
        with self.assertRaises(f.MukerrerKayit):
            self.db.kullanici_ekle("y", "gecici1234")

    def test_oturumlar_sifre_degisince_ve_pasifte_duser(self):
        self.db.ilk_yonetici_olustur("Y", "gizli1234")
        self.db.kullanici_ekle("Ayse", "gecici1234", sifre_degismeli=False)
        a = self.db.giris_dene("ayse", "gecici1234")[0]["id"]
        t1, t2 = self.db.oturum_olustur(a), self.db.oturum_olustur(a)
        self.assertTrue(self.db.oturum_dogrula(t1))
        self.db.sifre_degistir(a, "yeni-sifre-9", eski="gecici1234")
        self.assertIsNone(self.db.oturum_dogrula(t1))
        self.assertIsNone(self.db.oturum_dogrula(t2))
        t3 = self.db.oturum_olustur(a)
        self.db.sifre_sifirla(a, "sifir-1234")
        self.assertIsNone(self.db.oturum_dogrula(t3))
        t4 = self.db.oturum_olustur(a)
        self.db.kullanici_durumu(a, aktif=False)
        self.assertIsNone(self.db.oturum_dogrula(t4))
        self.assertIn("pasif", self.db.giris_dene("ayse", "sifir-1234")[1])

    def test_oturum_suresi_dolar(self):
        self.db.ilk_yonetici_olustur("Y", "gizli1234")
        kid = self.db.giris_dene("y", "gizli1234")[0]["id"]
        t = self.db.oturum_olustur(kid)
        con = sqlite3.connect(self.db.yol); con.execute("UPDATE oturumlar SET olusturma_tarihi='2000-01-01 00:00:00'"); con.commit(); con.close()
        self.assertIsNone(self.db.oturum_dogrula(t))

    def test_son_aktif_yonetici_korunur(self):
        self.db.ilk_yonetici_olustur("Y", "gizli1234")
        y = self.db.giris_dene("y", "gizli1234")[0]["id"]
        with self.assertRaises(f.VeritabaniHatasi):
            self.db.kullanici_durumu(y, aktif=False)
        with self.assertRaises(f.VeritabaniHatasi):
            self.db.kullanici_durumu(y, yonetici=False)
        self.assertEqual(self.db.giris_dene("y", "gizli1234")[0]["aktif"], 1)     # işlem geri alındı


class Arama(unittest.TestCase):
    def test_fatura_ara(self):
        tum = [fatura(1, firma_adi="ALE YÜK A.Ş.", fatura_no="ALE-001", tutar_haric=3250000),
               fatura(2, firma_adi="Özer Doğrayan", fatura_no="GIB2026000000015", tutar_haric=2000000),
               fatura(3, firma_adi="ÇİMENTO LTD.", fatura_no="ÇİM 77", tutar_haric=None)]
        ids = lambda l: sorted(x["id"] for x in l)
        self.assertEqual(ids(f.fatura_ara(tum)), [1, 2, 3])
        self.assertEqual(ids(f.fatura_ara(tum, firma="ÖZER")), [2])
        self.assertEqual(ids(f.fatura_ara(tum, firma="çimento")), [3])
        self.assertEqual(ids(f.fatura_ara(tum, fatura_no="çim77")), [3])               # boşluk önemsiz
        self.assertEqual(ids(f.fatura_ara(tum, en_az=3250000, en_cok=3250000)), [1])
        self.assertEqual(ids(f.fatura_ara(tum, en_az=1)), [1, 2])                      # tutarsız kayıt sınırla çıkmaz
        self.assertEqual(ids(f.fatura_ara(tum, firma="ale", en_az=9999999)), [])


class PdfKurallari(unittest.TestCase):
    ust = "T.C.\nTÜRKİYE ELEKTRİK İLETİM ANONİM ŞİRKETİ GENEL MÜDÜRLÜĞÜ\n3.Bölge Müdürlüğü (İzmir)\nXYZ Birimi\nSayı : E-1-3570050\n"

    def test_birim_dort_belge_turu(self):
        for baslik in ("TALEP FORMU", "Malzeme Talep Formu", "Hizmet   İhtiyaç  Formu", "Malzeme İhtiyaç Formu"):
            self.assertEqual(f.birim_coz(self.ust + baslik), ("XYZ Birimi", ""), baslik)

    def test_birim_yanlis_belge_okunmaz(self):
        deger, uyari = f.birim_coz(self.ust + "Teklif Mektubu")
        self.assertIsNone(deger)
        self.assertTrue(uyari)

    def test_talep_no_tarih_sadece_talep_formundan(self):
        metin = "23.01.2026\nSayı : E-60994212-934.01.99-3570050\nMal Alımı Talep Formu"
        self.assertEqual(f.talep_coz(metin)[:2], ("3570050", "2026-01-23"))
        self.assertEqual(f.talep_coz(metin.replace("Mal Alımı Talep Formu", "Hizmet İhtiyaç Formu"))[:2], (None, None))

    def test_is_adinin_sonundaki_parantez(self):
        a = f.is_adi_parantez_at
        self.assertEqual(a("Sepetli Vinç Kiralama (1 Adet) (Gölmarmara)"), "Sepetli Vinç Kiralama")
        self.assertEqual(a("Vinç (Sepetli) Kiralama"), "Vinç (Sepetli) Kiralama")
        self.assertEqual(a("(Sadece parantez)"), "(Sadece parantez)")


class Guvenlik(unittest.TestCase):
    def test_guvenli_uzantilar(self):
        for ok in ("a.pdf", "A.PDF", "x.docx", "r.jpeg", "C:\\a b\\c.xlsx"):
            self.assertTrue(f.guvenli_acilir(ok), ok)
        for kotu in ("a.exe", "a.bat", "a.cmd", "a.lnk", "a.js", "a.vbs", "a.ps1", "a.msi", "a.scr", "a.xlsm", "a.docm",
                     "a.pdf.exe", "uzantisiz", "", None):
            self.assertFalse(f.guvenli_acilir(kotu), kotu)

    def test_pdf_ac_calistirilabilir_dosyayi_acmaz(self):
        d = tempfile.mkdtemp()
        yol = os.path.join(d, "zararli.exe")
        open(yol, "wb").write(b"MZ")
        with self.assertRaises(f.GuvensizDosyaTuru):
            f.pdf_ac(yol)
        self.assertTrue(issubclass(f.GuvensizDosyaTuru, OSError))       # arayüzdeki `except OSError` yakalar

    def test_siskinlik_bombasi_sinirlanir(self):
        eski = f._AKIS_EN_COK
        f._AKIS_EN_COK = 1024 * 1024
        try:
            bomba = zlib.compress(b"\0" * (8 * 1024 * 1024))
            self.assertIsNone(f._akis_coz(b"<< /Filter /FlateDecode >>", bomba))
            normal = zlib.compress(b"merhaba")
            self.assertEqual(f._akis_coz(b"<< /Filter /FlateDecode >>", normal), b"merhaba")
        finally:
            f._AKIS_EN_COK = eski

    def test_bozuk_config_ezilmez(self):
        d = tempfile.mkdtemp()
        yol = os.path.join(d, "config.ini")
        icerik = "[ayarlar\nveritabani = benim\\ozel.db\n"          # bozuk bölüm başlığı
        open(yol, "w", encoding="utf-8").write(icerik)
        ayar = f.ayarlari_oku(yol)
        self.assertTrue(ayar["veritabani"].endswith("fatura.db"))     # varsayılanla devam eder, çökmez
        self.assertEqual(open(yol, encoding="utf-8").read(), icerik)   # kullanıcının dosyası ezilmedi


@unittest.skipIf(f.openpyxl is None, "openpyxl kurulu değil")
class ExcelCiktisi(unittest.TestCase):
    def test_formul_enjeksiyonu_ve_baglanti(self):
        import openpyxl
        d = tempfile.mkdtemp()
        pdf = os.path.join(d, "f.pdf"); open(pdf, "wb").write(b"%PDF")
        exe = os.path.join(d, "x.exe"); open(exe, "wb").write(b"MZ")
        x = os.path.join(d, "a.xlsx")
        f.xlsx_yaz(x, [fatura(1, firma_adi='=HYPERLINK("http://x")', pdf_yolu=pdf),
                       fatura(2, fatura_no="F-2", pdf_yolu=exe)], "Test")
        ws = openpyxl.load_workbook(x).active
        self.assertEqual(ws["C4"].data_type, "s")                      # '=...' formül değil metin
        self.assertTrue(ws["B4"].hyperlink)                            # güvenli türde fatura no bağlantılı
        self.assertIsNone(ws["B5"].hyperlink)                          # .exe'ye bağlantı verilmez

    def test_satinalma_sablonu_eslesmesi(self):
        import openpyxl
        x = os.path.join(tempfile.mkdtemp(), "s.xlsx")
        f.satinalma_xlsx_yaz(x, [fatura(2), fatura(1, firma_adi="=1+1")], "2026 EKİM HARCAMA LİSTESİ", "2026")
        ws = openpyxl.load_workbook(x).active
        self.assertEqual(ws["A1"].value, "2026 EKİM HARCAMA LİSTESİ")
        satir = [ws.cell(3, k).value for k in range(1, 12)]
        self.assertEqual(satir[0], 1)                                  # SIRA NO
        self.assertEqual(ws["D3"].value, ws["F3"].value)               # talep no = belgenet no
        self.assertEqual(ws["H3"].data_type, "s")                      # '=...' firma adı formül olmaz
        self.assertEqual(ws["K4"].value, 1000.02)                      # KDV hariç tutar (TL)


if __name__ == "__main__":
    unittest.main()
