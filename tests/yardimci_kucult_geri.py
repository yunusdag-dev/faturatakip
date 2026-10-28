# -*- coding: utf-8 -*-
"""Gerçek uygulamada küçült / geri yükle testi (yalnız Windows). test_kucult_geri.py bunu ayrı bir süreçte çalıştırır.

Neden: Tk'nin kilidi (grab) ve transient mekanizması, küçültülen uygulamanın görev çubuğundan / Alt+Tab ile geri
getirilmesini engelleyebiliyordu (modal pencere gizli kalır, kilit onda, uygulama 'kilitlenmiş' görünür). Küçültme
Windows'un sistem komutuyla (Win+D gibi), geri yükleme görev çubuğunun gönderdiği SC_RESTORE ile, ayrı bir iş
parçacığından yapılır."""
import ctypes
import faulthandler
import os
import sys
import tempfile
import threading
import time
import traceback

faulthandler.dump_traceback_later(170, exit=True)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import tkinter as tk
from tkinter import ttk
import tkinter.messagebox as mb
import fatura_takip as f

d = tempfile.mkdtemp()
os.environ["APPDATA"] = d
DBYOL = os.path.join(d, "test.db")
_asil = f.ayarlari_oku
f.ayarlari_oku = lambda *a, **k: dict(_asil(), veritabani=DBYOL, tema="Açık")
db = f.FaturaDB(DBYOL)
db.kullanici_ekle("testci", "gizli1234", yonetici=True, sifre_degismeli=False)
kid = db.giris_dene("testci", "gizli1234")[0]["id"]
f.oturum_yaz(DBYOL, db.oturum_olustur(kid))                   # giriş ekranı çıkmadan açılsın
db.ekle(dict(fatura_tarihi="2026-10-02", fatura_no="A1", firma_adi="Firma", yapilan_is="", tutar_haric=100, pdf_yolu="",
             kategori="Genel Alım", talep_no=None, talep_tarihi=None, dt_no=None, talep_eden_birim=None), "testci")
mb.showinfo = mb.showerror = mb.showwarning = lambda *a, **k: None

u = ctypes.windll.user32
u.GetParent.restype = ctypes.c_void_p
u.GetParent.argtypes = [ctypes.c_void_p]
for fn in (u.IsIconic, u.IsWindowVisible):
    fn.argtypes = [ctypes.c_void_p]
u.ShowWindow.argtypes = [ctypes.c_void_p, ctypes.c_int]
u.SendMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p, ctypes.c_void_p]
u.GetWindow.restype = ctypes.c_void_p
u.GetWindow.argtypes = [ctypes.c_void_p, ctypes.c_uint]
WM_SYSCOMMAND, SC_RESTORE = 0x0112, 0xF120
sonuc = {"hata": None, "satirlar": [], "basarisiz": []}


def hw(w):
    return u.GetParent(w.winfo_id()) or w.winfo_id()


def desc(w):
    for c in w.winfo_children():
        yield c
        yield from desc(c)


def toplevels(kok):
    return [w for w in kok.winfo_children() if isinstance(w, tk.Toplevel)]


def bul_dugme(w, metin):
    return [x for x in desc(w) if isinstance(x, ttk.Button) and str(x.cget("text")) == metin][0]


def kontrol(ad, kosul, ayrinti=""):
    sonuc["satirlar"].append(("OK " if kosul else "XX ") + ad + (("  [%s]" % ayrinti) if ayrinti else ""))
    if not kosul:
        sonuc["basarisiz"].append(ad)


def dongu(kok, pencereler, ad, grab_beklenen, tur=1):
    """Küçült -> sahipli pencereler gizlenmeli -> görev çubuğu komutuyla geri yükle -> hepsi görünür, kilit geri gelmeli."""
    h_ana = hw(kok)
    hp = [hw(p) for p in pencereler]
    kayit = {}

    def isci():
        for _ in range(tur):
            time.sleep(0.5)
            u.ShowWindow(h_ana, 6)                                       # sistem küçültmesi (Win+D / Masaüstünü göster)
            time.sleep(1.0)
            kayit.setdefault("kucuk", []).append((bool(u.IsIconic(h_ana)), [bool(u.IsWindowVisible(x)) for x in hp]))
            u.SendMessageW(h_ana, WM_SYSCOMMAND, SC_RESTORE, 0)           # görev çubuğu / Alt+Tab
            time.sleep(1.6)
            kayit.setdefault("geri", []).append((bool(u.IsIconic(h_ana)), [bool(u.IsWindowVisible(x)) for x in hp]))
        kayit["bitti"] = True
    threading.Thread(target=isci, daemon=True).start()
    bas = time.time()
    while not kayit.get("bitti") and time.time() - bas < 40:
        kok.update()
        time.sleep(0.03)
    kok.update()
    grab = str(kok.grab_current() or "-")
    for i in range(tur):
        k_k, k_p = kayit["kucuk"][i]
        g_k, g_p = kayit["geri"][i]
        kontrol("%s [döngü %d] küçültünce ana küçülür, sahipli pencereler gizlenir" % (ad, i + 1), k_k and not any(k_p),
                "ana küçük=%s pencereler=%s" % (k_k, k_p))
        kontrol("%s [döngü %d] geri yüklenince ana + pencereler geri gelir" % (ad, i + 1), (not g_k) and all(g_p),
                "ana küçük=%s pencereler=%s" % (g_k, g_p))
    if grab_beklenen:
        kontrol("%s kilit (grab) geri alındı" % ad, grab == str(grab_beklenen), "grab=%s beklenen=%s" % (grab, grab_beklenen))
    return grab


def senaryo(kok):
    kok.update()
    time.sleep(0.5)
    kok.update()
    # S1: modal form, 3 döngü
    bul_dugme(kok, "Yeni Kayıt").invoke()
    kok.update()
    form = [w for w in toplevels(kok) if w.title() == "Yeni Kayıt"][0]
    kontrol("S1 form açıldı ve modal (kilit formda)", str(kok.grab_current()) == str(form))
    dongu(kok, [form], "S1", form, tur=3)
    kontrol("S1 form ana pencerenin sahipliğinde", bool(u.GetWindow(hw(form), 4)))
    form.destroy()
    kok.update()
    kontrol("S1 form kapanınca kilit kalkar", str(kok.grab_current() or "-") == "-")
    # S2: modal olmayan pencere
    bul_dugme(kok, "🔍 Fatura Bul").invoke()
    kok.update()
    bul = [w for w in toplevels(kok) if w.title() == "Fatura Bul"][0]
    dongu(kok, [bul], "S2", None, tur=2)
    bul.destroy()
    kok.update()
    kok.update()
    kontrol("S2 pencere kapanınca klavye odağı ana pencereye döner", kok.focus_get() is not None, str(kok.focus_get()))
    # S3: iç içe modal (Kullanıcılar -> Yeni Kullanıcı)
    m = kok.nametowidget(kok["menu"])
    dosya = m.nametowidget(m.entrycget(0, "menu"))
    for i in range(dosya.index("end") + 1):
        if dosya.type(i) != "separator" and dosya.entrycget(i, "label") == "Kullanıcılar...":
            dosya.invoke(i)
    kok.update()
    kul = [w for w in toplevels(kok) if w.title() == "Kullanıcılar"][0]

    def ic_diyalog():
        try:
            dlg = [w for w in toplevels(kok) if w.title() == "Yeni Kullanıcı"][0]
            dongu(kok, [kul, dlg], "S3", dlg, tur=2)
            dlg.destroy()
        except Exception:
            sonuc["hata"] = traceback.format_exc()
            for w in toplevels(kok):
                w.destroy()
    kok.after(600, ic_diyalog)
    bul_dugme(kul, "Yeni Kullanıcı...").invoke()                     # diyalog wait_window ile bloklar; geri çağrı kapatır
    kok.update()
    kul.destroy()
    kok.update()
    kontrol("S3 tüm modal pencereler kapanınca kilit kalmaz", str(kok.grab_current() or "-") == "-")
    print("\n".join(sonuc["satirlar"]))
    print("SONUÇ: %d kontrol, %d başarısız" % (len(sonuc["satirlar"]), len(sonuc["basarisiz"])))
    kok.destroy()


Tk0 = tk.Tk


class Tk2(Tk0):
    def __init__(self, *a, **k):
        super().__init__(*a, **k)

        def ba():
            try:
                senaryo(self)
            except Exception:
                sonuc["hata"] = traceback.format_exc()
                print(sonuc["hata"])
                print("\n".join(sonuc["satirlar"]))
                try:
                    self.destroy()
                except tk.TclError:
                    pass
        self.after(1200, ba)


tk.Tk = Tk2
f.arayuzu_calistir()
sys.exit(1 if sonuc["hata"] or sonuc["basarisiz"] else 0)
