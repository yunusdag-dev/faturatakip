# Üçüncü taraf bileşenler / Third-party components

Bu uygulamanın kaynak kodu [MIT lisansıyla](LICENSE) yayınlanmıştır (kurumlar dahil, kullanım, değiştirme ve dağıtım serbesttir).
Derlenmiş `.exe` paketi aşağıdaki bileşenleri içerir; her biri kendi lisansıyla gelir.

| Bileşen | Lisans | Not |
|---|---|---|
| [Python](https://www.python.org/) ve standart kütüphane (tkinter, sqlite3 …) | PSF License | Pakete gömülüdür |
| [PyMuPDF](https://github.com/pymupdf/PyMuPDF) | **AGPL-3.0** veya ticari lisans (Artifex) | PDF metin okuma. İsteğe bağlıdır; kaynak kodda `import` hatasına karşı korumalıdır |
| [openpyxl](https://openpyxl.readthedocs.io/) | MIT | Excel (.xlsx) dışa aktarma |
| [sv-ttk](https://github.com/rdbende/Sun-Valley-ttk-theme) | MIT | Arayüz teması |
| [PyInstaller](https://pyinstaller.org/) | GPLv2+ (derlenen programları sınırlamayan özel istisna ile) | Yalnızca exe üretiminde kullanılır |

## PyMuPDF (AGPL-3.0) hakkında

AGPL-3.0, kullanımı kısıtlamaz: bu uygulamayı bir kurumda, ücretsiz ya da dahili olarak çalıştırmak serbesttir.
AGPL'nin şartları **dağıtım** sırasında devreye girer. PyMuPDF içeren derlenmiş paketi başkalarına dağıtırsanız,
paketin AGPL-3.0 koşullarına uyması gerekir; bu kaynak kodun tamamı bu depoda herkese açıktır.

PyMuPDF'siz çalışmak isterseniz kütüphaneyi kurmadan (`pip install pymupdf` yapmadan) programı kaynaktan çalıştırabilir
ya da `derle.bat` içindeki `--collect-all pymupdf` seçeneğini çıkararak exe üretebilirsiniz; uygulama yerleşik PDF okuyucuya
geçer (okuma kalitesi daha düşük olabilir). Ticari/kapalı kaynak dağıtım planlıyorsanız Artifex'ten ticari lisans alın.
