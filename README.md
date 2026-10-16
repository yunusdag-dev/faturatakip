# Alış Faturası Kayıt ve Takip

Ekip içinde ortak ağ üzerinden kullanılan, alış faturalarını kaydeden ve takip eden masaüstü uygulaması (Python, tkinter, SQLite).

## Özellikler

- Fatura kaydı: tarih, fatura no, firma, yapılan iş, KDV hariç/dahil tutar, PDF konumu
- **PDF'den otomatik doldurma:** PyMuPDF, yerleşik PDF okuyucu ve Windows OCR (taranmış belgeler için)
- **Çok sayfalı belgede fatura sayfasını otomatik bulma** (taramalarda ilk 20 sayfa) veya "Fatura sayfası" kutusundan elle seçme
- Faturaya ek evrak bağlama (talep, yazı, irsaliye vb.); bir faturaya çift tıklayınca tüm evraklar açılır
- Yıl/ay/firma/tarih filtresi ve arama, aylık özet
- Biçimlendirilmiş Excel (`.xlsx`) dışa aktarma (CSV de desteklenir)
- Açık / Koyu / Mavi tema
- Mükerrer kayıt kontrolü (aynı firma + fatura no, aynı PDF)

Tutarlar veritabanında kuruş (tam sayı) olarak tutulur. Dosyalar kopyalanmaz, yalnızca konumları kaydedilir.

## Çalıştırma (geliştirme)

```
pip install pymupdf openpyxl sv_ttk
python fatura_takip.py
```

`pymupdf`, `openpyxl` ve `sv_ttk` isteğe bağlıdır; yoksa program sırasıyla yerleşik PDF okuyucu, CSV ve klasik temayla çalışmaya devam eder.

## Ayarlar

Programın yanındaki `config.ini` dosyası kullanılır (ilk açılışta oluşturulur). Örnek: [config.ornek.ini](config.ornek.ini)

| Ayar | Açıklama |
|---|---|
| `veritabani` | Veritabanı yolu. Ağ için `\\SUNUCU-PC\FaturaTakip\Veri\fatura.db` (sürücü harfi değil, UNC yolu) |
| `tarama_klasoru` | PDF "Gözat" penceresinin ilk klasörü |
| `tema` | Açık / Koyu / Mavi |
| `kurum_anahtar` | Kendi kurumunuzun adından kelimeler; PDF'te alıcıyı firma sanmasın diye |
| `kurum_vkn` | Kendi kurumunuzun vergi numarası |

## Exe oluşturma

Bu bilgisayarda bir kez: `pip install pyinstaller pymupdf openpyxl sv_ttk`

Sonra [derle.bat](derle.bat) dosyasına çift tıklayın. Çıktı `dist\FaturaTakip\` klasörüdür (Python kurulumu gerektirmez). Kullanıcılara `.exe` tek başına değil, **klasörün tamamı** dağıtılır.

## Ortak ağda kurulum

1. Ortak bilgisayarda `C:\FaturaTakip\Program`, `Veri` ve `Evraklar` klasörlerini açın; `dist\FaturaTakip` içeriğini `Program` içine kopyalayın.
2. `C:\FaturaTakip` klasörünü paylaşıma açın; kullanıcılara **Okuma + Değiştirme** izni verin (veritabanı yazma izni ister).
3. `Program\config.ini` içinde `veritabani` ve `tarama_klasoru` yollarını `\\SUNUCU-PC\...` biçiminde yazın.
4. Diğer bilgisayarlarda `\\SUNUCU-PC\FaturaTakip\Program\FaturaTakip.exe` için masaüstü kısayolu oluşturun.
5. Güncelleme: herkes programı kapattıktan sonra yeni dosyaları `Program` klasörüne kopyalayın; `config.ini` ve `Veri` klasörüne dokunmayın.
6. `Veri\fatura.db` dosyasını düzenli yedekleyin.

Taranmış (resim) PDF'leri okumak için Windows'ta Türkçe OCR dil paketi kurulu olmalıdır.

## Dikkat

- `fatura.db`, `config.ini` ve fatura PDF'leri `.gitignore` ile repo dışında tutulur; gerçek veriyi commit etmeyin.
- Sonuçları kontrol edin: OCR ile okunan taramalarda hata payı vardır (ör. `I` yerine `L`).
