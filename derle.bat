@echo off
rem Alis Faturasi Kayit ve Takip - exe olusturma
rem Bu dosyayi cift tiklayin. Sonuc: dist\FaturaTakip klasoru (icinde FaturaTakip.exe)
rem Gerekenler (bir kez): pip install pyinstaller pymupdf openpyxl sv_ttk
cd /d "%~dp0"

python -m PyInstaller --noconfirm --clean --onedir --windowed --name FaturaTakip ^
  --collect-all sv_ttk --collect-all pymupdf --hidden-import openpyxl ^
  fatura_takip.py
if errorlevel 1 (
  echo.
  echo HATA: exe olusturulamadi. Yukaridaki mesajlari kontrol edin.
  pause
  exit /b 1
)

rem Ornek ayar dosyasi: yoksa program ilk acilista kendisi olusturur, burada ag yolunu onceden yazabilirsiniz
if not exist "dist\FaturaTakip\config.ini" (
  > "dist\FaturaTakip\config.ini" (
    echo [ayarlar]
    echo veritabani = fatura.db
    echo tarama_klasoru =
    echo tema = Açık
    echo kurum_anahtar =
    echo kurum_vkn =
  )
)

echo.
echo Tamam: dist\FaturaTakip\FaturaTakip.exe
echo Bu klasorun tamamini (sadece exe'yi degil) flash bellege kopyalayin.
pause
