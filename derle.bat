@echo off
rem Alis Faturasi Kayit ve Takip - exe olusturma
rem Bu dosyayi cift tiklayin. Sonuc: dist\FaturaTakip klasoru (icinde FaturaTakip.exe)
rem Gerekenler (bir kez): pip install pyinstaller pymupdf openpyxl sv_ttk
cd /d "%~dp0"

python -m PyInstaller --noconfirm --clean --onedir --windowed --name FaturaTakip ^
  --icon ikon.ico --add-data "ikon.ico;." --add-data "sablon\satinalma_sablonu.xlsx;sablon" ^
  --collect-all sv_ttk --collect-all pymupdf --hidden-import openpyxl ^
  fatura_takip.py
if errorlevel 1 (
  echo.
  echo HATA: exe olusturulamadi. Yukaridaki mesajlari kontrol edin.
  pause
  exit /b 1
)

rem Pakete config.ini KONMAZ: tum klasoru ag klasorundeki kurulumun ustune kopyalamak, ayarli config.ini'yi (ag yolu) ezip
rem uygulamayi bos bir veritabaniyla actirirdi. Ilk kurulumda config.ornek.ini dosyasini config.ini adiyla kopyalayip yollari yazin.
copy /Y "config.ornek.ini" "dist\FaturaTakip\config.ornek.ini" >nul
if exist "dist\FaturaTakip\config.ini" del "dist\FaturaTakip\config.ini"
copy /Y "LICENSE" "dist\FaturaTakip\" >nul
copy /Y "THIRD_PARTY_NOTICES.md" "dist\FaturaTakip\" >nul

echo.
echo Tamam: dist\FaturaTakip\FaturaTakip.exe
echo Bu klasorun tamamini (sadece exe'yi degil) flash bellege kopyalayin.
echo Ag klasorundeki kurulumu GUNCELLERKEN config.ini ve Veri klasorune dokunmayin (pakette config.ini yoktur).
pause
