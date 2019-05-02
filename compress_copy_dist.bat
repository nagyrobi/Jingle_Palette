del Jingle_Palette_Vb.exe /F /Q
rename Jingle_Palette.exe Jingle_Palette_Vb.exe
rem upx --best --crp-ms=999999 --nrv2b -o Jingle_Palette.exe Jingle_Palette_Vb.exe
upx --best --crp-ms=999999 --nrv2d -o Jingle_Palette.exe Jingle_Palette_Vb.exe
copy "Jingle_Palette.exe" "C:\Program Files\Jingle Palette\" /Y
copy "language.ini" "C:\Program Files\Jingle Palette\" /Y
pause
