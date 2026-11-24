@echo off
chcp 65001 >nul
REM 在「有 Windows 與 Python 的電腦」上執行一次,即可產生免安裝版。
REM 若這台電腦已安裝 LibreOffice,會自動一併打包。

pip install pyinstaller tkinterdnd2 pywin32 || goto :err
pyinstaller --noconsole --onedir --name Doc2Docx --collect-all tkinterdnd2 ^
  --hidden-import win32com.client --hidden-import pythoncom doc2docx_gui.py || goto :err

if exist "%ProgramFiles%\LibreOffice" (
  xcopy "%ProgramFiles%\LibreOffice" "dist\Doc2Docx\LibreOffice\" /E /I /Q /Y
) else (
  echo [警告] 找不到 LibreOffice,請先安裝後再執行,否則 exe 需依賴使用者電腦上的 Word / LibreOffice。
)
echo.
echo 完成!請將整個 dist\Doc2Docx 資料夾打包給使用者,執行其中的 Doc2Docx.exe。
pause
exit /b 0
:err
echo 建置失敗
pause
exit /b 1
