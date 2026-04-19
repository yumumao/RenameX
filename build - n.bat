@echo off
REM 打包命令（根据你的实际路径修改参数）
pyinstaller --noconfirm --noconsole ^
--upx-dir="C:\tools\upx-5.0.0-win64" ^
--hidden-import=sip ^
--icon="%~dp0rename.ico" ^
--add-data="%~dp0rename.ico;." ^
--add-data="C:\Users\ZGVi-HX90\AppData\Local\Programs\Python\Python313\Lib\site-packages\PyQt5\Qt5\plugins\platforms;PyQt5\Qt5\plugins\platforms" ^
"%~dp0RenameX.pyw"

REM 执行完成后暂停，方便查看结果
pause