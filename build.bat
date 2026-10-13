@echo off
chcp 65001 >nul

pyinstaller --onefile --console ^
    --hidden-import=win32com ^
    --hidden-import=win32com.client ^
    --hidden-import=pythoncom ^
    --hidden-import=pywintypes ^
    --hidden-import=win32timezone ^
    --hidden-import=shapely ^
    --hidden-import=shapely.geometry ^
    --hidden-import=shapely.ops ^
    --hidden-import=shapely.geos ^
    --collect-all=shapely ^
    --collect-all=ezdxf ^
    --collect-all=pywin32 ^
    --copy-metadata=shapely ^
    --copy-metadata=ezdxf ^
    --name FlatPatternExport ^
    scripts_kmp.py

echo.
echo Готово. Файл: dist\FlatPatternExport.exe
pause