@echo off
REM ============================================================
REM  Script build AutoHelper tren Windows
REM  Ho tro ca 2 trinh bien dich: MSVC (cl) hoac MinGW (g++).
REM ============================================================

where cl >nul 2>nul
if %ERRORLEVEL%==0 (
    echo [*] Dang build bang MSVC cl...
    cl /EHsc /O2 /Fe:AutoHelper.exe main.cpp user32.lib gdi32.lib
    goto done
)

where g++ >nul 2>nul
if %ERRORLEVEL%==0 (
    echo [*] Dang build bang MinGW g++...
    g++ -O2 -std=c++11 main.cpp -o AutoHelper.exe -lgdi32 -luser32 -static
    goto done
)

echo [LOI] Khong tim thay cl (MSVC) hay g++ (MinGW).
echo        Cai Visual Studio Build Tools hoac MinGW-w64 roi chay lai.
exit /b 1

:done
echo.
echo [*] Xong. Chay: AutoHelper.exe
