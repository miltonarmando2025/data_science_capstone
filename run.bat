@echo off
:: =============================================
:: Data Science Pipeline - CMD Compatible Version
:: =============================================

:: 1. Set working directory to script location
cd /d "%~dp0"

:: 1. Generate presentation
echo Generating report...
python src\helper.py
if errorlevel 1 (
    echo ERROR: Failed to generate presentation
    goto :error
)

:: 6. Verify outputs
if not exist "results\income_plot.png" (
    echo ERROR: Missing results\income_plot.png
    goto :error
)

if not exist "presentation.pdf" (
    echo ERROR: Missing presentation.pdf
    goto :error
)

:: Success message
echo.
echo ==================================
echo SUCCESS: PIPELINE COMPLETED
echo ==================================
echo Output files created:
dir results /b
echo.

start "" "presentation.pdf" 2>nul
goto :end

:error
echo.
echo ==================================
echo ERROR: PIPELINE FAILED
echo ==================================
pause
exit /b 1

:end
popd
pause