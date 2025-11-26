@echo off
REM Script batch per convertire Word in PDF su Windows
REM Trascina e rilascia un file .docx o una cartella su questo file .bat

echo ====================================
echo Word to PDF Converter
echo ====================================
echo.

REM Verifica se è stato passato un argomento
if "%~1"=="" (
    echo Utilizzo: Trascina un file .docx o una cartella su questo file
    echo.
    echo Oppure usa da linea di comando:
    echo   convert_word_to_pdf.bat file.docx
    echo   convert_word_to_pdf.bat cartella
    echo.
    pause
    exit /b 1
)

REM Esegui la conversione
python word_to_pdf_converter.py "%~1"

echo.
echo ====================================
echo Conversione completata!
echo ====================================
pause
