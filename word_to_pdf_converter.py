#!/usr/bin/env python3
"""
Word to PDF Converter
Converte documenti Word (.docx) in PDF
Supporta singoli file o intere cartelle
"""

import os
import sys
from pathlib import Path
from docx2pdf import convert
import argparse


def convert_single_file(input_path, output_path=None):
    """
    Converte un singolo file DOCX in PDF
    
    Args:
        input_path: Path del file .docx da convertire
        output_path: Path del file PDF di output (opzionale)
    """
    input_file = Path(input_path)
    
    if not input_file.exists():
        print(f"❌ Errore: Il file '{input_path}' non esiste")
        return False
    
    if input_file.suffix.lower() != '.docx':
        print(f"❌ Errore: Il file '{input_path}' non è un file .docx")
        return False
    
    # Se non specificato, crea il PDF nella stessa cartella con lo stesso nome
    if output_path is None:
        output_path = input_file.with_suffix('.pdf')
    else:
        output_path = Path(output_path)
    
    try:
        print(f"📄 Conversione: {input_file.name} -> {output_path.name}")
        convert(str(input_file), str(output_path))
        print(f"✅ Completato: {output_path}")
        return True
    except Exception as e:
        print(f"❌ Errore durante la conversione di {input_file.name}: {str(e)}")
        return False


def convert_folder(input_folder, output_folder=None):
    """
    Converte tutti i file .docx in una cartella
    
    Args:
        input_folder: Path della cartella contenente i file .docx
        output_folder: Path della cartella di output (opzionale)
    """
    input_path = Path(input_folder)
    
    if not input_path.exists() or not input_path.is_dir():
        print(f"❌ Errore: '{input_folder}' non è una cartella valida")
        return
    
    # Trova tutti i file .docx
    docx_files = list(input_path.glob('*.docx'))
    
    if not docx_files:
        print(f"⚠️  Nessun file .docx trovato in '{input_folder}'")
        return
    
    print(f"📁 Trovati {len(docx_files)} file .docx")
    
    # Crea la cartella di output se specificata
    if output_folder:
        output_path = Path(output_folder)
        output_path.mkdir(parents=True, exist_ok=True)
    else:
        output_path = input_path
    
    # Converti ogni file
    successful = 0
    failed = 0
    
    for docx_file in docx_files:
        pdf_file = output_path / docx_file.with_suffix('.pdf').name
        if convert_single_file(docx_file, pdf_file):
            successful += 1
        else:
            failed += 1
    
    print(f"\n📊 Riepilogo:")
    print(f"   ✅ Convertiti con successo: {successful}")
    print(f"   ❌ Falliti: {failed}")


def main():
    parser = argparse.ArgumentParser(
        description='Converte documenti Word (.docx) in PDF',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Esempi d'uso:
  # Converti un singolo file
  python word_to_pdf_converter.py documento.docx
  
  # Converti un file specificando l'output
  python word_to_pdf_converter.py documento.docx -o output.pdf
  
  # Converti tutti i file in una cartella
  python word_to_pdf_converter.py -f cartella_documenti
  
  # Converti una cartella specificando la cartella di output
  python word_to_pdf_converter.py -f cartella_input -o cartella_output
        """
    )
    
    parser.add_argument('input', nargs='?', help='File .docx o cartella da convertire')
    parser.add_argument('-f', '--folder', help='Cartella contenente i file .docx da convertire')
    parser.add_argument('-o', '--output', help='File o cartella di output')
    
    args = parser.parse_args()
    
    # Verifica che sia stato fornito almeno un input
    if not args.input and not args.folder:
        parser.print_help()
        sys.exit(1)
    
    # Modalità cartella
    if args.folder:
        convert_folder(args.folder, args.output)
    # Modalità file singolo
    elif args.input:
        input_path = Path(args.input)
        if input_path.is_dir():
            convert_folder(args.input, args.output)
        else:
            convert_single_file(args.input, args.output)


if __name__ == '__main__':
    main()
