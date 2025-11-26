# Word to PDF Converter 📄➡️📕

Tool Python per convertire documenti Word (.docx) in PDF, con supporto per singoli file o elaborazione batch di intere cartelle.

## 🚀 Installazione

### Requisiti
- Python 3.6 o superiore
- Microsoft Word (su Windows) o LibreOffice (su Linux/Mac)

### Installazione dipendenze

```bash
# Crea un ambiente virtuale
python3 -m venv venv

# Attivalo
source venv/bin/activate

# Installa
pip install -r requirements.txt

# Quando hai finito, disattiva l'ambiente
deactivate
```

Oppure installare direttamente:

```bash
pip install docx2pdf
```

**Nota per Linux/Mac:** È necessario avere LibreOffice installato:
```bash
# Ubuntu/Debian
sudo apt-get install libreoffice

# macOS (con Homebrew)
brew install --cask libreoffice
```

## 📖 Utilizzo

### Converti un singolo file

```bash
python word_to_pdf_converter.py documento.docx
```

Il PDF verrà creato nella stessa cartella con lo stesso nome.

### Specifica il nome del file di output

```bash
python word_to_pdf_converter.py documento.docx -o mio_pdf.pdf
```

### Converti tutti i file in una cartella

```bash
python word_to_pdf_converter.py -f cartella_documenti
```

I PDF verranno creati nella stessa cartella dei file originali.

### Converti una cartella specificando la cartella di output

```bash
python word_to_pdf_converter.py -f cartella_input -o cartella_output
```

### Opzioni disponibili

```
positional arguments:
  input                 File .docx o cartella da convertire

optional arguments:
  -h, --help            Mostra questo messaggio di aiuto
  -f, --folder FOLDER   Cartella contenente i file .docx da convertire
  -o, --output OUTPUT   File o cartella di output
```

## 📝 Esempi

### Esempio 1: Convertire un documento
```bash
python word_to_pdf_converter.py relazione_annuale.docx
```
Output: `relazione_annuale.pdf` nella stessa cartella

### Esempio 2: Convertire con nome personalizzato
```bash
python word_to_pdf_converter.py bozza.docx -o versione_finale.pdf
```

### Esempio 3: Batch conversion
```bash
python word_to_pdf_converter.py -f ./documenti -o ./pdf_output
```
Converte tutti i file .docx da `./documenti` e salva i PDF in `./pdf_output`

## 🔧 Troubleshooting

### Windows
Se ricevi errori su Windows, assicurati che Microsoft Word sia installato e attivato.

### Linux/Mac
Assicurati che LibreOffice sia installato correttamente:
```bash
libreoffice --version
```

### Errori comuni
- **"File not found"**: Verifica che il path del file sia corretto
- **"Not a .docx file"**: Il tool supporta solo file .docx (non .doc)
- **Errori di conversione**: Verifica che il file Word non sia corrotto e possa essere aperto normalmente

## ⚙️ Come funziona

Il tool utilizza la libreria `docx2pdf` che:
- Su **Windows**: usa l'installazione di Microsoft Word via COM
- Su **Linux/Mac**: usa LibreOffice in modalità headless

## 📄 Licenza

Questo tool è fornito "as is" per uso personale e commerciale.

## 🤝 Contributi

Sentiti libero di modificare e migliorare questo tool secondo le tue esigenze!
