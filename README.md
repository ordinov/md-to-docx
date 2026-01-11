# md-docx-converter

Converte file Markdown in documenti Word (.docx) e viceversa.

**Zero configurazione** - gli script gestiscono tutto automaticamente.

## Requisiti

Solo Python 3.6+ (gia installato su Mac/Linux, scaricabile da python.org per Windows).

## Uso

### Markdown -> Word

```bash
python3 md2docx.py /path/to/documento.md
```

Crea `documento.docx` nella stessa cartella del file originale.

Per comodità si può usare la cartella `docs` di questo repository per lo storage dei documenti da convertire, è untracked.

### Word -> Markdown

```bash
python3 docx2md.py documento.docx
```

Crea `documento.md` nella stessa cartella del file originale.

### Prima esecuzione

Al primo avvio lo script:
1. Crea un ambiente virtuale (`.venv/`) nella cartella dello script
2. Installa automaticamente le dipendenze
3. Esegue la conversione

Le esecuzioni successive sono immediate.

## Compatibilita

| Sistema | Testato |
|---------|---------|
| Windows | Si |
| macOS   | Si |
| Linux   | Si |

## Funzionalita supportate

### md2docx.py (Markdown -> Word)

| Markdown | Risultato |
|----------|-----------|
| `# Titolo` | Titolo principale |
| `## Sezione` | Intestazione livello 1 |
| `### Sottosezione` | Intestazione livello 2 |
| `#### Paragrafo` | Intestazione livello 3 |
| `**testo**` | **Grassetto** |
| `*testo*` | *Corsivo* |
| `[link](url)` | Link sottolineato |
| `- elemento` | Lista puntata |
| `1. elemento` | Lista numerata |
| `> citazione` | Blocco citazione |
| `---` | Linea orizzontale |
| Tabelle | Tabelle con intestazione in grassetto |

### docx2md.py (Word -> Markdown)

| Word | Risultato |
|------|-----------|
| Titoli/Intestazioni | `#`, `##`, `###`, `####` |
| **Grassetto** | `**testo**` |
| *Corsivo* | `*testo*` |
| Liste puntate | `- elemento` |
| Liste numerate | `1. elemento` |
| Paragrafi indentati | `> citazione` |
| Linee orizzontali | `---` |
| Tabelle | Tabelle markdown |

## Esempi

### Convertire Markdown in Word

```bash
python3 md2docx.py ~/Documents/report.md
```

Output:
```
Setting up environment (first run only)...
Installing dependencies...
Setup complete!

Converted: report.md -> report.docx
```

### Convertire Word in Markdown

```bash
python3 docx2md.py ~/Documents/report.docx
```

Output:
```
Converted: report.docx -> report.md
```

## Struttura cartella

```
md-docx-converter/
├── md2docx.py     # Markdown -> Word
├── docx2md.py     # Word -> Markdown
├── README.md      # Questo file
└── .venv/         # Creato automaticamente al primo avvio
```

## Limitazioni

### md2docx.py
- Code blocks non supportati (trattati come testo)
- Immagini non supportate
- Liste annidate solo a un livello

### docx2md.py
- Immagini non estratte
- Formattazione complessa potrebbe non essere preservata
- Link ipertestuali convertiti come testo semplice
- Stili personalizzati potrebbero non essere riconosciuti
