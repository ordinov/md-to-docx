# md2docx

Converte file Markdown in documenti Word (.docx).

**Zero configurazione** - lo script gestisce tutto automaticamente.

## Requisiti

Solo Python 3.6+ (gia installato su Mac/Linux, scaricabile da python.org per Windows).

## Uso

Eseguire semplicemente:

```bash
python md2docx.py documento.md
```

Crea `documento.docx` nella stessa cartella del file originale.

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

## Esempio

```bash
python md2docx.py ~/Documents/report.md
```

Output:
```
Setting up environment (first run only)...
Installing dependencies...
Setup complete!

Converted: report.md -> report.docx
```

## Struttura cartella

```
md-to-pdf/
├── md2docx.py     # Lo script
├── README.md      # Questo file
└── .venv/         # Creato automaticamente al primo avvio
```

## Limitazioni

- Code blocks non supportati (trattati come testo)
- Immagini non supportate
- Liste annidate solo a un livello
