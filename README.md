# HKEX Rulebook EN/ZH Aligner

Aligns English and Traditional Chinese HKEX rulebook PDF chapters entry-by-entry and writes the results to Excel.

## Requirements

```
pip install pdfplumber openpyxl
```

## Usage

```
python3 compare.py <en_folder> <zh_folder> [out_folder]
```

| Argument | Description |
|---|---|
| `en_folder` | Folder containing English PDF chapters |
| `zh_folder` | Folder containing Traditional Chinese PDF chapters |
| `out_folder` | Output folder for `.xlsx` files (default: `./output`) |

PDFs are matched by filename across the two folders. One `.xlsx` file is written per matched pair.

**Example:**

```
python3 compare.py hkex-rulebook-pdfs hkex-rulebook-pdfs-zh
```

Produces `output/chapter_1_aligned.xlsx`, `output/chapter_2_aligned.xlsx`, etc.

## Output

Each Excel file has one row per aligned entry:

| Column | Description |
|---|---|
| Line | Row number |
| Page (EN) | Page the entry was found on in the English PDF |
| Page (ZH) | Page the entry was found on in the Chinese PDF |
| English | Full reconstructed text of the English entry |
| Chinese | Full reconstructed text of the Chinese entry |
| Match? | Sanity-check result (see below) |

### Match column values

| Value | Meaning |
|---|---|
| `OK (anchor)` | English term found in parentheses inside the ZH entry — strong signal |
| `OK (1.02A)` | Both entries start with the same rule number |
| `X anchor en=... zh=...` | Anchor check failed — terms don't match |
| `X rule en=... zh=...` | Rule ID mismatch |
| `-` | No checkable signal in either entry |

## How it works

### Entry extraction

Words are extracted with their `(x, y)` positions using `pdfplumber`. The PDF's two-column layout is exploited: words with `x0 < 180` (EN) or `x0 < 145` (ZH) are the left column (term/rule number); the rest are the right column (definition).

Lines are grouped by y-coordinate proximity. A **new entry boundary** is detected only when both conditions hold:

1. Vertical gap to the previous line exceeds `14px`
2. The line's leftmost word starts with a smart open-quote (`"`, `《`) or a numbered rule token (`1.02A`)

This keeps sub-notes (`Note:`), lettered sub-items (`(a)`), and inline quoted aliases (`"AFRC"`) merged into their parent entry rather than split off.

### Alignment

Entries are aligned positionally — entry *N* in the English PDF is paired with entry *N* in the Chinese PDF. A console summary is printed showing how many entries were extracted from each PDF; large discrepancies indicate a segmentation issue worth investigating.

### Sanity checks

Each row is checked with two independent signals:

1. **Glossary anchor** — the English term appears in parentheses inside the ZH entry (e.g. `"申請版本" (Application Proof)` ↔ `"Application Proof"`). Comparison is case-insensitive and ignores whitespace/dashes.
2. **Rule-ID identity** — both entries start with the same numbered rule (e.g. `4.21`).

If neither signal is available the row is marked `-` (no signal), which is normal for plain-prose entries.

## Known limitations

- Alignment is positional (index-based). If the segmentation produces a different entry count for EN vs ZH, pairs after the first mismatch will be wrong. Check the `X` rows and the console entry counts to detect this.
- Tuned for HKEX Main Board rulebook PDFs. The column-split x-coordinates (`EN_COL_SPLIT = 180`, `ZH_COL_SPLIT = 145`) and gap threshold (`ENTRY_GAP = 14`) may need adjustment for other documents.
- Page 1 header decoration is excluded up to y=140. If a document has an unusually tall title block, adjust `HEADER_Y_MAX`.
