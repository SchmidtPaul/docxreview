# docxreview — Implementierungsplan

## Status: Bereit zur Implementierung (Plan genehmigt)

## Context

Paul rendert regelmäßig QMD-Dateien zu DOCX, schickt diese an Reviewer, die in Microsoft Word Kommentare und Korrekturvorschläge (Tracked Changes) einfügen. Aktuell gibt es keinen automatisierten Weg, dieses Feedback strukturiert und mit Kontext zu extrahieren. Der bestehende `docx`-Skill von Claude Code ist auf das *Erstellen/Bearbeiten* von DOCX ausgelegt, nicht auf *Extraktion*.

**Ziel:** Ein schlankes R-Paket `docxreview`, das Kommentare und Tracked Changes aus einer gegengelesenen DOCX-Datei extrahiert und als strukturierte Markdown-Liste ausgibt.

## Technische Grundlage (Recherche abgeschlossen)

- `officer::docx_comments()` liefert Kommentare inkl. markiertem Text (`commented_text`), Autor, Datum — aber keinen erweiterten Kontext
- `officer` hat **keine** eingebaute Tracked-Changes-Extraktion
- Tracked Changes (`w:ins`, `w:del`) müssen via `xml2` direkt aus dem Body-XML geparst werden
- Kontext: Der umgebende Absatz (Parent `w:p`) liefert den nötigen Zusammenhang

## Paketstruktur

```
C:/GitHub/docxreview/
├── R/
│   ├── extract_comments.R         # Kommentare mit Kontext extrahieren
│   ├── extract_tracked_changes.R  # Tracked Changes mit Kontext extrahieren
│   ├── extract_review.R           # Hauptfunktion: kombiniert beides → Markdown
│   ├── utils.R                    # Hilfsfunktionen (XML-Kontext, Text-Cleaning)
│   └── docxreview-package.R       # Package-Level Doku
├── tests/
│   └── testthat/
│       ├── test-extract_comments.R
│       ├── test-extract_tracked_changes.R
│       └── fixtures/              # Test-DOCX mit bekannten Kommentaren/Changes
├── DESCRIPTION
├── NAMESPACE
├── LICENSE
└── README.md
```

## Funktionen im Detail

### 1. `extract_comments(docx_path)`

- Nutzt `officer::read_docx()` + `officer::docx_comments()`
- Erweitert um Paragraph-Kontext via `xml2`: Für jeden Kommentar wird der volle Absatztext extrahiert, in dem `commentRangeStart`/`commentRangeEnd` verankert sind
- **Rückgabe:** tibble mit Spalten: `comment_id`, `author`, `date`, `comment_text`, `commented_text`, `paragraph_context`

### 2. `extract_tracked_changes(docx_path)`

- Liest DOCX via `officer::read_docx()`, greift auf Body-XML zu via `officer::docx_body_xml()`
- Findet alle `w:ins` und `w:del` Nodes mit `xml2::xml_find_all()`
- Für jeden Change:
  - **Typ**: "insertion" oder "deletion"
  - **Geänderter Text**: `w:t` (bei ins) bzw. `w:delText` (bei del)
  - **Autor/Datum**: aus `w:author` / `w:date` Attributen
  - **Kontext**: Voller Text des Parent-`w:p` (Absatz), wobei die Änderung markiert wird
- **Rückgabe:** tibble mit Spalten: `change_id`, `type`, `author`, `date`, `changed_text`, `paragraph_context`

### 3. `extract_review(docx_path, output = NULL)`

- Hauptfunktion, ruft `extract_comments()` und `extract_tracked_changes()` auf
- Generiert strukturierten Markdown-Output:

```markdown
# Review Feedback: bericht_v2_reviewed.docx

## Kommentare (3)

### 1. Max Mustermann (2024-01-15)
> **Markierter Text:** "the primary endpoint was overall survival"
> **Absatz:** "In this study, the primary endpoint was overall survival, measured from randomization to death from any cause."

**Kommentar:** "Sollten wir hier auch die sekundären Endpunkte erwähnen?"

---

## Tracked Changes (5)

### 1. [Löschung] Max Mustermann (2024-01-15)
> **Gelöscht:** "approximately"
> **Absatz:** "The sample size was ~~approximately~~ 200 patients."

### 2. [Einfügung] Max Mustermann (2024-01-15)
> **Eingefügt:** "statistically significant"
> **Absatz:** "The difference was **statistically significant** (p < 0.05)."
```

- Wenn `output` angegeben → schreibt in Datei; sonst gibt den Markdown-String zurück

### 4. `utils.R` — Hilfsfunktionen

- `get_paragraph_text(p_node)`: Extrahiert vollen Text eines `w:p` Nodes (alle `w:t` + `w:delText` zusammen)
- `get_paragraph_context(change_node)`: Navigiert zum Parent-Paragraph und gibt Kontext zurück
- `format_context_with_markup(paragraph_text, changed_text, type)`: Markiert die Änderung im Absatztext (~~deletion~~ / **insertion**)

## Kritische technische Details

### XML-Zugriff auf den Body
```r
doc <- officer::read_docx(path)
body_xml <- officer::docx_body_xml(doc)
# oder: xml2::read_xml(doc$doc_obj$get())
```

### Tracked Changes finden
```r
# Insertions
ins_nodes <- xml2::xml_find_all(body_xml, "//w:ins")
# Deletions
del_nodes <- xml2::xml_find_all(body_xml, "//w:del")
# Attribute
xml2::xml_attr(ins_nodes, "w:author")
xml2::xml_attr(ins_nodes, "w:date")
```

### Kontext via Parent-Paragraph
```r
parent_p <- xml2::xml_find_first(change_node, "ancestor::w:p")
all_text <- xml2::xml_find_all(parent_p, ".//w:t | .//w:delText")
paragraph_text <- paste(xml2::xml_text(all_text), collapse = "")
```

## Implementierungsschritte

1. **Paket-Gerüst erstellen** via `usethis::create_package("C:/GitHub/docxreview")`
2. **Dependencies** einrichten: `officer`, `xml2`, `cli`, `tibble` in DESCRIPTION
3. **`utils.R`** implementieren (XML-Hilfsfunktionen)
4. **`extract_comments.R`** implementieren
5. **`extract_tracked_changes.R`** implementieren
6. **`extract_review.R`** implementieren (Kombination + Markdown-Formatierung)
7. **Test-Fixture erstellen**: Kleine DOCX-Datei mit bekannten Kommentaren und Tracked Changes
8. **Tests schreiben** für alle drei Kernfunktionen
9. **Dokumentation** (roxygen2) für alle exportierten Funktionen
10. **README** mit Workflow-Beispiel: QMD → DOCX → Review → `extract_review()`

## Verifizierung

1. **Manueller Test**: Eine DOCX mit bekannten Kommentaren und Tracked Changes erstellen, `extract_review()` ausführen, Output prüfen
2. **Unit Tests**: `testthat`-Tests gegen Fixture-DOCX
3. **`devtools::check()`**: Package-Check bestehen
4. **Praxistest**: Mit einer echten gegengelesenen DOCX aus Pauls Workflow testen

## Offene Punkte / Spätere Erweiterungen

- Verschachtelte Tracked Changes (Insertion innerhalb einer Deletion) — Erstversion: flach behandeln
- Formatierungs-Changes (w:rPrChange) — Erstversion: nur Text-Insertions/Deletions
- Mehrere Reviewer: Gruppierung nach Autor optional
- Export als tibble statt Markdown (für programmatische Weiterverarbeitung)
