# AGENTS.md

## Cursor Cloud specific instructions

This repository contains a Russian diploma thesis (ВКР) for АО «ВИТА ЛАЙФ». It is a document-processing project, not a traditional application.

### Environment

- Python 3.12 with `python-docx`, `pdf2image`, `Pillow`, `PyPDF2`
- `poppler-utils` (system package for PDF-to-image conversion)
- LibreOffice Writer (for verifying page counts via headless PDF export)

### Working with the thesis document

- The main document is `работа/DIPLOMNAYA_rabota .docx` (note the space before `.docx`)
- Format: Times New Roman 14pt, 1.5 line spacing, margins 3.0/1.5/2.0/2.0 cm
- To verify page count: `libreoffice --headless --convert-to pdf --outdir /tmp "работа/DIPLOMNAYA_rabota .docx"`
- After modifying the document, the Table of Contents (Оглавление) must be updated manually in Word

### Key gotchas

- File names contain Cyrillic characters and spaces — always quote paths
- The document uses custom styles (`Body Text`, `First Paragraph`, `Compact`, `Heading 2`, `Heading 3`) — preserve these when editing programmatically
- Character-based page estimates (~1800 chars/page) are very conservative; LibreOffice PDF conversion gives accurate page counts
- Tables occupy significant vertical space not captured by character counts
