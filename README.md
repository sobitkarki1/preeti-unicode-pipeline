# Preeti batch conversion pipeline

`preeti_pipeline.py` converts legacy Nepali-font documents to Unicode
Devanagari without changing the source files. It uses `npttf2utf` for DOCX,
TXT, and HTML content, and `pdfplumber` plus `npttf2utf` for extracted PDF
text, with `preeti-unicode-converter` as a fallback.

## Install and run

```powershell
python -m pip install -r requirements.txt
python preeti_pipeline.py --input .\docs\ --output .\converted\ --threshold 0.7
python preeti_pipeline.py --manifest .\sources.csv --output .\converted\
```

The input may be a file, a directory (scanned recursively), or a CSV/YAML
manifest. CSV columns can be `path`, `url`, `source`, or `source_url`; an
optional `filename` column controls the downloaded output name. YAML can be a
list or `sources:` list. URLs are downloaded to a temporary directory and are
never modified. A PDF is emitted as `converted/<name>.pdf.txt`, because PDF
layout and font embedding cannot be reliably recreated from extracted text.

DOCX detection inspects run font metadata for Preeti, Sagarmatha, and Kalimati
and conversion uses `DocxHandler(from_font="auto")`, preserving the DOCX
container, tables, shapes, and formatting supported by the library. TXT and
HTML use the configurable n-gram heuristic. HTML `<font face="Preeti">` and
CSS containing `Preeti` are converted even when the text is too short for the
heuristic. Mixed TXT/HTML content is handled line/text-node by line/text-node.

Files are written under the output directory with the same relative filename
where possible. Already-Unicode or undetected files are skipped. Every file is
logged with source, destination, status, and validation warnings. Validation
checks for Devanagari and likely residual Preeti sequences; warnings are
reported rather than silently hiding questionable output. PDF detection first
inspects extracted text so ordinary English or already-Unicode PDFs are
skipped instead of being incorrectly mapped.

## Regression tests

```powershell
pytest -q
```

The fixtures in `test/fixtures` include TXT, HTML, DOCX, and PDF samples.
DOCX/PDF integration tests are skipped automatically when their optional
converters are not installed.
