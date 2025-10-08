Notes for forUdyog-assignment extractor

- This is a small demo extractor. It supports CSV, JSON, and XLSX inputs.
- It will normalize rows into a list of dictionaries and write them to `outputs/structuredJSON.xlsx` and `outputs/structuredJSON.json`.
- If you want PDF or image parsing, we can add OCR later.



# notes.md — quick summary

## What I built
A small pipeline that:
1. Reads an Excel with PDF URLs (column pdf).
2. Downloads each PDF to data/raw/.
3. Tries text extraction using PyMuPDF (fast) -> pdfplumber fallback -> OCR fallback (pdf2image+pytesseract).
4. Parses a small set of fields heuristically into JSON.
5. Writes two Excel outputs: structuredJSON.xlsx (one JSON column) and mapped_output.xlsx (columns per field).

## Sample test
I used the sample links listed in the README. On a 6-worker run typical timings per PDF (text-extract only) were ~1–4s depending on network and PDF complexity; OCR fallback is much slower (10–30s/page). For a conservative 50-sample run expect ~3–6 minutes.

## Limitations
- Parsing uses regex heuristics and will not be perfect for every template.
- OCR is slow and requires system installs.
- LLM-based extraction is not used here due to cost and time constraints.

## Scaling notes (summary)
- For millions of PDFs: use a queue (SQS/Kafka), autoscaled workers (K8s), object store for PDFs (S3), and a human-in-loop step for low-confidence records.
- Use schema validation and a monitoring dashboard (success rates, OCR fallback rate).

## What to include when submitting
- extractor.py, requirements.txt, README.md, notes.md, and a sample outputs/structuredJSON.xlsx with 10–50 processed rows.
