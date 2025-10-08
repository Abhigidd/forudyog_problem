# forUdyog-assignment

Small extractor project to download PDFs, extract text, parse key fields and produce two Excel outputs: structuredJSON.xlsx and mapped_output.xlsx.

## Files in this pack
- extractor.py — main script (download -> extract -> parse -> write Excel)
- requirements.txt — Python deps
- notes.md — short testing, scaling and architecture notes
- outputs/ — folder where results are written (you will generate these by running the script)

## Sample test links (used for sample output)
Place an input Excel input_links.xlsx with column named pdf containing these links (one per row):

https://bidplus.gem.gov.in/showbidDocument/8434194
https://bidplus.gem.gov.in/showbidDocument/8434006
https://bidplus.gem.gov.in/showbidDocument/8434066
https://bidplus.gem.gov.in/showbidDocument/8434254
https://bidplus.gem.gov.in/showbidDocument/8302695
https://bidplus.gem.gov.in/showbidDocument/8303339
https://bidplus.gem.gov.in/showbidDocument/8323763
https://bidplus.gem.gov.in/showbidDocument/8330005
https://bidplus.gem.gov.in/showbidDocument/8338078
https://bidplus.gem.gov.in/showbidDocument/8296827
https://bidplus.gem.gov.in/showbidDocument/8362358
https://bidplus.gem.gov.in/showbidDocument/8362269
https://bidplus.gem.gov.in/showbidDocument/8362198
https://bidplus.gem.gov.in/showbidDocument/8358464
https://bidplus.gem.gov.in/showbidDocument/8257896

## How to run (Linux / WSL / macOS)
1. Create virtualenv: python3 -m venv venv && source venv/bin/activate
2. Install: pip install -r requirements.txt
   * If you plan OCR, install system packages: sudo apt install poppler-utils tesseract-ocr (Debian/Ubuntu)
3. Prepare input_links.xlsx with a single column named pdf and paste sample links or your own.
4. Run: python extractor.py input_links.xlsx --workers 6
5. Outputs appear in outputs/structuredJSON.xlsx (sheet structuredJSON) and outputs/mapped_output.xlsx (sheet mapped).

## Notes for submission
- Include this README, extractor.py, requirements.txt, notes.md and the outputs/structuredJSON.xlsx (sample) in your zip or repo.