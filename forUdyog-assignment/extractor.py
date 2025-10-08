#!/usr/bin/env python3
"""
extractor.py
Usage: python extractor.py input_links.xlsx --workers 6

Reads an Excel with column pdf containing URLs, downloads PDFs, extracts text,
parses fields heuristically, and writes two Excel outputs under outputs/.
"""
import os
import re
import json
import time
import argparse
from io import BytesIO
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Optional

import requests
import pandas as pd
from dateutil import parser as dateparser
from tqdm import tqdm

# Text extraction libs (imported lazily where used to avoid heavy deps on help)
_HAS_PYMUPDF = False
_HAS_PDFPLUMBER = False
_HAS_OCR = False
try:
    import fitz  # PyMuPDF
    _HAS_PYMUPDF = True
except Exception:
    _HAS_PYMUPDF = False
try:
    import pdfplumber
    _HAS_PDFPLUMBER = True
except Exception:
    _HAS_PDFPLUMBER = False

# OCR dependencies will be imported when OCR fallback is invoked

# === Config ===
OUT_DIR = "outputs"
RAW_DIR = "data/raw"
os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(RAW_DIR, exist_ok=True)

# === Helpers ===

def download_pdf(url, dest_folder=RAW_DIR, timeout=30):
    """Download a PDF and return local path or raise."""
    r = requests.get(url, timeout=timeout)
    r.raise_for_status()
    name = url.rstrip("/\n").split("/")[-1] or str(int(time.time() * 1000))
    if not name.lower().endswith(".pdf"):
        name = name + ".pdf"
    path = os.path.join(dest_folder, name)
    with open(path, "wb") as f:
        f.write(r.content)
    return path


def extract_text_pymupdf(path: str) -> str:
    if not _HAS_PYMUPDF:
        return ""
    try:
        doc = fitz.open(path)
        pages = [p.get_text("text") for p in doc]
        return "\n".join(pages).strip()
    except Exception:
        return ""


def extract_text_pdfplumber(path: str) -> str:
    if not _HAS_PDFPLUMBER:
        return ""
    try:
        texts = []
        with pdfplumber.open(path) as pdf:
            for p in pdf.pages:
                texts.append(p.extract_text() or "")
        return "\n".join(texts).strip()
    except Exception:
        return ""


def ocr_pdf_bytes(pdf_bytes: bytes, max_pages: Optional[int] = None) -> str:
    # Import OCR libs lazily to avoid requiring poppler/tesseract at top-level
    global _HAS_OCR
    try:
        from pdf2image import convert_from_bytes
        import pytesseract
        _HAS_OCR = True
    except Exception:
        _HAS_OCR = False
        return ""

    images = convert_from_bytes(pdf_bytes, dpi=200)
    if max_pages:
        images = images[:max_pages]
    texts = [pytesseract.image_to_string(img) for img in images]
    return "\n".join(texts)


# === Simple heuristic parsing ===

def find_first(regexes, text):
    for rx in regexes:
        m = re.search(rx, text, re.IGNORECASE | re.DOTALL)
        if m:
            # return first non-empty group or whole match
            for g in m.groups():
                if g:
                    return g.strip()
            return m.group(0).strip()
    return None


def extract_fields(text):
    d = {
        "bid_number": None,
        "bid_date": None,
        "ministry": None,
        "department": None,
        "organisation": None,
        "item_category": None,
        "quantity": None,
        "estimated_value_in_inr": None,
        "end_datetime": None,
        "open_datetime": None,
        "validity_days": None,
        "type_of_bid": None,
        "reverse_auction": None,
        "emd_amount_in_inr": None,
        "epbg_percent": None,
        "epbg_months": None,
        "mse_exemption": None,
        "startup_exemption": None,
        "mii_purchase_preference": None,
        "mse_purchase_preference": None,
        "evaluation_method": None,
        "prebid": {"datetime": None, "venue": None},
        "delivery": {"qty": None, "days": None, "consignee": None, "address": None},
    }

    d["bid_number"] = find_first([
        r"Bid\s*No[:\s]([A-Za-z0-9\-/]+)",
        r"Bid Number[:\s]([A-Za-z0-9\-/]+)",
        r"Tender No[:\s]([A-Za-z0-9\-/]+)",
        r"Tender No\.[:\s]([A-Za-z0-9\-/]+)"
    ], text)

    bd = find_first([
        r"Bid Date[:\s]([0-9]{1,2}[\-/]\d{1,2}[\-/]\d{2,4})",
        r"Date[:\s]([0-9]{1,2}\s+\w+\s+\d{4})",
        r"Date of publication[:\s]([0-9]{1,2}[\s\w\-\/,]+)"
    ], text)
    if bd:
        try:
            d["bid_date"] = dateparser.parse(bd, dayfirst=True).date().isoformat()
        except Exception:
            d["bid_date"] = bd

    est = find_first([r"Estimated Value[:\s]([₹Rs\.,\s0-9A-Za-z/-]+)", r"Estimated Cost[:\s]([₹Rs\.,\s0-9A-Za-z/-]+)", r"Estimated Contract Value[:\s]([₹Rs\.,\s0-9A-Za-z/-]+)"], text)
    if est:
        d["estimated_value_in_inr"] = est

    qty = find_first([r"Quantity[:\s]([0-9,]+)", r"Qty[:\s]([0-9,]+)", r"Quantity\(Nos\)[:\s]*([0-9,]+)", r"Total Quantity[:\s]*([0-9,]+)"], text)
    if qty:
        try:
            d["quantity"] = int(qty.replace(",", ""))
        except Exception:
            d["quantity"] = qty

    end_dt = find_first([r"End Date[:\s]([0-9A-Za-z:,\- ]+)", r"Closing Date[:\s]([0-9A-Za-z:,\- ]+)", r"Submission Ends[:\s]*([0-9A-Za-z:,\- ]+)", r"Bid Submission End Date[:\s]*([0-9A-Za-z:,\- ]+)"] , text)
    if end_dt:
        d["end_datetime"] = end_dt

    open_dt = find_first([r"Opening Date[:\s]([0-9A-Za-z:,\- ]+)", r"Bid Opening[:\s]([0-9A-Za-z:,\- ]+)", r"Bid Opening Date[:\s]*([0-9A-Za-z:,\- ]+)"] , text)

    # Additional heuristics
    d["ministry"] = find_first([r"Ministry[:\s]*(.+)", r"Ministry/Department[:\s]*(.+)"], text)
    d["department"] = find_first([r"Department[:\s]*(.+)", r"Dept[:\s]*(.+)"], text)
    d["item_category"] = find_first([r"Category[:\s]*(.+)", r"Item Category[:\s]*(.+)", r"Item\(s\)[:\s]*(.+)"], text)
    d["type_of_bid"] = find_first([r"Type of Bid[:\s]*(.+)", r"Bid Type[:\s]*(.+)"], text)

    # EMD / EPBG
    emd = find_first([r"EMD[:\s]*([₹Rs\.,\s0-9A-Za-z/-]+)", r"EMD Amount[:\s]*([₹Rs\.,\s0-9A-Za-z/-]+)"], text)
    if emd:
        try:
            d["emd_amount_in_inr"] = float(re.sub(r"[^0-9.]", "", emd))
        except Exception:
            d["emd_amount_in_inr"] = emd

    epbg = find_first([r"EPBG[:\s]*([0-9\.]+)%", r"EPBG[:\s]*([0-9\.]+) percent"], text)
    if epbg:
        try:
            d["epbg_percent"] = float(epbg)
        except Exception:
            d["epbg_percent"] = epbg

    epbg_months = find_first([r"EPBG Months[:\s]*([0-9]+)", r"EPBG[:\s]*([0-9]+) months"], text)
    if epbg_months:
        try:
            d["epbg_months"] = int(epbg_months)
        except Exception:
            d["epbg_months"] = epbg_months

    # boolean flags
    d["reverse_auction"] = bool(find_first([r"Reverse Auction[:\s]*(Yes|No)", r"Reverse Auction"], text))
    d["mse_exemption"] = bool(find_first([r"MSE Exemption[:\s]*(Yes|No)", r"MSE Exemption"], text))
    d["startup_exemption"] = bool(find_first([r"Startup Exemption[:\s]*(Yes|No)", r"Startup Exemption"], text))
    d["mii_purchase_preference"] = bool(find_first([r"MII Purchase Preference[:\s]*(Yes|No)", r"MII Purchase Preference"], text))
    d["mse_purchase_preference"] = bool(find_first([r"MSE Purchase Preference[:\s]*(Yes|No)", r"MSE Purchase Preference"], text))

    # prebid and delivery
    pre_dt = find_first([r"Pre-bid Meeting[:\s]*Date[:\s]*([0-9A-Za-z:,\- ]+)", r"Pre-bid Meeting Date[:\s]*([0-9A-Za-z:,\- ]+)"], text)
    if pre_dt:
        d["prebid"]["datetime"] = pre_dt
    pre_venue = find_first([r"Pre-bid Venue[:\s]*(.+)", r"Pre-bid Meeting Venue[:\s]*(.+)", r"Pre-bid Meeting[:\s]*Venue[:\s]*(.+)"], text)
    if pre_venue:
        d["prebid"]["venue"] = pre_venue

    delivery_qty = find_first([r"Delivery Qty[:\s]*([0-9,]+)", r"Delivery Quantity[:\s]*([0-9,]+)", r"Qty to be supplied[:\s]*([0-9,]+)"], text)
    if delivery_qty:
        try:
            d["delivery"]["qty"] = int(delivery_qty.replace(",", ""))
        except Exception:
            d["delivery"]["qty"] = delivery_qty
    delivery_days = find_first([r"Delivery Period[:\s]*([0-9]+) days", r"Delivery within[:\s]*([0-9]+) days"], text)
    if delivery_days:
        try:
            d["delivery"]["days"] = int(re.sub(r"[^0-9]", "", delivery_days))
        except Exception:
            d["delivery"]["days"] = delivery_days
    consignee = find_first([r"Consignee[:\s]*(.+)", r"Delivery Address[:\s]*(.+)", r"Consignee Name[:\s]*(.+)"], text)
    if consignee:
        d["delivery"]["consignee"] = consignee
    addr = find_first([r"Address[:\s]*(.+)", r"Delivery Address[:\s]*(.+)"], text)
    if addr:
        d["delivery"]["address"] = addr
    if open_dt:
        d["open_datetime"] = open_dt

    return d


# === Process single URL ===

def process_url(url, ocr_fallback=True, ocr_max_pages=2):
    rec = {"url": url, "pdf_path": None, "json": None, "error": None}
    try:
        path = download_pdf(url)
        rec["pdf_path"] = path
    except Exception as e:
        rec["error"] = f"download_error:{e}"
        return rec

    text = extract_text_pymupdf(path)
    if not text:
        text = extract_text_pdfplumber(path)
    if not text and ocr_fallback:
        try:
            with open(path, "rb") as f:
                pdf_bytes = f.read()
            text = ocr_pdf_bytes(pdf_bytes, max_pages=ocr_max_pages)
        except Exception as e:
            rec["error"] = f"ocr_error:{e}"
            return rec

    parsed = extract_fields(text or "")
    rec["json"] = parsed
    return rec


# === Main ===

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input_excel", help="input Excel file with a column named 'pdf'", nargs="?", default="input_links.xlsx")
    ap.add_argument("--workers", type=int, default=6)
    args = ap.parse_args()

    df = pd.read_excel(args.input_excel)
    if "pdf" not in df.columns:
        raise SystemExit("Input excel must contain column named 'pdf'")
    urls = df["pdf"].dropna().astype(str).tolist()

    results = []
    with ThreadPoolExecutor(max_workers=args.workers) as exe:
        futures = {exe.submit(process_url, u): u for u in urls}
        for fut in tqdm(as_completed(futures), total=len(futures)):
            results.append(fut.result())

    # build outputs
    out_rows = []
    mapped_rows = []
    for r in results:
        out_rows.append({"url": r["url"], "json": json.dumps(r.get("json"), ensure_ascii=False)})
        mr = r.get("json") or {}
        mr.update({"url": r.get("url"), "pdf_path": r.get("pdf_path"), "error": r.get("error")})
        mapped_rows.append(mr)

    df_out = pd.DataFrame(out_rows)
    df_mapped = pd.DataFrame(mapped_rows)

    # Sanitize values to avoid illegal Excel characters (openpyxl raises on certain control chars)
    def sanitize_string(s: str) -> str:
        if not isinstance(s, str):
            return s
        # remove control chars except tab/newline/carriage return
        return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", " ", s)

    def clean_dataframe_for_excel(df: pd.DataFrame) -> pd.DataFrame:
        df2 = df.copy()
        for col in df2.columns:
            def clean_cell(v):
                if v is None:
                    return v
                # convert complex types to JSON string
                if isinstance(v, (dict, list)):
                    try:
                        return json.dumps(v, ensure_ascii=False)
                    except Exception:
                        return str(v)
                # convert bytes
                if isinstance(v, (bytes, bytearray)):
                    try:
                        v = v.decode('utf-8', errors='ignore')
                    except Exception:
                        v = str(v)
                # sanitize strings
                if isinstance(v, str):
                    return sanitize_string(v)
                return v

            df2[col] = df2[col].apply(clean_cell)
        return df2

    df_out_clean = clean_dataframe_for_excel(df_out)
    df_mapped_clean = clean_dataframe_for_excel(df_mapped)

    df_out_clean.to_excel(os.path.join(OUT_DIR, "structuredJSON.xlsx"), sheet_name="structuredJSON", index=False)
    df_mapped_clean.to_excel(os.path.join(OUT_DIR, "mapped_output.xlsx"), sheet_name="mapped", index=False)
    print("Done. Outputs written to outputs/")


if __name__ == "__main__":
    main()