import csv
import os
import re
import subprocess
import uuid
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape

from flask import Flask, Response, jsonify, render_template_string, request, send_file, url_for
from PIL import Image, ImageOps


BASE_DIR = Path(__file__).resolve().parent
REFERENCE_DIR = Path(r"C:\Users\ddoch\Desktop\지출결의서")
TEMPLATE_XLSX = REFERENCE_DIR / "4월" / "26년_04월_지출결의서_정시내.xlsx"
WORK_DIR = BASE_DIR / "work"
OUTPUT_DIR = BASE_DIR / "generated"
OCR_SCRIPT = BASE_DIR / "ocr_windows.ps1"
OCR_LOG_PATH = BASE_DIR / "ocr_debug.log"
PADDLE_CACHE_DIR = BASE_DIR / "paddle_cache" / "paddlex"
PADDLE_OCR = None

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
ET.register_namespace("", NS["m"])

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tif", ".tiff"}
UPLOAD_EXTS = IMAGE_EXTS | {".pdf"}
RESULT_NAME_MARKERS = ("지출결의서", "지출증빙", "거래목록", "검토용")

ACCOUNT_ROWS = {
    10: 18,
    29: 37,
    36: 44,
    48: 56,
    65: 73,
}

ACCOUNT_LABELS = {
    10: "회의식대",
    29: "우편",
    36: "인쇄비",
    48: "사무용품비",
    65: "재료비",
}

PURPOSES = {
    10: "회의식대(중식)",
    29: "등기우편",
    36: "인쇄비",
    48: "사무용품",
    65: "정기결제비용",
}


@dataclass
class Entry:
    date: str
    account_no: int
    account: str
    amount: int
    purpose: str
    attendee: str
    user: str
    vendor: str
    filename: str
    ocr_amount: int = 0
    ocr_text: str = ""


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 300 * 1024 * 1024


def ensure_dirs():
    WORK_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)


def clean_name(name):
    name = Path(name).name
    return re.sub(r'[<>:"/\\|?*]', "_", name)


def receipt_key(filename):
    path = Path(filename)
    if path.suffix.lower() not in UPLOAD_EXTS:
        return None
    stem = path.stem
    if any(marker in stem for marker in RESULT_NAME_MARKERS):
        return None
    match = re.match(r"^(\d{4})[_\-\s]*(.+)$", stem)
    if not match:
        return None
    mmdd, vendor = match.groups()
    month = int(mmdd[:2])
    day = int(mmdd[2:])
    if not 1 <= month <= 12 or not 1 <= day <= 31:
        return None
    vendor = re.sub(r"\s+", "", vendor).casefold()
    return f"{mmdd}_{vendor}"


def prefer_receipt(existing_name, new_name):
    existing_ext = Path(existing_name).suffix.lower()
    new_ext = Path(new_name).suffix.lower()
    if existing_ext == ".pdf" and new_ext in IMAGE_EXTS:
        return True
    return False


def infer_account(vendor):
    text = vendor.lower()
    if any(k in text for k in ["chatgpt", "챗", "claude", "클로드", "cursor", "커서", "kiri"]):
        return 65
    if any(k in text for k in ["office", "오피스", "문구", "사무"]):
        return 48
    if any(k in text for k in ["소프트웨어", "인쇄", "출력", "프린트"]):
        return 36
    if any(k in text for k in ["등기", "우편"]):
        return 29
    return 10


def parse_file(filename, year, month):
    stem = Path(filename).stem
    match = re.search(r"(?<!\d)(\d{4})(?!\d)", stem)
    day = 1
    vendor = stem
    if match:
        mmdd = match.group(1)
        day = int(mmdd[2:])
        vendor = stem[match.end():].strip(" _-")
    vendor = vendor or stem
    account_no = infer_account(vendor)
    purpose = PURPOSES[account_no]
    if account_no == 65 and vendor:
        purpose = f"{vendor.upper()} {purpose}"
    return Entry(
        date=f"{year:04d}.{month:02d}.{day:02d}",
        account_no=account_no,
        account=ACCOUNT_LABELS[account_no],
        amount=0,
        purpose=purpose,
        attendee="정시내",
        user="정시내(개인)",
        vendor=vendor,
        filename=filename,
    )


def run_windows_ocr(path):
    if not OCR_SCRIPT.exists():
        return ""
    completed = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(OCR_SCRIPT),
            "-Path",
            str(path),
        ],
        cwd=str(BASE_DIR),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        timeout=45,
        check=False,
    )
    output = completed.stdout
    for encoding in ("utf-8-sig", "cp949", "utf-16"):
        try:
            return output.decode(encoding).strip()
        except UnicodeDecodeError:
            continue
    return output.decode("utf-8", errors="ignore").strip()


def get_paddle_ocr():
    global PADDLE_OCR
    if PADDLE_OCR is not None:
        return PADDLE_OCR
    os.environ.setdefault("HOME", str(BASE_DIR))
    os.environ.setdefault("USERPROFILE", str(BASE_DIR))
    os.environ.setdefault("PADDLE_PDX_CACHE_HOME", str(PADDLE_CACHE_DIR))
    os.environ.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")
    from paddleocr import PaddleOCR

    PADDLE_OCR = PaddleOCR(
        lang="korean",
        text_detection_model_name="PP-OCRv5_mobile_det",
        text_recognition_model_name="korean_PP-OCRv5_mobile_rec",
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=False,
    )
    return PADDLE_OCR


def run_paddle_ocr(path):
    try:
        ocr = get_paddle_ocr()
        result = ocr.predict(str(path))
        texts = []
        for page in result:
            data = getattr(page, "json", None)
            if data:
                texts.extend(data.get("res", {}).get("rec_texts", []))
            elif isinstance(page, dict):
                texts.extend(page.get("rec_texts", []))
        return " ".join(str(text) for text in texts if text)
    except Exception as exc:
        return ""


def normalize_money_token(token):
    token = token.strip()
    if re.search(r"[.,]\d{1,2}$", token):
        base, decimals = re.split(r"[.,](?=\d{1,2}$)", token, maxsplit=1)
        if decimals != "00":
            return None
        token = base
    if re.search(r"[.,]\d{3}(?:[.,]\d{3})*$", token):
        value = re.sub(r"[.,]", "", token)
    else:
        value = re.sub(r"\D", "", token)
    if not value:
        return None
    amount = int(value)
    if amount < 1000 or amount > 500000:
        return None
    return amount


def has_number_context(text, start, end):
    prefix = text[max(0, start - 24): start]
    if re.search(r"(?:KRW|원)\s*$", prefix, flags=re.IGNORECASE):
        return False
    return bool(re.search(r"(?:가맹점|승인|사업자|카드|회원)\s*번호\s*$", prefix))


def first_money_after_label(text, label_pattern, max_chars=80):
    labels = list(re.finditer(label_pattern, text, flags=re.IGNORECASE))
    label = None
    for candidate in labels:
        prefix = text[max(0, candidate.start() - 16): candidate.start()]
        if re.search(r"금\s*액\s*부\s*가\s*세\s*$", prefix):
            continue
        label = candidate
        break
    if not label:
        return 0
    window = text[label.end(): label.end() + max_chars]
    window_start = label.end()
    for match in re.finditer(r"\d{1,3}(?:[,.]\d{3})+(?:[,.]00)?", window):
        start = window_start + match.start()
        end = window_start + match.end()
        if has_number_context(text, start, end):
            continue
        amount = normalize_money_token(match.group(0))
        if amount:
            return amount
    for match in re.finditer(r"(\d{4,9})\s*(?:원|\(?KRW\)?)", window, flags=re.IGNORECASE):
        start = window_start + match.start(1)
        end = window_start + match.end(1)
        if has_number_context(text, start, end):
            continue
        amount = normalize_money_token(match.group(1))
        if amount:
            return amount
    return 0


def money_candidates(text, start=0, max_chars=None):
    end = len(text) if max_chars is None else min(len(text), start + max_chars)
    window = text[start:end]
    candidates = []
    patterns = (
        r"(\d{1,3}(?:[,.]\d{3})+(?:[,.]00)?)\s*(?:원|\(?KRW\)?)?",
        r"(\d{4,9})\s*(?:원|\(?KRW\)?)",
    )
    for pattern in patterns:
        for match in re.finditer(pattern, window, flags=re.IGNORECASE):
            abs_start = start + match.start(1)
            abs_end = start + match.end(1)
            if has_number_context(text, abs_start, abs_end):
                continue
            amount = normalize_money_token(match.group(1))
            if amount:
                candidates.append((abs_start, amount))
    candidates.sort(key=lambda item: item[0])
    return [amount for _, amount in candidates]


def total_from_amount_vat_total(text):
    label = re.search(r"금\s*액\s*부\s*가\s*세\s*합\s*계", text)
    if not label:
        return 0
    window_start = label.end()
    window = text[window_start: window_start + 120]
    three_amounts = re.search(
        r"(\d{1,3}(?:[,.]\d{3})+|\d{4,9})\s*(?:원)?\s+"
        r"(\d{1,3}(?:[,.]\d{3})+|\d{3,9})\s*(?:원)?\s+"
        r"(\d{1,3}(?:[,.]\d{3})+|\d{4,9})\s*(?:원|\(?KRW\)?)?",
        window,
        flags=re.IGNORECASE,
    )
    if three_amounts:
        total = normalize_money_token(three_amounts.group(3))
        if total:
            return total
    amounts = money_candidates(text, window_start, max_chars=120)
    if amounts:
        return amounts[-1]
    return 0


def direct_krw_amount(text):
    for pattern in (
        r"KRW\s*(\d{1,3}(?:[,.]\d{3})+(?:[,.]00)?|\d{4,9})",
        r"(\d{1,3}(?:[,.]\d{3})+(?:[,.]00)?|\d{4,9})\s*\(?KRW\)?",
    ):
        for match in re.finditer(pattern, text, flags=re.IGNORECASE):
            if has_number_context(text, match.start(1), match.end(1)):
                continue
            amount = normalize_money_token(match.group(1))
            if amount:
                return amount
    return 0


def extract_ocr_amount(text):
    amount, _ = explain_ocr_amount(text)
    return amount


def explain_ocr_amount(text):
    if not text:
        return 0, "empty_text"
    total_won = first_money_after_label(text, r"총\s*결제\s*원화\s*금액")
    if total_won:
        return total_won, "총 결제 원화 금액"
    vat_total = total_from_amount_vat_total(text)
    if vat_total:
        return vat_total, "금액/부가세/합계 표"
    total = first_money_after_label(text, r"합\s*계")
    if total:
        return total, "합계"
    krw_amount = direct_krw_amount(text)
    if krw_amount:
        return krw_amount, "KRW"
    return 0, "no_match"


def write_ocr_log(entry, path, source, text, amount, rule):
    try:
        OCR_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with OCR_LOG_PATH.open("a", encoding="utf-8") as log:
            log.write("\n" + "=" * 80 + "\n")
            log.write(f"time: {timestamp}\n")
            log.write(f"file: {path.name}\n")
            log.write(f"vendor: {entry.vendor}\n")
            log.write(f"source: {source}\n")
            log.write(f"selected_amount: {amount}\n")
            log.write(f"rule: {rule}\n")
            log.write("ocr_text:\n")
            log.write((text or "").strip() + "\n")
    except Exception:
        pass


def apply_business_rules(entry, amount):
    if amount <= 0:
        return 0
    if entry.account_no == 10:
        return min(amount, 10000)
    return amount


def apply_ocr(entry, path):
    source = "paddleocr"
    text = run_paddle_ocr(path)
    if not text:
        source = "windows_ocr"
        text = run_windows_ocr(path)
    amount, rule = explain_ocr_amount(text)
    entry.ocr_amount = amount
    entry.ocr_text = text[:1000]
    entry.amount = apply_business_rules(entry, amount)
    write_ocr_log(entry, path, source, text, amount, rule)
    return entry


def read_shared_strings(zf):
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    strings = []
    for si in root.findall("m:si", NS):
        strings.append("".join(t.text or "" for t in si.findall(".//m:t", NS)))
    return strings


def shared_string_xml(strings):
    items = []
    for value in strings:
        preserve = ' xml:space="preserve"' if value != value.strip() else ""
        items.append(f"<si><t{preserve}>{escape(value)}</t></si>")
    body = "".join(items)
    count = len(strings)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<sst xmlns="{NS["m"]}" count="{count}" uniqueCount="{count}">{body}</sst>'
    ).encode("utf-8")


def column_row(ref):
    match = re.match(r"([A-Z]+)(\d+)", ref)
    return match.group(1), int(match.group(2))


def get_or_create_cell(sheet, row_el, ref):
    for cell in row_el.findall("m:c", NS):
        if cell.attrib.get("r") == ref:
            return cell
    cell = ET.Element(f"{{{NS['m']}}}c", {"r": ref})
    row_el.append(cell)
    return cell


def get_or_create_row(sheet, row_no):
    sheet_data = sheet.find("m:sheetData", NS)
    for row in sheet_data.findall("m:row", NS):
        if int(row.attrib.get("r", 0)) == row_no:
            return row
    row = ET.Element(f"{{{NS['m']}}}row", {"r": str(row_no)})
    sheet_data.append(row)
    return row


def clear_cell(cell):
    for child in list(cell):
        if child.tag.endswith("}v") or child.tag.endswith("}is"):
            cell.remove(child)
    cell.attrib.pop("t", None)


def set_cell(cell, value, strings):
    clear_cell(cell)
    if isinstance(value, int):
        v = ET.SubElement(cell, f"{{{NS['m']}}}v")
        v.text = str(value)
        return
    if value is None:
        return
    strings.append(str(value))
    cell.attrib["t"] = "s"
    v = ET.SubElement(cell, f"{{{NS['m']}}}v")
    v.text = str(len(strings) - 1)


def update_xlsx(entries, year, month, output_path):
    with zipfile.ZipFile(TEMPLATE_XLSX, "r") as zin:
        strings = read_shared_strings(zin)
        sheet = ET.fromstring(zin.read("xl/worksheets/sheet2.xml"))

        set_cell(get_or_create_cell(sheet, get_or_create_row(sheet, 2), "B2"), f"{year}년 {month}월분 법인(개인) 카드 정산서", strings)

        for row_no in range(9, 77):
            row = get_or_create_row(sheet, row_no)
            for col in ["B", "C", "D", "E", "F", "G", "H"]:
                clear_cell(get_or_create_cell(sheet, row, f"{col}{row_no}"))
            n_cell = get_or_create_cell(sheet, row, f"N{row_no}")
            if row_no in ACCOUNT_ROWS.values() or row_no == 77:
                set_cell(n_cell, 0, strings)

        totals = {key: 0 for key in ACCOUNT_ROWS}
        for index, entry in enumerate(entries, start=9):
            if index > 76:
                break
            row = get_or_create_row(sheet, index)
            account_no = int(entry["account_no"])
            amount = int(entry["amount"] or 0)
            totals[account_no] = totals.get(account_no, 0) + amount
            values = {
                "B": entry["date"],
                "C": account_no,
                "D": ACCOUNT_LABELS.get(account_no, entry.get("account", "")),
                "E": amount,
                "F": entry.get("purpose", ""),
                "G": entry.get("attendee", "정시내"),
                "H": entry.get("user", "정시내(개인)"),
            }
            for col, value in values.items():
                set_cell(get_or_create_cell(sheet, row, f"{col}{index}"), value, strings)

        grand_total = sum(totals.values())
        for account_no, row_no in ACCOUNT_ROWS.items():
            set_cell(get_or_create_cell(sheet, get_or_create_row(sheet, row_no), f"N{row_no}"), totals.get(account_no, 0), strings)
        set_cell(get_or_create_cell(sheet, get_or_create_row(sheet, 77), "E77"), grand_total, strings)
        set_cell(get_or_create_cell(sheet, get_or_create_row(sheet, 77), "N77"), grand_total, strings)

        sheet_xml = ET.tostring(sheet, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "xl/worksheets/sheet2.xml":
                    zout.writestr(item, sheet_xml)
                elif item.filename == "xl/sharedStrings.xml":
                    zout.writestr(item, shared_string_xml(strings))
                else:
                    zout.writestr(item, zin.read(item.filename))


def make_images_pdf(files, pdf_path):
    pages = []
    for file_path in files:
        if file_path.suffix.lower() not in IMAGE_EXTS:
            continue
        with Image.open(file_path) as image:
            img = ImageOps.exif_transpose(image).convert("RGB")
            canvas = Image.new("RGB", (1240, 1754), "white")
            img.thumbnail((1120, 1634), Image.Resampling.LANCZOS)
            x = (canvas.width - img.width) // 2
            y = (canvas.height - img.height) // 2
            canvas.paste(img, (x, y))
            pages.append(canvas)
    if not pages:
        return False
    first, rest = pages[0], pages[1:]
    first.save(pdf_path, "PDF", resolution=150, save_all=True, append_images=rest)
    return True


def merge_pdfs(pdf_paths, output_path):
    object_pattern = re.compile(rb"(?ms)^(\d+)\s+(\d+)\s+obj\s*(.*?)\s*^endobj")
    ref_pattern = re.compile(rb"(?<!\d)(\d+)\s+(\d+)\s+R")
    copied = []
    page_keys = []
    next_id = 3

    for pdf_path in pdf_paths:
        raw = pdf_path.read_bytes()
        matches = list(object_pattern.finditer(raw))
        if not matches:
            raise ValueError(f"PDF 구조를 읽을 수 없습니다: {pdf_path.name}")

        objects = {(int(m.group(1)), int(m.group(2))): m.group(3) for m in matches}
        excluded = set()
        for key, body in objects.items():
            if re.search(rb"/Type\s*/Catalog\b", body) or re.search(rb"/Type\s*/Pages\b", body):
                excluded.add(key)

        mapping = {}
        for key, body in objects.items():
            if key in excluded:
                continue
            mapping[key] = next_id
            next_id += 1
            if re.search(rb"/Type\s*/Page\b", body):
                page_keys.append((pdf_path, key))

        def replace_ref(match):
            key = (int(match.group(1)), int(match.group(2)))
            if key in mapping:
                return f"{mapping[key]} 0 R".encode("ascii")
            return match.group(0)

        for key, body in objects.items():
            if key in excluded:
                continue
            body = ref_pattern.sub(replace_ref, body)
            if re.search(rb"/Type\s*/Page\b", body):
                body = re.sub(rb"/Parent\s+\d+\s+\d+\s+R", b"/Parent 2 0 R", body)
            copied.append((mapping[key], body))

    page_ids = []
    copied_map = dict(copied)
    for _, key in page_keys:
        for new_id, body in copied:
            if new_id in page_ids:
                continue
            if re.search(rb"/Type\s*/Page\b", body):
                page_ids.append(new_id)
                break

    page_ids = [new_id for new_id, body in copied if re.search(rb"/Type\s*/Page\b", body)]
    if not page_ids:
        raise ValueError("병합할 PDF 페이지가 없습니다.")

    kids = b" ".join(f"{page_id} 0 R".encode("ascii") for page_id in page_ids)
    objects_out = [
        (1, b"<< /Type /Catalog /Pages 2 0 R >>"),
        (2, b"<< /Type /Pages /Kids [" + kids + b"] /Count " + str(len(page_ids)).encode("ascii") + b" >>"),
    ]
    objects_out.extend(sorted(copied_map.items()))

    with output_path.open("wb") as out:
        out.write(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
        offsets = {0: 0}
        for obj_id, body in objects_out:
            offsets[obj_id] = out.tell()
            out.write(f"{obj_id} 0 obj\n".encode("ascii"))
            out.write(body)
            out.write(b"\nendobj\n")
        xref_at = out.tell()
        max_id = max(offsets)
        out.write(f"xref\n0 {max_id + 1}\n".encode("ascii"))
        out.write(b"0000000000 65535 f \n")
        for obj_id in range(1, max_id + 1):
            out.write(f"{offsets.get(obj_id, 0):010d} 00000 n \n".encode("ascii"))
        out.write(
            b"trailer\n<< /Size "
            + str(max_id + 1).encode("ascii")
            + b" /Root 1 0 R >>\nstartxref\n"
            + str(xref_at).encode("ascii")
            + b"\n%%EOF\n"
        )


def make_evidence_pdf(files, pdf_path):
    image_files = [path for path in files if path.suffix.lower() in IMAGE_EXTS]
    pdf_files = [path for path in files if path.suffix.lower() == ".pdf"]
    parts = []
    image_pdf = pdf_path.with_name(f"{pdf_path.stem}_images.pdf")
    if make_images_pdf(image_files, image_pdf):
        parts.append(image_pdf)
    parts.extend(pdf_files)
    if not parts:
        raise ValueError("PDF로 만들 이미지 또는 PDF 파일이 없습니다.")
    if len(parts) == 1:
        if parts[0] != pdf_path:
            pdf_path.write_bytes(parts[0].read_bytes())
    else:
        merge_pdfs(parts, pdf_path)
    if image_pdf.exists():
        image_pdf.unlink()


def write_csv(entries, csv_path):
    fieldnames = ["date", "account_no", "account", "amount", "ocr_amount", "purpose", "attendee", "user", "vendor", "filename", "ocr_text"]
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(entries)


def session_paths(session_id):
    root = WORK_DIR / session_id
    return root, root / "uploads"


PAGE = """
<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>지출결의서 자동 생성</title>
  <style>
    :root { color-scheme: light; --line:#d8dee4; --ink:#1f2328; --muted:#667085; --brand:#0f766e; --bg:#f6f8fa; }
    * { box-sizing: border-box; }
    body { margin: 0; font-family: "Malgun Gothic", Arial, sans-serif; color: var(--ink); background: var(--bg); }
    header { background: #fff; border-bottom: 1px solid var(--line); padding: 18px 24px; }
    h1 { margin: 0; font-size: 22px; letter-spacing: 0; }
    main { max-width: 1280px; margin: 0 auto; padding: 20px; }
    section { background: #fff; border: 1px solid var(--line); border-radius: 8px; padding: 16px; margin-bottom: 16px; }
    label { display: block; font-size: 13px; color: var(--muted); margin-bottom: 6px; }
    input, select, button { font: inherit; }
    input, select { width: 100%; border: 1px solid var(--line); border-radius: 6px; padding: 9px 10px; background: #fff; }
    .grid { display: grid; grid-template-columns: 120px 120px 1fr auto; gap: 12px; align-items: end; }
    button { border: 1px solid #0d5f58; border-radius: 6px; padding: 9px 14px; background: var(--brand); color: #fff; cursor: pointer; white-space: nowrap; }
    button.secondary { background: #fff; color: var(--ink); border-color: var(--line); }
    table { width: 100%; border-collapse: collapse; table-layout: fixed; }
    th, td { border-bottom: 1px solid var(--line); padding: 7px; text-align: left; vertical-align: top; }
    th { font-size: 12px; color: var(--muted); background: #f9fafb; }
    td input, td select { padding: 7px; }
    .amount { text-align: right; }
    .actions { display: flex; gap: 8px; justify-content: flex-end; align-items: center; }
    .status { color: var(--muted); font-size: 13px; }
    .totals { display: flex; gap: 14px; margin: 10px 0 4px; font-size: 14px; justify-content: flex-end; }
    .totals strong { font-size: 16px; }
    .downloads a { display: inline-block; margin-right: 10px; color: #075985; }
    @media (max-width: 900px) { .grid { grid-template-columns: 1fr 1fr; } table { min-width: 980px; } .table-wrap { overflow:auto; } }
  </style>
</head>
<body>
  <header><h1>지출결의서 자동 생성</h1></header>
  <main>
    <section>
      <form id="uploadForm" class="grid">
        <div><label>연도</label><input name="year" type="number" value="2026" min="2020" max="2099"></div>
        <div><label>월</label><input name="month" type="number" value="5" min="1" max="12"></div>
        <div><label>매출전표 이미지/PDF 폴더</label><input name="files" type="file" webkitdirectory directory multiple accept="image/*,.pdf,application/pdf"></div>
        <button type="submit">불러오기</button>
      </form>
    </section>
    <section>
      <div class="actions">
        <span id="status" class="status">파일을 불러오면 거래 목록이 표시됩니다.</span>
        <button id="generateBtn" type="button" disabled>엑셀/PDF 생성</button>
      </div>
      <div class="totals">
        <span>전표 원금액 합계 <strong id="ocrTotal">0</strong>원</span>
        <span>지출결의서 합계 <strong id="amountTotal">0</strong>원</span>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th style="width:105px">일자</th><th style="width:110px">계정</th><th style="width:110px">금액</th><th style="width:110px">OCR</th>
              <th>목적</th><th style="width:90px">참석자</th><th style="width:115px">사용자</th><th style="width:150px">가맹점</th><th>파일</th>
            </tr>
          </thead>
          <tbody id="rows"></tbody>
        </table>
      </div>
    </section>
    <section class="downloads" id="downloads" hidden></section>
  </main>
  <script>
    let sessionId = null;
    const accounts = {{ accounts|tojson }};
    const rowsEl = document.getElementById('rows');
    const statusEl = document.getElementById('status');
    const generateBtn = document.getElementById('generateBtn');
    const downloadsEl = document.getElementById('downloads');
    const ocrTotalEl = document.getElementById('ocrTotal');
    const amountTotalEl = document.getElementById('amountTotal');

    document.getElementById('uploadForm').addEventListener('submit', async (event) => {
      event.preventDefault();
      const formEl = event.currentTarget;
      const files = filterReceiptFiles([...formEl.elements.files.files]);
      if (files.length === 0) {
        statusEl.textContent = 'MMDD_가맹점 형식의 이미지/PDF 전표가 없습니다.';
        return;
      }
      const form = new FormData();
      form.append('year', formEl.elements.year.value);
      form.append('month', formEl.elements.month.value);
      files.forEach((file) => form.append('files', file, file.name));
      statusEl.textContent = `${files.length}개 전표만 업로드 중입니다.`;
      try {
        const res = await fetch('/upload', { method: 'POST', body: form });
        const data = await res.json();
        if (!res.ok) { statusEl.textContent = data.error || '업로드 실패'; return; }
        sessionId = data.session_id;
        renderRows(data.entries);
        generateBtn.disabled = true;
        statusEl.textContent = `${data.entries.length}건을 불러왔습니다. OCR을 시작합니다.`;
        downloadsEl.hidden = true;
        runOcr();
      } catch (error) {
        statusEl.textContent = `업로드 실패: ${error.message}`;
      }
    });

    generateBtn.addEventListener('click', async () => {
      const entries = collectRows();
      statusEl.textContent = '생성 중입니다.';
      const res = await fetch('/generate', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ session_id: sessionId, entries })
      });
      const data = await res.json();
      if (!res.ok) { statusEl.textContent = data.error || '생성 실패'; return; }
      statusEl.textContent = `완료: ${entries.length}건`;
      downloadsEl.hidden = false;
      downloadsEl.innerHTML = `<a href="${data.xlsx}">지출결의서 엑셀 다운로드</a><a href="${data.pdf}">지출증빙 PDF 다운로드</a><a href="${data.csv}">검토용 CSV 다운로드</a>`;
    });

    function renderRows(entries) {
      rowsEl.innerHTML = '';
      entries.forEach((entry) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td><input data-k="date" value="${esc(entry.date)}"></td>
          <td><select data-k="account_no">${Object.entries(accounts).map(([no, label]) => `<option value="${no}" ${Number(no) === entry.account_no ? 'selected' : ''}>${no} ${label}</option>`).join('')}</select></td>
          <td><input data-k="amount" class="amount" type="number" min="0" step="1" value="${entry.amount || 0}"></td>
          <td><input data-k="ocr_amount" class="amount" type="number" value="${entry.ocr_amount || 0}" readonly title="${esc(entry.ocr_text || '')}"><input data-k="ocr_text" type="hidden" value="${esc(entry.ocr_text || '')}"></td>
          <td><input data-k="purpose" value="${esc(entry.purpose)}"></td>
          <td><input data-k="attendee" value="${esc(entry.attendee)}"></td>
          <td><input data-k="user" value="${esc(entry.user)}"></td>
          <td><input data-k="vendor" value="${esc(entry.vendor)}"></td>
          <td><input data-k="filename" value="${esc(entry.filename)}" readonly></td>`;
        rowsEl.appendChild(tr);
      });
      rowsEl.querySelectorAll('input, select').forEach((el) => el.addEventListener('input', updateTotals));
      updateTotals();
    }

    function collectRows() {
      return [...rowsEl.querySelectorAll('tr')].map((tr) => {
        const row = {};
        tr.querySelectorAll('[data-k]').forEach((el) => row[el.dataset.k] = el.value);
        row.account_no = Number(row.account_no);
        row.account = accounts[row.account_no];
        row.amount = Number(row.amount || 0);
        row.ocr_amount = Number(row.ocr_amount || 0);
        return row;
      });
    }
    function receiptKey(file) {
      const name = file.name;
      const ext = name.slice(name.lastIndexOf('.')).toLowerCase();
      if (!['.jpg', '.jpeg', '.png', '.bmp', '.webp', '.tif', '.tiff', '.pdf'].includes(ext)) return null;
      const stem = name.slice(0, name.length - ext.length);
      if (['지출결의서', '지출증빙', '거래목록', '검토용'].some((marker) => stem.includes(marker))) return null;
      const match = stem.match(/^(\d{4})[_\-\s]*(.+)$/);
      if (!match) return null;
      const month = Number(match[1].slice(0, 2));
      const day = Number(match[1].slice(2, 4));
      if (month < 1 || month > 12 || day < 1 || day > 31) return null;
      return `${match[1]}_${match[2].replace(/\s+/g, '').toLowerCase()}`;
    }
    function preferFile(existing, next) {
      const existingExt = existing.name.slice(existing.name.lastIndexOf('.')).toLowerCase();
      const nextExt = next.name.slice(next.name.lastIndexOf('.')).toLowerCase();
      const imageExts = ['.jpg', '.jpeg', '.png', '.bmp', '.webp', '.tif', '.tiff'];
      return existingExt === '.pdf' && imageExts.includes(nextExt);
    }
    function filterReceiptFiles(files) {
      const selected = new Map();
      files.forEach((file) => {
        const key = receiptKey(file);
        if (!key) return;
        if (!selected.has(key) || preferFile(selected.get(key), file)) {
          selected.set(key, file);
        }
      });
      return [...selected.values()].sort((a, b) => a.name.localeCompare(b.name, 'ko'));
    }
    async function runOcr() {
      const rowEls = [...rowsEl.querySelectorAll('tr')];
      generateBtn.disabled = true;
      for (let index = 0; index < rowEls.length; index++) {
        statusEl.textContent = `OCR 처리 중입니다. ${index + 1}/${rowEls.length}`;
        const entry = collectRows()[index];
        const res = await fetch('/ocr', {
          method: 'POST',
          headers: {'Content-Type': 'application/json'},
          body: JSON.stringify({ session_id: sessionId, entry })
        });
        const data = await res.json();
        if (!res.ok) {
          statusEl.textContent = data.error || `OCR 실패: ${entry.filename}`;
          continue;
        }
        const tr = rowEls[index];
        tr.querySelector('[data-k="amount"]').value = data.entry.amount || 0;
        tr.querySelector('[data-k="ocr_amount"]').value = data.entry.ocr_amount || 0;
        tr.querySelector('[data-k="ocr_amount"]').title = data.entry.ocr_text || '';
        tr.querySelector('[data-k="ocr_text"]').value = data.entry.ocr_text || '';
        updateTotals();
      }
      statusEl.textContent = `${rowEls.length}건 OCR 완료. 금액을 확인한 뒤 생성하세요.`;
      generateBtn.disabled = rowEls.length === 0;
    }
    function updateTotals() {
      const rows = collectRows();
      const amountTotal = rows.reduce((sum, row) => sum + Number(row.amount || 0), 0);
      const ocrTotal = rows.reduce((sum, row) => sum + Number(row.ocr_amount || 0), 0);
      amountTotalEl.textContent = amountTotal.toLocaleString('ko-KR');
      ocrTotalEl.textContent = ocrTotal.toLocaleString('ko-KR');
    }
    function esc(value) {
      return String(value ?? '').replace(/[&<>"']/g, (c) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    }
  </script>
</body>
</html>
"""


@app.route("/")
def index():
    response = Response(render_template_string(PAGE, accounts={str(k): v for k, v in ACCOUNT_LABELS.items()}))
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    return response


@app.post("/upload")
def upload():
    ensure_dirs()
    year = int(request.form.get("year", datetime.now().year))
    month = int(request.form.get("month", datetime.now().month))
    files = request.files.getlist("files")
    if not files:
        return jsonify(error="업로드할 이미지 폴더를 선택하세요."), 400
    session_id = uuid.uuid4().hex
    root, upload_dir = session_paths(session_id)
    upload_dir.mkdir(parents=True, exist_ok=True)
    selected = {}
    for item in files:
        name = clean_name(item.filename)
        key = receipt_key(name)
        if key is None:
            continue
        if key not in selected or prefer_receipt(selected[key][0], name):
            selected[key] = (name, item)

    saved = []
    for name, item in selected.values():
        target = upload_dir / name
        if target.exists():
            target = upload_dir / f"{target.stem}_{len(saved) + 1}{target.suffix}"
        item.save(target)
        saved.append(target)

    saved.sort(key=lambda p: p.name)
    entries = [parse_file(path.name, year, month).__dict__ for path in saved]
    return jsonify(session_id=session_id, entries=entries)


@app.post("/ocr")
def ocr_entry():
    payload = request.get_json(force=True)
    session_id = payload.get("session_id")
    entry = payload.get("entry", {})
    if not session_id or not re.fullmatch(r"[0-9a-f]{32}", session_id):
        return jsonify(error="세션 정보가 올바르지 않습니다."), 400
    filename = clean_name(entry.get("filename", ""))
    if not filename:
        return jsonify(error="OCR 파일명이 없습니다."), 400
    _, upload_dir = session_paths(session_id)
    path = upload_dir / filename
    if not path.exists():
        return jsonify(error=f"OCR 파일을 찾을 수 없습니다: {filename}"), 404
    parsed = Entry(
        date=entry.get("date", ""),
        account_no=int(entry.get("account_no", 10)),
        account=entry.get("account", ""),
        amount=int(entry.get("amount") or 0),
        purpose=entry.get("purpose", ""),
        attendee=entry.get("attendee", "정시내"),
        user=entry.get("user", "정시내(개인)"),
        vendor=entry.get("vendor", ""),
        filename=filename,
    )
    parsed = apply_ocr(parsed, path)
    return jsonify(entry=parsed.__dict__)


@app.post("/generate")
def generate():
    ensure_dirs()
    payload = request.get_json(force=True)
    session_id = payload.get("session_id")
    entries = payload.get("entries", [])
    if not session_id or not re.fullmatch(r"[0-9a-f]{32}", session_id):
        return jsonify(error="세션 정보가 올바르지 않습니다."), 400
    root, upload_dir = session_paths(session_id)
    if not upload_dir.exists():
        return jsonify(error="업로드 파일을 찾을 수 없습니다. 다시 불러오세요."), 400
    if not TEMPLATE_XLSX.exists():
        return jsonify(error=f"템플릿 파일이 없습니다: {TEMPLATE_XLSX}"), 500
    if len(entries) > 68:
        return jsonify(error="엑셀 템플릿의 거래 입력 가능 행은 최대 68건입니다."), 400

    dates = [e.get("date", "") for e in entries if e.get("date")]
    year, month = 2026, 1
    if dates and re.match(r"\d{4}\.\d{2}\.\d{2}", dates[0]):
        year = int(dates[0][:4])
        month = int(dates[0][5:7])

    out_root = OUTPUT_DIR / session_id
    out_root.mkdir(parents=True, exist_ok=True)
    xlsx_name = f"{str(year)[2:]}년_{month:02d}월_지출결의서_정시내.xlsx"
    pdf_name = f"{str(year)[2:]}년_{month:02d}월_지출결의서_정시내_증빙.pdf"
    csv_name = f"{str(year)[2:]}년_{month:02d}월_거래목록_검토용.csv"
    update_xlsx(entries, year, month, out_root / xlsx_name)
    image_files = sorted(upload_dir.iterdir(), key=lambda p: p.name)
    make_evidence_pdf(image_files, out_root / pdf_name)
    write_csv(entries, out_root / csv_name)
    return jsonify(
        xlsx=url_for("download", session_id=session_id, filename=xlsx_name),
        pdf=url_for("download", session_id=session_id, filename=pdf_name),
        csv=url_for("download", session_id=session_id, filename=csv_name),
    )


@app.get("/download/<session_id>/<path:filename>")
def download(session_id, filename):
    if not re.fullmatch(r"[0-9a-f]{32}", session_id):
        return Response("bad session", status=400)
    path = OUTPUT_DIR / session_id / clean_name(filename)
    if not path.exists():
        return Response("not found", status=404)
    return send_file(path, as_attachment=True, download_name=path.name)


if __name__ == "__main__":
    ensure_dirs()
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", "5000")), debug=False)
