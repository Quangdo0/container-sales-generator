from pathlib import Path
from datetime import datetime, date, timedelta
import re
import copy
import unicodedata
from collections import OrderedDict

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side


# =========================================================
# CONFIG
# These fixed paths are only used when running this file directly by terminal.
# In Streamlit website, uploaded files will be passed into generate_excel().
# =========================================================

BASE_DIR = Path(__file__).parent

TARGET_FILE = BASE_DIR / "KK柜销售单 - BẢNG GIÁ CHI TIẾT CONTAINER T101-200.xlsx"
SOURCE_FILE = BASE_DIR / "BÁO CÁO KHO TỔNG HỢP.xlsx"
TRACKING_FILE = BASE_DIR / "THEODOICONTAINERVANCHUYEN-DESKTOP-3B36BRT.xlsx"

SOURCE_REPORT_SHEET = "BÁO CÁO KHO TỔNG HỢP"
SOURCE_EXPORT_SHEET = "Xuất kho"
TRACKING_SHEET = "MATOU-KK"

RUN_DATE = "01/07/2026"
# RUN_DATE = None

REPLACE_EXISTING_SHEETS = True


# =========================================================
# SOURCE REPORT SHEET STRUCTURE
# B = Xe / ngày / K number
# C = Tổng số thùng
# D = Tổng khối lượng
# =========================================================

REPORT_COL_XE = 2
REPORT_COL_QTY = 3
REPORT_COL_WEIGHT = 4


# =========================================================
# XUẤT KHO SHEET STRUCTURE
# A = Ngày đóng
# B = Mã Container
# C = Chuyến
# D = Ngày xuất phát
# E = Loại sầu riêng
# F = Quy cách
# G = Số KG
# H = Số thùng
# I = Tổng khối lượng (KG)
# =========================================================

XUATKHO_COL_RECEIPT_DATE = 1
XUATKHO_COL_CONTAINER_NO = 2
XUATKHO_COL_TRIP = 3
XUATKHO_COL_FULL_NAME_TRIP = 4
XUATKHO_COL_EXPORT_DATE = 5
XUATKHO_COL_FRUIT_TYPE = 6
XUATKHO_COL_PRODUCT_CODE = 7
XUATKHO_COL_KG = 8
XUATKHO_COL_QTY = 9
XUATKHO_COL_WEIGHT = 10
XUATKHO_COL_TEM = 11


# =========================================================
# TRACKING FILE STRUCTURE
# C = STT, ví dụ KK093
# D = Biển số xe
# E = No Container
# F = No Rơ-moóc
# =========================================================

TRACKING_COL_STT = 3
TRACKING_COL_TRUCK_PLATE = 4
TRACKING_COL_CONTAINER_NO = 5
TRACKING_COL_TRAILER_NO = 6


# =========================================================
# TARGET CELLS
# =========================================================

TARGET_CELLS = {
    "stt_label": "A2",
    "container_no": "B2",
    "truck_plate": "D2",
    "factory": "B3",
    "trailer_no": "D3",
    "packaging": "B4",
    "export_date": "H2",
}

TARGET_ITEM_START_ROW = 7

TARGET_ITEM_COLUMNS = {
    "product_name": "A",
    "qty_boxes": "B",
    "fruit_count": "C",
    "kg_per_box": "D",
    "total_weight": "E",
    "unit_price": "F",
    "amount": "G",
    "receipt_date": "H",
    "item_no": "I",
}

TEMPLATE_ITEM_ROWS = 5


# =========================================================
# FIXED VALUES + STYLE
# =========================================================

FACTORY_VALUE = "林同 \nLâm Đồng"
PACKAGING_VALUE = " 油皮纸箱"
VAT_NOTE = "不包含增值税 Chưa bao gồm VAT"

SONGTI_FONT = "宋体"

ITEM_FILL = PatternFill("solid", fgColor="FFF2CC")
GRAND_TOTAL_FILL = PatternFill("solid", fgColor="FFFF00")

RED_FONT = Font(name=SONGTI_FONT, color="FF0000", bold=True)
BOLD_FONT = Font(name=SONGTI_FONT, color="000000", bold=True)
NORMAL_SONGTI_FONT = Font(name=SONGTI_FONT, color="000000", bold=False)
BLACK_BOLD_FONT = Font(name=SONGTI_FONT, color="000000", bold=True)
BLACK_NORMAL_FONT = Font(name=SONGTI_FONT, color="000000", bold=False)

CENTER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)

THIN_SIDE = Side(style="thin", color="000000")
THIN_BORDER = Border(
    left=THIN_SIDE,
    right=THIN_SIDE,
    top=THIN_SIDE,
    bottom=THIN_SIDE,
)


# =========================================================
# BASIC HELPERS
# =========================================================

def safe_sheet_name(name: str) -> str:
    name = re.sub(r"[\\/*?:\[\]]", "-", str(name))
    return name[:31]


def clean_text(value):
    if value is None:
        return ""
    return str(value).strip()


def remove_vietnamese_accents(text):
    text = str(text)
    normalized = unicodedata.normalize("NFD", text)

    without_accents = "".join(
        char for char in normalized
        if unicodedata.category(char) != "Mn"
    )

    return without_accents


def normalize_sheet_name(name):
    text = remove_vietnamese_accents(name)
    text = text.lower().strip()
    text = re.sub(r"\s+", " ", text)
    return text


def find_sheet_by_keywords(wb, required_keywords, display_name):
    for sheet_name in wb.sheetnames:
        normalized = normalize_sheet_name(sheet_name)

        if all(keyword in normalized for keyword in required_keywords):
            print(f"DEBUG SHEET FOUND for {display_name}: {sheet_name}")
            return wb[sheet_name]

    available_sheets = ", ".join(wb.sheetnames)

    raise ValueError(
        f"Không tìm thấy sheet {display_name}. "
        f"Available sheets: {available_sheets}"
    )


def get_report_sheet(wb):
    return find_sheet_by_keywords(
        wb,
        required_keywords=["bao", "cao", "kho", "tong", "hop"],
        display_name="BÁO CÁO KHO TỔNG HỢP",
    )


def get_xuat_kho_sheet(wb):
    return find_sheet_by_keywords(
        wb,
        required_keywords=["xuat", "kho"],
        display_name="XUẤT KHO",
    )


def get_tracking_sheet(wb):
    if TRACKING_SHEET in wb.sheetnames:
        return wb[TRACKING_SHEET]

    for sheet_name in wb.sheetnames:
        normalized = normalize_sheet_name(sheet_name)

        if "matou" in normalized or "kk" in normalized:
            print(f"DEBUG SHEET FOUND for tracking: {sheet_name}")
            return wb[sheet_name]

    available_sheets = ", ".join(wb.sheetnames)

    raise ValueError(
        f"Không tìm thấy sheet tracking. Available sheets: {available_sheets}"
    )


def move_sheet_to_index(wb, sheet_name, target_index):
    """
    Move sheet to target index.

    target_index = 0 means move to first sheet.

    This avoids wb._move_sheet(), because some openpyxl versions
    do not support _move_sheet().
    """
    ws = wb[sheet_name]

    sheets = wb._sheets
    sheets.remove(ws)
    sheets.insert(target_index, ws)


def normalize_text(value):
    return clean_text(value).upper().replace("-", "").replace(" ", "")


def to_number(value):
    if value is None:
        return 0

    if isinstance(value, (int, float)):
        return value

    text = str(value).strip().replace(",", "")

    if text == "":
        return 0

    try:
        return float(text)
    except ValueError:
        return 0


def display_number(value):
    num = to_number(value)

    if float(num).is_integer():
        return int(num)

    return num


def parse_date_value(value):
    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    if isinstance(value, (int, float)):
        if 30000 <= value <= 60000:
            return (datetime(1899, 12, 30) + timedelta(days=int(value))).date()

    text = str(value).strip()

    if not text:
        return None

    formats = [
        "%d/%m/%Y",
        "%d/%m/%y",
        "%d-%m-%Y",
        "%d-%m-%y",
        "%Y-%m-%d",
        "%d.%m.%Y",
        "%d.%m.%y",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass

    return None


def is_export_date_row(value):
    """
    Supports:
    - 01/07/2026
    - 1/7/2026
    - 01/07/26
    - Excel datetime/date
    """
    if value is None:
        return False

    if isinstance(value, (datetime, date)):
        return True

    text = str(value).strip()

    return bool(re.fullmatch(r"\d{1,2}/\d{1,2}/\d{2,4}", text))


def get_run_date():
    if RUN_DATE is None:
        return date.today()

    parsed = parse_date_value(RUN_DATE)

    if not parsed:
        raise ValueError(f"RUN_DATE không hợp lệ: {RUN_DATE}")

    return parsed


def get_report_bcd(row):
    xe = row[REPORT_COL_XE - 1] if len(row) >= REPORT_COL_XE else None
    qty = row[REPORT_COL_QTY - 1] if len(row) >= REPORT_COL_QTY else None
    weight = row[REPORT_COL_WEIGHT - 1] if len(row) >= REPORT_COL_WEIGHT else None

    return xe, qty, weight


def is_container_stt(value):
    if value is None:
        return False

    text = clean_text(value).upper()

    return bool(re.fullmatch(r"K\d+", text))


def normalize_container_stt(value):
    return clean_text(value).upper()


def is_tracking_stt(value):
    if value is None:
        return False

    text = clean_text(value).upper()

    return bool(re.fullmatch(r"KK\d+", text))


def is_product_code(value):
    if value is None:
        return False

    text = clean_text(value).upper()

    return bool(re.fullmatch(r"[AB][34]", text))


def format_chinese_month_day(d: date):
    if not d:
        return ""

    return f"{d.month}月{d.day}日"


def display_stt_without_kk(tracking_stt):
    text = clean_text(tracking_stt).upper()

    if text.startswith("KK"):
        return text[2:]

    return text


def format_receipt_date_range(dates):
    clean_dates = sorted({d for d in dates if d})

    if not clean_dates:
        return ""

    if len(clean_dates) == 1:
        d = clean_dates[0]
        return f"{d.month}月{d.day}日"

    first = clean_dates[0]
    last = clean_dates[-1]

    if first.month == last.month:
        return f"{first.month}月{first.day}-{last.day}日"

    return f"{first.month}月{first.day}日-{last.month}月{last.day}日"


def format_receipt_day_range_only(dates):
    clean_dates = sorted({d for d in dates if d})

    if not clean_dates:
        return ""

    if len(clean_dates) == 1:
        return str(clean_dates[0].day)

    first = clean_dates[0]
    last = clean_dates[-1]

    if first.month == last.month:
        return f"{first.day}-{last.day}"

    return f"{first.day}/{first.month}-{last.day}/{last.month}"


def get_all_receipt_dates_from_items(items):
    all_dates = set()

    for item in items:
        for d in item.get("receipt_dates", []):
            if d:
                all_dates.add(d)

    return sorted(all_dates)


# =========================================================
# TRACKING FILE
# =========================================================

def load_tracking_data(tracking_file: Path):
    wb = load_workbook(tracking_file, data_only=True)
    ws = get_tracking_sheet(wb)

    mapping = {}

    for row in ws.iter_rows(values_only=True):
        tracking_stt = row[TRACKING_COL_STT - 1] if len(row) >= TRACKING_COL_STT else None
        truck_plate = row[TRACKING_COL_TRUCK_PLATE - 1] if len(row) >= TRACKING_COL_TRUCK_PLATE else None
        container_no = row[TRACKING_COL_CONTAINER_NO - 1] if len(row) >= TRACKING_COL_CONTAINER_NO else None
        trailer_no = row[TRACKING_COL_TRAILER_NO - 1] if len(row) >= TRACKING_COL_TRAILER_NO else None

        tracking_stt_text = clean_text(tracking_stt).upper()

        if not is_tracking_stt(tracking_stt_text):
            continue

        report_key = "K" + tracking_stt_text[2:]

        mapping[report_key] = {
            "tracking_stt": tracking_stt_text,
            "truck_plate": clean_text(truck_plate),
            "container_no": normalize_text(container_no),
            "trailer_no": clean_text(trailer_no),
        }

    return mapping


# =========================================================
# PRODUCT
# =========================================================

def build_product_name(product_code, kg_per_box):
    product_code = clean_text(product_code).upper()
    grade = product_code[0]
    fruit_count = 3 if product_code.endswith("3") else 4
    kg_per_box = int(kg_per_box)

    if grade == "A":
        color_vi = "trắng"
        color_cn = "白带"
    else:
        color_vi = "vàng"
        color_cn = "黄带"

    belt_count = 2 if kg_per_box == 9 else 3

    return (
        f"{grade} - {fruit_count} trái - {belt_count} đai {color_vi}\n"
        f"{grade}-{fruit_count}头-{belt_count}{color_cn}"
    )


def make_item(product_code, kg_per_box, qty_boxes, total_weight, receipt_dates):
    product_code = clean_text(product_code).upper()
    kg_per_box = int(to_number(kg_per_box))
    qty_boxes = to_number(qty_boxes)
    total_weight = to_number(total_weight)

    if not is_product_code(product_code):
        return None

    if qty_boxes == 0 or total_weight == 0:
        return None

    fruit_count = 3 if product_code.endswith("3") else 4
    receipt_date_text = format_receipt_date_range(receipt_dates)

    return {
        "product_code": product_code,
        "product_name": build_product_name(product_code, kg_per_box),
        "qty_boxes": display_number(qty_boxes),
        "fruit_count": fruit_count,
        "kg_per_box": kg_per_box,
        "total_weight": display_number(total_weight),
        "unit_price": None,
        "amount": None,
        "receipt_dates": sorted({d for d in receipt_dates if d}),
        "receipt_date_text": receipt_date_text,
    }


# =========================================================
# READ K LIST
# =========================================================

def get_k_list_for_export_date(source_wb, run_date: date):
    ws = get_report_sheet(source_wb)

    inside_target_date = False
    k_list = []

    for row in ws.iter_rows(values_only=True):
        xe, qty, weight = get_report_bcd(row)

        if xe is None:
            continue

        if is_export_date_row(xe):
            parsed = parse_date_value(xe)
            inside_target_date = parsed == run_date

            if inside_target_date:
                print(f"DEBUG EXPORT DATE FOUND: {parsed}")

            continue

        if not inside_target_date:
            continue

        if is_container_stt(xe):
            k = normalize_container_stt(xe)

            if k not in k_list:
                k_list.append(k)
                print(f"DEBUG K FOUND UNDER EXPORT DATE: {k}")

            continue

    return k_list


def get_k_list_from_xuat_kho_for_export_date(source_wb, run_date: date):
    ws = get_xuat_kho_sheet(source_wb)

    k_list = []

    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row_idx <= 2:
            continue

        trip = row[XUATKHO_COL_TRIP - 1] if len(row) >= XUATKHO_COL_TRIP else None
        export_date = row[XUATKHO_COL_EXPORT_DATE - 1] if len(row) >= XUATKHO_COL_EXPORT_DATE else None

        trip_text = clean_text(trip).upper()
        export_date_parsed = parse_date_value(export_date)

        if export_date_parsed != run_date:
            continue

        if not is_container_stt(trip_text):
            continue

        if trip_text not in k_list:
            k_list.append(trip_text)
            print(f"DEBUG K FOUND FROM XUẤT KHO: {trip_text}")

    return k_list


# =========================================================
# READ XUẤT KHO DETAILS
# =========================================================

def read_xuat_kho_items_by_k(source_wb, k_stt: str, run_date: date):
    ws = get_xuat_kho_sheet(source_wb)

    grouped = OrderedDict()

    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row_idx <= 2:
            continue

        receipt_date = row[XUATKHO_COL_RECEIPT_DATE - 1] if len(row) >= XUATKHO_COL_RECEIPT_DATE else None
        trip = row[XUATKHO_COL_TRIP - 1] if len(row) >= XUATKHO_COL_TRIP else None
        export_date = row[XUATKHO_COL_EXPORT_DATE - 1] if len(row) >= XUATKHO_COL_EXPORT_DATE else None
        product_code = row[XUATKHO_COL_PRODUCT_CODE - 1] if len(row) >= XUATKHO_COL_PRODUCT_CODE else None
        kg_per_box = row[XUATKHO_COL_KG - 1] if len(row) >= XUATKHO_COL_KG else None
        qty_boxes = row[XUATKHO_COL_QTY - 1] if len(row) >= XUATKHO_COL_QTY else None
        total_weight = row[XUATKHO_COL_WEIGHT - 1] if len(row) >= XUATKHO_COL_WEIGHT else None

        trip_text = clean_text(trip).upper()
        export_date_parsed = parse_date_value(export_date)
        receipt_date_parsed = parse_date_value(receipt_date)

        if trip_text != k_stt:
            continue

        if export_date_parsed != run_date:
            continue

        if not is_product_code(product_code):
            continue

        kg_num = int(to_number(kg_per_box))

        if kg_num not in (9, 10):
            continue

        key = (
            receipt_date_parsed,
            clean_text(product_code).upper(),
            kg_num,
        )

        if key not in grouped:
            grouped[key] = {
                "receipt_dates": set(),
                "product_code": clean_text(product_code).upper(),
                "kg_per_box": kg_num,
                "qty_boxes": 0,
                "total_weight": 0,
            }

        if receipt_date_parsed:
            grouped[key]["receipt_dates"].add(receipt_date_parsed)

        grouped[key]["qty_boxes"] += to_number(qty_boxes)
        grouped[key]["total_weight"] += to_number(total_weight)

    items = []

    for data in grouped.values():
        item = make_item(
            product_code=data["product_code"],
            kg_per_box=data["kg_per_box"],
            qty_boxes=data["qty_boxes"],
            total_weight=data["total_weight"],
            receipt_dates=data["receipt_dates"],
        )

        if item:
            items.append(item)

    return items


# =========================================================
# MAIN DATA READ
# =========================================================

def read_containers_for_date(source_file: Path, run_date: date, tracking_map: dict):
    source_wb = load_workbook(source_file, data_only=True)

    k_list = get_k_list_from_xuat_kho_for_export_date(source_wb, run_date)

    print(f"DEBUG K LIST FROM REPORT: {k_list}")

    # if not k_list:
    #     print("WARNING: Không tìm thấy K list trong BÁO CÁO KHO TỔNG HỢP.")
    #     print("Trying fallback: scan K list from sheet Xuất kho...")

    #     k_list = get_k_list_from_xuat_kho_for_export_date(source_wb, run_date)

    #     print(f"DEBUG K LIST FROM XUẤT KHO: {k_list}")

    if not k_list:
        print(f"WARNING: Không tìm thấy chuyến K nào cho ngày {run_date.strftime('%d/%m/%Y')}")
        return []

    results = []

    for k_stt in k_list:
        tracking_info = tracking_map.get(k_stt)

        if not tracking_info:
            print(f"WARNING: {k_stt} không tìm thấy trong file theo dõi container. Bỏ qua.")
            continue

        items = read_xuat_kho_items_by_k(source_wb, k_stt, run_date)

        print(f"DEBUG ITEMS FOR {k_stt}: {len(items)}")

        if not items:
            print(
                f"SKIP EMPTY: {k_stt} | "
                f"tracking={tracking_info['tracking_stt']} | "
                f"container={tracking_info['container_no']}"
            )
            continue

        results.append({
            "report_stt": k_stt,
            "tracking_stt": tracking_info["tracking_stt"],
            "truck_plate": tracking_info["truck_plate"],
            "container_no": tracking_info["container_no"],
            "trailer_no": tracking_info["trailer_no"],
            "export_date": run_date,
            "items": items,
        })

        print(
            f"DEBUG SAVE: {tracking_info['tracking_stt']} | "
            f"container={tracking_info['container_no']} | "
            f"items={len(items)}"
        )

    return results


# =========================================================
# WRITE TARGET WORKBOOK
# =========================================================

def copy_row_style(ws, source_row, target_row, max_col=12):
    for col in range(1, max_col + 1):
        source_cell = ws.cell(source_row, col)
        target_cell = ws.cell(target_row, col)

        if source_cell.has_style:
            target_cell._style = copy.copy(source_cell._style)

        target_cell.number_format = source_cell.number_format
        target_cell.alignment = copy.copy(source_cell.alignment)
        target_cell.font = copy.copy(source_cell.font)
        target_cell.fill = copy.copy(source_cell.fill)
        target_cell.border = copy.copy(source_cell.border)


def get_template_sheet(target_wb):
    max_stt = -1
    template_name = target_wb.sheetnames[-1]

    for name in target_wb.sheetnames:
        match = re.match(r"^(\d+)-", str(name))

        if match:
            stt = int(match.group(1))

            if stt > max_stt:
                max_stt = stt
                template_name = name

    return target_wb[template_name]


def prepare_item_area(ws, item_count):
    extra_needed = max(0, item_count - TEMPLATE_ITEM_ROWS)

    if extra_needed > 0:
        insert_at = TARGET_ITEM_START_ROW + TEMPLATE_ITEM_ROWS
        ws.insert_rows(insert_at, extra_needed)

        source_style_row = insert_at - 1

        for i in range(extra_needed):
            copy_row_style(ws, source_style_row, insert_at + i)

    rows_to_clear = max(TEMPLATE_ITEM_ROWS, item_count)

    for row_idx in range(TARGET_ITEM_START_ROW, TARGET_ITEM_START_ROW + rows_to_clear):
        for col_letter in TARGET_ITEM_COLUMNS.values():
            ws[f"{col_letter}{row_idx}"] = None


def style_item_row(ws, row_idx):
    for col in range(1, 10):
        cell = ws.cell(row_idx, col)
        cell.fill = ITEM_FILL
        cell.border = THIN_BORDER
        cell.alignment = CENTER_ALIGN
        cell.font = BLACK_NORMAL_FONT

    ws[f"A{row_idx}"].alignment = CENTER_ALIGN
    ws[f"H{row_idx}"].alignment = CENTER_ALIGN
    ws[f"I{row_idx}"].alignment = CENTER_ALIGN


def ensure_summary_rows(ws, first_summary_row):
    labels = [
        "总合计Tổng cộng",
        "单证费Phí Chứng từ",
        "运费Phí vận chuyển",
        "越南盾总合计Tổng cộng (VNĐ)",
    ]

    style_source_row = max(first_summary_row - 1, 1)

    for i, label in enumerate(labels):
        row_idx = first_summary_row + i

        copy_row_style(ws, style_source_row, row_idx, max_col=9)

        for c in range(1, 10):
            ws.cell(row_idx, c).value = None

        ws[f"A{row_idx}"] = label

    transport_fee_row = first_summary_row + 2
    ws[f"H{transport_fee_row}"] = VAT_NOTE

    return {
        "total": first_summary_row,
        "document_fee": first_summary_row + 1,
        "transport_fee": first_summary_row + 2,
        "grand_total": first_summary_row + 3,
    }


def style_summary_rows(ws, summary_rows):
    total_row = summary_rows["total"]
    document_fee_row = summary_rows["document_fee"]
    transport_fee_row = summary_rows["transport_fee"]
    grand_total_row = summary_rows["grand_total"]

    for r in [total_row, document_fee_row, transport_fee_row]:
        for col in range(1, 10):
            cell = ws.cell(r, col)
            cell.fill = PatternFill(fill_type=None)
            cell.font = BLACK_BOLD_FONT
            cell.border = THIN_BORDER
            cell.alignment = CENTER_ALIGN

    for col in range(1, 10):
        cell = ws.cell(grand_total_row, col)
        cell.fill = GRAND_TOTAL_FILL
        cell.font = RED_FONT
        cell.border = THIN_BORDER
        cell.alignment = CENTER_ALIGN

    ws[f"H{transport_fee_row}"] = VAT_NOTE
    ws[f"H{transport_fee_row}"].font = BLACK_BOLD_FONT
    ws[f"H{transport_fee_row}"].alignment = CENTER_ALIGN
    ws[f"H{transport_fee_row}"].fill = PatternFill(fill_type=None)


def update_summary_formulas(ws, first_item_row, last_item_row):
    first_summary_row = last_item_row + 1

    summary_rows = ensure_summary_rows(ws, first_summary_row)

    total_row = summary_rows["total"]
    document_fee_row = summary_rows["document_fee"]
    transport_fee_row = summary_rows["transport_fee"]
    grand_total_row = summary_rows["grand_total"]

    if last_item_row < first_item_row:
        ws[f"B{total_row}"] = 0
        ws[f"E{total_row}"] = 0
        ws[f"G{total_row}"] = 0
    else:
        ws[f"B{total_row}"] = f"=SUM(B{first_item_row}:B{last_item_row})"
        ws[f"E{total_row}"] = f"=SUM(E{first_item_row}:E{last_item_row})"
        ws[f"G{total_row}"] = f"=SUM(G{first_item_row}:G{last_item_row})"

    ws[f"G{document_fee_row}"] = None
    ws[f"G{transport_fee_row}"] = None
    ws[f"G{grand_total_row}"] = f"=G{total_row}+G{document_fee_row}+G{transport_fee_row}"
    ws[f"H{transport_fee_row}"] = VAT_NOTE

    style_summary_rows(ws, summary_rows)


def write_container_to_sheet(target_wb, container_data):
    template_ws = get_template_sheet(target_wb)

    tracking_stt = container_data["tracking_stt"]
    truck_plate = container_data["truck_plate"]
    container_no = container_data["container_no"]
    trailer_no = container_data["trailer_no"]
    export_date = container_data["export_date"]
    items = container_data["items"]

    new_sheet_name = safe_sheet_name(f"{tracking_stt}-{container_no}")

    if new_sheet_name in target_wb.sheetnames:
        if REPLACE_EXISTING_SHEETS:
            old_ws = target_wb[new_sheet_name]
            target_wb.remove(old_ws)
            print(f"Replaced old sheet: {new_sheet_name}")
        else:
            print(f"Sheet đã tồn tại, bỏ qua: {new_sheet_name}")
            return False

    ws = target_wb.copy_worksheet(template_ws)
    ws.title = new_sheet_name

    display_stt = display_stt_without_kk(tracking_stt)

    ws[TARGET_CELLS["stt_label"]] = f"陆运柜号: {display_stt}\nSố cont: {display_stt}"
    ws[TARGET_CELLS["container_no"]] = container_no
    ws[TARGET_CELLS["truck_plate"]] = truck_plate
    ws[TARGET_CELLS["factory"]] = FACTORY_VALUE
    ws[TARGET_CELLS["trailer_no"]] = trailer_no
    ws[TARGET_CELLS["packaging"]] = PACKAGING_VALUE
    ws[TARGET_CELLS["export_date"]] = format_chinese_month_day(export_date)

    ws[TARGET_CELLS["truck_plate"]].font = BOLD_FONT
    ws[TARGET_CELLS["trailer_no"]].font = BOLD_FONT
    ws[TARGET_CELLS["truck_plate"]].alignment = CENTER_ALIGN
    ws[TARGET_CELLS["trailer_no"]].alignment = CENTER_ALIGN

    all_receipt_dates = get_all_receipt_dates_from_items(items)
    receipt_day_range = format_receipt_day_range_only(all_receipt_dates)

    ws["G4"] = f"收货日:{receipt_day_range}"
    ws["G5"] = f"Ngày nhận hàng: {receipt_day_range}"

    ws["G4"].font = BOLD_FONT
    ws["G5"].font = BOLD_FONT
    ws["G4"].alignment = CENTER_ALIGN
    ws["G5"].alignment = CENTER_ALIGN

    prepare_item_area(ws, len(items))

    row_idx = TARGET_ITEM_START_ROW

    for item_index, item in enumerate(items, start=1):
        product_col = TARGET_ITEM_COLUMNS["product_name"]
        qty_col = TARGET_ITEM_COLUMNS["qty_boxes"]
        fruit_col = TARGET_ITEM_COLUMNS["fruit_count"]
        kg_col = TARGET_ITEM_COLUMNS["kg_per_box"]
        weight_col = TARGET_ITEM_COLUMNS["total_weight"]
        price_col = TARGET_ITEM_COLUMNS["unit_price"]
        amount_col = TARGET_ITEM_COLUMNS["amount"]
        receipt_col = TARGET_ITEM_COLUMNS["receipt_date"]
        item_no_col = TARGET_ITEM_COLUMNS["item_no"]

        ws[f"{product_col}{row_idx}"] = item["product_name"]
        ws[f"{qty_col}{row_idx}"] = item["qty_boxes"]
        ws[f"{fruit_col}{row_idx}"] = item["fruit_count"]
        ws[f"{kg_col}{row_idx}"] = item["kg_per_box"]
        ws[f"{weight_col}{row_idx}"] = item["total_weight"]

        ws[f"{price_col}{row_idx}"] = None
        ws[f"{amount_col}{row_idx}"] = f"={weight_col}{row_idx}*{price_col}{row_idx}"

        receipt_date_text = item.get("receipt_date_text", "")

        if receipt_date_text:
            ws[f"{receipt_col}{row_idx}"] = (
                f"{receipt_date_text} 金枕\n"
                f"东部货/Miền Đông"
            )
        else:
            ws[f"{receipt_col}{row_idx}"] = None

        ws[f"{item_no_col}{row_idx}"] = item_index

        style_item_row(ws, row_idx)

        row_idx += 1

    last_item_row = row_idx - 1

    update_summary_formulas(ws, TARGET_ITEM_START_ROW, last_item_row)

    print(
        f"CREATED: {new_sheet_name} | "
        f"truck={truck_plate} | trailer={trailer_no} | items={len(items)}"
    )

    return new_sheet_name


# =========================================================
# PUBLIC FUNCTION FOR STREAMLIT WEBSITE
# =========================================================

def generate_excel(
    target_file_path,
    source_file_path,
    tracking_file_path,
    run_date_text,
    output_file_path,
):
    global TARGET_FILE
    global SOURCE_FILE
    global TRACKING_FILE
    global RUN_DATE

    TARGET_FILE = Path(target_file_path)
    SOURCE_FILE = Path(source_file_path)
    TRACKING_FILE = Path(tracking_file_path)
    RUN_DATE = run_date_text

    run_date = get_run_date()

    tracking_map = load_tracking_data(TRACKING_FILE)
    containers = read_containers_for_date(SOURCE_FILE, run_date, tracking_map)

    if not containers:
        raise ValueError(
            f"No valid containers found for export date {run_date.strftime('%d/%m/%Y')}"
        )

    target_wb = load_workbook(TARGET_FILE)

    created_count = 0
    created_sheet_names = []

    for container_data in containers:
        created_sheet_name = write_container_to_sheet(target_wb, container_data)

        if created_sheet_name:
            created_count += 1
            created_sheet_names.append(created_sheet_name)

    # Stack order:
    # If created K145, K146, K147,
    # final order will be K147, K146, K145, old sheets...
    for sheet_name in created_sheet_names:
        move_sheet_to_index(target_wb, sheet_name, 0)

    target_wb.save(output_file_path)

    return {
        "created_count": created_count,
        "containers": containers,
        "output_file": str(output_file_path),
    }


# =========================================================
# OPTIONAL TERMINAL RUNNER
# =========================================================

def create_sheets():
    run_date = get_run_date()

    if not SOURCE_FILE.exists():
        raise FileNotFoundError(f"Không tìm thấy file nguồn: {SOURCE_FILE}")

    if not TARGET_FILE.exists():
        raise FileNotFoundError(f"Không tìm thấy file đích: {TARGET_FILE}")

    if not TRACKING_FILE.exists():
        raise FileNotFoundError(f"Không tìm thấy file theo dõi container: {TRACKING_FILE}")

    print(f"Running for export date: {run_date.strftime('%d/%m/%Y')}")

    tracking_map = load_tracking_data(TRACKING_FILE)

    print(f"Loaded tracking rows: {len(tracking_map)}")

    containers = read_containers_for_date(SOURCE_FILE, run_date, tracking_map)

    if not containers:
        print(f"No valid containers found for export date {run_date.strftime('%d/%m/%Y')}")
        return

    print(f"Valid containers with item data: {len(containers)}")

    target_wb = load_workbook(TARGET_FILE)

    created_count = 0
    created_sheet_names = []

    for container_data in containers:
        created_sheet_name = write_container_to_sheet(target_wb, container_data)

        if created_sheet_name:
            created_count += 1
            created_sheet_names.append(created_sheet_name)

    # Stack order:
    # If created K145, K146, K147,
    # final order will be K147, K146, K145, old sheets...
    for sheet_name in created_sheet_names:
        move_sheet_to_index(target_wb, sheet_name, 0)

    backup_file = TARGET_FILE.with_name(
        TARGET_FILE.stem
        + f"_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        + TARGET_FILE.suffix
    )

    target_wb.save(backup_file)
    target_wb.save(TARGET_FILE)

    print("Done.")
    print(f"Created sheets: {created_count}")
    print(f"Backup saved at: {backup_file}")


if __name__ == "__main__":
    create_sheets()