import re
import csv
import sys
from collections import OrderedDict
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Side, Font

# ================= 公共判定/清洗 =================

DATE_MMDDSHORT = re.compile(
    r'^\s*(?:1[0-2]|0[1-9])/(?:3[01]|[12]\d|0[1-9])\s*$'
)

DATE_MMDDYY = re.compile(
    r'^\s*(?:1[0-2]|0[1-9])/(?:3[01]|[12]\d|0[1-9])/\d{2}\s*$'
)

DATE_MMDDYYYY = re.compile(
    r'^\s*(?:1[0-2]|0[1-9])/(?:3[01]|[12]\d|0[1-9])/\d{4}\s*$'
)

DEFAULT_DATE_ANY_IN_TEXT_PATTERN = (
    r'(?:1[0-2]|0[1-9])/'
    r'(?:3[01]|[12]\d|0[1-9])'
    r'(?:/(?:\d{2}|\d{4}))?'
)

DATE_ANY_IN_TEXT_PATTERN = DEFAULT_DATE_ANY_IN_TEXT_PATTERN

AMOUNT_RE = re.compile(
    r'^\s*(?:'
    r'-\s*\$?\s*[\d,]+\.\d{2}'
    r'|'
    r'\$?\s*-\s*[\d,]+\.\d{2}'
    r'|'
    r'\(?\s*\$?\s*[\d,]+\.\d{2}\s*\)?'
    r')\s*$'
)

PHONE_RE = re.compile(r'^\s*\+?\d[\d\s\-]{6,}\s*$')
STATE_RE = re.compile(r'^\s*[A-Z]{2}\s*$')
DOT_RE = re.compile(r'^\s*\.\s*$')

MERCHANT_CUT_TOKENS = (" DES:", " ID:", " INDN:", " CO ID:")

DATE_IN_TEXT_RE = re.compile(DATE_ANY_IN_TEXT_PATTERN, re.IGNORECASE)

DATE_FULL_RE = re.compile(
    rf'^\s*(?:{DATE_ANY_IN_TEXT_PATTERN})\s*$', re.IGNORECASE
)

DATE_MERCHANT_LINE_RE = re.compile(
    rf'^\s*({DATE_ANY_IN_TEXT_PATTERN})\s+(.*\S)\s*$', re.IGNORECASE
)

def convert_custom_date_format_to_regex(date_format: str) -> str:
    """Convert one or more user date formats into a regex pattern.

    Unified notation used by the UI:
      M       month: numeric or English month name
      MM      numeric month only
      DD      day of month
      YY      two-digit year
      YYYY    four-digit year

    Examples:
      M-DD       -> 4-22
      M DD       -> May 17
      DD M       -> 17 May
      M-DD-YYYY  -> 4-22-2025

    Multiple formats may be separated by commas, Chinese commas, semicolons,
    or vertical bars. Old tokens D, MON, MMM and MONTH remain accepted for
    backward compatibility, but the UI only presents the unified notation.
    """
    raw = str(date_format or "").strip()
    if not raw:
        return DEFAULT_DATE_ANY_IN_TEXT_PATTERN

    formats = [
        part.strip()
        for part in re.split(r'[,，;；|]+', raw)
        if part.strip()
    ]
    if not formats:
        return DEFAULT_DATE_ANY_IN_TEXT_PATTERN

    month_name_pattern = (
        r'(?:JAN(?:UARY)?|FEB(?:RUARY)?|MAR(?:CH)?|APR(?:IL)?|MAY|'
        r'JUN(?:E)?|JUL(?:Y)?|AUG(?:UST)?|SEP(?:T(?:EMBER)?)?|'
        r'OCT(?:OBER)?|NOV(?:EMBER)?|DEC(?:EMBER)?)'
    )
    numeric_month_pattern = r'(?:1[0-2]|0?[1-9])'
    day_pattern = r'(?:3[01]|[12]\d|0?[1-9])'

    # M means Month and can therefore recognize either a numeric month
    # (4) or an English month name (May/September).
    any_month_pattern = rf'(?:{numeric_month_pattern}|{month_name_pattern})'

    token_patterns = {
        "YYYY": r"\d{4}",
        "MONTH": any_month_pattern,
        "MMM": any_month_pattern,
        "MON": any_month_pattern,
        "YY": r"\d{2}",
        "MM": numeric_month_pattern,
        "DD": day_pattern,
        "M": any_month_pattern,
        "D": day_pattern,  # legacy alias
    }
    token_order = ("YYYY", "MONTH", "MMM", "MON", "YY", "MM", "DD", "M", "D")

    def one_format_to_regex(fmt_text: str) -> str:
        fmt = fmt_text.strip().upper()
        result = []
        i = 0

        while i < len(fmt):
            matched = False
            for token in token_order:
                if fmt.startswith(token, i):
                    result.append(token_patterns[token])
                    i += len(token)
                    matched = True
                    break

            if matched:
                continue

            current_char = fmt[i]
            if current_char.isalpha():
                raise ValueError(
                    "无法识别的日期格式。请使用 M、MM、DD、YY、YYYY，"
                    "例如 M-DD、M DD、DD M、M-DD-YYYY。"
                )

            if current_char.isspace():
                while i < len(fmt) and fmt[i].isspace():
                    i += 1
                result.append(r"\s+")
                continue

            result.append(re.escape(current_char))
            i += 1

        return "".join(result)

    patterns = [one_format_to_regex(fmt) for fmt in formats]
    return patterns[0] if len(patterns) == 1 else r"(?:" + "|".join(patterns) + r")"



def infer_date_format_from_sample(sample_text: str) -> str:
    """Infer a unified date format from one user-entered date example.

    Examples:
      4-22       -> M-DD
      04/22/25   -> M/DD/YY
      2025-4-22  -> YYYY-M-DD
      May 17     -> M DD
      17 May     -> DD M
    """
    sample = " ".join(str(sample_text or "").strip().split())
    if not sample:
        raise ValueError("日期示例不能为空。")

    month_names = (
        r"JAN(?:UARY)?|FEB(?:RUARY)?|MAR(?:CH)?|APR(?:IL)?|MAY|"
        r"JUN(?:E)?|JUL(?:Y)?|AUG(?:UST)?|SEP(?:T(?:EMBER)?)?|"
        r"OCT(?:OBER)?|NOV(?:EMBER)?|DEC(?:EMBER)?"
    )

    # English month first: May 17, May 17 2025, May-17-25
    m = re.fullmatch(
        rf"(?i)({month_names})([\s./-]+)(\d{{1,2}})(?:([\s./-]+)(\d{{2}}|\d{{4}}))?",
        sample,
    )
    if m:
        sep1 = m.group(2)
        year = m.group(5)
        fmt = f"M{sep1}DD"
        if year:
            fmt += f"{m.group(4)}{'YYYY' if len(year) == 4 else 'YY'}"
        return fmt

    # English month second: 17 May, 17 May 2025, 17-May-25
    m = re.fullmatch(
        rf"(?i)(\d{{1,2}})([\s./-]+)({month_names})(?:([\s./-]+)(\d{{2}}|\d{{4}}))?",
        sample,
    )
    if m:
        sep1 = m.group(2)
        year = m.group(5)
        fmt = f"DD{sep1}M"
        if year:
            fmt += f"{m.group(4)}{'YYYY' if len(year) == 4 else 'YY'}"
        return fmt

    # Numeric formats. A four-digit first field is treated as the year.
    m = re.fullmatch(r"(\d{1,4})([./-])(\d{1,2})(?:([./-])(\d{2}|\d{4}))?", sample)
    if m:
        first, sep1, second, sep2, third = m.groups()
        if len(first) == 4:
            if third is None:
                raise ValueError("年份在前的日期需要包含年、月、日，例如 2025-4-22。")
            return f"YYYY{sep1}M{sep2}DD"

        fmt = f"M{sep1}DD"
        if third:
            fmt += f"{sep2}{'YYYY' if len(third) == 4 else 'YY'}"
        return fmt

    raise ValueError(
        "无法从日期示例中判断格式。请输入类似 4-22、04/22/2025、"
        "2025-4-22、May 17 或 17 May。"
    )


def resolve_date_input_to_format(date_input: str) -> str:
    """Accept either format tokens or real date examples and return formats.

    Multiple entries may be separated by commas, semicolons, or vertical bars.
    """
    raw = str(date_input or "").strip()
    if not raw:
        return ""

    parts = [p.strip() for p in re.split(r"[,，;；|]+", raw) if p.strip()]
    resolved = []
    format_token_re = re.compile(r"(?i)^(?:YYYY|YY|MM|M|DD|D|MMM|MON|MONTH|[\s./-])+$")

    for part in parts:
        if format_token_re.fullmatch(part):
            resolved.append(part.upper())
        else:
            resolved.append(infer_date_format_from_sample(part))

    return ", ".join(resolved)

def configure_date_format(custom_date_format: str):
    """Apply the custom date format, or restore defaults when blank."""
    global DATE_ANY_IN_TEXT_PATTERN
    global DATE_IN_TEXT_RE
    global DATE_FULL_RE
    global DATE_MERCHANT_LINE_RE

    custom_date_format = resolve_date_input_to_format(custom_date_format)
    DATE_ANY_IN_TEXT_PATTERN = (
        convert_custom_date_format_to_regex(custom_date_format)
        if custom_date_format
        else DEFAULT_DATE_ANY_IN_TEXT_PATTERN
    )

    DATE_IN_TEXT_RE = re.compile(DATE_ANY_IN_TEXT_PATTERN, re.IGNORECASE)
    DATE_FULL_RE = re.compile(rf'^\s*(?:{DATE_ANY_IN_TEXT_PATTERN})\s*$', re.IGNORECASE)
    DATE_MERCHANT_LINE_RE = re.compile(
        rf'^\s*({DATE_ANY_IN_TEXT_PATTERN})\s+(.*\S)\s*$', re.IGNORECASE
    )

EMAIL_RE = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$', re.IGNORECASE)
URLISH_RE = re.compile(r'(https?://|www\.|\.com\b|squareup\.com\b)', re.IGNORECASE)

# ================= 常量 =================

MONTH_HEADERS = ["Jan", "Feb", "March", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

UI_TO_SHEET_MONTH = {
    "JAN": "Jan",
    "FEB": "Feb",
    "MAR": "March",
    "APR": "Apr",
    "MAY": "May",
    "JUN": "Jun",
    "JUL": "Jul",
    "AUG": "Aug",
    "SEP": "Sep",
    "OCT": "Oct",
    "NOV": "Nov",
    "DEC": "Dec",
}

CREDIT_DEFAULT_HEADER = "CASH / CHECK DEPOSIT"
CREDIT_TOTAL_HEADER = "TOTAL CREDIT"
CREDIT_DEBIT_HEADER = "DEBIT"
CREDIT_ENDING_HEADER = "ENDING BALANCE"
CREDIT_BEGIN_HEADER = "BEGIN BALANCE"
DEFAULT_BANK_NAME = "CHASE #3444"

CREDIT_MONTH_ROWS = {
    "JAN": 3,
    "FEB": 4,
    "MAR": 5,
    "APR": 6,
    "MAY": 7,
    "JUN": 8,
    "JUL": 9,
    "AUG": 10,
    "SEP": 11,
    "OCT": 12,
    "NOV": 13,
    "DEC": 14,
}

CREDIT_MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]

DEBIT_SHEET_NAME = "debit summary"
DEBIT_MONTH_HEADERS = ["Jan", "Feb", "March", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

DEFAULT_UI_MONTH_LABELS = list(DEBIT_MONTH_HEADERS)

def normalize_month_labels(labels) -> List[str]:
    """Return exactly 12 non-empty display labels, preserving Excel order."""
    cleaned = []
    for value in list(labels or [])[:12]:
        text = str(value).strip() if value is not None else ""
        cleaned.append(text)
    while len(cleaned) < 12:
        cleaned.append(DEFAULT_UI_MONTH_LABELS[len(cleaned)])
    return [label or DEFAULT_UI_MONTH_LABELS[i] for i, label in enumerate(cleaned)]

def read_month_labels_from_wb(wb) -> List[str]:
    """Read the 12 display labels already stored in the Monthly workbook."""
    if DEBIT_SHEET_NAME in wb.sheetnames:
        ws = wb[DEBIT_SHEET_NAME]
        labels = [ws.cell(row=1, column=col).value for col in range(2, 14)]
        if any(v is not None and str(v).strip() for v in labels):
            return normalize_month_labels(labels)
    if "Credit Summary" in wb.sheetnames:
        ws = wb["Credit Summary"]
        labels = [ws.cell(row=row, column=1).value for row in range(3, 15)]
        if any(v is not None and str(v).strip() for v in labels):
            return normalize_month_labels(labels)
    return list(DEFAULT_UI_MONTH_LABELS)

def read_month_labels_from_path(path: Path) -> List[str]:
    if not path.exists():
        return list(DEFAULT_UI_MONTH_LABELS)
    try:
        wb = load_workbook(path, data_only=False, read_only=True)
        labels = read_month_labels_from_wb(wb)
        wb.close()
        return labels
    except Exception:
        return list(DEFAULT_UI_MONTH_LABELS)

# ================= 基础工具 =================

def is_date_short(s: str) -> bool:
    return bool(DATE_MMDDSHORT.match(s))

def is_date_yy(s: str) -> bool:
    return bool(DATE_MMDDYY.match(s))

def is_date_yyyy(s: str) -> bool:
    return bool(DATE_MMDDYYYY.match(s))

def is_date_any(s: str) -> bool:
    return bool(DATE_FULL_RE.match(str(s)))

def is_amount(s: str) -> bool:
    return bool(AMOUNT_RE.match(s))

def is_phone(s: str) -> bool:
    return bool(PHONE_RE.match(s))

def is_state(s: str) -> bool:
    return bool(STATE_RE.match(s))

def is_dot(s: str) -> bool:
    return bool(DOT_RE.match(s))

def is_email(s: str) -> bool:
    return bool(EMAIL_RE.match(s.strip()))

def looks_like_url_or_site(s: str) -> bool:
    return bool(URLISH_RE.search(s.strip()))

def clean_amount(s: str) -> str:
    s = s.strip()
    neg = False

    if s.startswith("(") and s.endswith(")"):
        neg = True
    if "-" in s:
        neg = True

    s = (
        s.replace("$", "")
        .replace(",", "")
        .replace("(", "")
        .replace(")", "")
        .replace("-", "")
        .strip()
    )

    if neg and s:
        return f"-{s}"
    return s

def clean_merchant(line: str) -> str:
    raw = ' '.join(line.split())
    for t in MERCHANT_CUT_TOKENS:
        idx = raw.find(t)
        if idx != -1:
            return raw[:idx].rstrip()
    return raw

def extract_date_and_merchant(line: str):
    m = DATE_MERCHANT_LINE_RE.match(line.strip())
    if not m:
        return None, None
    return m.group(1), m.group(2).strip()

def normalize_amount_string(amount) -> str:
    value = float(str(amount).replace(",", "").strip())
    return f"{value:.2f}"

def safe_float(x):
    try:
        if x is None:
            return 0.0
        if isinstance(x, (int, float)):
            return float(x)
        s = str(x).strip()
        if not s:
            return 0.0
        if s.startswith("="):
            nums = re.findall(r'[-+]?\d+(?:\.\d+)?', s)
            total = 0.0
            for n in nums:
                total += float(n)
            return total
        return float(s.replace(",", ""))
    except Exception:
        return 0.0

def build_plus_formula(parts: List[str]) -> str:
    clean_parts = []
    for p in parts:
        p = str(p).strip()
        if not p:
            continue
        if p.startswith("+"):
            p = p[1:]
        clean_parts.append(p)

    if not clean_parts:
        return ""

    expr = clean_parts[0]
    for p in clean_parts[1:]:
        if p.startswith("-"):
            expr += p
        else:
            expr += f"+{p}"
    return f"={expr}"

def split_formula_parts(formula_or_value) -> List[str]:
    if formula_or_value is None:
        return []

    if isinstance(formula_or_value, (int, float)):
        return [f"{float(formula_or_value):.2f}"]

    s = str(formula_or_value).strip()
    if not s:
        return []

    if s.startswith("="):
        s = s[1:].strip()

    if not s:
        return []

    nums = re.findall(r'[-+]?\d+(?:\.\d+)?', s)
    out = []
    for n in nums:
        if n.startswith("+"):
            n = n[1:]
        out.append(f"{float(n):.2f}")
    return out

def script_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

def excel_col_letter(col_num: int) -> str:
    result = ""
    while col_num > 0:
        col_num, rem = divmod(col_num - 1, 26)
        result = chr(65 + rem) + result
    return result

# ================= Category Rules (XLSX) =================

def category_rules_folder() -> Path:
    folder = script_dir() / "Category Rules"
    folder.mkdir(parents=True, exist_ok=True)
    return folder

def category_rules_path() -> Path:
    return category_rules_folder() / "category_rules.xlsx"

def normalize_merchant_for_category(name: str) -> str:
    s = str(name or "").upper().strip()
    s = re.sub(r'[^A-Z0-9]+', ' ', s)

    stop_words = {
        "STORE", "STORES", "MARKET", "ONLINE", "PAYMENT", "PURCHASE",
        "DEBIT", "CREDIT", "CHECKCARD", "CHECK", "POS", "AUTH", "CARD",
        "WITHDRAWAL", "DBT", "ACH", "VISA", "MASTERCARD", "MC"
    }
    parts = [p for p in s.split() if p and p not in stop_words]
    parts = [p for p in parts if not p.isdigit()]

    return " ".join(parts).strip()

def create_default_category_rules_xlsx(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Category Rules"

    ws.cell(row=1, column=1, value="merchant")
    ws.cell(row=1, column=2, value="category")

    sample_rows = [
        ("TARGET", "Office Expense"),
        ("WHOLEFOODS", "Grocery"),
        ("WHOLE FOODS", "Grocery"),
        ("COSTCO", "Grocery"),
        ("UBER", "Travel"),
        ("SHELL", "Gasoline"),
    ]

    row_idx = 2
    for merchant, category in sample_rows:
        ws.cell(row=row_idx, column=1, value=merchant)
        ws.cell(row=row_idx, column=2, value=category)
        row_idx += 1

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 20

    wb.save(path)

def load_category_rules() -> List[Tuple[str, str]]:
    path = category_rules_path()

    if not path.exists():
        create_default_category_rules_xlsx(path)

    wb = load_workbook(path, data_only=True)
    ws = wb.active

    rules = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        merchant = str(row[0]).strip() if len(row) > 0 and row[0] is not None else ""
        category = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""

        if merchant and category:
            rules.append((normalize_merchant_for_category(merchant), category))

    return rules

def get_category_for_merchant(merchant: str, rules: List[Tuple[str, str]]) -> str:
    raw = str(merchant or "").strip()
    norm = normalize_merchant_for_category(raw)

    if not norm:
        return ""

    for rule_merchant, category in rules:
        if norm == rule_merchant:
            return category

    for rule_merchant, category in rules:
        if rule_merchant and rule_merchant in norm:
            return category

    for rule_merchant, category in rules:
        if norm and norm in rule_merchant:
            return category

    return ""

# ================= 预处理删除内容 =================

def parse_remove_items(raw: str) -> List[str]:
    if not raw:
        return []

    parts = re.split(r'[,\n;，；]+', raw)
    cleaned = []
    seen = set()

    for p in parts:
        item = p.strip()
        if not item:
            continue
        low = item.lower()
        if low not in seen:
            seen.add(low)
            cleaned.append(item)

    return cleaned

def preprocess_statement_text(text: str, remove_items: List[str]) -> str:
    if not text or not remove_items:
        return text

    new_text = text

    for item in remove_items:
        if not item.strip():
            continue
        pattern = re.compile(re.escape(item), re.IGNORECASE)
        new_text = pattern.sub("", new_text)

    cleaned_lines = []
    for line in new_text.splitlines():
        line = re.sub(r'[ \t]+', ' ', line).strip()
        cleaned_lines.append(line)

    return "\n".join(cleaned_lines)

# ================= 把整串文本拆成“伪多行” =================

def build_amount_at_end_pattern() -> str:
    return (
        r'('
        r'(?:-\s*\$?\s*[\d,]+\.\d{2})'
        r'|'
        r'(?:\$?\s*-\s*[\d,]+\.\d{2})'
        r'|'
        r'(?:\(\s*\$?\s*[\d,]+\.\d{2}\s*\))'
        r'|'
        r'(?:\$?\s*[\d,]+\.\d{2})'
        r')\s*$'
    )

def expand_compact_transactions(text: str) -> List[str]:
    text = ' '.join(text.split())
    if not text:
        return []

    matches = list(DATE_IN_TEXT_RE.finditer(text))
    if not matches:
        return [text]

    amount_at_end_re = re.compile(build_amount_at_end_pattern())
    out_lines = []

    for i, m in enumerate(matches):
        start = m.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        chunk = text[start:end].strip()

        dm = DATE_IN_TEXT_RE.match(chunk)
        if not dm:
            out_lines.append(chunk)
            continue

        date_str = dm.group(0)
        rest = chunk[dm.end():].strip()

        amt_match = amount_at_end_re.search(rest)
        if not amt_match:
            out_lines.append(chunk)
            continue

        amount_str = amt_match.group(1).strip()
        merchant_str = rest[:amt_match.start()].strip()

        if date_str and merchant_str:
            out_lines.append(f"{date_str} {merchant_str}")
        elif date_str:
            out_lines.append(date_str)

        if amount_str:
            out_lines.append(amount_str)

    return out_lines

def normalize_lines(text: str) -> List[str]:
    text = text.replace('\r\n', '\n').replace('\r', '\n').strip()
    if not text:
        return []

    lines = [ln.strip() for ln in text.split('\n') if ln.strip()]

    if len(lines) > 1:
        rebuilt = []
        for ln in lines:
            expanded = expand_compact_transactions(ln)
            rebuilt.extend(expanded)
        return rebuilt

    return expand_compact_transactions(text)

# ================= 解析器 =================

@dataclass
class ParseResult:
    span: Tuple[int, int]
    merchant: str
    amount: str
    who: str

class BaseExtractor:
    name = "base"

    def extract(self, text: str) -> List['ParseResult']:
        raise NotImplementedError

class SimpleDateStoreAmountExtractor(BaseExtractor):
    name = "simple_date_merchant_then_amount"

    def extract(self, text: str) -> List['ParseResult']:
        lines = normalize_lines(text)
        results: List[ParseResult] = []

        offsets = []
        off = 0
        for ln in lines:
            offsets.append(off)
            off += len(ln) + 1

        i = 0
        n = len(lines)

        while i < n:
            line = lines[i].strip()

            if not line:
                i += 1
                continue

            date_str, merchant_raw = extract_date_and_merchant(line)

            if date_str and merchant_raw:
                start_idx = i
                merchant = clean_merchant(merchant_raw)

                i += 1
                amount = None
                amount_idx = None

                while i < n:
                    cur = lines[i].strip()

                    if not cur:
                        i += 1
                        continue

                    next_date, _ = extract_date_and_merchant(cur)
                    if next_date:
                        break

                    if is_amount(cur):
                        amount = clean_amount(cur)
                        amount_idx = i
                        i += 1
                        break

                    i += 1

                if amount is not None and amount_idx is not None:
                    span = (offsets[start_idx], offsets[amount_idx] + len(lines[amount_idx]))
                    results.append(ParseResult(span, merchant, amount, self.name))

                continue

            i += 1

        return results

EXTRACTORS = {
    "a": SimpleDateStoreAmountExtractor(),
}

def merge_non_overlapping(results: List[ParseResult]) -> List[ParseResult]:
    results = sorted(results, key=lambda r: r.span[0])
    merged: List[ParseResult] = []
    last_end = -1
    for r in results:
        if r.span[0] >= last_end:
            merged.append(r)
            last_end = r.span[1]
    return merged

def parse_with_extractors(text: str, keys: List[str]) -> List[ParseResult]:
    hits: List[ParseResult] = []
    for k in keys:
        ext = EXTRACTORS.get(k)
        if not ext:
            continue
        hits.extend(ext.extract(text))
    return merge_non_overlapping(hits)

def parse_auto(text: str) -> List[ParseResult]:
    return parse_with_extractors(text, ["a"])

# ================= 数据合并 =================

def merge_same_merchants(rows):
    merged = OrderedDict()

    for merchant, amount, who in rows:
        merchant = merchant.strip()
        amount = normalize_amount_string(amount)

        if merchant not in merged:
            merged[merchant] = {
                "amounts": [],
                "who_list": []
            }

        merged[merchant]["amounts"].append(amount)
        merged[merchant]["who_list"].append(who)

    return merged

# ================= 明细合并表 =================

def write_merged_xlsx(xlsx_path: Path, rows):
    merged = merge_same_merchants(rows)
    rules = load_category_rules()

    wb = Workbook()
    ws = wb.active
    ws.title = "Merged Transactions"

    ws["A1"] = "Merchant"
    ws["B1"] = "Amount"
    ws["C1"] = "Category"

    sorted_items = sorted(
        merged.items(),
        key=lambda item: item[0].strip().lower()
    )

    r = 2
    for merchant, data in sorted_items:
        amounts = data["amounts"]
        formula = build_plus_formula(amounts)
        category = get_category_for_merchant(merchant, rules)

        ws.cell(row=r, column=1, value=merchant)
        ws.cell(row=r, column=2, value=formula)
        ws.cell(row=r, column=3, value=category)
        ws.cell(row=r, column=2).number_format = '0.00'
        r += 1

    ws.column_dimensions["A"].width = 60
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 20

    wb.save(xlsx_path)

# ================= Credit 月总表 =================

def normalize_credit_header_name(name: str) -> str:
    s = " ".join(str(name).strip().split())
    return s.upper()

def is_bad_credit_header_name(name: str) -> bool:
    if name is None:
        return True

    s = str(name).strip()
    if not s:
        return True

    s_upper = s.upper()

    allowed_fixed = {
        CREDIT_BEGIN_HEADER,
        CREDIT_TOTAL_HEADER,
        CREDIT_DEBIT_HEADER,
        CREDIT_ENDING_HEADER,
        CREDIT_DEFAULT_HEADER,
    }
    if s_upper in allowed_fixed:
        return False

    if re.fullmatch(r'-?\d+(\.\d+)?', s):
        return True

    if s_upper in {"COUNT", "TOTAL", "AMOUNT", "MERCHANT", "CATEGORY"}:
        return True

    return False

def classify_credit_column(merchant: str) -> str:
    m = normalize_credit_header_name(merchant)

    default_keywords = [
        "CASH",
        "CHECK",
        "DEPOSIT",
        "CASH APP",
        "CASH DEPOSIT",
        "CHECK DEPOSIT",
        "CASH / CHECK DEPOSIT",
    ]

    for kw in default_keywords:
        if kw in m:
            return CREDIT_DEFAULT_HEADER

    return m

def read_existing_credit_summary_from_wb(wb):
    existing_dynamic_headers = []
    existing_month_values = {m: {} for m in CREDIT_MONTHS}
    bank_name = DEFAULT_BANK_NAME

    if "Credit Summary" not in wb.sheetnames:
        return bank_name, existing_dynamic_headers, existing_month_values

    ws = wb["Credit Summary"]

    bank_cell = ws.cell(row=1, column=2).value
    if bank_cell:
        bank_name = str(bank_cell).strip()

    headers = {}
    for col in range(2, ws.max_column + 1):
        val = ws.cell(row=2, column=col).value
        if val is None:
            continue

        raw_header = str(val).strip()
        if not raw_header:
            continue

        normalized = normalize_credit_header_name(raw_header)

        if is_bad_credit_header_name(normalized):
            continue

        headers[col] = normalized

    fixed_headers = {
        CREDIT_BEGIN_HEADER,
        CREDIT_TOTAL_HEADER,
        CREDIT_DEBIT_HEADER,
        CREDIT_ENDING_HEADER,
    }

    dynamic_cols = []
    for col in sorted(headers.keys()):
        h = headers[col]
        if h not in fixed_headers:
            dynamic_cols.append((col, h))

    existing_dynamic_headers = [h for _, h in dynamic_cols]

    for month, row_idx in CREDIT_MONTH_ROWS.items():
        for col, h in dynamic_cols:
            val = ws.cell(row=row_idx, column=col).value
            if val is not None and str(val).strip() != "":
                existing_month_values[month][h] = safe_float(val)

    return bank_name, existing_dynamic_headers, existing_month_values

def build_credit_month_matrix(existing_month_values, new_rows_by_month):
    month_values = {m: dict(existing_month_values.get(m, {})) for m in CREDIT_MONTHS}

    for month, rows in new_rows_by_month.items():
        for merchant, amount, _ in rows:
            col_name = classify_credit_column(merchant)
            old_val = month_values[month].get(col_name, 0.0)
            month_values[month][col_name] = old_val + safe_float(amount)

    return month_values

def collect_all_credit_headers(month_values, existing_dynamic_headers):
    all_dynamic = []

    all_dynamic.append(CREDIT_DEFAULT_HEADER)

    for h in existing_dynamic_headers:
        h2 = normalize_credit_header_name(h)
        if h2 != CREDIT_DEFAULT_HEADER and h2 not in all_dynamic:
            all_dynamic.append(h2)

    for month in CREDIT_MONTHS:
        for h in month_values.get(month, {}):
            h2 = normalize_credit_header_name(h)
            if h2 != CREDIT_DEFAULT_HEADER and h2 not in all_dynamic:
                all_dynamic.append(h2)

    return all_dynamic

def style_credit_sheet(ws, last_col):
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in range(1, 16):
        for col in range(1, last_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.cell(row=1, column=2).font = Font(bold=True)
    ws.cell(row=1, column=2).alignment = Alignment(horizontal="center", vertical="center")

    for col in range(1, last_col + 1):
        ws.cell(row=2, column=col).font = Font(bold=True)

    for row in range(3, 16):
        for col in range(2, last_col + 1):
            ws.cell(row=row, column=col).number_format = "0.00"

    ws.column_dimensions["A"].width = 16
    for col in range(2, last_col + 1):
        ws.column_dimensions[excel_col_letter(col)].width = 22

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 24

# ================= Debit Summary Sheet =================

def read_existing_debit_summary_from_wb(wb):
    data = {}

    if DEBIT_SHEET_NAME not in wb.sheetnames:
        return data

    ws = wb[DEBIT_SHEET_NAME]

    for row in range(2, ws.max_row + 1):
        merchant = ws.cell(row=row, column=1).value
        if merchant is None:
            continue

        merchant = str(merchant).strip()
        if not merchant or merchant.lower() == "total":
            continue

        if merchant not in data:
            data[merchant] = {m: [] for m in CREDIT_MONTHS}

        for idx, month in enumerate(CREDIT_MONTHS, start=2):
            cell_val = ws.cell(row=row, column=idx).value
            parts = split_formula_parts(cell_val)
            data[merchant][month] = parts

    return data

def read_existing_debit_categories_from_wb(wb) -> Dict[str, str]:
    """Read Merchant -> Category from the current debit summary.

    This preserves categories imported from an existing company report instead
    of replacing them every time new bank-statement data is appended.
    """
    categories: Dict[str, str] = {}

    if DEBIT_SHEET_NAME not in wb.sheetnames:
        return categories

    ws = wb[DEBIT_SHEET_NAME]
    for row in range(2, ws.max_row + 1):
        merchant = ws.cell(row=row, column=1).value
        if merchant is None:
            continue

        merchant_text = str(merchant).strip()
        if not merchant_text or merchant_text.lower() == "total":
            continue

        category = ws.cell(row=row, column=15).value
        category_text = str(category).strip() if category is not None else ""
        if category_text:
            categories[merchant_text] = category_text

    return categories


def style_debit_summary_sheet(ws, last_row: int):
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    max_col = 15  # A~O

    for row in range(1, last_row + 1):
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in range(1, max_col + 1):
        ws.cell(row=1, column=col).font = Font(bold=True)

    ws.column_dimensions["A"].width = 42
    for col in range(2, 14):
        ws.column_dimensions[excel_col_letter(col)].width = 12
    ws.column_dimensions["N"].width = 14
    ws.column_dimensions["O"].width = 18

    for row in range(2, last_row + 1):
        for col in range(2, 15):
            ws.cell(row=row, column=col).number_format = "0.00"

    ws.auto_filter.ref = f"A1:O{last_row}"

def sync_credit_debit_from_debit_sheet(wb):
    if "Credit Summary" not in wb.sheetnames:
        return

    if DEBIT_SHEET_NAME not in wb.sheetnames:
        return

    ws_credit = wb["Credit Summary"]
    ws_debit = wb[DEBIT_SHEET_NAME]

    debit_total_row = ws_debit.max_row

    header_col_map = {}
    for col in range(2, ws_credit.max_column + 1):
        val = ws_credit.cell(row=2, column=col).value
        if val:
            header_col_map[str(val).strip()] = col

    if CREDIT_DEBIT_HEADER not in header_col_map:
        return

    debit_col_credit = header_col_map[CREDIT_DEBIT_HEADER]

    for idx, month in enumerate(CREDIT_MONTHS, start=2):
        credit_row = CREDIT_MONTH_ROWS[month]
        debit_month_col_letter = excel_col_letter(idx)
        ws_credit.cell(
            row=credit_row,
            column=debit_col_credit,
            value=f"='{DEBIT_SHEET_NAME}'!{debit_month_col_letter}{debit_total_row}"
        )
        ws_credit.cell(row=credit_row, column=debit_col_credit).number_format = "0.00"

    ws_credit.cell(
        row=15,
        column=debit_col_credit,
        value=f"='{DEBIT_SHEET_NAME}'!N{debit_total_row}"
    )
    ws_credit.cell(row=15, column=debit_col_credit).number_format = "0.00"

def write_or_update_debit_summary_sheet(xlsx_path: Path, rows, selected_month_ui: str):
    selected_month_ui = selected_month_ui.upper().strip()
    if selected_month_ui not in CREDIT_MONTHS:
        selected_month_ui = "JAN"

    if xlsx_path.exists():
        wb = load_workbook(xlsx_path)
    else:
        wb = Workbook()
        default_ws = wb.active
        default_ws.title = "Sheet"

    month_labels = read_month_labels_from_wb(wb)
    existing = read_existing_debit_summary_from_wb(wb)
    existing_categories = read_existing_debit_categories_from_wb(wb)
    merged = merge_same_merchants(rows)
    rules = load_category_rules()

    for merchant, data in merged.items():
        if merchant not in existing:
            existing[merchant] = {m: [] for m in CREDIT_MONTHS}

        new_parts = [normalize_amount_string(a) for a in data["amounts"]]
        existing[merchant][selected_month_ui].extend(new_parts)

    if DEBIT_SHEET_NAME in wb.sheetnames:
        old_ws = wb[DEBIT_SHEET_NAME]
        wb.remove(old_ws)

    ws = wb.create_sheet(DEBIT_SHEET_NAME)

    headers = ["Merchant"] + normalize_month_labels(month_labels) + ["Total", "Category"]
    for col_idx, h in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=h)

    merchants_sorted = sorted(existing.keys(), key=lambda x: x.strip().lower())

    row_idx = 2
    for merchant in merchants_sorted:
        category = existing_categories.get(merchant, "")
        if not category:
            category = get_category_for_merchant(merchant, rules)

        ws.cell(row=row_idx, column=1, value=merchant)

        for month_i, month in enumerate(CREDIT_MONTHS, start=2):
            parts = existing[merchant].get(month, [])
            formula = build_plus_formula(parts)
            if formula:
                ws.cell(row=row_idx, column=month_i, value=formula)
                ws.cell(row=row_idx, column=month_i).number_format = "0.00"
            else:
                ws.cell(row=row_idx, column=month_i, value=None)

        ws.cell(row=row_idx, column=14, value=f"=SUM(B{row_idx}:M{row_idx})")
        ws.cell(row=row_idx, column=14).number_format = "0.00"
        ws.cell(row=row_idx, column=15, value=category)

        row_idx += 1

    total_row = row_idx
    ws.cell(row=total_row, column=1, value="Total")

    for col in range(2, 14):
        col_letter = excel_col_letter(col)
        if total_row == 2:
            ws.cell(row=total_row, column=col, value=None)
        else:
            ws.cell(row=total_row, column=col, value=f"=SUM({col_letter}2:{col_letter}{total_row-1})")
            ws.cell(row=total_row, column=col).number_format = "0.00"

    if total_row == 2:
        ws.cell(row=total_row, column=14, value=None)
    else:
        ws.cell(row=total_row, column=14, value=f"=SUM(B{total_row}:M{total_row})")
        ws.cell(row=total_row, column=14).number_format = "0.00"

    ws.cell(row=total_row, column=15, value="")

    style_debit_summary_sheet(ws, total_row)
    sync_credit_debit_from_debit_sheet(wb)

    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        maybe_sheet = wb["Sheet"]
        if maybe_sheet.max_row == 1 and maybe_sheet.max_column == 1 and maybe_sheet["A1"].value is None:
            wb.remove(maybe_sheet)

    wb.save(xlsx_path)

# ================= Credit Summary 写入 =================

def write_or_update_credit_summary_xlsx(
    xlsx_path: Path,
    rows,
    selected_month_ui: str,
    bank_name: str = DEFAULT_BANK_NAME
):
    selected_month_ui = selected_month_ui.upper().strip()
    if selected_month_ui not in CREDIT_MONTHS:
        selected_month_ui = "JAN"

    if xlsx_path.exists():
        wb = load_workbook(xlsx_path)
    else:
        wb = Workbook()
        default_ws = wb.active
        default_ws.title = "Sheet"

    month_labels = read_month_labels_from_wb(wb)
    old_bank_name, existing_dynamic_headers, existing_month_values = read_existing_credit_summary_from_wb(wb)
    final_bank_name = (bank_name or "").strip() or old_bank_name or DEFAULT_BANK_NAME

    new_rows_by_month = {m: [] for m in CREDIT_MONTHS}
    new_rows_by_month[selected_month_ui] = list(rows)

    month_values = build_credit_month_matrix(existing_month_values, new_rows_by_month)
    dynamic_headers = collect_all_credit_headers(month_values, existing_dynamic_headers)

    if "Credit Summary" in wb.sheetnames:
        old_ws = wb["Credit Summary"]
        wb.remove(old_ws)

    ws = wb.create_sheet("Credit Summary", 0)

    ws.cell(row=1, column=2, value=final_bank_name)

    headers = [CREDIT_BEGIN_HEADER] + dynamic_headers + [
        CREDIT_TOTAL_HEADER,
        CREDIT_DEBIT_HEADER,
        CREDIT_ENDING_HEADER,
    ]

    for idx, h in enumerate(headers, start=2):
        ws.cell(row=2, column=idx, value=h)

    for index, month in enumerate(CREDIT_MONTHS):
        ws.cell(row=CREDIT_MONTH_ROWS[month], column=1, value=month_labels[index])

    ws.cell(row=15, column=1, value="TOTAL")

    header_col_map = {}
    for col in range(2, len(headers) + 2):
        header_col_map[str(ws.cell(row=2, column=col).value).strip()] = col

    begin_col = header_col_map[CREDIT_BEGIN_HEADER]
    total_credit_col = header_col_map[CREDIT_TOTAL_HEADER]
    debit_col = header_col_map[CREDIT_DEBIT_HEADER]
    ending_col = header_col_map[CREDIT_ENDING_HEADER]

    for month in CREDIT_MONTHS:
        row_idx = CREDIT_MONTH_ROWS[month]

        for h in dynamic_headers:
            raw_val = month_values.get(month, {}).get(h, None)
            col_idx = header_col_map[h]

            if raw_val is None or abs(safe_float(raw_val)) < 1e-12:
                ws.cell(row=row_idx, column=col_idx, value=None)
            else:
                ws.cell(row=row_idx, column=col_idx, value=safe_float(raw_val))
                ws.cell(row=row_idx, column=col_idx).number_format = "0.00"

    jan_row = CREDIT_MONTH_ROWS["JAN"]
    ws.cell(row=jan_row, column=begin_col, value=0)
    ws.cell(row=jan_row, column=begin_col).number_format = "0.00"

    for i in range(1, len(CREDIT_MONTHS)):
        month = CREDIT_MONTHS[i]
        prev_month = CREDIT_MONTHS[i - 1]

        row_idx = CREDIT_MONTH_ROWS[month]
        prev_row = CREDIT_MONTH_ROWS[prev_month]

        prev_ending_ref = ws.cell(row=prev_row, column=ending_col).coordinate
        ws.cell(row=row_idx, column=begin_col, value=f"={prev_ending_ref}")
        ws.cell(row=row_idx, column=begin_col).number_format = "0.00"

    dynamic_start_col = header_col_map[dynamic_headers[0]]
    dynamic_end_col = header_col_map[dynamic_headers[-1]]

    for month in CREDIT_MONTHS:
        row_idx = CREDIT_MONTH_ROWS[month]
        start_ref = ws.cell(row=row_idx, column=dynamic_start_col).coordinate
        end_ref = ws.cell(row=row_idx, column=dynamic_end_col).coordinate

        ws.cell(
            row=row_idx,
            column=total_credit_col,
            value=f"=SUM({start_ref}:{end_ref})"
        )
        ws.cell(row=row_idx, column=total_credit_col).number_format = "0.00"

    for month in CREDIT_MONTHS:
        row_idx = CREDIT_MONTH_ROWS[month]
        ws.cell(row=row_idx, column=debit_col, value=None)

    for month in CREDIT_MONTHS:
        row_idx = CREDIT_MONTH_ROWS[month]

        begin_ref = ws.cell(row=row_idx, column=begin_col).coordinate
        total_credit_ref = ws.cell(row=row_idx, column=total_credit_col).coordinate
        debit_ref = ws.cell(row=row_idx, column=debit_col).coordinate

        ws.cell(
            row=row_idx,
            column=ending_col,
            value=f"={begin_ref}+{total_credit_ref}-{debit_ref}"
        )
        ws.cell(row=row_idx, column=ending_col).number_format = "0.00"

    total_row = 15
    ws.cell(row=total_row, column=begin_col, value=None)

    for h in dynamic_headers:
        col_idx = header_col_map[h]
        start_ref = ws.cell(row=CREDIT_MONTH_ROWS["JAN"], column=col_idx).coordinate
        end_ref = ws.cell(row=CREDIT_MONTH_ROWS["DEC"], column=col_idx).coordinate
        ws.cell(row=total_row, column=col_idx, value=f"=SUM({start_ref}:{end_ref})")
        ws.cell(row=total_row, column=col_idx).number_format = "0.00"

    start_ref = ws.cell(row=CREDIT_MONTH_ROWS["JAN"], column=total_credit_col).coordinate
    end_ref = ws.cell(row=CREDIT_MONTH_ROWS["DEC"], column=total_credit_col).coordinate
    ws.cell(row=total_row, column=total_credit_col, value=f"=SUM({start_ref}:{end_ref})")
    ws.cell(row=total_row, column=total_credit_col).number_format = "0.00"

    ws.cell(row=total_row, column=debit_col, value=None)

    dec_ending_ref = ws.cell(row=CREDIT_MONTH_ROWS["DEC"], column=ending_col).coordinate
    ws.cell(row=total_row, column=ending_col, value=f"={dec_ending_ref}")
    ws.cell(row=total_row, column=ending_col).number_format = "0.00"

    last_col = 1 + len(headers)
    style_credit_sheet(ws, last_col)
    ws.auto_filter.ref = f"A2:{excel_col_letter(last_col)}15"

    sync_credit_debit_from_debit_sheet(wb)

    if DEBIT_SHEET_NAME not in wb.sheetnames:
        start_ref = ws.cell(row=CREDIT_MONTH_ROWS["JAN"], column=debit_col).coordinate
        end_ref = ws.cell(row=CREDIT_MONTH_ROWS["DEC"], column=debit_col).coordinate
        ws.cell(row=15, column=debit_col, value=f"=SUM({start_ref}:{end_ref})")
        ws.cell(row=15, column=debit_col).number_format = "0.00"

    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        maybe_sheet = wb["Sheet"]
        if maybe_sheet.max_row == 1 and maybe_sheet.max_column == 1 and maybe_sheet["A1"].value is None:
            wb.remove(maybe_sheet)

    wb.save(xlsx_path)

# ================= 已有公司报表模板导入 =================

FIXED_CREDIT_HEADERS = {
    CREDIT_BEGIN_HEADER,
    CREDIT_TOTAL_HEADER,
    CREDIT_DEBIT_HEADER,
    CREDIT_ENDING_HEADER,
}


def normalized_cell_text(value) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().split())


def normalized_header_key(value) -> str:
    return normalized_cell_text(value).upper()


def find_header_row(ws, required_headers: List[str], max_scan_rows: int = 80) -> Optional[int]:
    required = {h.upper() for h in required_headers}
    limit = min(ws.max_row, max_scan_rows)

    for row in range(1, limit + 1):
        values = {
            normalized_header_key(ws.cell(row=row, column=col).value)
            for col in range(1, ws.max_column + 1)
        }
        if required.issubset(values):
            return row
    return None


def find_column_by_header(ws, header_row: int, aliases: List[str]) -> Optional[int]:
    alias_set = {a.upper() for a in aliases}
    for col in range(1, ws.max_column + 1):
        if normalized_header_key(ws.cell(row=header_row, column=col).value) in alias_set:
            return col
    return None


def extract_expense_template_from_sheet(ws) -> List[Tuple[str, str]]:
    """Extract only Merchant and Category; all historical amounts are ignored."""
    header_row = find_header_row(ws, ["Merchant", "Category"])
    if header_row is None:
        return []

    merchant_col = find_column_by_header(ws, header_row, ["Merchant", "Vendor", "Description"])
    category_col = find_column_by_header(ws, header_row, ["Category", "Class"])
    if merchant_col is None or category_col is None:
        return []

    rows: List[Tuple[str, str]] = []
    seen = set()

    for row in range(header_row + 1, ws.max_row + 1):
        merchant = normalized_cell_text(ws.cell(row=row, column=merchant_col).value)
        category = normalized_cell_text(ws.cell(row=row, column=category_col).value)

        if not merchant:
            continue
        if merchant.upper() in {"TOTAL", "GRAND TOTAL"}:
            continue

        key = merchant.casefold()
        if key in seen:
            continue
        seen.add(key)
        rows.append((merchant, category))

    return rows


def extract_income_headers_from_credit_layout(ws) -> List[str]:
    """Find dynamic income columns from a Credit Summary-style report."""
    limit = min(ws.max_row, 80)
    for row in range(1, limit + 1):
        values = [
            normalized_header_key(ws.cell(row=row, column=col).value)
            for col in range(1, ws.max_column + 1)
        ]
        value_set = set(values)

        # Strong signal for the report layout already used by this program.
        if CREDIT_TOTAL_HEADER in value_set and CREDIT_ENDING_HEADER in value_set:
            headers: List[str] = []
            for value in values:
                if not value or value in FIXED_CREDIT_HEADERS:
                    continue
                if value in {"MONTH", "MERCHANT", "CATEGORY", "TOTAL", "AMOUNT"}:
                    continue
                if value not in headers:
                    headers.append(value)
            return headers
    return []


def extract_income_headers_from_merchant_layout(ws) -> List[str]:
    """Support an Income sheet arranged like Merchant | Jan..Dec | Total | Category."""
    header_row = find_header_row(ws, ["Merchant"])
    if header_row is None:
        return []

    merchant_col = find_column_by_header(ws, header_row, ["Merchant", "Income", "Source", "Description"])
    if merchant_col is None:
        return []

    headers: List[str] = []
    seen = set()
    for row in range(header_row + 1, ws.max_row + 1):
        name = normalized_cell_text(ws.cell(row=row, column=merchant_col).value)
        if not name or name.upper() in {"TOTAL", "GRAND TOTAL"}:
            continue
        key = name.casefold()
        if key in seen:
            continue
        seen.add(key)
        headers.append(normalize_credit_header_name(name))
    return headers


def display_cell_text(cell) -> str:
    """Return a readable Excel cell value while preserving date-like labels."""
    value = cell.value
    if value is None:
        return ""

    # Excel date/datetime cells should be displayed using their own number format
    # where possible, instead of Python's ``2026-01-01 00:00:00`` representation.
    try:
        if getattr(cell, "is_date", False) and hasattr(value, "strftime"):
            fmt = str(cell.number_format or "").lower()
            if "yyyy" in fmt:
                return value.strftime("%m/%d/%Y")
            if "yy" in fmt:
                return value.strftime("%m/%d/%y")
            return value.strftime("%m/%d")
    except Exception:
        pass

    return normalized_cell_text(value)




def looks_like_period_header(cell) -> bool:
    """Return True when a header cell looks like a month/date/period label.

    This includes normal month names (AUG), date-like labels (31-Jul, 6-Aug,
    08/06/2026), and real Excel date cells. It allows valid period columns
    to appear after Total/Category instead of requiring one contiguous block.
    """
    value = cell.value
    if value is None:
        return False
    try:
        if getattr(cell, "is_date", False):
            return True
    except Exception:
        pass

    text = display_cell_text(cell).strip()
    if not text:
        return False
    key = text.upper().replace(".", "")
    if key in {
        "JAN", "JANUARY", "FEB", "FEBRUARY", "MAR", "MARCH",
        "APR", "APRIL", "MAY", "JUN", "JUNE", "JUL", "JULY",
        "AUG", "AUGUST", "SEP", "SEPT", "SEPTEMBER",
        "OCT", "OCTOBER", "NOV", "NOVEMBER", "DEC", "DECEMBER"
    }:
        return True

    month_word = r"(?:JAN(?:UARY)?|FEB(?:RUARY)?|MAR(?:CH)?|APR(?:IL)?|MAY|JUN(?:E)?|JUL(?:Y)?|AUG(?:UST)?|SEP(?:T(?:EMBER)?)?|OCT(?:OBER)?|NOV(?:EMBER)?|DEC(?:EMBER)?)"
    patterns = [
        rf"^\d{{1,2}}[-/\s]{month_word}(?:[-/\s]\d{{2,4}})?$",
        rf"^{month_word}[-/\s]\d{{1,2}}(?:[-/\s]\d{{2,4}})?$",
        r"^\d{1,2}[-/]\d{1,2}(?:[-/]\d{2,4})?$",
        r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}$",
    ]
    return any(re.fullmatch(p, text, re.IGNORECASE) for p in patterns)


def build_sum_formula_for_columns(ws, row: int, columns: List[int]) -> str:
    """Build a SUM formula for possibly non-contiguous period columns."""
    refs = [ws.cell(row=row, column=c).coordinate for c in columns]
    return f"=SUM({','.join(refs)})" if refs else ""


def extract_month_labels_from_sheet(ws) -> List[str]:
    """Read every horizontal period/date header from a Merchant-style sheet.

    Period columns are all columns between Merchant and Total (or Category).
    This intentionally supports more than 12 periods, for example an extra
    ``31-Jul`` column after the normal monthly columns.
    """
    header_row = find_header_row(ws, ["Merchant"])
    if header_row is None:
        return []
    merchant_col = find_column_by_header(ws, header_row, ["Merchant", "Vendor", "Description"])
    if merchant_col is None:
        return []

    total_col = find_column_by_header(ws, header_row, ["Total", "Amount Total", "Grand Total"])
    category_col = find_column_by_header(ws, header_row, ["Category", "Class"])

    candidates: List[str] = []
    for col in range(merchant_col + 1, ws.max_column + 1):
        if col in {total_col, category_col}:
            continue
        cell = ws.cell(row=header_row, column=col)
        # Columns before Total remain compatible with the original layout.
        # Columns after Category/Total are included only when they look like dates.
        if (total_col is not None and col < total_col) or looks_like_period_header(cell):
            text = display_cell_text(cell)
            if text:
                candidates.append(text)
    return candidates


def find_credit_period_rows(ws, header_row: int, begin_col: int) -> Tuple[List[int], Optional[int]]:
    """Return all labeled period rows and the TOTAL row for a Credit layout."""
    if begin_col <= 1:
        return [], None

    label_col = begin_col - 1
    period_rows: List[int] = []
    total_row: Optional[int] = None

    for row in range(header_row + 1, ws.max_row + 1):
        label = display_cell_text(ws.cell(row=row, column=label_col)).strip()
        key = label.upper()
        if key in {"TOTAL", "GRAND TOTAL"}:
            total_row = row
            break
        if label:
            period_rows.append(row)

    # Some templates may not already contain a TOTAL row.
    if total_row is None and period_rows:
        total_row = period_rows[-1] + 1
    return period_rows, total_row


def extract_month_labels_from_income_sheet(ws) -> List[str]:
    """Read every date/period label from an Income/Credit page."""
    limit = min(ws.max_row, 100)
    begin_row = None
    begin_col = None
    for row in range(1, limit + 1):
        for col in range(1, ws.max_column + 1):
            if normalized_header_key(ws.cell(row=row, column=col).value) == CREDIT_BEGIN_HEADER:
                begin_row, begin_col = row, col
                break
        if begin_row is not None:
            break

    if begin_row is not None and begin_col is not None and begin_col > 1:
        period_rows, _ = find_credit_period_rows(ws, begin_row, begin_col)
        labels = [display_cell_text(ws.cell(row=r, column=begin_col - 1)) for r in period_rows]
        if labels:
            return labels

    return extract_month_labels_from_sheet(ws)


def discover_source_report_structure(source_path: Path):
    """Discover income names and expense Merchant/Category rows.

    Amount cells from the source workbook are deliberately never imported.
    """
    wb = load_workbook(source_path, data_only=False)

    income_headers: List[str] = []
    expense_rows: List[Tuple[str, str]] = []
    month_labels: List[str] = []

    # Income is scanned first because its 12 date/period labels are the
    # authoritative labels for both generated sheets and for the UI dropdown.
    ordered_income_sheets = sorted(
        wb.worksheets,
        key=lambda ws: (0 if any(k in ws.title.lower() for k in ("income", "credit")) else 1)
    )
    for ws in ordered_income_sheets:
        found = extract_income_headers_from_credit_layout(ws)
        if not found and any(k in ws.title.lower() for k in ("income", "credit")):
            found = extract_income_headers_from_merchant_layout(ws)

        detected_labels = extract_month_labels_from_income_sheet(ws)
        if found:
            income_headers = found
            if detected_labels:
                month_labels = detected_labels
            break

    # Expense contributes only Merchant and Category. Its own date headers are
    # deliberately ignored because the Income page controls the date labels.
    ordered_expense_sheets = sorted(
        wb.worksheets,
        key=lambda ws: (0 if any(k in ws.title.lower() for k in ("expense", "debit")) else 1)
    )
    for ws in ordered_expense_sheets:
        found = extract_expense_template_from_sheet(ws)
        if found:
            expense_rows = found
            break

    # Always keep the program's standard deposit column.
    final_income_headers = [CREDIT_DEFAULT_HEADER]
    for header in income_headers:
        normalized = normalize_credit_header_name(header)
        if normalized and normalized not in FIXED_CREDIT_HEADERS and normalized not in final_income_headers:
            final_income_headers.append(normalized)

    return final_income_headers, expense_rows, normalize_month_labels(month_labels)


def create_credit_template_sheet(wb, income_headers: List[str], bank_name: str, month_labels: List[str]):
    if "Credit Summary" in wb.sheetnames:
        wb.remove(wb["Credit Summary"])

    ws = wb.create_sheet("Credit Summary", 0)
    ws.cell(row=1, column=2, value=(bank_name or DEFAULT_BANK_NAME).strip() or DEFAULT_BANK_NAME)

    dynamic_headers = []
    for header in income_headers:
        normalized = normalize_credit_header_name(header)
        if normalized and normalized not in FIXED_CREDIT_HEADERS and normalized not in dynamic_headers:
            dynamic_headers.append(normalized)
    if CREDIT_DEFAULT_HEADER not in dynamic_headers:
        dynamic_headers.insert(0, CREDIT_DEFAULT_HEADER)

    headers = [CREDIT_BEGIN_HEADER] + dynamic_headers + [
        CREDIT_TOTAL_HEADER,
        CREDIT_DEBIT_HEADER,
        CREDIT_ENDING_HEADER,
    ]

    for col, header in enumerate(headers, start=2):
        ws.cell(row=2, column=col, value=header)

    month_labels = normalize_month_labels(month_labels)
    for index, month in enumerate(CREDIT_MONTHS):
        ws.cell(row=CREDIT_MONTH_ROWS[month], column=1, value=month_labels[index])
    ws.cell(row=15, column=1, value="TOTAL")

    header_col_map = {
        normalized_cell_text(ws.cell(row=2, column=col).value): col
        for col in range(2, len(headers) + 2)
    }
    begin_col = header_col_map[CREDIT_BEGIN_HEADER]
    total_credit_col = header_col_map[CREDIT_TOTAL_HEADER]
    debit_col = header_col_map[CREDIT_DEBIT_HEADER]
    ending_col = header_col_map[CREDIT_ENDING_HEADER]
    dynamic_start_col = header_col_map[dynamic_headers[0]]
    dynamic_end_col = header_col_map[dynamic_headers[-1]]

    # All imported historical amounts are intentionally blank.
    ws.cell(row=CREDIT_MONTH_ROWS["JAN"], column=begin_col, value=0)
    ws.cell(row=CREDIT_MONTH_ROWS["JAN"], column=begin_col).number_format = "0.00"

    for index, month in enumerate(CREDIT_MONTHS):
        row = CREDIT_MONTH_ROWS[month]
        if index > 0:
            previous_row = CREDIT_MONTH_ROWS[CREDIT_MONTHS[index - 1]]
            previous_ending = ws.cell(row=previous_row, column=ending_col).coordinate
            ws.cell(row=row, column=begin_col, value=f"={previous_ending}")

        start_ref = ws.cell(row=row, column=dynamic_start_col).coordinate
        end_ref = ws.cell(row=row, column=dynamic_end_col).coordinate
        ws.cell(row=row, column=total_credit_col, value=f"=SUM({start_ref}:{end_ref})")

        begin_ref = ws.cell(row=row, column=begin_col).coordinate
        total_credit_ref = ws.cell(row=row, column=total_credit_col).coordinate
        debit_ref = ws.cell(row=row, column=debit_col).coordinate
        ws.cell(row=row, column=ending_col, value=f"={begin_ref}+{total_credit_ref}-{debit_ref}")

    # Annual formulas remain formulas, not pasted values.
    for header in dynamic_headers:
        col = header_col_map[header]
        start_ref = ws.cell(row=CREDIT_MONTH_ROWS["JAN"], column=col).coordinate
        end_ref = ws.cell(row=CREDIT_MONTH_ROWS["DEC"], column=col).coordinate
        ws.cell(row=15, column=col, value=f"=SUM({start_ref}:{end_ref})")

    start_ref = ws.cell(row=CREDIT_MONTH_ROWS["JAN"], column=total_credit_col).coordinate
    end_ref = ws.cell(row=CREDIT_MONTH_ROWS["DEC"], column=total_credit_col).coordinate
    ws.cell(row=15, column=total_credit_col, value=f"=SUM({start_ref}:{end_ref})")

    start_ref = ws.cell(row=CREDIT_MONTH_ROWS["JAN"], column=debit_col).coordinate
    end_ref = ws.cell(row=CREDIT_MONTH_ROWS["DEC"], column=debit_col).coordinate
    ws.cell(row=15, column=debit_col, value=f"=SUM({start_ref}:{end_ref})")

    dec_ending_ref = ws.cell(row=CREDIT_MONTH_ROWS["DEC"], column=ending_col).coordinate
    ws.cell(row=15, column=ending_col, value=f"={dec_ending_ref}")

    last_col = 1 + len(headers)
    style_credit_sheet(ws, last_col)
    ws.auto_filter.ref = f"A2:{excel_col_letter(last_col)}15"


def create_debit_template_sheet(wb, expense_rows: List[Tuple[str, str]], month_labels: List[str]):
    if DEBIT_SHEET_NAME in wb.sheetnames:
        wb.remove(wb[DEBIT_SHEET_NAME])

    ws = wb.create_sheet(DEBIT_SHEET_NAME)
    headers = ["Merchant"] + month_labels + ["Total", "Category"]
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=header)

    row = 2
    for merchant, category in expense_rows:
        ws.cell(row=row, column=1, value=merchant)
        # Jan-Dec remain blank. Total is still a live formula.
        ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
        ws.cell(row=row, column=14).number_format = "0.00"
        ws.cell(row=row, column=15, value=category)
        row += 1

    total_row = row
    ws.cell(row=total_row, column=1, value="Total")
    for col in range(2, 14):
        letter = excel_col_letter(col)
        if total_row > 2:
            ws.cell(row=total_row, column=col, value=f"=SUM({letter}2:{letter}{total_row - 1})")
            ws.cell(row=total_row, column=col).number_format = "0.00"
    if total_row > 2:
        ws.cell(row=total_row, column=14, value=f"=SUM(B{total_row}:M{total_row})")
        ws.cell(row=total_row, column=14).number_format = "0.00"

    style_debit_summary_sheet(ws, total_row)


def clear_credit_amounts_keep_layout(ws) -> int:
    """Clear historical Credit amounts in-place while preserving sheet name/layout."""
    header_row, header_map = find_credit_layout(ws)
    if header_row is None:
        return 0

    begin_col = header_map[CREDIT_BEGIN_HEADER]
    total_credit_col = header_map[CREDIT_TOTAL_HEADER]
    debit_col = header_map[CREDIT_DEBIT_HEADER]
    ending_col = header_map[CREDIT_ENDING_HEADER]
    period_rows, total_row = find_credit_period_rows(ws, header_row, begin_col)
    if not period_rows:
        return 0
    first_data_row = period_rows[0]
    last_data_row = period_rows[-1]
    if total_row is None:
        total_row = last_data_row + 1

    dynamic_cols = list(range(begin_col + 1, total_credit_col))
    for row in period_rows:
        for col in dynamic_cols:
            ws.cell(row=row, column=col).value = None
        ws.cell(row=row, column=debit_col).value = None

    # Rebuild the controlled formulas without renaming or recreating the sheet.
    if first_data_row <= ws.max_row:
        ws.cell(row=first_data_row, column=begin_col, value=0)
        ws.cell(row=first_data_row, column=begin_col).number_format = "0.00"

    for row_index, row in enumerate(period_rows):
        if row_index > 0:
            prev_end = ws.cell(row=period_rows[row_index - 1], column=ending_col).coordinate
            ws.cell(row=row, column=begin_col, value=f"={prev_end}")
        if dynamic_cols:
            start_ref = ws.cell(row=row, column=dynamic_cols[0]).coordinate
            end_ref = ws.cell(row=row, column=dynamic_cols[-1]).coordinate
            ws.cell(row=row, column=total_credit_col, value=f"=SUM({start_ref}:{end_ref})")
        else:
            ws.cell(row=row, column=total_credit_col, value=0)
        begin_ref = ws.cell(row=row, column=begin_col).coordinate
        credit_ref = ws.cell(row=row, column=total_credit_col).coordinate
        debit_ref = ws.cell(row=row, column=debit_col).coordinate
        ws.cell(row=row, column=ending_col, value=f"={begin_ref}+{credit_ref}-{debit_ref}")

    for col in dynamic_cols + [total_credit_col, debit_col]:
        start_ref = ws.cell(row=first_data_row, column=col).coordinate
        end_ref = ws.cell(row=last_data_row, column=col).coordinate
        ws.cell(row=total_row, column=col, value=f"=SUM({start_ref}:{end_ref})")
    ws.cell(row=total_row, column=ending_col,
            value=f"={ws.cell(row=last_data_row, column=ending_col).coordinate}")
    return len(dynamic_cols)


def clear_debit_amounts_keep_layout(ws) -> int:
    """Clear historical Debit amounts in-place; keep Merchant/Category and sheet name."""
    layout = find_debit_layout(ws)
    if layout is None:
        return 0
    header_row, merchant_col, month_cols, total_col, category_col = layout

    total_row = None
    merchant_count = 0
    for row in range(header_row + 1, ws.max_row + 1):
        merchant = normalized_cell_text(ws.cell(row=row, column=merchant_col).value)
        if not merchant:
            continue
        if merchant.upper() in {"TOTAL", "GRAND TOTAL"}:
            total_row = row
            break
        merchant_count += 1
        for col in month_cols:
            ws.cell(row=row, column=col).value = None
        ws.cell(row=row, column=total_col, value=build_sum_formula_for_columns(ws, row, month_cols))
        ws.cell(row=row, column=total_col).number_format = "0.00"

    if total_row is None:
        total_row = ws.max_row + 1
        ws.cell(row=total_row, column=merchant_col, value="Total")

    first_data_row = header_row + 1
    last_data_row = total_row - 1
    if last_data_row >= first_data_row:
        for col in month_cols:
            letter = excel_col_letter(col)
            ws.cell(row=total_row, column=col,
                    value=f"=SUM({letter}{first_data_row}:{letter}{last_data_row})")
            ws.cell(row=total_row, column=col).number_format = "0.00"
        ws.cell(row=total_row, column=total_col,
                value=build_sum_formula_for_columns(ws, total_row, month_cols))
        ws.cell(row=total_row, column=total_col).number_format = "0.00"
    return merchant_count


def import_existing_report_template(
    source_path: Path,
    target_path: Path,
    bank_name: str = DEFAULT_BANK_NAME,
    clear_existing_data: bool = True,
):
    """完整复制源工作簿的所有工作表，并保留原始名称和顺序。

    clear_existing_data=True：模板模式，清空可识别 Credit/Debit 页面中的旧月份金额，
    但保留工作表名称、顺序、格式、公式框架、Merchant 与 Category。

    clear_existing_data=False：完整保留模式，源 Excel 中所有原有数据、公式、格式、
    工作表名称和顺序均保持不变，之后可在 UI 中继续追加新数据。
    """
    import shutil
    import tempfile

    source_path = Path(source_path).expanduser().resolve()
    target_path = Path(target_path).expanduser().resolve()

    if not source_path.exists():
        raise FileNotFoundError(f"找不到源 Excel：{source_path}")
    if source_path.suffix.lower() != ".xlsx":
        raise ValueError("目前只支持 .xlsx 文件。")
    if source_path == target_path:
        raise ValueError("源报表和目标 Excel 不能是同一个文件。")

    # 先读取源工作簿并记录所有工作表名称。这里包含隐藏工作表。
    source_wb = load_workbook(source_path, read_only=True, data_only=False)
    try:
        source_sheet_names = list(source_wb.sheetnames)
    finally:
        source_wb.close()

    if not source_sheet_names:
        raise ValueError("源 Excel 中没有可读取的工作表。")

    target_path.parent.mkdir(parents=True, exist_ok=True)

    # 使用同目录临时文件，避免覆盖过程中出现只剩部分工作表的损坏文件。
    fd, temp_name = tempfile.mkstemp(
        prefix=target_path.stem + "_import_",
        suffix=".xlsx",
        dir=str(target_path.parent),
    )
    import os
    os.close(fd)
    temp_path = Path(temp_name)

    try:
        shutil.copyfile(source_path, temp_path)

        # 第一次验证：整本复制后，所有 sheet 名称和顺序必须完全一致。
        check_wb = load_workbook(temp_path, read_only=True, data_only=False)
        try:
            copied_sheet_names = list(check_wb.sheetnames)
        finally:
            check_wb.close()

        if copied_sheet_names != source_sheet_names:
            raise RuntimeError(
                "工作表复制验证失败。\n"
                f"源文件工作表({len(source_sheet_names)}): {source_sheet_names}\n"
                f"复制后工作表({len(copied_sheet_names)}): {copied_sheet_names}"
            )

        wb = load_workbook(temp_path, data_only=False)
        try:
            credit_item_count = 0
            debit_merchant_count = 0
            month_labels: List[str] = []

            # 注意：这里只读取/修改工作表内容，绝不 remove/create/rename worksheet。
            for ws in list(wb.worksheets):
                if not month_labels:
                    labels = extract_month_labels_from_income_sheet(ws)
                    if not labels:
                        labels = extract_month_labels_from_sheet(ws)
                    if labels:
                        month_labels = list(labels)

                header_row, _ = find_credit_layout(ws)
                if header_row is not None:
                    if clear_existing_data:
                        credit_item_count += clear_credit_amounts_keep_layout(ws)
                    else:
                        # 完整保留模式只统计可识别的 Credit 项目，不改变任何单元格。
                        layout = find_credit_layout(ws)
                        if layout[0] is not None:
                            credit_item_count += max(0, len(layout[1]))
                    continue

                debit_layout = find_debit_layout(ws)
                if debit_layout is not None:
                    if clear_existing_data:
                        debit_merchant_count += clear_debit_amounts_keep_layout(ws)
                    else:
                        # 完整保留模式只统计 Merchant，不改变任何单元格。
                        header_row_d, merchant_col, _, _, _ = debit_layout
                        for row in range(header_row_d + 1, ws.max_row + 1):
                            merchant = normalized_cell_text(ws.cell(row=row, column=merchant_col).value)
                            if not merchant or merchant.upper() in {"TOTAL", "GRAND TOTAL"}:
                                continue
                            debit_merchant_count += 1

            # 保存前再次确认代码没有意外删除、增加或改名。
            before_save_names = list(wb.sheetnames)
            if before_save_names != source_sheet_names:
                raise RuntimeError(
                    "处理过程中工作表名称或数量发生变化。\n"
                    f"源文件: {source_sheet_names}\n处理后: {before_save_names}"
                )
            wb.save(temp_path)
        finally:
            wb.close()

        # 第二次验证：保存并重新打开后，仍然必须保留全部工作表。
        final_check_wb = load_workbook(temp_path, read_only=True, data_only=False)
        try:
            final_sheet_names = list(final_check_wb.sheetnames)
        finally:
            final_check_wb.close()

        if final_sheet_names != source_sheet_names:
            raise RuntimeError(
                "保存后的工作表验证失败。\n"
                f"源文件工作表({len(source_sheet_names)}): {source_sheet_names}\n"
                f"保存后工作表({len(final_sheet_names)}): {final_sheet_names}"
            )

        # 所有验证通过后再替换正式目标文件。
        os.replace(temp_path, target_path)

        return (
            credit_item_count,
            debit_merchant_count,
            (month_labels if month_labels else list(DEFAULT_UI_MONTH_LABELS)),
            final_sheet_names,
        )
    finally:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass


# ================= 多工作表读取与原位追加 =================

def list_workbook_sheet_names(path: Path) -> List[str]:
    path = Path(path)
    if not path.exists():
        return []
    wb = load_workbook(path, read_only=True, data_only=False)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def read_month_labels_from_selected_sheet(path: Path, sheet_name: str, account_type: str) -> List[str]:
    path = Path(path)
    if not path.exists() or not sheet_name:
        return list(DEFAULT_UI_MONTH_LABELS)
    wb = load_workbook(path, read_only=True, data_only=False)
    try:
        if sheet_name not in wb.sheetnames:
            return list(DEFAULT_UI_MONTH_LABELS)
        ws = wb[sheet_name]
        if str(account_type).strip().lower() == "credit":
            labels = extract_month_labels_from_income_sheet(ws)
        else:
            labels = extract_month_labels_from_sheet(ws)
        return labels if labels else list(DEFAULT_UI_MONTH_LABELS)
    finally:
        wb.close()


def append_amount_to_cell(cell, amounts: List[str]):
    parts = split_formula_parts(cell.value)
    parts.extend(normalize_amount_string(a) for a in amounts)
    formula = build_plus_formula(parts)
    cell.value = formula if formula else None
    cell.number_format = "0.00"


def copy_row_style(ws, source_row: int, target_row: int, max_col: int):
    from copy import copy
    if source_row < 1:
        return
    for col in range(1, max_col + 1):
        src = ws.cell(row=source_row, column=col)
        dst = ws.cell(row=target_row, column=col)
        if src.has_style:
            dst._style = copy(src._style)
        dst.alignment = copy(src.alignment)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.font = copy(src.font)
        dst.protection = copy(src.protection)
        dst.number_format = src.number_format
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def copy_column_style(ws, source_col: int, target_col: int, max_row: int):
    from copy import copy
    if source_col < 1:
        return
    source_letter = excel_col_letter(source_col)
    target_letter = excel_col_letter(target_col)
    ws.column_dimensions[target_letter].width = ws.column_dimensions[source_letter].width
    for row in range(1, max_row + 1):
        src = ws.cell(row=row, column=source_col)
        dst = ws.cell(row=row, column=target_col)
        if src.has_style:
            dst._style = copy(src._style)
        dst.alignment = copy(src.alignment)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.font = copy(src.font)
        dst.protection = copy(src.protection)
        dst.number_format = src.number_format


def find_credit_layout(ws):
    limit = min(ws.max_row, 100)
    for row in range(1, limit + 1):
        header_map = {}
        for col in range(1, ws.max_column + 1):
            key = normalized_header_key(ws.cell(row=row, column=col).value)
            if key:
                header_map[key] = col
        required = {CREDIT_BEGIN_HEADER, CREDIT_TOTAL_HEADER, CREDIT_DEBIT_HEADER, CREDIT_ENDING_HEADER}
        if required.issubset(header_map):
            return row, header_map
    return None, {}


def find_debit_layout(ws):
    header_row = find_header_row(ws, ["Merchant"])
    if header_row is None:
        return None
    merchant_col = find_column_by_header(ws, header_row, ["Merchant", "Vendor", "Description"])
    total_col = find_column_by_header(ws, header_row, ["Total", "Amount Total", "Grand Total"])
    category_col = find_column_by_header(ws, header_row, ["Category", "Class"])
    if merchant_col is None or total_col is None:
        return None
    # Standard period columns are between Merchant and Total. In addition,
    # date-like headers anywhere after Total/Category are also writable periods.
    month_cols = list(range(merchant_col + 1, total_col))
    for col in range(total_col + 1, ws.max_column + 1):
        if col == category_col:
            continue
        if looks_like_period_header(ws.cell(row=header_row, column=col)):
            month_cols.append(col)
    if not month_cols:
        return None
    return header_row, merchant_col, month_cols, total_col, category_col


def update_credit_sheet_in_place(ws, rows, selected_month_index: int, bank_name: str = DEFAULT_BANK_NAME):
    header_row, header_map = find_credit_layout(ws)
    if header_row is None:
        raise ValueError(f"工作表“{ws.title}”不是可识别的 Credit 格式。")
    begin_col = header_map[CREDIT_BEGIN_HEADER]
    period_rows, total_row = find_credit_period_rows(ws, header_row, begin_col)
    if selected_month_index < 0 or selected_month_index >= len(period_rows):
        raise ValueError(
            f"工作表“{ws.title}”只有 {len(period_rows)} 个可写入日期/期间，"
            f"无法写入第 {selected_month_index + 1} 个。"
        )
    month_row = period_rows[selected_month_index]
    if begin_col > 1:
        bank_cell = ws.cell(row=max(1, header_row - 1), column=begin_col)
        if bank_cell.value is None or str(bank_cell.value).strip() == "":
            bank_cell.value = bank_name
    merged = merge_same_merchants(rows)
    fixed = {CREDIT_BEGIN_HEADER, CREDIT_TOTAL_HEADER, CREDIT_DEBIT_HEADER, CREDIT_ENDING_HEADER}
    def refresh_map():
        return {normalized_header_key(ws.cell(row=header_row, column=c).value): c
                for c in range(1, ws.max_column + 1)
                if normalized_header_key(ws.cell(row=header_row, column=c).value)}
    for merchant, data in merged.items():
        target_header = classify_credit_column(merchant)
        current_map = refresh_map()
        target_col = current_map.get(target_header)
        if target_col is None:
            current_total_col = current_map[CREDIT_TOTAL_HEADER]
            ws.insert_cols(current_total_col, 1)
            copy_column_style(ws, max(begin_col + 1, current_total_col - 1), current_total_col,
                              max(ws.max_row, (total_row or header_row + len(period_rows) + 1)))
            ws.cell(row=header_row, column=current_total_col, value=target_header)
            for r in period_rows:
                ws.cell(row=r, column=current_total_col, value=None)
                ws.cell(row=r, column=current_total_col).number_format = "0.00"
            target_col = current_total_col
        append_amount_to_cell(ws.cell(row=month_row, column=target_col), data["amounts"])
    current_map = refresh_map()
    begin_col = current_map[CREDIT_BEGIN_HEADER]
    total_credit_col = current_map[CREDIT_TOTAL_HEADER]
    debit_col = current_map[CREDIT_DEBIT_HEADER]
    ending_col = current_map[CREDIT_ENDING_HEADER]
    dynamic_cols = [c for c in range(begin_col + 1, total_credit_col)
                    if normalized_header_key(ws.cell(row=header_row, column=c).value) not in fixed]
    if not dynamic_cols:
        raise ValueError(f"工作表“{ws.title}”没有可写入的 Credit 收入栏。")
    if not period_rows:
        raise ValueError(f"工作表“{ws.title}”没有可识别的日期/期间行。")
    if total_row is None:
        total_row = period_rows[-1] + 1
        if begin_col > 1:
            ws.cell(total_row, begin_col - 1, "TOTAL")

    for r in period_rows:
        a = ws.cell(r, dynamic_cols[0]).coordinate
        b = ws.cell(r, dynamic_cols[-1]).coordinate
        ws.cell(r, total_credit_col, f"=SUM({a}:{b})")
        ws.cell(r, ending_col, f"={ws.cell(r, begin_col).coordinate}+{ws.cell(r, total_credit_col).coordinate}-{ws.cell(r, debit_col).coordinate}")

    first_data_row, last_data_row = period_rows[0], period_rows[-1]
    for c in dynamic_cols + [total_credit_col, debit_col]:
        ws.cell(total_row, c, f"=SUM({ws.cell(first_data_row,c).coordinate}:{ws.cell(last_data_row,c).coordinate})")
    ws.cell(total_row, ending_col, f"={ws.cell(last_data_row, ending_col).coordinate}")


def update_debit_sheet_in_place(ws, rows, selected_month_index: int):
    layout = find_debit_layout(ws)
    if layout is None:
        raise ValueError(f"工作表“{ws.title}”不是可识别的 Debit 格式。")
    header_row, merchant_col, month_cols, total_col, category_col = layout
    if selected_month_index < 0 or selected_month_index >= len(month_cols):
        raise ValueError(
            f"工作表“{ws.title}”只有 {len(month_cols)} 个可写入日期/期间，"
            f"无法写入第 {selected_month_index + 1} 个。"
        )
    selected_col = month_cols[selected_month_index]
    rules = load_category_rules()
    merged = merge_same_merchants(rows)
    total_row = None
    merchant_rows = {}
    for r in range(header_row + 1, ws.max_row + 1):
        name = normalized_cell_text(ws.cell(r, merchant_col).value)
        if not name:
            continue
        if name.upper() in {"TOTAL", "GRAND TOTAL"}:
            total_row = r
            break
        merchant_rows[name.casefold()] = r
    if total_row is None:
        total_row = ws.max_row + 1
        ws.cell(total_row, merchant_col, "Total")
    for merchant, data in merged.items():
        r = merchant_rows.get(merchant.casefold())
        if r is None:
            ws.insert_rows(total_row, 1)
            copy_row_style(ws, max(header_row + 1, total_row - 1), total_row, ws.max_column)
            r = total_row
            total_row += 1
            ws.cell(r, merchant_col, merchant)
            for c in month_cols:
                ws.cell(r, c, None)
            if category_col is not None:
                ws.cell(r, category_col, get_category_for_merchant(merchant, rules))
            merchant_rows[merchant.casefold()] = r
        append_amount_to_cell(ws.cell(r, selected_col), data["amounts"])
        ws.cell(r, total_col, build_sum_formula_for_columns(ws, r, month_cols))
    ws.cell(total_row, merchant_col, "Total")
    first_data_row, last_data_row = header_row + 1, total_row - 1
    if last_data_row >= first_data_row:
        for c in month_cols:
            letter = excel_col_letter(c)
            ws.cell(total_row, c, f"=SUM({letter}{first_data_row}:{letter}{last_data_row})")
        ws.cell(total_row, total_col,
                build_sum_formula_for_columns(ws, total_row, month_cols))


def append_rows_to_selected_sheet(xlsx_path: Path, sheet_name: str, rows,
                                  selected_month_index: int, account_type: str,
                                  bank_name: str = DEFAULT_BANK_NAME):
    xlsx_path = Path(xlsx_path)
    if not xlsx_path.exists():
        raise FileNotFoundError(f"找不到 Excel 文件：{xlsx_path}")
    wb = load_workbook(xlsx_path, data_only=False)
    try:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Excel 中找不到工作表：{sheet_name}")
        ws = wb[sheet_name]
        if str(account_type).strip().lower() == "credit":
            update_credit_sheet_in_place(ws, rows, selected_month_index, bank_name)
        elif str(account_type).strip().lower() == "debit":
            update_debit_sheet_in_place(ws, rows, selected_month_index)
        else:
            raise ValueError("Account Type 必须是 Credit 或 Debit。")
        wb.save(xlsx_path)
    finally:
        wb.close()

# ================= 输出路径 =================

def get_summary_output_path(base_path_str: str) -> Path:
    """Return the exact selected .xlsx path, or the default file in a folder."""
    raw = (base_path_str or "").strip()

    if raw:
        p = Path(raw).expanduser()
        if p.suffix.lower() == ".xlsx":
            p.parent.mkdir(parents=True, exist_ok=True)
            return p
        folder = p
    else:
        folder = script_dir()

    folder.mkdir(parents=True, exist_ok=True)
    return folder / "credit_monthly_summary.xlsx"

# ================= GUI =================

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

BG = "#1e1e1e"
FG = "#ffffff"
BTN = "#2d2d2d"

def mk_label(parent, text="", **kw):
    params = {"bg": BG, "fg": FG}
    params.update(kw)
    return tk.Label(parent, text=text, **params)

def mk_button(parent, text, cmd):
    return tk.Button(
        parent,
        text=text,
        command=cmd,
        bg=BTN,
        fg=FG,
        activebackground="#3a3a3a",
        relief="raised",
        padx=10,
        pady=6,
        bd=1,
        highlightthickness=0
    )

def run_parser_ui():
    root = tk.Tk()
    root.title("BSDP - Bank Statement Data Processing")
    root.geometry("1120x950")
    root.configure(bg=BG)

    mk_label(root, "BSDP - Bank Statement Data Processing", font=("Helvetica", 16, "bold")).pack(anchor="w", padx=16, pady=(12, 6))

    default_monthly_summary = script_dir() / "credit_monthly_summary.xlsx"

    summary_var = tk.StringVar(value=str(default_monthly_summary.resolve()))
    remove_var = tk.StringVar(value="")
    date_format_var = tk.StringVar(value="")
    initial_month_labels = read_month_labels_from_path(default_monthly_summary)
    month_var = tk.StringVar(value=initial_month_labels[0])
    account_type_var = tk.StringVar(value="Credit")
    bank_name_var = tk.StringVar(value=DEFAULT_BANK_NAME)
    sheet_var = tk.StringVar(value="")
    sheet_names: List[str] = []

    filebar = tk.Frame(root, bg=BG)
    filebar.pack(fill="x", padx=16, pady=(2, 8))
    filebar.columnconfigure(1, weight=1)

    mk_label(filebar, "月份总表 Excel 保存位置(Monthly summary location)：").grid(row=0, column=0, sticky="w")
    ent_summary = tk.Entry(filebar, textvariable=summary_var, width=78, bg="#111", fg=FG, insertbackground=FG, relief="flat")
    ent_summary.grid(row=0, column=1, padx=6, sticky="we")

    def refresh_sheet_and_month_options(preferred_sheet: str = ""):
        path = get_summary_output_path(summary_var.get())
        names = list_workbook_sheet_names(path)
        sheet_names[:] = names
        sheet_box["values"] = sheet_names

        selected = preferred_sheet if preferred_sheet in sheet_names else ""
        if not selected and sheet_var.get() in sheet_names:
            selected = sheet_var.get()
        if not selected and sheet_names:
            selected = sheet_names[0]
        sheet_var.set(selected)

        detected = read_month_labels_from_selected_sheet(
            path, selected, account_type_var.get()
        ) if selected else list(DEFAULT_UI_MONTH_LABELS)
        months[:] = detected
        month_box["values"] = months
        month_var.set(months[0])

    def choose_summary():
        p = filedialog.askopenfilename(
            title="选择需要继续写入的 Excel（支持多个工作表）",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if p:
            summary_var.set(p)
            refresh_sheet_and_month_options()

    mk_button(filebar, "选择 Excel(Browse)", choose_summary).grid(row=0, column=2)

    mk_label(filebar, "预处理删除内容(Preprocess remove text)：").grid(row=1, column=0, sticky="w", pady=(8, 0))
    ent_remove = tk.Entry(
        filebar,
        textvariable=remove_var,
        width=78,
        bg="#111",
        fg=FG,
        insertbackground=FG,
        relief="flat"
    )
    ent_remove.grid(row=1, column=1, padx=6, pady=(8, 0), sticky="we")

    mk_label(
        filebar,
        "例如: target, Mcdonald || 不同的关键字请使用逗号隔开",
        fg="#9cdcfe"
    ).grid(row=2, column=1, sticky="w", padx=6, pady=(4, 0))

    mk_label(filebar, "日期示例或格式(Date example / format)：").grid(row=3, column=0, sticky="w", pady=(8, 0))
    ent_date_format = tk.Entry(
        filebar,
        textvariable=date_format_var,
        width=78,
        bg="#111",
        fg=FG,
        insertbackground=FG,
        relief="flat"
    )
    ent_date_format.grid(row=3, column=1, padx=6, pady=(8, 0), sticky="we")

    date_help = (
        "直接输入账单里的日期示例，程序会自动转换格式：\n"
        "4-22 → M-DD    May 17 → M DD    17 May → DD M    "
        "2025-4-22 → YYYY-M-DD\n"
    )
    mk_label(
        filebar,
        date_help,
        fg="#9cdcfe",
        justify="left",
        anchor="w"
    ).grid(row=4, column=1, columnspan=2, sticky="w", padx=6, pady=(4, 0))

    mk_label(filebar, "Bank Name（仅 Credit 表第一行使用）：").grid(row=5, column=0, sticky="w", pady=(8, 0))
    ent_bank = tk.Entry(
        filebar,
        textvariable=bank_name_var,
        width=78,
        bg="#111",
        fg=FG,
        insertbackground=FG,
        relief="flat"
    )
    ent_bank.grid(row=5, column=1, padx=6, pady=(8, 0), sticky="we")

    center = tk.Frame(root, bg=BG)
    center.pack(fill="both", expand=True, padx=16, pady=(6, 6))

    title_row = tk.Frame(center, bg=BG)
    title_row.pack(fill="x", pady=(0, 6))

    txt_frame = tk.Frame(center, bg=BG)
    txt_frame.pack(fill="both", expand=True)

    txt_input = tk.Text(
        txt_frame,
        wrap="word",
        bg="#111",
        fg=FG,
        insertbackground=FG,
        relief="flat",
        undo=True
    )

    status = tk.StringVar(value="尚未开始(Not started)")

    def clear_textbox():
        txt_input.delete("1.0", tk.END)
        txt_input.focus_set()
        status.set("文本框已清空 (Text box cleared)")

    def on_month_selected(event=None):
        m = month_var.get().strip()
        t = account_type_var.get().strip()
        status.set(f"已选择月份: {m} | 类型: {t}")

    def on_account_type_selected(event=None):
        refresh_sheet_and_month_options(sheet_var.get())
        m = month_var.get().strip()
        t = account_type_var.get().strip()
        status.set(f"已选择工作表: {sheet_var.get()} | 日期: {m} | 类型: {t}")

    def on_sheet_selected(event=None):
        refresh_sheet_and_month_options(sheet_var.get())
        status.set(
            f"已选择工作表: {sheet_var.get()} | 日期: {month_var.get()} | "
            f"类型: {account_type_var.get()}"
        )

    mk_label(
        title_row,
        "在下方文本框输入账单(Please paste your bank statement data into the text box below.)：",
        font=("Helvetica", 12)
    ).pack(side="left", anchor="w")

    months = list(initial_month_labels)

    mk_label(title_row, "工作表/Page:").pack(side="left", padx=(10, 2))
    sheet_box = ttk.Combobox(
        title_row,
        textvariable=sheet_var,
        values=sheet_names,
        width=18,
        state="readonly"
    )
    sheet_box.pack(side="left", padx=(2, 6))
    sheet_box.bind("<<ComboboxSelected>>", on_sheet_selected)

    mk_label(title_row, "日期/Month:").pack(side="left", padx=(2, 2))
    month_box = ttk.Combobox(
        title_row,
        textvariable=month_var,
        values=months,
        width=8,
        state="readonly"
    )
    month_box.pack(side="left", padx=(10, 6))
    month_box.set(months[0])
    month_var.set(months[0])
    month_box.bind("<<ComboboxSelected>>", on_month_selected)

    account_type_box = ttk.Combobox(
        title_row,
        textvariable=account_type_var,
        values=["Credit", "Debit"],
        width=10,
        state="readonly"
    )
    account_type_box.pack(side="left", padx=(6, 6))
    account_type_box.set("Credit")
    account_type_var.set("Credit")
    account_type_box.bind("<<ComboboxSelected>>", on_account_type_selected)

    refresh_sheet_and_month_options()

    mk_button(title_row, "清空文本框", clear_textbox).pack(side="right", padx=(10, 0))

    txt_input.pack(side="left", fill="both", expand=True)

    sb = tk.Scrollbar(txt_frame, command=txt_input.yview)
    sb.pack(side="right", fill="y")
    txt_input.configure(yscrollcommand=sb.set)

    mk_label(root, "状态(Status)：").pack(anchor="w", padx=16)
    mk_label(root, "", fg="#9cdcfe", textvariable=status).pack(anchor="w", padx=16)

    def keep_cursor_visible(event=None):
        txt_input.see("insert")
        return None

    txt_input.bind("<KeyRelease>", keep_cursor_visible)
    txt_input.bind("<ButtonRelease-1>", keep_cursor_visible)
    txt_input.bind("<MouseWheel>", keep_cursor_visible)
    txt_input.bind("<Return>", keep_cursor_visible)
    txt_input.bind("<<Paste>>", keep_cursor_visible)

    def import_existing_report(clear_existing_data: bool):
        mode_name = "清空旧金额（模板模式）" if clear_existing_data else "保留全部原有数据"
        source = filedialog.askopenfilename(
            title=f"选择已有公司报表 Excel - {mode_name}",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not source:
            return

        try:
            target = get_summary_output_path(summary_var.get())
            action_text = (
                "可识别页面中的旧月份金额会被清空，但格式、公式框架、Merchant、Category、"
                "工作表名称和顺序会保留。"
                if clear_existing_data
                else
                "源 Excel 中的所有原有数据、公式、格式、工作表名称和顺序都会完整保留。"
            )

            if target.exists():
                overwrite = messagebox.askyesno(
                    "目标 Excel 已存在",
                    f"目标文件已经存在：\n{target}\n\n"
                    f"导入模式：{mode_name}\n"
                    f"{action_text}\n\n"
                    "目标文件将被这次导入结果覆盖，是否继续？"
                )
                if not overwrite:
                    return
            else:
                proceed = messagebox.askyesno(
                    "确认导入模式",
                    f"导入模式：{mode_name}\n\n{action_text}\n\n是否继续？"
                )
                if not proceed:
                    return

            status.set(f"正在导入已有报表：{mode_name}...")
            income_count, expense_count, imported_month_labels, imported_sheet_names = import_existing_report_template(
                Path(source),
                target,
                bank_name=bank_name_var.get().strip() or DEFAULT_BANK_NAME,
                clear_existing_data=clear_existing_data,
            )
            summary_var.set(str(target))
            refresh_sheet_and_month_options(imported_sheet_names[0] if imported_sheet_names else "")
            status.set(
                f"导入完成：{target.name} | 模式: {mode_name} | "
                f"共 {len(imported_sheet_names)} 个工作表 | 当前: {sheet_var.get()}"
            )

            result_note = (
                "可识别页面中的历史月份金额已清空；其他内容已保留。"
                if clear_existing_data
                else
                "所有原有数据、公式和格式均已完整保留，可直接继续追加。"
            )
            messagebox.showinfo(
                "导入完成",
                f"目标 Excel：\n{target}\n\n"
                f"导入模式：{mode_name}\n"
                f"导入工作表：{len(imported_sheet_names)} 个\n"
                f"工作表名称：{', '.join(imported_sheet_names)}\n"
                f"Credit 项目：{income_count}\n"
                f"Debit Merchant：{expense_count}\n\n"
                f"{result_note}\n"
                f"UI 日期顺序：{', '.join(months)}"
            )
        except Exception as e:
            status.set(f"模板导入失败：{type(e).__name__}: {e}")
            messagebox.showerror("导入失败", f"{type(e).__name__}: {e}")



    def start():
        try:
            # 直接处理文本框中的内容，不再生成 statement.txt。
            content = txt_input.get("1.0", "end-1c")

            custom_date_input = date_format_var.get().strip()
            resolved_date_format = resolve_date_input_to_format(custom_date_input)
            configure_date_format(resolved_date_format)

            selected_month_display = month_var.get().strip()
            selected_month_index = month_box.current()
            if selected_month_index < 0 or selected_month_index >= len(months):
                messagebox.showerror("错误", "请选择有效日期/月份")
                return

            selected_sheet = sheet_var.get().strip()
            if not selected_sheet:
                messagebox.showerror("错误", "请选择要修改的工作表/Page")
                return

            selected_account_type = account_type_var.get().strip().lower()
            if selected_account_type not in ("credit", "debit"):
                messagebox.showerror("错误", "请选择 Credit 或 Debit")
                return

            bank_name = bank_name_var.get().strip() or DEFAULT_BANK_NAME
            remove_items = parse_remove_items(remove_var.get())

            status.set(f"正在处理... 日期: {selected_month_display} | 类型: {selected_account_type.title()}")

            # 全部在内存中完成预处理和解析，不再生成以下中间文件：
            # statement.txt、parsed_transactions.csv、parsed_transactions.txt、
            # parsed_transactions.xlsx。
            preprocessed_text = preprocess_statement_text(content, remove_items)

            hits = parse_auto(preprocessed_text)
            rows = [(h.merchant, h.amount, h.who) for h in hits]

            summary_xlsx = get_summary_output_path(summary_var.get())

            append_rows_to_selected_sheet(
                summary_xlsx,
                selected_sheet,
                rows,
                selected_month_index,
                selected_account_type,
                bank_name=bank_name,
            )

            status.set(
                f"已完成 / Completed | 工作表: {selected_sheet} | 日期: {selected_month_display} | "
                f"类型: {selected_account_type.title()} | 总表: {summary_xlsx.name}"
            )
            date_format_display = (
                resolved_date_format
                if resolved_date_format
                else "默认格式 MM/DD、MM/DD/YY、MM/DD/YYYY"
            )

            messagebox.showinfo(
                "Completed",
                f"提取成功: {len(rows)} 笔交易\n"
                f"当前工作表: {selected_sheet}\n"
                f"当前日期: {selected_month_display}\n"
                f"当前类型: {selected_account_type.title()}\n"
                "写入方式: 保留原有数据并继续追加\n"
                f"日期格式: {date_format_display}\n"
                f"总表文件: {summary_xlsx.name}\n"
                f"分类规则文件: {category_rules_path()}\n"
            )

        except Exception as e:
            messagebox.showerror("异常 / Error", f"{type(e).__name__}: {e}")
            status.set(f"解析失败 / Failed: {type(e).__name__}: {e}")

    def stop():
        root.destroy()

    btns = tk.Frame(root, bg=BG)
    btns.pack(pady=12)
    mk_button(
        btns,
        "导入并清空旧金额",
        lambda: import_existing_report(True)
    ).grid(row=0, column=0, padx=8)
    mk_button(
        btns,
        "导入并保留全部数据",
        lambda: import_existing_report(False)
    ).grid(row=0, column=1, padx=8)
    mk_button(btns, "Start", start).grid(row=0, column=2, padx=8)
    mk_button(btns, "Stop", stop).grid(row=0, column=3, padx=8)

    root.mainloop()

if __name__ == "__main__":
    try:
        run_parser_ui()
    except Exception as e:
        print(f"程序异常：{e}", file=sys.stderr)
        raise