import re
import csv
import sys
from collections import OrderedDict
from dataclasses import dataclass
from pathlib import Path
from typing import List, Tuple

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

    existing = read_existing_debit_summary_from_wb(wb)
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

    headers = ["Merchant"] + DEBIT_MONTH_HEADERS + ["Total", "Category"]
    for col_idx, h in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=h)

    merchants_sorted = sorted(existing.keys(), key=lambda x: x.strip().lower())

    row_idx = 2
    for merchant in merchants_sorted:
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

    for month, row_idx in CREDIT_MONTH_ROWS.items():
        ws.cell(row=row_idx, column=1, value=month)

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

# ================= 输出路径 =================

def get_summary_output_path(base_path_str: str) -> Path:
    raw = (base_path_str or "").strip()

    if raw:
        p = Path(raw)
        if p.suffix.lower() == ".xlsx":
            folder = p.parent
        else:
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
    root.title("BSDP")
    root.geometry("1120x950")
    root.configure(bg=BG)

    mk_label(root, "BSDP", font=("Helvetica", 16, "bold")).pack(anchor="w", padx=16, pady=(12, 6))

    default_monthly_summary = script_dir() / "credit_monthly_summary.xlsx"

    summary_var = tk.StringVar(value=str(default_monthly_summary.resolve()))
    remove_var = tk.StringVar(value="")
    date_format_var = tk.StringVar(value="")
    month_var = tk.StringVar(value="JAN")
    account_type_var = tk.StringVar(value="Credit")
    bank_name_var = tk.StringVar(value=DEFAULT_BANK_NAME)

    filebar = tk.Frame(root, bg=BG)
    filebar.pack(fill="x", padx=16, pady=(2, 8))
    filebar.columnconfigure(1, weight=1)

    mk_label(filebar, "月份总表 Excel 保存位置(Monthly summary location)：").grid(row=0, column=0, sticky="w")
    ent_summary = tk.Entry(filebar, textvariable=summary_var, width=78, bg="#111", fg=FG, insertbackground=FG, relief="flat")
    ent_summary.grid(row=0, column=1, padx=6, sticky="we")

    def choose_summary():
        p = filedialog.asksaveasfilename(
            title="选择总表保存位置（文件名会自动生成为 credit_monthly_summary.xlsx）",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if p:
            summary_var.set(p)

    mk_button(filebar, "更改(Browse)", choose_summary).grid(row=0, column=2)

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
        m = month_var.get().strip()
        t = account_type_var.get().strip()
        status.set(f"已选择月份: {m} | 类型: {t}")

    mk_label(
        title_row,
        "在下方文本框输入账单(Please paste your bank statement data into the text box below.)：",
        font=("Helvetica", 12)
    ).pack(side="left", anchor="w")

    months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]

    month_box = ttk.Combobox(
        title_row,
        textvariable=month_var,
        values=months,
        width=8,
        state="readonly"
    )
    month_box.pack(side="left", padx=(10, 6))
    month_box.set("JAN")
    month_var.set("JAN")
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

    def start():
        try:
            # 直接处理文本框中的内容，不再生成 statement.txt。
            content = txt_input.get("1.0", "end-1c")

            custom_date_input = date_format_var.get().strip()
            resolved_date_format = resolve_date_input_to_format(custom_date_input)
            configure_date_format(resolved_date_format)

            selected_month = month_var.get().strip().upper()
            if selected_month not in months:
                messagebox.showerror("错误", "请选择有效月份")
                return

            selected_account_type = account_type_var.get().strip().lower()
            if selected_account_type not in ("credit", "debit"):
                messagebox.showerror("错误", "请选择 Credit 或 Debit")
                return

            bank_name = bank_name_var.get().strip() or DEFAULT_BANK_NAME
            remove_items = parse_remove_items(remove_var.get())

            status.set(f"正在处理... 月份: {selected_month} | 类型: {selected_account_type.title()}")

            # 全部在内存中完成预处理和解析，不再生成以下中间文件：
            # statement.txt、parsed_transactions.csv、parsed_transactions.txt、
            # parsed_transactions.xlsx。
            preprocessed_text = preprocess_statement_text(content, remove_items)

            hits = parse_auto(preprocessed_text)
            rows = [(h.merchant, h.amount, h.who) for h in hits]

            summary_xlsx = get_summary_output_path(summary_var.get())

            if selected_account_type == "credit":
                write_or_update_credit_summary_xlsx(
                    summary_xlsx,
                    rows,
                    selected_month,
                    bank_name=bank_name
                )
            else:
                write_or_update_debit_summary_sheet(
                    summary_xlsx,
                    rows,
                    selected_month
                )

            status.set(
                f"已完成 / Completed | 月份: {selected_month} | 类型: {selected_account_type.title()} | 总表: {summary_xlsx.name}"
            )
            date_format_display = (
                resolved_date_format
                if resolved_date_format
                else "默认格式 MM/DD、MM/DD/YY、MM/DD/YYYY"
            )

            messagebox.showinfo(
                "Completed",
                f"提取成功: {len(rows)} 笔交易\n"
                f"当前月份: {selected_month}\n"
                f"当前类型: {selected_account_type.title()}\n"
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
    mk_button(btns, "Start", start).grid(row=0, column=0, padx=8)
    mk_button(btns, "Stop", stop).grid(row=0, column=1, padx=8)

    root.mainloop()

if __name__ == "__main__":
    try:
        run_parser_ui()
    except Exception as e:
        print(f"程序异常：{e}", file=sys.stderr)
        raise