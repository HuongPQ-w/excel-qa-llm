#!/usr/bin/env python3
"""
excelqa.py - Hỏi đáp Excel đa ngôn ngữ (query + sheet EN/VI/ZH/JA/KO/FR)

Luồng:
  1) Đọc xlsx → Sheet (merge, header path, multi-table, FormulaEngine cho ô công thức chưa có giá trị tính sẵn)
  2) Dump (+ routing sheet khi workbook lớn)
  3) Bước REASONING (plan): 1 lượt LLM cho kế hoạch phân tích + format spec JSON
       {"rounding_mode", "decimals", "date_format", "extra"}
  4) LLM sinh code (có kế hoạch làm gợi ý) → sandbox chạy; lỗi thì repair 1 lần
  5) Vote + judge khi lệch
  6) Formatter xác định áp dụng format spec JSON (làm tròn, ngày); regex trên câu hỏi chỉ là DỰ PHÒNG
     khi JSON thiếu/sai

Chạy: --xlsx-dir / --queries / --output  hoặc --input / --output
Khoá API: biến môi trường OPENROUTER_API_KEY hoặc --openrouter-key (không hardcode trong mã).
Env: OPENROUTER_MODEL, WORKERS, VOTES, JUDGE, PLAN (=0 tắt reasoning), PLAN_THINK, WB_CACHE_MAX, ...
"""
import argparse, csv, datetime as dt, json, math, os, re, statistics, sys, threading, time, traceback, unicodedata
from concurrent.futures import ThreadPoolExecutor
from collections import Counter
from decimal import (Decimal, ROUND_HALF_UP, ROUND_HALF_EVEN, ROUND_UP, ROUND_DOWN, ROUND_CEILING,
                     ROUND_FLOOR, InvalidOperation)
from functools import lru_cache
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string, range_boundaries
from openpyxl.utils.cell import coordinate_from_string
import requests
START = time.time()
API_URL = os.environ.get("OPENROUTER_API_URL", "https://openrouter.ai/api/v1/chat/completions")
MODEL = os.environ.get("OPENROUTER_MODEL", "qwen/qwen3.7-flash")
WORKERS = int(os.environ.get("WORKERS", "6"))
TIME_BUDGET = float(os.environ.get("TIME_BUDGET_SEC", "2100"))
PRICE_IN = float(os.environ.get("PRICE_IN_PER_M", "0.03"))
PRICE_OUT = float(os.environ.get("PRICE_OUT_PER_M", "0.13"))
MAX_COST_PER_Q = float(os.environ.get("MAX_COST_PER_Q", "0.08"))
MAX_TOKENS = int(os.environ.get("LLM_MAX_TOKENS", "3000"))
REASONING_OFF = os.environ.get("REASONING", "off").lower() == "off"
DEBUG_LOG = os.environ.get("DEBUG_LOG", "debug_log.jsonl")
VOTES = max(1, int(os.environ.get("VOTES", "3")))
VOTE_TEMP = float(os.environ.get("VOTE_TEMP", "0.3"))
JUDGE_ON = os.environ.get("JUDGE", "1") != "0"
PARALLEL_VOTES = os.environ.get("PARALLEL_VOTES", "1") != "0"
MULTI_TABLE = os.environ.get("MULTI_TABLE", "1") != "0"
ROUTE_MIN_CHARS = int(os.environ.get("ROUTE_MIN_CHARS", "12000"))
ROUTING_ON = os.environ.get("ROUTING", "1") != "0"
# --- ngưỡng/tham số (mặc định = giá trị cũ; ghi đè bằng biến môi trường để A/B) ---
ROUTE_SHRINK_RATIO = float(os.environ.get("ROUTE_SHRINK_RATIO", "0.8"))  # chỉ dùng routing khi dump giảm dưới tỉ lệ này
JUDGE_TIME_MARGIN = float(os.environ.get("JUDGE_TIME_MARGIN", "15"))     # giây còn lại tối thiểu để gọi judge
JUDGE_COST_RATIO = float(os.environ.get("JUDGE_COST_RATIO", "0.85"))     # chỉ gọi judge khi chi phí < tỉ lệ này của MAX_COST_PER_Q
SANDBOX_TIMEOUT = float(os.environ.get("SANDBOX_TIMEOUT", "15"))
MAX_HEADER_ROWS = int(os.environ.get("MAX_HEADER_ROWS", "12"))           # số dòng tối đa được coi là tiêu đề
WIDE_MERGE_MIN_COLS = int(os.environ.get("WIDE_MERGE_MIN_COLS", "4"))    # ô gộp rộng (tiêu đề) tối thiểu số cột
WIDE_MERGE_RATIO = float(os.environ.get("WIDE_MERGE_RATIO", "0.8"))      # ... hoặc tỉ lệ bề rộng bảng
LLM_RETRIES = int(os.environ.get("LLM_RETRIES", "4"))
HTTP_TIMEOUT = float(os.environ.get("HTTP_TIMEOUT", "90"))
VOTE_ATTEMPT_FACTOR = int(os.environ.get("VOTE_ATTEMPT_FACTOR", "2"))    # tối đa VOTES*factor lượt sinh phiếu
# --- cờ hành vi (đặt =0 để tắt) ---
REJECT_ERROR_ANSWERS = os.environ.get("REJECT_ERROR_ANSWERS", "1") != "0"  # chuỗi "Error..." không được tính là đáp án
FULL_DUMP_ON_RETRY = os.environ.get("FULL_DUMP_ON_RETRY", "1") != "0"      # repair/direct dùng dump đầy đủ, bỏ routing
LOG_CANDS = os.environ.get("LOG_CANDS", "1") != "0"                        # ghi mọi phiếu (answer+code) vào debug_log
DUMP_HEADER_CELLS = os.environ.get("DUMP_HEADER_CELLS", "1") != "0"        # dump thêm các dòng tiêu đề dạng ô thô (giữ địa chỉ ô)
CELL_FALLBACK = os.environ.get("CELL_FALLBACK", "1") != "0"                # tra cứu ô theo địa chỉ + đọc công thức thô + xử lý ô trống
JUDGE_OVERRIDE = os.environ.get("JUDGE_OVERRIDE", "0") == "1"
# --- bước reasoning (plan) trước khi sinh code ---
PLAN_ON = os.environ.get("PLAN", "1") != "0"                       # =0 để tắt bước reasoning
PLAN_THINK = os.environ.get("PLAN_THINK", "0") == "1"              # =1 bật thinking của nhà cung cấp cho riêng bước plan
PLAN_MAX_TOKENS = int(os.environ.get("PLAN_MAX_TOKENS", "1500"))
PLAN_MAX_CHARS = int(os.environ.get("PLAN_MAX_CHARS", "1800"))     # cắt phần phân tích đưa vào prompt sinh code
WB_CACHE_MAX = max(1, int(os.environ.get("WB_CACHE_MAX", "4")))    # số workbook giữ trong RAM (giới hạn 1 GB)

_log_lock = threading.Lock()

def log(*a):
    with _log_lock:
        print(*a, file=sys.stderr, flush=True)

# =============================================================================
# 1. TIỆN ÍCH: số, chuỗi, ngày + làm sạch giá trị (từ mau-btc)
# =============================================================================
# --- markers & number parse (đa dạng dash / fullwidth / kế toán) ---
_DASH_CHARS = "-\u2010\u2011\u2012\u2013\u2014\u2015\u2212\ufe58\ufe63\uff0d"
MARKERS = {
    "", "undefine", "undefined", "#", "n/a", "na", "n.a.", "n.a", "null", "none", "nan",
    "-", "--", "---", "?", "ø", "..", "...", "…", "‥", "x", "X", "*", "†", "‡",
    "(d)", "(D)", "s", "S", "u", "U", "nd", "ND", "np", "NP",
}
for _d in _DASH_CHARS:
    MARKERS.add(_d)
    MARKERS.add(_d * 2)
    MARKERS.add(_d * 3)
ERRORS = {"#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!", "#N/A!", "N/A"}
KEEP_MARKERS = {"nd", "np", "n/a", "na", "n.a", "n.a."} | {e.lower() for e in ERRORS}
_NUM_RE = re.compile(
    r"^[\s\$\u20ac\u00a3\u00a5+\-\u2212\u2013\u2014]*"
    r"\(?[\s\$\u20ac\u00a3\u00a5+\-\u2212\u2013\u2014]*"
    r"[\d\uff10-\uff19]"
    r"[\d\uff10-\uff19,\.\s\u00a0\u202f\u2009]*"
    r"\)?\s*%?\s*[\-\u2212]?$"
)

def _to_ascii_digits(s):
    out = []
    for ch in s:
        o = ord(ch)
        if 0xFF10 <= o <= 0xFF19:
            out.append(chr(o - 0xFF10 + ord("0")))
        elif ch in "\u2212\u2013\u2014\u2010\u2011":
            out.append("-")
        else:
            out.append(ch)
    return "".join(out)

def num(x):
    """Ô → float | None. Marker thiếu → None (không = 0)."""
    if x is None or isinstance(x, bool):
        return None
    if isinstance(x, (int, float)):
        return None if (isinstance(x, float) and (math.isnan(x) or math.isinf(x))) else float(x)
    if not isinstance(x, str):
        return None
    s = _to_ascii_digits(x.strip())
    s = s.replace("\u00a0", " ").replace("\u202f", " ").replace("\u2009", " ")
    if not s or s.lower() in MARKERS or s.upper() in ERRORS:
        return None
    if re.fullmatch(r"[\s\-–—−…\.]+", s):
        return None
    if not _NUM_RE.match(s):
        return None
    neg = (
        (s.startswith("(") and s.rstrip("%").rstrip().endswith(")"))
        or s.lstrip("$€£¥ ").startswith("-")
        or s.rstrip("%").rstrip().endswith("-")
    )
    t = re.sub(r"[^\d,\.]", "", s)
    if not t:
        return None
    if "," in t and "." in t:
        if t.rfind(",") > t.rfind("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            t = t.replace(",", "")
    elif "," in t:
        parts = t.split(",")
        if len(parts) == 2 and len(parts[1]) != 3:
            t = t.replace(",", ".")
        else:
            t = t.replace(",", "")
    elif t.count(".") > 1:
        t = t.replace(".", "")
    try:
        v = float(t)
    except ValueError:
        return None
    return -v if neg else v

def norm(s):
    """Chuẩn hoá để so khớp nhãn: bỏ dấu Latin/tiếng Việt, bỏ chú thích (1), bỏ ký tự lạ, hạ chữ thường.
    Giữ nguyên chữ Hán/Nhật/Hàn và các bảng chữ cái khác (trước đây bị xoá sạch)."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))  # gộp dấu tách rời + ký tự full-width (ＡＢＣ１２ -> ABC12)
    s = re.sub(r"\(\s*(\d+|x)\s*\)", " ", s, flags=re.I)
    out = []
    for ch in s:
        d = unicodedata.normalize("NFD", ch)
        if len(d) > 1 and unicodedata.name(d[0], "").startswith("LATIN"):
            ch = "".join(c for c in d if unicodedata.category(c) != "Mn")  # chỉ bỏ dấu của chữ Latin
        out.append(ch)
    s = "".join(out).replace("\u0111", "d").replace("\u0110", "D")
    s = re.sub(r"[\W_]+", " ", s.lower())
    return s.strip()


_DATE_FORMATS = (
    "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y/%m/%d", "%Y.%m.%d",
    "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%Y %H:%M:%S", "%d/%m/%y", "%d-%m-%y",
    "%m/%d/%Y", "%m-%d-%Y",  # chỉ khớp khi dd/mm/yyyy không hợp lệ (tháng > 12)
    "%Y年%m月%d日", "%B %d, %Y", "%b %d, %Y", "%d %B %Y", "%d %b %Y", "%d-%b-%Y", "%d-%b-%y",
)

def to_date(x):
    """datetime hoặc None. Chuỗi dd/mm/yyyy được hiểu theo kiểu ngày-trước."""
    if x is None:
        return None
    if isinstance(x, dt.datetime):
        return x
    if isinstance(x, dt.date):
        return dt.datetime(x.year, x.month, x.day)
    if not isinstance(x, str):
        return None
    s = re.sub(r"\s+", " ", x.strip())
    for f in _DATE_FORMATS:
        try:
            return dt.datetime.strptime(s, f)
        except ValueError:
            pass
    return None

# =============================================================================
# 2. ĐỌC EXCEL → Sheet (mini-eval công thức từ mau-btc)
# =============================================================================
_KEEP_HYPHEN = {"self", "non", "full", "part", "semi", "anti", "ex", "pre", "post", "multi", "co", "re", "two",
                "three", "one", "year", "short", "long", "low", "high", "middle", "e", "x"}

def clean_cell(v):
    if isinstance(v, str):
        s = v.replace("\r", "").replace("\u00a0", " ")
        s = re.sub(r"([A-Za-z]+)-[ \t]*\n[ \t]*(?=[a-z])",
                   lambda m: m.group(1) + ("-" if m.group(1).lower() in _KEEP_HYPHEN else ""), s)
        s = re.sub(r"\s*\n\s*", " ", s)
        s = re.sub(r"[ \t]{2,}", " ", s).strip()
        if s.lower() in KEEP_MARKERS:
            return s
        if s.lower() in MARKERS or s.upper() in ERRORS:
            return None
        return s if s != "" else None
    return v

def fmt_val(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d") if (v.hour, v.minute, v.second) == (0, 0, 0) else v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, dt.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, dt.time):
        return v.strftime("%H:%M:%S")
    if isinstance(v, float):
        return f"{v:.12g}"
    return str(v)

def _is_year(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and float(v).is_integer() and 1900 <= v <= 2100

def _data_like(row):
    n = nnum = ntxt = 0
    for v in row:
        if v is None:
            continue
        if isinstance(v, (dt.datetime, dt.date)):
            n += 1
        elif isinstance(v, (int, float)) and not isinstance(v, bool):
            if not _is_year(v):
                n += 1
                nnum += 1
        elif isinstance(v, str):
            if num(v) is not None and not _is_year(num(v)):
                n += 1
                nnum += 1
            elif num(v) is None:
                ntxt += 1
    if nnum == 0 and ntxt >= 5:  # nhiều nhãn chữ, không có số nào → dòng tiêu đề (kể cả khi các cột được đặt tên bằng ngày)
        return False
    return n >= 2

# =============================================================================
# 2b. BỘ TÍNH CÔNG THỨC (xuyên sheet + cache liên kết ngoài) cho ô công thức không có giá trị tính sẵn
# =============================================================================
import operator
from collections import namedtuple
from openpyxl.formula.tokenizer import Tokenizer, Token

class XlErr(Exception):
    """Lỗi Excel thật (#N/A, #DIV/0!...): IFERROR bắt được."""
    def __init__(self, code="#VALUE!"):
        super().__init__(code)
        self.code = code

class XlUnsupported(Exception):
    """Giới hạn của bộ tính (hàm lạ, liên kết ngoài không cache...): IFERROR KHÔNG được nuốt."""
XFAIL = object()
XRange = namedtuple("XRange", "sheet r1 c1 r2 c2")
_ERR_CODES = {"#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!"}
_PREC = {"=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5}
_CMP = {"=": operator.eq, "<>": operator.ne, "<": operator.lt, ">": operator.gt, "<=": operator.le, ">=": operator.ge}
_REF_RE = re.compile(r"^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?"
                     r"|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)$")
_EXT_RE = re.compile(r"^\[(\d+)\](.+)$")
_STR_FN = {"LEN": lambda s: float(len(s)), "TRIM": lambda s: re.sub(r" +", " ", s).strip(), "UPPER": lambda s: s.upper(),
           "LOWER": lambda s: s.lower(), "LEFT": lambda s, n=1: s[:int(n)],
           "RIGHT": lambda s, n=1: s[-int(n):] if int(n) > 0 else "",
           "MID": lambda s, a, n: s[max(int(a) - 1, 0): max(int(a) - 1, 0) + int(n)]}

def _parse(formula):
    t = [x for x in Tokenizer(formula).items if x.type != Token.WSPACE]
    i = [0]
    def peek():
        return t[i[0]] if i[0] < len(t) else None
    def take():
        x = peek()
        i[0] += 1
        return x
    def expr(minp=1):
        left = unary()
        while True:
            p = peek()
            if p is None or p.type != Token.OP_IN or _PREC.get(p.value, 0) < minp:
                return left
            take()
            left = ("bin", p.value, left, expr(_PREC[p.value] + 1))
    def unary():
        p = peek()
        if p is not None and p.type == Token.OP_PRE:
            take()
            return ("un", p.value, unary())
        n = primary()
        while peek() is not None and peek().type == Token.OP_POST:
            take()
            n = ("pct", n)
        return n
    def primary():
        p = take()
        if p is None:
            raise XlUnsupported("unexpected end")
        if p.type == Token.OPERAND:
            if p.subtype == Token.NUMBER:
                return ("val", float(p.value))
            if p.subtype == Token.TEXT:
                return ("val", p.value[1:-1].replace('""', '"'))
            if p.subtype == Token.LOGICAL:
                return ("val", p.value.upper() == "TRUE")
            return ("err", p.value) if p.subtype == Token.ERROR else ("ref", p.value)
        if p.type == Token.FUNC and p.subtype == Token.OPEN:
            args = []
            if peek() is not None and peek().type == Token.FUNC and peek().subtype == Token.CLOSE:
                take()
                return ("fn", p.value[:-1].upper().replace("_XLFN.", ""), args)
            while True:
                q = peek()
                args.append(("val", None) if q is not None and (q.type == Token.SEP or q.subtype == Token.CLOSE) else expr())
                q = take()
                if q is None:
                    raise XlUnsupported("unclosed function")
                if q.type == Token.FUNC and q.subtype == Token.CLOSE:
                    return ("fn", p.value[:-1].upper().replace("_XLFN.", ""), args)
                if q.type != Token.SEP:
                    raise XlUnsupported("bad arguments")
        if p.type == Token.PAREN and p.subtype == Token.OPEN:
            n = expr()
            q = take()
            if q is None or q.type != Token.PAREN:
                raise XlUnsupported("unclosed paren")
            return n
        raise XlUnsupported(f"token {p.value!r}")
    n = expr()
    if peek() is not None:
        raise XlUnsupported("trailing tokens")
    return n

def _serial(d):
    d = d if isinstance(d, dt.datetime) else dt.datetime(d.year, d.month, d.day)
    return (d - dt.datetime(1899, 12, 30)).total_seconds() / 86400.0

def _xl_str(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return str(int(v)) if v == int(v) and abs(v) < 1e15 else f"{v:.15g}"
    return fmt_val(v) if isinstance(v, (dt.datetime, dt.date)) else str(v)

def _cmp(op, a, b):
    def norm_(v, other):
        if v is None:
            return "" if isinstance(other, str) else (False if isinstance(other, bool) else 0.0)
        return _serial(v) if isinstance(v, (dt.datetime, dt.date)) else v
    a, b = norm_(a, b), norm_(b, a)
    key = lambda v: (1, v.lower()) if isinstance(v, str) else ((2, v) if isinstance(v, bool) else (0, v))  # noqa: E731
    return _CMP[op](key(a), key(b))

def _crit(c):
    """Tiêu chí SUMIF/COUNTIF: số hoặc chuỗi (hỗ trợ * ?), kèm toán tử >,<,>=,<=,<>,=."""
    op, rest = "=", c
    m = re.match(r"^(<=|>=|<>|<|>|=)(.*)$", c) if isinstance(c, str) else None
    if m:
        op, rest = m.groups()
    try:
        rn = None if isinstance(rest, bool) else float(rest)
    except (TypeError, ValueError):
        rn = None
    pat = str(rest).lower()
    rx = re.compile("^" + re.escape(pat).replace(r"\*", ".*").replace(r"\?", ".") + "$", re.S)
    def pred(v):
        if rn is not None:
            try:
                return _CMP[op](float(v), rn)
            except (TypeError, ValueError):
                return op == "<>"
        ok = bool(rx.match(_xl_str(v).lower()))
        return (not ok) if op == "<>" else (ok if op == "=" else False)
    return pred

class FormulaEngine:
    def __init__(self, wb_v, wb_f):
        self.wb_v, self.wb_f = wb_v, wb_f
        self.names = {n.lower(): n for n in wb_v.sheetnames}
        self.vals, self.fmls, self.memo, self.busy, self.ast = {}, {}, {}, set(), {}
        self.ext = self._load_ext(wb_f)
    @staticmethod
    def _load_ext(wb):
        """Cache liên kết ngoài: [ {tên sheet thường: ({(r,c): giá trị}, maxr, maxc)} ... ] theo thứ tự [1], [2]..."""
        out = []
        for link in (getattr(wb, "_external_links", None) or []):
            sheets = {}
            try:
                book = link.externalBook
                names = list(book.sheetNames.sheetName)
                for sd in (book.sheetDataSet.sheetData if book.sheetDataSet is not None else []):
                    cells = {}
                    for row in sd.row:
                        for c in row.cell:
                            if c.v is None:
                                continue
                            try:
                                v = c.v if c.t == "str" else (c.v in ("1", "true", "TRUE") if c.t == "b" else
                                                              XlErr(c.v) if c.t == "e" else float(c.v))
                                col, r = coordinate_from_string(c.r)
                                cells[(r, column_index_from_string(col))] = v
                            except Exception:
                                continue
                    mr = max((k[0] for k in cells), default=0)
                    mc = max((k[1] for k in cells), default=0)
                    sheets[names[int(sd.sheetId)].lower()] = (cells, mr, mc)
            except Exception:
                pass
            out.append(sheets)
        return out
    def _g(self, store, wb, sheet):
        if sheet not in store:
            try:
                store[sheet] = [list(r) for r in wb[sheet].iter_rows(values_only=True)]
            except Exception:
                store[sheet] = []
        return store[sheet]
    def _ext_sheet(self, key):
        try:
            return self.ext[key[1] - 1][key[2]]
        except (IndexError, KeyError):
            raise XlUnsupported("external sheet not cached")
    def _dims(self, sheet):
        if isinstance(sheet, tuple):
            return self._ext_sheet(sheet)[1:]
        g = self._g(self.vals, self.wb_v, sheet)
        return len(g), max((len(r) for r in g), default=0)
    def _cell(self, sheet, r, c):
        """Giá trị ô: giá trị tính sẵn nếu có, không thì tính từ công thức (đệ quy + memo)."""
        if isinstance(sheet, tuple):
            return self._ext_sheet(sheet)[0].get((r, c))
        key = (sheet, r, c)
        if key in self.memo:
            if self.memo[key] is XFAIL:
                raise XlUnsupported("unresolved")
            return self.memo[key]
        at = lambda g: g[r - 1][c - 1] if 0 < r <= len(g) and 0 < c <= len(g[r - 1]) else None  # noqa: E731
        v = at(self._g(self.vals, self.wb_v, sheet))
        if v is None:
            f = at(self._g(self.fmls, self.wb_f, sheet))
            f = getattr(f, "text", f)
            if isinstance(f, str) and len(f) > 1 and f.startswith("="):
                if key in self.busy:
                    raise XlUnsupported("circular")
                self.busy.add(key)
                try:
                    v = self._sc(self._ev(self.ast.setdefault(f, _parse(f)), key), key)
                except XlErr as e:
                    v = XlErr(e.code)
                except XlUnsupported:
                    self.memo[key] = XFAIL
                    raise
                finally:
                    self.busy.discard(key)
        elif isinstance(v, str) and v.strip().upper() in _ERR_CODES:
            v = XlErr(v.strip().upper())
        self.memo[key] = v
        return v
    def value(self, sheet, r, c):
        """Giá trị ô đã tính; XFAIL nếu chưa hỗ trợ; None nếu rỗng/lỗi Excel thật."""
        try:
            v = self._cell(sheet, r, c)
        except (XlUnsupported, RecursionError, XlErr):
            return XFAIL
        if isinstance(v, XlErr):
            return v.code  # lỗi Excel thật (#DIV/0!...) là giá trị hiển thị của ô
        return None if v == "" else (XFAIL if isinstance(v, XRange) else v)
    def _ref(self, text, ctx):
        m = _REF_RE.match(text)
        if not m:
            raise XlUnsupported(f"name/ref {text!r}")
        name = m.group(1) if m.group(1) is not None else m.group(2)
        if name is None:
            sheet = ctx[0]
        else:
            name = name.replace("''", "'")
            e = _EXT_RE.match(name)
            sheet = ("ext", int(e.group(1)), e.group(2).lower()) if e else self.names.get(name.lower())
            if sheet is None:
                raise XlUnsupported("3D reference") if ":" in name else XlErr("#REF!")
        a = m.group(3).replace("$", "").upper()
        nr, nc = self._dims(sheet)
        if re.fullmatch(r"[A-Z]{1,3}:[A-Z]{1,3}", a):
            x, y = a.split(":")
            return XRange(sheet, 1, column_index_from_string(x), nr or 1, column_index_from_string(y))
        if re.fullmatch(r"\d+:\d+", a):
            x, y = a.split(":")
            return XRange(sheet, int(x), 1, int(y), nc or 1)
        c1, r1, c2, r2 = range_boundaries(a)
        return XRange(sheet, r1, c1, r2, c2)
    def _sc(self, v, ctx):
        """Về giá trị đơn: giao điểm ngầm định của vùng; lỗi thì ném."""
        if isinstance(v, XRange):
            if v.r1 == v.r2 and v.c1 <= (ctx[2] if v.c1 != v.c2 else v.c1) <= v.c2:
                v = self._cell(v.sheet, v.r1, ctx[2] if v.c1 != v.c2 else v.c1)
            elif v.c1 == v.c2 and v.r1 <= ctx[1] <= v.r2:
                v = self._cell(v.sheet, ctx[1], v.c1)
            else:
                raise XlErr("#VALUE!")
        if isinstance(v, XlErr):
            raise XlErr(v.code)
        return v
    def _cells(self, rg):
        mr, mc = self._dims(rg.sheet)
        for r in range(rg.r1, min(rg.r2, mr) + 1):
            for c in range(rg.c1, min(rg.c2, mc) + 1):
                yield r, c, self._cell(rg.sheet, r, c)
    def _rng(self, node, ctx):
        v = self._ev(node, ctx)
        if not isinstance(v, XRange):
            raise XlErr("#VALUE!")
        return v
    @staticmethod
    def _n(v):
        if v is None or v == "":
            return 0.0
        if isinstance(v, (bool, int, float)):
            return float(v)
        if isinstance(v, (dt.datetime, dt.date)):
            return _serial(v)
        try:
            return float(str(v).strip().replace(",", ""))
        except ValueError:
            raise XlErr("#VALUE!")
    def _flat(self, v, ctx):
        return [x for _, _, x in self._cells(v)] if isinstance(v, XRange) else [self._sc(v, ctx)]
    def _nums(self, args, ctx):
        out = []
        for a in args:
            v = self._ev(a, ctx)
            if isinstance(v, XRange):
                for x in self._flat(v, ctx):
                    if isinstance(x, XlErr):
                        raise XlErr(x.code)
                    if isinstance(x, (int, float)) and not isinstance(x, bool):
                        out.append(float(x))
            elif self._sc(v, ctx) is not None:
                out.append(self._n(self._sc(v, ctx)))
        return out
    def _ev(self, n, ctx):
        k = n[0]
        if k == "val":
            return n[1]
        if k == "err":
            raise XlErr(n[1])
        if k == "ref":
            return self._ref(n[1], ctx)
        if k == "fn":
            return self._fn(n[1], n[2], ctx)
        if k == "un":
            v = self._sc(self._ev(n[2], ctx), ctx)
            return v if n[1] == "+" else -self._n(v)  # '+' đơn ngôi giữ nguyên giá trị (cả chữ/ngày)
        if k == "pct":
            return self._n(self._sc(self._ev(n[1], ctx), ctx)) / 100.0
        if k == "bin":
            op, a, b = n[1], self._sc(self._ev(n[2], ctx), ctx), self._sc(self._ev(n[3], ctx), ctx)
            if op == "&":
                return _xl_str(a) + _xl_str(b)
            if op in _CMP:
                return _cmp(op, a, b)
            x, y = self._n(a), self._n(b)
            if op == "/" and y == 0:
                raise XlErr("#DIV/0!")
            try:
                r = {"+": x + y, "-": x - y, "*": x * y}[op] if op in "+-*" else (x / y if op == "/" else x ** y)
            except (ZeroDivisionError, OverflowError):
                raise XlErr("#NUM!")
            if isinstance(r, complex):
                raise XlErr("#NUM!")
            return r
        raise XlUnsupported(k)
    def _fn(self, name, args, ctx):
        ev = lambda i: self._sc(self._ev(args[i], ctx), ctx)  # noqa: E731
        if name in ("SUM", "AVERAGE", "MIN", "MAX", "COUNT"):
            v = self._nums(args, ctx)
            if name == "AVERAGE" and not v:
                raise XlErr("#DIV/0!")
            return {"SUM": sum, "COUNT": lambda x: float(len(x)), "AVERAGE": lambda x: sum(x) / len(x),
                    "MIN": lambda x: min(x, default=0.0), "MAX": lambda x: max(x, default=0.0)}[name](v)
        if name == "COUNTA":
            return float(sum(1 for a in args for x in self._flat(self._ev(a, ctx), ctx) if x is not None and x != ""))
        if name == "IF":
            c = ev(0)
            if isinstance(c, str):
                raise XlErr("#VALUE!")
            i = 1 if self._n(c) else 2
            return self._ev(args[i], ctx) if len(args) > i else bool(i == 1)
        if name in ("IFERROR", "IFNA"):
            try:
                return ev(0)
            except XlErr as e:
                if name == "IFNA" and e.code != "#N/A":
                    raise
                return self._ev(args[1], ctx)
        if name in ("AND", "OR"):
            vals = []
            for a in args:
                for x in self._flat(self._ev(a, ctx), ctx):
                    if isinstance(x, XlErr):
                        raise XlErr(x.code)
                    if x is not None and not isinstance(x, str):
                        vals.append(bool(self._n(x)))
            if not vals:
                raise XlErr("#VALUE!")
            return all(vals) if name == "AND" else any(vals)
        if name == "NOT":
            return not self._n(ev(0))
        if name in ("ISNUMBER", "ISTEXT", "ISBLANK", "ISERROR", "ISNA"):
            try:
                v = ev(0)
            except XlErr as e:
                return name == "ISERROR" or (name == "ISNA" and e.code == "#N/A")
            return {"ISNUMBER": isinstance(v, (int, float, dt.datetime, dt.date)) and not isinstance(v, bool),
                    "ISTEXT": isinstance(v, str) and v != "", "ISBLANK": v is None}.get(name, False)
        if name in ("ROUND", "ROUNDUP", "ROUNDDOWN"):
            d = int(self._n(ev(1))) if len(args) > 1 else 0
            mode = {"ROUND": ROUND_HALF_UP, "ROUNDUP": ROUND_UP, "ROUNDDOWN": ROUND_DOWN}[name]
            try:
                return float(Decimal(repr(self._n(ev(0)))).quantize(Decimal(1).scaleb(-d), rounding=mode))
            except (InvalidOperation, ValueError):
                raise XlErr("#NUM!")
        if name == "ABS":
            return abs(self._n(ev(0)))
        if name == "INT":
            return float(math.floor(self._n(ev(0))))
        if name in ("CONCATENATE", "CONCAT"):
            return "".join(_xl_str(x) for a in args for x in self._flat(self._ev(a, ctx), ctx) if not isinstance(x, XlErr))
        if name in _STR_FN:
            return _STR_FN[name](_xl_str(ev(0)), *[self._n(ev(i)) for i in range(1, len(args))])
        if name in ("SUMIF", "COUNTIF", "SUMIFS", "COUNTIFS"):
            return self._cond(name, args, ctx)
        if name == "VLOOKUP":
            look, rg, idx = ev(0), self._rng(args[1], ctx), int(self._n(ev(2)))
            if not 1 <= idx <= rg.c2 - rg.c1 + 1:
                raise XlErr("#REF!")
            keys = [self._cell(rg.sheet, r, rg.c1) for r in range(rg.r1, min(rg.r2, self._dims(rg.sheet)[0]) + 1)]
            exact = len(args) > 3 and not self._n(ev(3))
            v = self._cell(rg.sheet, rg.r1 + self._find(look, keys, 0 if exact else 1), rg.c1 + idx - 1)
            if isinstance(v, XlErr):
                raise XlErr(v.code)
            return v
        if name == "MATCH":
            return float(self._find(ev(0), self._flat(self._ev(args[1], ctx), ctx), int(self._n(ev(2))) if len(args) > 2 else 1) + 1)
        if name == "INDEX":
            rg = self._rng(args[0], ctx)
            r = int(self._n(ev(1))) if len(args) > 1 else 1
            c = int(self._n(ev(2))) if len(args) > 2 else 1
            if rg.r1 == rg.r2 and len(args) == 2:
                r, c = 1, r
            rr, cc = rg.r1 + (r or 1) - 1, rg.c1 + (c or 1) - 1
            if not (rg.r1 <= rr <= rg.r2 and rg.c1 <= cc <= rg.c2):
                raise XlErr("#REF!")
            return self._cell(rg.sheet, rr, cc)
        raise XlUnsupported(f"function {name}")
    @staticmethod
    def _find(look, keys, mt):
        """Vị trí (0-based) cho MATCH/VLOOKUP: mt=0 khớp đúng; mt>0 lớn nhất <= look; mt<0 nhỏ nhất >= look."""
        best = None
        for i, k in enumerate(keys):
            if k is None or isinstance(k, XlErr) or isinstance(k, str) != isinstance(look, str):
                continue
            if mt == 0:
                if _cmp("=", k, look):
                    return i
            elif _cmp("<=" if mt > 0 else ">=", k, look):
                best = i
            else:
                break
        if best is None:
            raise XlErr("#N/A")
        return best
    def _cond(self, name, args, ctx):
        tgt, rest = None, args
        if name == "SUMIFS":
            tgt, rest = self._rng(args[0], ctx), args[1:]
        elif name == "SUMIF":
            rest, tgt = args[:2], self._rng(args[2] if len(args) > 2 else args[0], ctx)
        pairs = [(self._rng(rest[i], ctx), _crit(self._sc(self._ev(rest[i + 1], ctx), ctx))) for i in range(0, len(rest) - 1, 2)]
        base, total, cnt = pairs[0][0], 0.0, 0
        for r, c, _ in self._cells(base):
            dr, dc = r - base.r1, c - base.c1
            vals = [self._cell(g.sheet, g.r1 + dr, g.c1 + dc) for g, _ in pairs]
            if all(not isinstance(v, XlErr) and p(v) for v, (_, p) in zip(vals, pairs)):
                cnt += 1
                if tgt is not None:
                    x = self._cell(tgt.sheet, tgt.r1 + dr, tgt.c1 + dc)
                    total += x if isinstance(x, (int, float)) and not isinstance(x, bool) else 0.0
        return float(cnt) if tgt is None else total

class Sheet:
    def __init__(self, name, ws, ws_f=None, engine=None):
        self.name = name
        nrows, ncols = self._used_bounds(ws_f if ws_f is not None else ws)
        grid = [[None] * ncols for _ in range(nrows)]
        if nrows and ncols:
            for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=nrows, max_col=ncols, values_only=True)):
                for ci, v in enumerate(row):
                    grid[ri][ci] = clean_cell(v)
        self._raw = {}
        try:
            src = ws_f if ws_f is not None else ws
            for ri, row in enumerate(src.iter_rows(min_row=1, max_row=nrows, max_col=ncols, values_only=True)):
                for ci, v in enumerate(row):
                    v = getattr(v, "text", v)  # công thức mảng
                    if v is None or (isinstance(v, str) and v.strip() == ""):
                        continue
                    self._raw[(ri, ci)] = v
        except Exception:
            pass
        if ws_f is not None and nrows and ncols:
            try:
                for ri, frow in enumerate(ws_f.iter_rows(min_row=1, max_row=nrows, max_col=ncols, values_only=True)):
                    for ci, fval in enumerate(frow):
                        if grid[ri][ci] is not None:
                            continue
                        fval = getattr(fval, "text", fval)  # công thức mảng
                        if isinstance(fval, str) and fval.startswith("="):
                            ev = XFAIL
                            if engine is not None:
                                try:
                                    ev = engine.value(self.name, ri + 1, ci + 1)
                                except Exception:
                                    ev = XFAIL
                            if ev is XFAIL:  # FormulaEngine không hỗ trợ -> để trống, không đoán số
                                ev = None
                            elif isinstance(ev, str):
                                ev = clean_cell(ev)
                            if ev is not None:
                                grid[ri][ci] = ev
            except Exception:
                pass
        self.nrows, self.ncols = nrows, ncols
        wide = []
        for rng in ws.merged_cells.ranges:
            r0, c0, r1, c1 = rng.min_row - 1, rng.min_col - 1, rng.max_row - 1, rng.max_col - 1
            if r0 >= nrows or c0 >= ncols:
                continue
            width = c1 - c0 + 1
            if width >= max(WIDE_MERGE_MIN_COLS, WIDE_MERGE_RATIO * ncols) and r1 == r0:
                wide.append((r0, c0))
                continue
            val = grid[r0][c0]
            for r in range(r0, min(r1, nrows - 1) + 1):
                for c in range(c0, min(c1, ncols - 1) + 1):
                    grid[r][c] = val
        self.grid = grid
        spans = [(0, nrows)]
        if MULTI_TABLE and nrows:
            try:
                spans = self._split_tables(grid, nrows)
            except Exception:
                spans = [(0, nrows)]
        self.tables = [self._analyze(grid, ncols, a, b, wide) for (a, b) in spans]
        t0 = self.tables[0]
        self.header_end, self.titles, self.title_rows, self.headers = t0["he"], t0["titles"], t0["title_rows"], t0["headers"]
    @staticmethod
    def _used_bounds(ws):
        """(nrows, ncols) của vùng thực sự có dữ liệu/công thức, không áp trần cứng."""
        last_r = last_c = 0
        for ri, row in enumerate(ws.iter_rows(values_only=True), 1):
            for ci in range(len(row) - 1, -1, -1):
                v = row[ci]
                if v is None or (isinstance(v, str) and v == ""):
                    continue
                last_r = ri
                if ci + 1 > last_c:
                    last_c = ci + 1
                break
        return last_r, last_c
    @staticmethod
    def _split_tables(grid, nrows):
        """Tách sheet thành các bảng xếp chồng, ngăn cách bởi >= 2 dòng trống."""
        empty = [all(v is None for v in row) for row in grid]
        blocks, r = [], 0
        while r < nrows:
            if empty[r]:
                r += 1
                continue
            s, e, gap = r, r, 0
            while r < nrows:
                if empty[r]:
                    gap += 1
                    if gap >= 2:
                        break
                else:
                    gap, e = 0, r
                r += 1
            blocks.append((s, e + 1))
        spans, pending = [], None
        for (s, e) in blocks:
            if any(_data_like(grid[i]) for i in range(s, e)):
                spans.append([pending if pending is not None else s, e])
                pending = None
            elif pending is None:
                pending = s
        if len(spans) < 2:
            return [(0, nrows)]
        spans[0][0] = 0
        spans[-1][1] = nrows
        return [tuple(sp) for sp in spans]
    @staticmethod
    def _analyze(grid, ncols, a, b, wide):
        """Dò tiêu đề cho khối dòng [a, b): dòng đầu có dạng dữ liệu, các dòng trước đó là tiêu đề."""
        he = a
        for r in range(a, b):
            if _data_like(grid[r]):
                he = r
                break
        if he - a > MAX_HEADER_ROWS:
            he = a
        titles = []
        for (r0, c0) in wide:
            if a <= r0 < he and grid[r0][c0] is not None:
                titles.append((r0 + 1, fmt_val(grid[r0][c0])))
                grid[r0][c0] = None
        title_rows = set()
        for r in range(a, he):
            filled = [c for c in range(ncols) if grid[r][c] is not None]
            if filled and all(c == 0 for c in filled):
                title_rows.add(r)
                titles.append((r + 1, fmt_val(grid[r][0])))
        headers = {}
        for c in range(ncols):
            parts = []
            for r in range(a, he):
                if r in title_rows:
                    continue
                v = grid[r][c]
                if v is None:
                    continue
                t = fmt_val(v)
                if not parts or parts[-1] != t:
                    parts.append(t)
            if parts:
                headers[get_column_letter(c + 1)] = " > ".join(parts)
        return {"start": a, "end": b, "he": he, "titles": titles, "title_rows": title_rows, "headers": headers}
    def v(self, addr):
        col, row = coordinate_from_string(addr.replace("$", "").upper())
        r, c = row - 1, column_index_from_string(col) - 1
        if 0 <= r < self.nrows and 0 <= c < self.ncols:
            return self.grid[r][c]
        return None
    def raw(self, addr):
        """Nội dung gốc của ô: chuỗi công thức đúng như viết (có '=') nếu là công thức, ngược lại giá trị gốc. Trống → None."""
        col, row = coordinate_from_string(addr.replace("$", "").upper())
        return self._raw.get((row - 1, column_index_from_string(col) - 1))
    def formula(self, addr):
        """Công thức đúng như viết trong ô (vd '=SUM(A1:A3)'); None nếu ô không phải công thức."""
        v = self.raw(addr)
        return v if isinstance(v, str) and v.startswith("=") else None
    def raw_row(self, r):
        return {get_column_letter(c + 1): v for (rr, c), v in sorted(self._raw.items(), key=lambda kv: kv[0][1]) if rr == r - 1}
    def raw_col(self, letter):
        ci = column_index_from_string(letter.upper()) - 1
        return {rr + 1: v for (rr, c), v in sorted(self._raw.items()) if c == ci}
    def row(self, r):
        if not (1 <= r <= self.nrows):
            return {}
        return {get_column_letter(c + 1): v for c, v in enumerate(self.grid[r - 1]) if v is not None}
    def col(self, letter):
        c = column_index_from_string(letter.upper()) - 1
        if not (0 <= c < self.ncols):
            return {}
        return {r + 1: self.grid[r][c] for r in range(self.nrows) if self.grid[r][c] is not None}
    
    def find_rows(self, text, col="A", exact=False):
        n = norm(text)
        nj = n.replace(" ", "")
        exact_hits, partial_hits = [], []
        for r, val in self.col(col).items():
            vn = norm(val)
            if not vn:
                continue
            if vn == n or vn.replace(" ", "") == nj:
                exact_hits.append(r)
            elif n in vn or nj in vn.replace(" ", ""):
                partial_hits.append(r)
        if exact:
            return exact_hits
        return exact_hits if exact_hits else partial_hits

    def find_cols(self, text):
        n = norm(text)
        nj = n.replace(" ", "")
        out = []
        for t in self.tables:
            for letter, path in t["headers"].items():
                pn = norm(path)
                if (n in pn or nj in pn.replace(" ", "")) and letter not in out:
                    out.append(letter)
        return out
    def _dump_table(self, t):
        lines = []
        seen = set()
        tl = []
        for r, title in t["titles"]:
            if title and title not in seen and title not in t["headers"].values():
                seen.add(title)
                tl.append(f"  R{r}: {title}")
        if tl:
            lines.append("TITLE / UNITS:")
            lines += tl
        if t["headers"]:
            lines.append("COLUMN HEADERS (letter: header path):")
            for letter, p in t["headers"].items():
                lines.append(f"  {letter}: {p}")
        if DUMP_HEADER_CELLS:
            hl = []
            for r in range(t["start"], t["he"]):
                if r in t["title_rows"]:
                    continue
                cells = [f"{get_column_letter(c + 1)}={fmt_val(v)}" for c, v in enumerate(self.grid[r]) if v is not None]
                if cells:
                    hl.append(f"R{r + 1} | " + " | ".join(cells))
            if hl:
                lines.append("HEADER ROWS (same cells as above, by address):")
                lines += hl
        lines.append("ROWS:")
        for r in range(t["he"], t["end"]):
            cells = [f"{get_column_letter(c + 1)}={fmt_val(v)}" for c, v in enumerate(self.grid[r]) if v is not None]
            if cells:
                lines.append(f"R{r + 1} | " + " | ".join(cells))
        return lines
    def dump(self):
        lines = [f"=== SHEET: {self.name} (rows 1-{self.nrows}, columns A-{get_column_letter(max(self.ncols, 1))}) ==="]
        if len(self.tables) == 1:
            lines += self._dump_table(self.tables[0])
        else:
            lines.append(f"NOTE: this sheet holds {len(self.tables)} separate tables stacked vertically; "
                         "each table has its own COLUMN HEADERS, so the same column letter may mean different things per table.")
            for k, t in enumerate(self.tables, 1):
                lines.append(f"--- TABLE {k} (rows {t['start'] + 1}-{t['end']}) ---")
                lines += self._dump_table(t)
        return "\n".join(lines)
from collections import OrderedDict
import gc
_wb_cache, _wb_lock = OrderedDict(), threading.Lock()  # LRU: giữ tối đa WB_CACHE_MAX workbook (RAM <= 1 GB)

def load_workbook_sheets(path):
    with _wb_lock:
        if path in _wb_cache:
            _wb_cache.move_to_end(path)
            return _wb_cache[path]
        wb_v = openpyxl.load_workbook(path, data_only=True)
        try:
            wb_f = openpyxl.load_workbook(path, data_only=False)
        except Exception:
            wb_f = None
        sheets = {}
        try:
            engine = FormulaEngine(wb_v, wb_f) if wb_f is not None else None
        except Exception:
            engine = None
        for ws in wb_v.worksheets:
            wsf = wb_f[ws.title] if (wb_f is not None and ws.title in wb_f.sheetnames) else None
            sh = Sheet(ws.title, ws, wsf, engine)
            if any(v is not None for row in sh.grid for v in row):
                sheets[ws.title] = sh
        text = "\n\n".join(s.dump() for s in sheets.values())
        del wb_v, wb_f, engine  # giải phóng workbook openpyxl ngay; chỉ giữ Sheet đã dựng
        _wb_cache[path] = (sheets, text)
        while len(_wb_cache) > WB_CACHE_MAX:
            _wb_cache.popitem(last=False)
        gc.collect()
        return _wb_cache[path]

# =============================================================================
# 3. SANDBOX CHẠY CODE DO LLM SINH
# =============================================================================
_SAFE_MODULES = {"datetime", "math", "re", "statistics", "itertools", "collections", "functools"}

def _safe_import(name, *a, **k):
    if name.split(".")[0] not in _SAFE_MODULES:
        raise ImportError(f"import {name} khong duoc phep")
    return __import__(name, *a, **k)
_BUILTIN_NAMES = ["abs", "all", "any", "bool", "dict", "enumerate", "filter", "float", "int", "len", "list", "map", "max",
                  "min", "range", "round", "set", "sorted", "str", "sum", "tuple", "zip", "isinstance", "reversed",
                  "divmod", "pow", "repr", "next", "iter", "slice", "frozenset", "Exception", "ValueError", "KeyError",
                  "IndexError", "TypeError", "ZeroDivisionError", "StopIteration", "True", "False", "None"]

def run_code(code, sheets, timeout=None):
    import builtins as _b
    safe = {n: getattr(_b, n) for n in _BUILTIN_NAMES if hasattr(_b, n)}
    safe["__import__"] = _safe_import
    safe["print"] = lambda *a, **k: None
    first = next(iter(sheets.values()))
    env = {"__builtins__": safe, "SHEETS": sheets, "sh": first, "num": num, "norm": norm, "to_date": to_date,
           "dt": dt, "math": math, "re": re, "statistics": statistics, "answer": None}
    res = {}
    def target():
        try:
            exec(compile(code, "<llm_code>", "exec"), env)
            res["answer"] = env.get("answer")
        except BaseException as e:  # noqa
            res["err"] = f"{type(e).__name__}: {e}"
    th = threading.Thread(target=target, daemon=True)
    th.start()
    th.join(timeout if timeout is not None else SANDBOX_TIMEOUT)
    if th.is_alive():
        return None, "Timeout: code chay qua lau"
    return res.get("answer"), res.get("err")

# =============================================================================
# 4. CHUẨN HOÁ ĐÁP ÁN (Exact Match) — format lấy từ QUERY đa ngôn ngữ
# =============================================================================
_WORD_NUM = {
    "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
    "seven": 7, "eight": 8, "nine": 9, "ten": 10,
    "một": 1, "mot": 1, "hai": 2, "ba": 3, "bốn": 4, "bon": 4, "năm": 5, "nam": 5,
}
_PLACES = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
           "1": 1, "2": 2, "3": 3, "4": 4, "5": 5, "0": 0}

def _nfc(s):
    """Chuẩn hoá Unicode về NFC (tiếng Việt có thể ở dạng NFD / dấu tách rời)."""
    return unicodedata.normalize("NFC", s) if isinstance(s, str) else s

def _parse_place_token(tok):
    tok = (tok or "").strip().lower()
    if tok in _PLACES:
        return _PLACES[tok]
    if tok in _WORD_NUM:
        return _WORD_NUM[tok]
    if tok.isdigit():
        return int(tok)
    return None

def _wants_integer(q):
    if not q:
        return False
    q = _nfc(q)  
    ql = q.lower()
    if re.search(r"\binteger\b|whole\s+number|số\s+nguyên", ql):
        return True
    if re.search(r"整数|整數|정수", q):
        return True
    return False
# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
MULTI_SEP = os.environ.get("MULTI_SEP", ", ")
_MODE_DEC = {"half_up": ROUND_HALF_UP, "up": ROUND_UP, "down": ROUND_DOWN, "ceil": ROUND_CEILING,
             "floor": ROUND_FLOOR, "half_even": ROUND_HALF_EVEN}
_MODE_LABEL = {"half_up": "round half up / nearest (Excel ROUND)", "up": "round up, away from zero (Excel ROUNDUP)",
               "down": "round down / truncate toward zero (Excel ROUNDDOWN)", "ceil": "ceiling (toward +inf)",
               "floor": "floor (toward -inf)", "half_even": "banker's rounding (half to even)"}
def _fn_mode(word):
    """'floor'/'ceiling' chỉ là kiểu làm tròn khi dùng như hàm/động từ (floor(x), use floor, floor the result...);
    không bắt danh từ thường như "floor area", "which floor", "ceiling height"."""
    not_noun = r"(?!\s*(?:area|plan|price|space|level|number|height|limit|rate|\d))"
    return (rf"\b{word}\s*\(|"
            rf"\b(?:use|using|apply|applying|take|taking|via|then)\s+(?:the\s+)?(?:excel\s+)?{word}\b{not_noun}|"
            rf"\bround(?:ed|ing|s)?\s+(?:\w+\s+){{0,5}}(?:using|with|by|via)\s+(?:the\s+)?{word}\b{not_noun}|"
            rf"\b{word}\s+(?:the\s+)?(?:result|value|number|average|mean|sum|total|figure|it)\b|"
            rf"\b{word}\s+to\s+(?:the\s+)?(?:nearest|\d|an?\s+integer|whole|integer)")

_MODE_PATTERNS = [
    ("half_even", r"banker|half[\s\-]*(?:to[\s\-]*)?even|round(?:ed|ing)?\s+to\s+even|làm\s*tròn\s*chẵn|银行家|銀行家|짝수"),
    ("ceil", _fn_mode(r"ceil(?:ing)?") + r"|向上取整|向上舍入"),
    ("floor", _fn_mode("floor") + r"|向下取整|向下舍入"),
    ("up", r"round(?:ed|ing|s)?\s+up\b|\broundup\b|làm\s*tròn\s*lên|tròn\s*lên|进一法|進一法|切り上げ|(?<!반)올림"),
    ("down", r"round(?:ed|ing|s)?\s+down\b|\brounddown\b|làm\s*tròn\s*xuống|tròn\s*xuống|truncat\w*|"
             r"cắt\s*(?:bỏ|bớt)|截断|截斷|截尾|舍去|去尾|切り捨て|切捨て|내림|버림"),
    ("half_up", r"round(?:ed|ing|s)?\b|làm\s*tròn|四舍五入|四捨五入|반올림|丸め"),
]
_MODE_RE = [(m, re.compile(p, re.I)) for m, p in _MODE_PATTERNS]
_CN_NUM = {"一": 1, "二": 2, "两": 2, "兩": 2, "三": 3, "四": 4, "五": 5}
_NUMW = r"zero|one|two|three|four|five|six|\d{1,2}"
_EXTRA_NUM = {"nhất": 1, "nhat": 1, "첫째": 1, "둘째": 2, "셋째": 3, "넷째": 4,
              "une": 1, "un": 1, "deux": 2, "trois": 3, "quatre": 4, "cinq": 5}

def _pl(tok):
    tok = (tok or "").strip().lower()
    if tok in _CN_NUM:
        return _CN_NUM[tok]
    if tok in _EXTRA_NUM:
        return _EXTRA_NUM[tok]
    return _parse_place_token(tok)
_PLACE_PATTERNS = [
    (re.compile(r"nearest\s+(tenth|hundredth|thousandth)", re.I), lambda t: {"tenth": 1, "hundredth": 2, "thousandth": 3}[t.lower()]),
    (re.compile(r"nearest\s+(ten|hundred|thousand)\b", re.I), lambda t: -{"ten": 1, "hundred": 2, "thousand": 3}[t.lower()]),
    (re.compile(r"hàng\s*(chục|trăm|nghìn|ngàn)", re.I), lambda t: -{"chục": 1, "trăm": 2, "nghìn": 3, "ngàn": 3}[t.lower()]),
    (re.compile(r"\b(" + _NUMW + r")[\s\-]*decimals?\b", re.I), _pl),
    (re.compile(r"decimal\s+places?\s*[:=]?\s*(\d+)", re.I), _pl),
    (re.compile(r"\b(\d+)\s*d\.?\s*p\.?\b", re.I), _pl),
    (re.compile(r"(một|hai|ba|bốn|bon|năm|nam|\d+)(?=\s*chữ\s*số\s*thập\s*phân)", re.I), _pl),
    (re.compile(r"làm\s*tròn\s*(?:đến|tới|về)?\s*(một|hai|ba|bốn|bon|\d+)\b", re.I), _pl),
    (re.compile(r"(\d+)\s*số\s*lẻ", re.I), _pl),
    (re.compile(r"thập\s*phân\s*[:=]?\s*(\d+)", re.I), _pl),
    (re.compile(r"保留\s*([一二两兩三四五\d]+)\s*位"), _pl),
    (re.compile(r"(\d+)\s*位小数"), _pl),
    (re.compile(r"四舍五入.*?(\d+)\s*位"), _pl),
    (re.compile(r"小数点以下\s*(\d+)\s*桁"), _pl),
    (re.compile(r"小数第\s*(\d+)\s*位"), _pl),
    (re.compile(r"소수점\s*(\d+)\s*자리"), _pl),
    (re.compile(r"(\d+)\s*자리\s*소수"), _pl),
    # --- bổ sung nhiều cách diễn đạt khác ---
    (re.compile(r"\b(" + _NUMW + r")[\s\-]*(?:digits?|numbers?|figures?)\s+(?:after|behind|following|to\s+the\s+right\s+of)\s+(?:the\s+)?(?:decimal|comma|dot|period|point)", re.I), _pl),
    (re.compile(r"(?:nearest|to|precision\s+of|increments?\s+of)\s+0\.(0*)1(?!\d)", re.I), lambda t: len(t) + 1),
    (re.compile(r"nearest\s+(10{1,3})\b", re.I), lambda t: -(len(t) - 1)),
    (re.compile(r"(?:format|pattern|as)\s*[:：]?\s*[\"'“]?[#,]*0\.(0+)", re.I), lambda t: len(t)),
    (re.compile(r"(một|hai|ba|bốn|bon|năm|nam|\d+)\s*chữ\s*số\s*(?:sau|đằng\s*sau)\s*dấu\s*(?:phẩy|thập\s*phân|chấm)", re.I), _pl),
    (re.compile(r"chữ\s*số\s*thập\s*phân\s*thứ\s*(nhất|nhat|hai|ba|bốn|bon)", re.I), _pl),
    (re.compile(r"(một|hai|ba|bốn|bon|\d+)\s*số\s*thập\s*phân", re.I), _pl),
    (re.compile(r"小数点[后後]\s*第?\s*([一二两兩三四五\d]+)\s*位"), _pl),
    (re.compile(r"([一二两兩三四五\d]+)\s*位小数"), _pl),
    (re.compile(r"精确到\s*0\.(0*)1(?!\d)"), lambda t: len(t) + 1),
    (re.compile(r"小数(?:点(?:以下)?)?第?\s*([一二三四五\d]+)\s*[位桁]"), _pl),
    (re.compile(r"소수(?:점)?\s*(?:이하\s*)?(첫째|둘째|셋째|넷째)\s*자리"), _pl),
    (re.compile(r"소수점\s*(?:이하\s*)?(\d+)\s*자리"), _pl),
    (re.compile(r"\b(\d+|une?|deux|trois|quatre|cinq)\s*décimales?\b", re.I), _pl),
    (re.compile(r"\b(integers?|whole\s+numbers?)\b|(số\s+nguyên)|(整数|整數|(?<![上下])取整|정수)", re.I), lambda t: 0),
]

def _scan(patterns, s):
    """Trả về [(start, end, giá trị)] không chồng nhau, theo thứ tự xuất hiện."""
    taken = []
    for rx, fn in patterns:
        for m in rx.finditer(s):
            a, b = m.span()
            if any(a < e and s_ < b for s_, e, _ in taken):
                continue
            g = next((x for x in m.groups() if x), "") if m.groups() else ""
            try:
                val = fn(g)
            except Exception:
                val = None
            if val is None:
                continue
            taken.append((a, b, val))
    return sorted(taken)

def _modes_in(s):
    return [v for _, _, v in _scan([(rx, (lambda m: (lambda _t: m))(m)) for m, rx in _MODE_RE], s)]

def _places_in(s):
    return [v for _, _, v in _scan(_PLACE_PATTERNS, s)]
_CLAUSE_SPLIT = re.compile(
    r"\s*(?:[;,，、；]|\band\b|\bas well as\b|\bthen\b|\bvà\b|\bcũng như\b|和|与|與|及|と|및)\s*", re.I)

@lru_cache(maxsize=4096)
def rounding_specs(q):
    """Danh sách (mode, places) theo thứ tự trong câu hỏi. () nếu không yêu cầu làm tròn."""
    if not q:
        return ()
    gm, gp = set(_modes_in(q)), set(_places_in(q))
    d_mode = next(iter(gm)) if len(gm) == 1 else None
    d_places = next(iter(gp)) if len(gp) == 1 else None
    specs = []
    for c in _CLAUSE_SPLIT.split(q):
        if not c or not c.strip():
            continue
        ms, ps = _modes_in(c), _places_in(c)
        if not ms and not ps:
            continue
        if not ps and d_places is None:
            if all(m == "half_up" for m in ms):
                continue  # "round the average" -> không đủ thông tin, giữ hành vi cũ
            ps = [0]    # "round up the average" -> ROUNDUP về số nguyên
        for i in range(max(len(ms), len(ps), 1)):
            mode = (ms[i] if i < len(ms) else ms[-1]) if ms else (d_mode or "half_up")
            places = (ps[i] if i < len(ps) else ps[-1]) if ps else d_places
            specs.append((mode, places))
    return tuple(dict.fromkeys(specs))

def _fmt_round(v, places, mode="half_up"):
    """Làm tròn và giữ đúng số chữ số thập phân theo yêu cầu."""
    places = 0 if places is None else int(places)
    try:
        r = Decimal(f"{float(v):.15g}").quantize(
            Decimal(1).scaleb(-places),
            rounding=_MODE_DEC.get(mode, ROUND_HALF_UP)
        )
    except (InvalidOperation, ValueError, OverflowError):
        return f"{round(float(v), places):.{max(places, 0)}f}"
    if r == 0:
        r = abs(r)
    if places <= 0:
        return str(int(r))
    return f"{r:.{places}f}"

def _post_multi(ans, specs):
    """Nhiều yêu cầu làm tròn. Trả None nếu không áp dụng được (để rơi về luồng cũ)."""
    if ans is None or isinstance(ans, bool):
        return None
    if isinstance(ans, (list, tuple)):
        items = list(ans)
    elif isinstance(ans, (int, float)):
        items = [ans]
    else:
        s = re.sub(r"\s+", " ", str(ans)).strip()
        m = re.search(r"<FINAL_ANSWER>\s*(.*?)\s*</FINAL_ANSWER>", s, flags=re.S | re.I)
        if m:
            s = m.group(1).strip()
        s = _PREFIX_JUNK.sub("", s).strip()
        items = [p for p in re.split(r"\s*[;|]\s*|,\s+", s) if p != ""]
    vals = []
    for x in items:
        v = float(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else _try_float(x)
        if v is None or math.isnan(v) or math.isinf(v):
            return None
        vals.append(v)
    if len(vals) == 1:
        pairs = [(vals[0], sp) for sp in specs]
    elif len(vals) == len(specs):
        pairs = list(zip(vals, specs))
    else:
        return None
    return MULTI_SEP.join(_fmt_round(v, p, m) for v, (m, p) in pairs)

def multi_round_hint(query):
    specs = rounding_specs(query)
    if len(specs) < 2:
        return ""
    lines = []
    for i, (m, p) in enumerate(specs, 1):
        pl = "integer" if p == 0 else (f"{p} decimal places" if p > 0 else f"nearest 10^{-p}")
        lines.append(f"  {i}. {_MODE_LABEL[m]}, {pl}")
    return ("\n\nOUTPUT FORMAT NOTE (overrides rule 7): the question asks for several rounded outputs, in this order:\n"
            + "\n".join(lines) +
            "\nDo NOT round or format yourself. Set `answer` to the FULL-PRECISION unrounded number when all outputs "
            "are the same quantity; if the outputs are DIFFERENT quantities, set `answer` to a Python list of "
            "full-precision numbers, one per output, in the order above. The post-processor applies the roundings and joins them.")
_PREFIX_JUNK = re.compile(
    r"^(?:the\s+answer\s+is\s*[:：]?|(?:final\s*answer|answer|result|đáp\s*án|kết\s*quả)\s*[:：])\s*",
    re.I,
)

def _try_float(s):
    if s is None:
        return None
    t = str(s).strip().replace("\u00a0", "").replace("\u202f", "").replace(" ", "")
    if not t or re.fullmatch(r"(?i)nan|none|null", t):
        return None
    if "," in t and "." in t:
        if t.rfind(",") > t.rfind("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            t = t.replace(",", "")
    elif "," in t:
        parts = t.split(",")
        if len(parts) == 2 and len(parts[1]) != 3:
            t = t.replace(",", ".")
        else:
            t = t.replace(",", "")
    try:
        return float(t)
    except ValueError:
        return None
_DATE_TOKEN = {"yyyy": "%Y", "yy": "%y", "mm": "%m", "dd": "%d"}
_DATE_TRIPLE_RE = re.compile(r"(?<![a-z])(yyyy|yy|mm|dd)\s*([/\-.])\s*(yyyy|yy|mm|dd)\s*\2\s*(yyyy|yy|mm|dd)(?![a-z])", re.I)

def _date_format_from_query(q):
    """Chuỗi strftime theo yêu cầu trong query; mặc định ISO YYYY-MM-DD."""
    if not q:
        return "%Y-%m-%d"
    q = _nfc(q) 
    if re.search(r"yyyy\s*年\s*mm\s*月\s*dd\s*日?", q, re.I):
        return "%Y年%m月%d日"
    m = _DATE_TRIPLE_RE.search(q)
    if m:
        toks = [m.group(i).lower() for i in (1, 3, 4)]
        if len(set(toks)) == 3:
            return m.group(2).join(_DATE_TOKEN[t] for t in toks)
    ql = q.lower()
    if re.search(r"ngày\s*[/\-.]\s*tháng\s*[/\-.]\s*năm", ql):
        return "%d/%m/%Y"
    if re.search(r"tháng\s*[/\-.]\s*ngày\s*[/\-.]\s*năm", ql):
        return "%m/%d/%Y"
    return "%Y-%m-%d"

def _fmt_date_answer(d, q, date_format=None):
    has_time = isinstance(d, dt.datetime) and (d.hour, d.minute, d.second) != (0, 0, 0)
    if not isinstance(d, dt.datetime):
        d = dt.datetime(d.year, d.month, d.day)
    out = d.strftime(date_format or _date_format_from_query(q))
    return out + d.strftime(" %H:%M:%S") if has_time else out
_ISO_MIDNIGHT = re.compile(r"^(\d{4}-\d{2}-\d{2})[ T]00:00:00(?:\.0+)?$")

# -----------------------------------------------------------------------------
# FORMAT SPEC: LLM trả JSON {rounding_mode, decimals, date_format, extra}; formatter bên dưới áp dụng
# một cách xác định. Regex (rounding_specs / _date_format_from_query) chỉ là DỰ PHÒNG khi JSON thiếu/sai.
# -----------------------------------------------------------------------------
_MODE_ALIASES = {
    "none": None, "null": None, "no": None, "": None,
    "half_up": "half_up", "round": "half_up", "nearest": "half_up", "halfup": "half_up",
    "up": "up", "round_up": "up", "roundup": "up",
    "down": "down", "round_down": "down", "rounddown": "down", "truncate": "down", "trunc": "down",
    "ceil": "ceil", "ceiling": "ceil", "floor": "floor",
    "half_even": "half_even", "bankers": "half_even", "banker": "half_even", "halfeven": "half_even",
}
_SAFE_DATE_FMT = re.compile(r"(?:%[YymdbBHMSjaAeIpT]|[\s\-/.:,T年月日])+")

def _coerce_places(x):
    """Số chữ số thập phân từ JSON; None nếu rỗng; _INVALID nếu có giá trị nhưng không hợp lệ."""
    if x is None or x == "" or (isinstance(x, str) and x.strip().lower() in ("null", "none")):
        return None
    if isinstance(x, bool):
        return _INVALID
    if isinstance(x, str):
        t = x.strip()
        v = int(t) if re.fullmatch(r"[-+]?\d+", t) else _parse_place_token(t)
        if v is None:
            return _INVALID
    elif isinstance(x, (int, float)) and float(x).is_integer():
        v = int(x)
    else:
        return _INVALID
    return v if -6 <= v <= 15 else _INVALID
_INVALID = object()

def _norm_date_fmt(s):
    """Chuỗi strftime an toàn từ JSON ('%d/%m/%Y'; cũng nhận 'dd/mm/yyyy'); None nếu rỗng/không hợp lệ."""
    if not isinstance(s, str) or not s.strip() or s.strip().lower() in ("null", "none"):
        return None
    t = s.strip()
    if "%" not in t:
        t = re.sub(r"(?i)yyyy", "%Y", t)
        t = re.sub(r"(?i)yy(?![a-z])", "%y", t)
        t = re.sub(r"(?i)mm", "%m", t)
        t = re.sub(r"(?i)dd", "%d", t)
    if not _SAFE_DATE_FMT.fullmatch(t) or not re.search(r"%[Yy]", t) or not re.search(r"%[mbBj]", t):
        return None
    try:
        dt.datetime(2001, 2, 3, 4, 5, 6).strftime(t)
    except Exception:
        return None
    return t

def fmt_from_obj(obj):
    """dict JSON -> {'specs': ((mode, places), ...), 'date_format': str|None} hoặc None nếu không hợp lệ."""
    if not isinstance(obj, dict) or not any(k in obj for k in ("rounding_mode", "decimals", "date_format")):
        return None
    items = [obj] + [e for e in (obj.get("extra") or obj.get("also") or []) if isinstance(e, dict)]
    specs = []
    for it in items:
        mk = str(it.get("rounding_mode") or "").strip().lower().replace("-", "_").replace(" ", "_")
        if mk not in _MODE_ALIASES:
            return None
        mode, places = _MODE_ALIASES[mk], _coerce_places(it.get("decimals"))
        if places is _INVALID:
            return None
        if mode is None and places is None:
            continue  # mục này không yêu cầu làm tròn
        specs.append((mode or "half_up", 0 if places is None else places))
    return {"specs": tuple(specs), "date_format": _norm_date_fmt(obj.get("date_format"))}

def parse_fmt(text):
    """Tìm object JSON format spec trong văn bản LLM (có/không hàng rào ```json). None nếu không có/không hợp lệ."""
    if not isinstance(text, str):
        return None
    dec = json.JSONDecoder()
    for m in re.finditer(r"\{", text):
        try:
            obj, _ = dec.raw_decode(text[m.start():])
        except ValueError:
            continue
        if isinstance(obj, dict) and any(k in obj for k in ("rounding_mode", "decimals", "date_format")):
            return fmt_from_obj(obj)
    return None

def fmt_mismatch(fmt, q):
    """So JSON với regex (chỉ để ghi log đối chiếu); None nếu khớp hoặc không có fmt."""
    if not fmt:
        return None
    rx = rounding_specs(q)
    return None if tuple(fmt["specs"]) == tuple(rx) else {"json": list(fmt["specs"]), "regex": list(rx)}

def format_note(fmt):
    """Ghi chú gắn vào prompt sinh code khi đã có format spec hợp lệ: code KHÔNG tự làm tròn/định dạng."""
    if not fmt or not (fmt["specs"] or fmt["date_format"]):
        return ""
    parts = []
    if fmt["specs"]:
        lines = []
        for i, (m, p) in enumerate(fmt["specs"], 1):
            pl = "integer" if p == 0 else (f"{p} decimal places" if p > 0 else f"nearest 10^{-p}")
            lines.append(f"  {i}. {_MODE_LABEL[m]}, {pl}")
        parts.append("rounding, in this order:\n" + "\n".join(lines) +
                     "\nSet `answer` to the FULL-PRECISION unrounded number when all outputs are the same quantity; "
                     "if the outputs are DIFFERENT quantities, set `answer` to a Python list of full-precision numbers, "
                     "one per output, in the order above.")
    if fmt["date_format"]:
        parts.append(f"date format {fmt['date_format']!r}: when the answer is a date, set `answer` to the datetime/date "
                     "object (use to_date(...)), not a formatted string.")
    return ("\n\nOUTPUT FORMAT NOTE (overrides rule 7): a deterministic post-processor applies the requested format. "
            "Do NOT round or format yourself.\n" + "\n".join(parts))
# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
_FOOTNOTE_TAIL = re.compile(r"\s*\(\d+\)\s*$")

def _lab_key(s):
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", str(s))).strip()

def _lab_loose(k):
    """Khoá nới lỏng: bỏ dấu câu cuối + không phân biệt hoa/thường (để khôi phục dấu chấm bị mất)."""
    return k.rstrip(" .,;:").casefold()

def sheet_labels(sh):
    """(tập nhãn chính xác, {khoá nới lỏng: {nhãn gốc}}) của 1 sheet; cache trên đối tượng Sheet."""
    cached = getattr(sh, "_labels", None)
    if cached is not None:
        return cached
    exact, loose, dotted = set(), {}, set()
    def add(v):
        if isinstance(v, str):
            k = _lab_key(v)
            if k and _try_float(k) is None:
                exact.add(k)
                loose.setdefault(_lab_loose(k), set()).add(k)
                dotted.update(w for w in k.split(" ") if len(w) > 1 and w.endswith(".") and not w[:-1].isdigit())
    for row in sh.grid:
        for v in row:
            add(v)
    for t in sh.tables:
        for _, title in t["titles"]:
            add(title)
    sh._labels = (exact, loose, dotted)
    return sh._labels

def build_labels(sheets):
    exact, loose, dotted = set(), {}, set()
    for sh in sheets.values():
        e, l, d = sheet_labels(sh)
        exact |= e
        dotted |= d
        for k, v in l.items():
            loose.setdefault(k, set()).update(v)
    return exact, loose, dotted

def _label_restore(s, q, labels):
    """Nếu chuỗi s khớp một nhãn trong workbook → trả đúng nhãn gốc (giữ dấu chấm/phẩy); ngược lại None."""
    if not labels or not isinstance(s, str):
        return None
    k = _lab_key(s)
    if not k or _try_float(k) is not None:
        return None
    exact, loose = labels[0], labels[1]
    hit = k if k in exact else None
    if hit is None:
        c = loose.get(_lab_loose(k))
        if c and len(c) == 1 and k.casefold() not in _lab_key(q or "").casefold():
            hit = next(iter(c))
    if hit is None:
        return None
    return _FOOTNOTE_TAIL.sub("", hit) or hit

def _split_labels(s, q, labels, max_join=6):
    """Tách 'a, b, c' thành các phần, nhưng ghép lại các mảnh cùng tạo thành MỘT nhãn có dấu phẩy"""
    toks = re.split(r",\s+", s.strip())
    if not labels:
        return toks
    out, i, n = [], 0, len(toks)
    while i < n:
        merged = None
        for j in range(min(n, i + max_join), i + 1, -1):
            cand = ", ".join(toks[i:j])
            if _label_restore(cand, q, labels) is not None:
                merged = (cand, j)
                break
        if merged:
            out.append(merged[0])
            i = merged[1]
        else:
            out.append(toks[i])
            i += 1
    return out

def _is_abbrev_token(tok, dotted=()):
    """Token cuối là từ viết tắt có dấu chấm? Nhận biết tổng quát bằng cấu trúc (L.L.C., U.S., chữ cái đầu 'J.')"""
    t = tok.strip("()[]\"'“”")
    if not t.endswith("."):
        return False
    body = t[:-1]
    if re.search(r"[^\W\d_]\.[^\W\d_]", t):                    # có dấu chấm xen giữa: L.L.C. / U.S. / e.g.
        return True
    if len(body) == 1 and body.isalpha() and body.isupper():       # chữ cái đầu: J.
        return True
    return t in dotted
_NO_PCT = re.compile(r"(?:without|no|omit)\s+(?:the\s+)?%|(?:không|khong)\s+(?:kèm|có|co|kem)\s+(?:dấu\s+)?%|不(?:带|含|要)\s*%|%\s*なし|%\s*없이")

def _post_scalar(x, q, spec=None, labels=None, fmt=None):
    q = _nfc(q)
    _df = fmt["date_format"] if fmt else None
    if x is None:
        return ""
    if isinstance(x, bool):
        return "Yes" if x else "No"
    if isinstance(x, (dt.datetime, dt.date)):
        return _fmt_date_answer(x, q, _df)
    if isinstance(x, str):
        _m = _ISO_MIDNIGHT.match(x.strip())
        if _m:
            return _fmt_date_answer(dt.datetime.strptime(_m.group(1), "%Y-%m-%d"), q, _df)
    if isinstance(x, str):
        _hit = _label_restore(x, q, labels)
        if _hit is not None:
            return _hit
    _sp = fmt["specs"] if fmt else rounding_specs(q)  # JSON của LLM là nguồn chính; regex chỉ khi không có fmt
    mode, places = spec if spec is not None else (_sp[0] if _sp else ("half_up", None))
    want_int = False if fmt else _wants_integer(q)
    ql = (q or "").lower()
    if isinstance(x, (int, float)):
        v = float(x)
        if math.isnan(v) or math.isinf(v):
            return ""
        if places is not None:
            return _fmt_round(v, places, mode)
        if want_int or abs(v - round(v)) < 1e-12:
            return str(int(round(v)))
        return f"{v:.10g}"
    s = re.sub(r"\s+", " ", str(x)).strip()
    s = _PREFIX_JUNK.sub("", s).strip()
    if re.fullmatch(r"[-+]?\d+\.", s):
        s = s[:-1]
    elif s.endswith(".") and not re.search(r"\d+\.\d", s):
        if not _is_abbrev_token(s.split(" ")[-1], labels[2] if labels else ()):  # giữ dấu chấm của từ viết tắt
            s = s[:-1].strip()
    if re.fullmatch(r"(?i)yes|no|true|false", s):
        if re.search(r"yes\s*(or|/)\s*no", ql) or "yes or no" in ql:
            return "Yes" if s.lower() in ("yes", "true") else "No"
        if re.search(r"\btrue\b|\bfalse\b", ql):
            return "True" if s.lower() in ("yes", "true") else "False"
    if _NO_PCT.search(ql) \
            and re.fullmatch(r"[-+]?[\d.,]+\s*%", s):
        s = s.rstrip("%").strip()
    v = _try_float(s)
    if v is not None and not re.match(r"^[-+]?0\d", s):
        if places is not None:
            return _fmt_round(v, places, mode)
        if not want_int and re.fullmatch(r"[-+]?\d+\.\d{1,6}", s) and s.endswith("0"):
            return s
        if want_int or abs(v - round(v)) < 1e-12:
            return str(int(round(v)))
        return f"{v:.10g}" if abs(v - round(v)) >= 1e-12 else str(int(round(v)))
    return re.sub(r"\s*\(\d+\)\s*$", "", s)

def postprocess(ans, q, labels=None, fmt=None):
    """fmt: dict từ parse_fmt (JSON của LLM). None -> dùng regex trên câu hỏi (dự phòng)."""
    if ans is None:
        return ""
    specs = fmt["specs"] if fmt else rounding_specs(q)
    if len(specs) >= 2:
        out = _post_multi(ans, specs)
        if out is not None:
            return out
    if isinstance(ans, (list, tuple)):
        return ", ".join(p for p in (_post_scalar(a, q, labels=labels, fmt=fmt) for a in ans) if p != "")
    if isinstance(ans, str):
        if _NO_PCT.search(_nfc(q).lower()) and re.fullmatch(r"\s*[-+]?[\d.,]+\s*%\s*", ans):  # "không kèm %" → bỏ % trước khi khớp nhãn
            ans = ans.strip().rstrip("%").strip()
        m = re.search(r"<FINAL_ANSWER>\s*(.*?)\s*</FINAL_ANSWER>", ans, flags=re.S | re.I)
        if m:
            ans = m.group(1).strip()
        whole = _label_restore(ans, q, labels)  # cả chuỗi là MỘT nhãn (có thể chứa dấu phẩy)
        if whole is not None:
            return whole
        if re.search(r",\s+", ans):
            parts = [_post_scalar(p, q, labels=labels, fmt=fmt) for p in _split_labels(ans, q, labels)]
            return ", ".join(p for p in parts if p != "")
    return _post_scalar(ans, q, labels=labels, fmt=fmt)

# =============================================================================
# 4b. TRA CỨU Ô THEO ĐỊA CHỈ (câu hỏi "ô D5 chứa gì / công thức nào")
# =============================================================================
_FORMULA_Q = re.compile(r"formula|công\s*thức|cong\s*thuc|公式|数式|數式|수식|fórmula|formule", re.I)
_QREF_RE = re.compile(
    r"(?:\bcells?\b|\bô\b|单元格|單元格|儲存格|储存格|セル|셀|!)\s*[:：]?\s*['\"“‘]?\$?([A-Za-z]{1,3})\$?(\d{1,7})(?![\w])"
    r"(?!\s*:\s*\$?[A-Za-z]{1,3}\$?\d)", re.I)

def _query_refs(query):
    out = []
    for m in _QREF_RE.finditer(query or ""):
        col, row = m.group(1).upper(), int(m.group(2))
        try:
            if column_index_from_string(col) > 16384 or row < 1:
                continue
        except Exception:
            continue
        if (col, row) not in out:
            out.append((col, row))
    return out

def _query_sheet_names(query, sheets):
    ql = (query or "").lower()
    return [k for k in sheets if k and k.lower() in ql]

def _formulaish(t):
    """0 = không giống công thức; 6 = công thức thật ('=...'); 2/1 = chuỗi chữ trông như công thức."""
    if not isinstance(t, str):
        return 0
    s_ = t.strip()
    if s_.startswith("="):
        return 6
    if "=" in s_ and re.search(r"[\w\)\]]\s*[-+*/^]\s*[\w\(\[]", s_):
        return 2
    if re.search(r"(?:\(\w{1,3}\)|\b[A-Za-z]{1,3}\$?\d+\b)\s*[-+*/^]\s*(?:\(\w{1,3}\)|\b[A-Za-z]{1,3}\$?\d+\b)", s_):
        return 1
    return 0

def _label_match(t, letter):
    """Ô bắt đầu bằng nhãn trùng chữ cột, vd '(d)', '(d) = ...', 'D:' (kiểu cột đánh nhãn (a),(b),(c)...)."""
    return isinstance(t, str) and bool(re.match(rf"^\s*\(?{re.escape(letter)}\)?\s*(?:[=:.\-–—]|$)", t, re.I))

def _fmt_raw(v, n=90):
    t = str(v).replace("\n", " ")
    return repr(t if len(t) <= n else t[:n] + "…")

def _empty_cell_candidates(sh, col, row, formula_intent):
    """Khi ô được hỏi TRỐNG: xếp hạng các ô cùng dòng theo mức khớp (công thức/nhãn trùng chữ cột)."""
    res = []
    for letter, v in sh.raw_row(row).items():
        sc = (_formulaish(v) if formula_intent else 0) + (3 if _label_match(v, col) else 0)
        if sc > 0:
            res.append((sc, f"{letter}{row}", v))
    res.sort(key=lambda x: (-x[0], column_index_from_string(re.match(r"[A-Z]+", x[1]).group(0))))
    return res

def cell_lookup_text(query, sheets):
    """Khối bằng chứng đưa vào prompt cho mọi câu hỏi có nêu địa chỉ ô cụ thể ('cell D5')."""
    if not CELL_FALLBACK:
        return ""
    refs = _query_refs(query)
    if not refs:
        return ""
    names = _query_sheet_names(query, sheets) or list(sheets)[:4]
    fq = bool(_FORMULA_Q.search(query))
    lines = ["CELL LOOKUP (exact addresses named in the question, read from the raw file; authoritative):"]
    for col, row in refs[:3]:
        for name in names[:4]:
            sh = sheets[name]
            addr = f"{col}{row}"
            v = sh.raw(addr)
            if v is None:
                lines.append(f"- {name}!{addr}: EMPTY (no value, no formula)")
                rc = sh.raw_row(row)
                if rc:
                    lines.append(f"  non-empty cells in row {row}: " + " | ".join(f"{k}={_fmt_raw(x)}" for k, x in list(rc.items())[:25]))
                cc = sh.raw_col(col)
                if cc:
                    lines.append(f"  non-empty cells in column {col}: " + " | ".join(f"{col}{k}={_fmt_raw(x)}" for k, x in list(cc.items())[:8]))
                cands = _empty_cell_candidates(sh, col, row, fq)
                if cands:
                    lines.append("  closest matches in the same row (score): " +
                                 " ; ".join(f"{a}={_fmt_raw(x)} [{sc}]" for sc, a, x in cands[:4]))
            elif isinstance(v, str) and v.startswith("="):
                lines.append(f"- {name}!{addr}: FORMULA {_fmt_raw(v, 300)}  (computed value: {_fmt_raw(sh.v(addr))})")
            else:
                lines.append(f"- {name}!{addr}: {type(v).__name__.upper()} {_fmt_raw(v, 300)}")
    return "\n".join(lines) + "\n\n"

def try_direct_cell(query, sheets):
    """Câu hỏi chỉ đọc công thức của MỘT ô nêu rõ địa chỉ → trả lời thẳng, không cần LLM."""
    if not CELL_FALLBACK or not _FORMULA_Q.search(query or ""):
        return None
    refs = _query_refs(query)
    if len(refs) != 1:
        return None
    names = _query_sheet_names(query, sheets)
    if len(names) > 1 or (not names and len(sheets) > 1):
        return None
    sh = sheets[names[0]] if names else next(iter(sheets.values()))
    col, row = refs[0]
    addr = f"{col}{row}"
    v = sh.raw(addr)
    if isinstance(v, str) and v.startswith("="):
        return v, "formula"
    if isinstance(v, str) and _formulaish(v) >= 2:
        return v, "text-formula"
    if v is None:  # ô được hỏi trống → ô cùng dòng khớp rõ ràng nhất (phải vượt trội, không đoán khi mơ hồ)
        c = _empty_cell_candidates(sh, col, row, True)
        if c and c[0][0] >= 3 and (len(c) == 1 or c[0][0] - c[1][0] >= 2):
            return str(c[0][2]), f"empty->{c[0][1]}"
    return None

class _Done(Exception):
    pass
SYSTEM_PROMPT = """You answer ONE question about an Excel workbook by writing a SHORT Python snippet. The snippet is executed; its variable `answer` is the final answer. Never compute in your head. Never translate labels.

WHAT YOU SEE
Text dump of every sheet: TITLE / UNITS, COLUMN HEADERS ("letter: path" with " > "), ROWS as "R<row> | A=val | B=val ...". Empty cells omitted. Merged cells expanded.
Header rows are shown both as header paths and as "HEADER ROWS" (raw cells by address). If the question names an exact cell (e.g. "cell D5"), a "CELL LOOKUP" block lists what that address really holds.
The message may also contain an "ANALYSIS" draft from a reasoning pass: use it as a hint, but verify it against the dump (the dump wins).

PYTHON API (do not import)
sh, SHEETS["name"], sh.v("B5"), sh.row(r), sh.col("B"),
sh.find_rows(text, col="A", exact=False), sh.find_cols(text),
num(x) → float|None (None for blank/-/n/a/#N/A/… — NEVER treat missing as 0),
to_date(x), modules: math, re, statistics, dt.
sh.raw("B5") → original content of the cell exactly as stored (a formula keeps its "=..." text; empty → None),
sh.formula("B5") → formula text exactly as written, or None if the cell is not a formula, sh.raw_row(r), sh.raw_col("B").

CELL-ADDRESS QUESTIONS ("what is in / which formula is in cell X")
- Read the exact address with sh.raw(); never answer from a different cell just because it looks similar.
- "Formula in cell X": real formula → sh.formula() exactly as written (keep "="); text cell → its COMPLETE text, never a header-path fragment or a part of it.
- If the addressed cell is EMPTY, use the CELL LOOKUP candidates: take the same-row/column cell whose content is what the question asks for (formula-like text; a label equal to the column letter such as "(d)") and return that cell's COMPLETE text. If nothing fits, answer with the empty cell's value (empty string).

MULTILINGUAL — sheet AND question may be EN/VI/ZH/JA/KO/FR
- Understand the question in any language; map its meaning to the sheet's labels even when the languages differ.
- Label answers: copy EXACT sheet string (keep its punctuation, e.g. abbreviation dots and commas inside names), do not translate.
- Format requests may be multilingual — obey them:
  "two decimal places" / "2 chữ số thập phân" / "保留2位小数" / "소수점 2자리" / "小数点以下2桁".

STRUCTURE
1. Match FULL header path + row label (not one ambiguous word like "Total").
2. Blocks: use the block the question names; else overall/TOTAL.
3. of which / trong đó / 其中 = parts of parent (do not double-count). Total rows are aggregates — use an existing Total, do not re-sum for the same quantity. For rank/max/min/count among categories, exclude Total/aggregate rows and columns.
4. Sums, counts, ranks, max/min over a group: use lowest-level columns, not aggregate siblings. AVERAGE / MEAN "across all X" (all levels, all categories, all years...) is different: take EVERY data cell on the axis the question names, exactly as the table lists them, subtotal columns/rows included (do not drop, merge or re-group cells by your own judgment), and take their simple arithmetic mean. Exclude a cell only if the question explicitly says to exclude totals/subtotals, or it is the single overall grand total of that whole axis. Never mix cells from different blocks or dimensions unless the question asks for it.
5. Orientation: answer may be column header text OR row label.
6. Missing values: ignore in sum/avg/min/max (`num(v) is not None`).
6b. If a sheet is split into "TABLE 1, TABLE 2 ...", each table has its own COLUMN HEADERS; the same column letter can mean different things per table. Use the table the question refers to; use R<row> numbers to stay inside that table's row range.

ANSWER FORMAT (Exact Match) — read the question carefully
7. Rounding is MANDATORY when asked:
   - Compute full precision, THEN format.
   - N decimal places → answer = f"{x:.Nf}" (show trailing zeros: 3 → "3.00").
   - Integer / số nguyên / 整数 → int, no decimal point.
   - Set `answer` to the formatted value, not a raw float.
   - EXCEPTION: if the message ends with an "OUTPUT FORMAT NOTE" (several roundings requested), follow that note instead.
8. Yes/No as the question requests. Labels: exact sheet string. EXCEPTION: if the question itself lists the items to output, write them exactly as worded in the QUESTION, not the sheet's longer label.
9. No commentary, no unit suffix unless required.
10. NEVER rescale units. Even if the title says "in thousands/millions", return the number exactly as stored in the cell, unless the question explicitly asks to convert.

Assign `answer`. Return exactly one ```python block, no explanation."""
REPAIR_PROMPT = """The previous code failed or produced an empty/invalid answer.
Problem: {problem}
Code:
```
{code}
```
Common causes: the label sits in a DIFFERENT sheet/table, is spelled differently, or is in another column than assumed. Search ALL sheets (SHEETS.items()), try shorter keywords, scan every column. Never set `answer` to an error message or placeholder text.
Fix it. Return exactly one ```python block, same rules."""
DIRECT_SYSTEM = ("You answer questions about an Excel sheet dump. Think privately, then output ONLY the final answer in "
                 "the exact format the question asks for. No explanation, no markdown.")
JUDGE_SYSTEM = """You are an arbitrator for Excel analysis answers. Several candidate answers exist for the same question; each comes with the Python code that produced it.
Vote counts are only a hint: a majority can still be wrong if every vote read the wrong column/row.
Cross-check against the sheet dump (each cell shows its column letter):
- Does each code use the cell/column/row whose label matches the question (multi-level headers, value type, units)?
- Is the filter / ranking / arithmetic logic correct for the question?
- Is the format correct?
- Never prefer an answer that rescales units (e.g. x1000 because the title says "in thousands") unless the question explicitly asks for it.
- Keep the majority answer unless you find a clear error in its column/row/logic/format.
Give a very short rationale (under 80 words), then a final line that is ONLY JSON: {"choice": <0-based candidate index>}"""
_reasoning_ok = {"v": REASONING_OFF}

class Budget:
    def __init__(self):
        self.tin = self.tout = 0
        self.cost = 0.0
        self.calls = 0
        self._lock = threading.Lock()
    def add(self, tin, tout, cost):
        with self._lock:
            self.tin += tin
            self.tout += tout
            self.cost += cost
            self.calls += 1

def llm_chat(messages, bud, deadline, temperature=0, max_tokens=None, think=False):
    key = os.environ.get("OPENROUTER_API_KEY") or os.environ.get("OPENAI_API_KEY") or ""
    if not key:
        raise RuntimeError("Thieu OPENROUTER_API_KEY")
    last = None
    tok_limit = max_tokens or MAX_TOKENS
    for attempt in range(LLM_RETRIES):
        remaining = deadline - time.time()
        if remaining < 10:
            raise TimeoutError("het ngan sach thoi gian")
        if bud.cost >= MAX_COST_PER_Q:
            raise RuntimeError("vuot ngan sach chi phi cau hoi")
        body = {"model": MODEL, "messages": messages, "temperature": temperature, "max_tokens": tok_limit}
        if think:
            body["reasoning"] = {"enabled": True}
        elif _reasoning_ok["v"]:
            body["reasoning"] = {"enabled": False}
        try:
            r = requests.post(API_URL, headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
                              json=body, timeout=min(HTTP_TIMEOUT, remaining))
        except requests.RequestException as e:
            last = e
            time.sleep(1.5 * (attempt + 1))
            continue
        if r.status_code in (429, 502, 503):
            last = RuntimeError(f"HTTP {r.status_code}")
            time.sleep(2 * (attempt + 1))
            continue
        if r.status_code != 200:
            last = RuntimeError(f"HTTP {r.status_code}: {r.text[:300]}")
            time.sleep(1)
            continue
        try:
            j = r.json()
        except Exception as e:
            last = e
            continue
        usage = j.get("usage") or {}
        tin = int(usage.get("prompt_tokens") or 0)
        tout = int(usage.get("completion_tokens") or 0)
        cost = tin * PRICE_IN / 1e6 + tout * PRICE_OUT / 1e6
        bud.add(tin, tout, cost)
        msg = (j.get("choices") or [{}])[0].get("message") or {}
        text = msg.get("content") or ""
        return re.sub(r"<think>.*?</think>", "", text, flags=re.S).strip()
    raise RuntimeError(f"LLM loi sau nhieu lan thu: {last}")

def extract_code(text):
    """Lấy code Python: ưu tiên khối ```python; khối không gắn nhãn chỉ nhận nếu có `answer` (bỏ qua khối ```json)."""
    m = re.findall(r"```(?:python|py)[ \t]*\n(.*?)```", text, flags=re.S | re.I)
    if not m:
        m = [b for b in re.findall(r"```[ \t]*\n(.*?)```", text, flags=re.S) if "answer" in b]
    if m:
        return max(m, key=len).strip()
    return text.strip() if "answer" in text and "```" not in text else ""

# =============================================================================
# 6. GIẢI MỘT CÂU HỎI (bỏ phiếu + trọng tài)
# =============================================================================
def resolve_file(data_dir, fname):
    p = os.path.join(data_dir, fname)
    if os.path.exists(p):
        return p
    for root, _, files in os.walk(data_dir):
        if fname in files:
            return os.path.join(root, fname)
    raise FileNotFoundError(fname)
_FAIL_RE = re.compile(
    r"^\W*(?:(?:error|exception|not\s+found|n/a|cannot|can't|unable|no\s+data|insufficient"
    r"|missing\s+(?:data|value|label|row|column))\b"
    r"|lỗi|không\s+tìm\s+thấy|thiếu\s+dữ\s+liệu|错误|錯誤|无法|無法|未找到|エラー|見つかりません|오류)",
    re.I)

def _bad(ans, sheets=None):
    """Đáp án không dùng được: rỗng/NaN/nhiều dòng (văn xuôi), hoặc là chuỗi báo lỗi do mô hình tự viết.
    Ngoại lệ: chuỗi trông như báo lỗi (vd "N/A") nhưng đúng là giá trị có thật trong workbook thì được giữ."""
    if ans is None or (isinstance(ans, float) and math.isnan(ans)):
        return True
    t = str(ans).strip()
    if t in ("", "None", "nan") or "\n" in t:
        return True
    if REJECT_ERROR_ANSWERS and isinstance(ans, str) and _FAIL_RE.match(t):
        if sheets and _lab_key(t) in build_labels(sheets)[0]:
            return False
        return True
    return False

def vote_key(s):
    """Khoá so sánh đáp án: số → giá trị làm tròn 6 chữ số; list số; chữ → lower + gộp khoảng trắng."""
    s = str(s).strip()
    try:
        return ("n", round(float(s), 6))
    except ValueError:
        pass
    parts = s.split(",")
    if len(parts) > 1 and all(re.fullmatch(r"\s*[-+]?\d+(\.\d+)?\s*", p) for p in parts):
        return ("l", tuple(round(float(p), 6) for p in parts))
    return ("s", re.sub(r"\s+", " ", s.lower()))

def gen_candidate(user, sheets, query, bud, deadline, temperature, user_full=None, labels=None, fmt=None):
    """Một bản: sinh code → chạy → (nếu lỗi) sửa 1 lần. Trả dict {answer, code, path, err, raw}."""
    msgs = [{"role": "system", "content": SYSTEM_PROMPT}, {"role": "user", "content": user}]
    code, raw, problem = "", None, "no attempt"
    user_retry = user_full if (FULL_DUMP_ON_RETRY and user_full) else user  # lần repair dùng dump đầy đủ
    for attempt in range(2):
        text = llm_chat(msgs, bud, deadline, temperature)
        code = extract_code(text)
        raw = None
        if not code:
            problem = "No Python code block found."
        else:
            raw, err = run_code(code, sheets)
            problem = err if err else ("`answer` is empty, invalid, or an error message (a needed value was not found)" if _bad(raw, sheets) else None)
        if not problem:
            return {"answer": postprocess(raw, query, labels, fmt), "code": code, "raw": str(raw),
                    "path": "code" if attempt == 0 else "repair", "err": None}
        msgs = [{"role": "system", "content": SYSTEM_PROMPT}, {"role": "user", "content": user_retry},
                {"role": "assistant", "content": text},
                {"role": "user", "content": REPAIR_PROMPT.format(problem=problem, code=code or "(none)")}]
    return {"answer": "", "code": code, "raw": None if raw is None else str(raw), "path": "none", "err": problem}

def judge_votes(query, dump, cands, bud, deadline, trace=None):
    """Trọng tài: khi phiếu lệch, nhờ LLM chọn ứng viên đúng nhất dựa trên code + bảng."""
    seen, body_parts = set(), []
    unique = []
    for c in cands:
        if not c.get("answer"):
            continue
        k = vote_key(c["answer"])
        if k in seen:
            continue
        seen.add(k)
        unique.append(c)
        n = sum(1 for x in cands if x.get("answer") and vote_key(x["answer"]) == k)
        body_parts.append(
            f"--- Candidate {len(unique) - 1} | answer: {c['answer']!r} | {n} vote(s) ---\n"
            f"{c.get('code') or ''}"
        )
    if len(unique) < 2:
        return None
    body = "\n\n".join(body_parts)
    view = dump
    text = llm_chat(
        [{"role": "system", "content": JUDGE_SYSTEM},
         {"role": "user", "content": f"{view}\n\nQUESTION: {query}\n\n{body}"}],
        bud, deadline, temperature=0,
    )
    if trace is not None:
        trace["text"] = text
    ms = re.findall(r'"choice"\s*:\s*(\d+)', text)
    if ms and 0 <= int(ms[-1]) < len(unique):
        return unique[int(ms[-1])]["answer"]
    return None
ROUTE_SYSTEM = ("You select which sheets of an Excel workbook are needed to answer a question. You see a catalog "
                "(sheet name, size, column headers, sample rows). Include EVERY sheet that might be needed, including "
                "sheets needed for lookups/joins/units; when unsure, include more. Reply with ONLY JSON: "
                '{"sheets": ["<exact sheet name>", ...]}')

def _catalog(sheets):
    parts = []
    for name, s in sheets.items():
        t = s.tables[0] if s.tables else None
        hdr = list((t["headers"] if t else {}).values())[:14]
        samples = []
        for r in range(t["he"] if t else 0, s.nrows):
            cells = [f"{get_column_letter(c + 1)}={fmt_val(v)[:40]}" for c, v in enumerate(s.grid[r]) if v is not None]
            if cells:
                samples.append(f"R{r + 1} | " + " | ".join(cells[:8]))
            if len(samples) >= 2:
                break
        parts.append(f"- {name!r}: {s.nrows} rows x {s.ncols} cols"
                     + (f", {len(s.tables)} tables" if len(s.tables) > 1 else "")
                     + "\n  headers: " + " ; ".join(h[:80] for h in hdr)
                     + "\n  " + "\n  ".join(samples))
    return "\n".join(parts)

def route_sheets(query, sheets, bud, deadline):
    """Chọn các sheet liên quan. Trả list tên sheet (giữ thứ tự gốc) hoặc None nếu lỗi/rỗng/không giảm được gì."""
    try:
        text = llm_chat([{"role": "system", "content": ROUTE_SYSTEM},
                         {"role": "user", "content": f"CATALOG:\n{_catalog(sheets)}\n\nQUESTION: {query}"}],
                        bud, deadline, temperature=0)
        m = re.search(r"\{.*\}", text, flags=re.S)
        names = json.loads(m.group(0)).get("sheets") if m else None
        if not isinstance(names, list):
            return None
        lower = {k.lower().strip(): k for k in sheets}
        picked = {lower[str(n).lower().strip()] for n in names if str(n).lower().strip() in lower}
        chosen = [k for k in sheets if k in picked]
        return chosen if chosen and len(chosen) < len(sheets) else None
    except Exception:
        return None

PLAN_SYSTEM = """You are the reasoning stage of an Excel question-answering pipeline. You see a text dump of a workbook and ONE question. Do NOT write code and do NOT compute the final number.
Think step by step, then output exactly two parts.

PART 1 - a short plan (max ~150 words, plain text, no code):
1) what exactly is asked; 2) which sheet/table, full header path, row labels and cell letters / R-numbers hold the data; 3) the aggregation and what to include or exclude (Total / aggregate rows, "of which" rows, missing values, blocks, units - never rescale units); 4) the answer type (label / number / date / yes-no / list).

PART 2 - a ```json block describing the OUTPUT FORMAT THE QUESTION ITSELF REQUESTS (not your preference):
```json
{"rounding_mode": "none", "decimals": null, "date_format": null, "extra": []}
```
- rounding_mode: "none" when the question asks for no rounding; "half_up" for plain rounding (round / làm tròn / 四舍五入 / 반올림); "up" for round up / ROUNDUP / làm tròn lên / 切り上げ / 올림; "down" for round down / truncate / ROUNDDOWN / làm tròn xuống / 切り捨て / 내림; "ceil" for ceiling; "floor" for floor; "half_even" for banker's rounding.
- decimals: integer number of decimal places (2 decimal places / 2 chữ số thập phân / 保留2位小数 / 소수점 2자리 -> 2). Integer / whole number / số nguyên / 整数 / 정수 -> 0. Nearest ten -> -1, nearest hundred -> -2. null when no rounding is requested.
- date_format: a Python strftime pattern ONLY when the question names a date format (dd/mm/yyyy -> "%d/%m/%Y", yyyy-mm-dd -> "%Y-%m-%d"); otherwise null.
- extra: when the question asks for SEVERAL outputs with their own rounding, put the first output in the top-level fields and the remaining ones, in order, as objects {"rounding_mode": ..., "decimals": ...}; otherwise [].
Never round or format numbers yourself in PART 1."""

def plan_step(query, dump, hint, bud, deadline):
    """Bước reasoning: 1 lượt LLM cho (kế hoạch phân tích, format spec JSON). Lỗi -> (\"\", None), không làm hỏng pipeline."""
    try:
        text = llm_chat([{"role": "system", "content": PLAN_SYSTEM},
                         {"role": "user", "content": f"{dump}\n\n{hint}QUESTION: {query}"}],
                        bud, deadline, temperature=0, max_tokens=PLAN_MAX_TOKENS, think=PLAN_THINK)
    except Exception as e:  # noqa
        log(f"plan_step loi: {type(e).__name__}: {e}")
        return "", None
    fmt = parse_fmt(text)
    cut = re.search(r"```\s*json|\{\s*\"rounding_mode\"", text)
    analysis = (text[:cut.start()] if cut else text).replace("PART 1", "").strip(" \n:-")
    return analysis[:PLAN_MAX_CHARS].strip(), fmt

def solve_one(q, data_dir, deadline):
    t0 = time.time()
    qid, fname = q["id"], q.get("file_name") or q.get("file name")
    query = q["query"]
    rec = {"id": qid, "file": fname, "query": query, "path": "none", "answer": "", "raw": None, "code": "", "err": None,
           "votes": [], "judge": None}
    bud = Budget()
    try:
        sheets, dump = load_workbook_sheets(resolve_file(data_dir, fname))
        if not sheets:
            raise RuntimeError("file khong co du lieu")
        dc = try_direct_cell(query, sheets)
        if dc is not None:  # đọc thẳng công thức của đúng ô được hỏi, không qua LLM
            rec.update(answer=dc[0], raw=dc[0], path="cell:" + dc[1], err=None)
            raise _Done()
        full_dump = dump  # bản đầy đủ, dùng lại khi repair/direct
        labels = build_labels(sheets)  # nhãn gốc của mọi ô chữ, dùng để giữ nguyên dấu chấm/phẩy khi chuẩn hoá đáp án
        if ROUTING_ON and len(sheets) > 1 and len(dump) > ROUTE_MIN_CHARS:
            chosen = route_sheets(query, sheets, bud, deadline)
            if chosen:
                routed = "\n\n".join(sheets[k].dump() for k in chosen)
                if len(routed) < ROUTE_SHRINK_RATIO * len(dump):  # chỉ dùng khi thật sự giảm đáng kể
                    dump = routed
                    sheets = {**{k: sheets[k] for k in chosen}, **{k: v for k, v in sheets.items() if k not in chosen}}
                    rec["routed"] = chosen
        hint = cell_lookup_text(query, sheets)
        fmt, analysis = None, ""
        if PLAN_ON and time.time() < deadline - 30:  # bước reasoning: kế hoạch phân tích + format spec JSON
            analysis, fmt = plan_step(query, f"FILE: {fname}\n\n{dump}", hint, bud, deadline)
        rec["plan"] = analysis
        if fmt:
            rec["fmt"] = {"specs": [list(x) for x in fmt["specs"]], "date_format": fmt["date_format"]}
            mm = fmt_mismatch(fmt, query)
            if mm:
                rec["fmt_mismatch"] = mm  # JSON khác regex: xem lại trong debug_log
        extra = ""
        if analysis:
            extra += ("\n\nANALYSIS (draft from a prior reasoning pass; verify it against the dump and trust the dump "
                      "if they disagree):\n" + analysis)
        extra += format_note(fmt) if fmt else multi_round_hint(query)  # không có JSON hợp lệ -> hint dựa trên regex
        user = f"FILE: {fname}\n\n{dump}\n\n{hint}QUESTION: {query}" + extra
        user_full = (f"FILE: {fname}\n\n{full_dump}\n\n{hint}QUESTION: {query}" + extra) if dump is not full_dump else None
        cands, need = [], VOTES // 2 + 1
        def _try(i):
            try:
                return gen_candidate(user, sheets, query, bud, deadline, 0 if i == 0 else VOTE_TEMP, user_full, labels, fmt), None
            except Exception as e:  # noqa
                return None, e
        def _agreed():
            counts = Counter(vote_key(x["answer"]) for x in cands if x["answer"])
            return bool(counts) and counts.most_common(1)[0][1] >= need
        first = min(VOTES, 2) if PARALLEL_VOTES else 1
        with ThreadPoolExecutor(max_workers=first) as vex:
            first_res = [f.result() for f in [vex.submit(_try, i) for i in range(first)]]
        for c, e in first_res:
            if e is not None:
                rec["err"] = f"{type(e).__name__}: {e}"
                continue
            cands.append(c)
            rec["votes"].append(c["answer"])
        if not cands:
            raise RuntimeError(rec["err"] or "khong co phieu nao")
        i = first
        while i < VOTES * VOTE_ATTEMPT_FACTOR:
            if _agreed():
                break
            if i >= VOTES and sum(1 for x in cands if x["answer"]) >= VOTES:
                break
            c, e = _try(i)
            if e is not None:
                rec["err"] = f"{type(e).__name__}: {e}"
                break
            cands.append(c)
            rec["votes"].append(c["answer"])
            i += 1
        rec["cand_errs"] = [x.get("err") for x in cands if not x["answer"]]  # lý do phiếu rỗng, xem trong debug_log
        if LOG_CANDS:
            rec["cands"] = [{"answer": x["answer"], "raw": x.get("raw"), "err": x.get("err"),
                             "code": x.get("code") or ""} for x in cands]
        valid = [x for x in cands if x["answer"]]
        if valid:
            counts = Counter(vote_key(x["answer"]) for x in valid)
            top_key = counts.most_common(1)[0][0]
            best = next(x for x in valid if vote_key(x["answer"]) == top_key)
            final = best["answer"]
            path = f"vote{len(cands)}"
            n_distinct = len({vote_key(x["answer"]) for x in valid})
            if JUDGE_ON and n_distinct >= 2 and (JUDGE_OVERRIDE or counts.most_common(1)[0][1] < need) and time.time() < deadline - JUDGE_TIME_MARGIN and bud.cost < MAX_COST_PER_Q * JUDGE_COST_RATIO:
                try:
                    jt = {}
                    picked = judge_votes(query, dump, cands, bud, deadline, jt)
                    if picked:
                        rec["judge"] = {"from": final, "to": picked, "why": jt.get("text")}
                        final = picked
                        path = f"judge{len(cands)}"
                except Exception as e:
                    rec["judge"] = {"error": f"{type(e).__name__}: {e}"}
            final = postprocess(final, query, labels, fmt)  # ép format lần cuối (formatter xác định)
            rec.update(answer=final, code=best["code"], raw=best["raw"], path=path, err=None)
        else:
            if cands:
                rec.update(code=cands[-1]["code"], err=cands[-1]["err"])
            text = llm_chat([{"role": "system", "content": DIRECT_SYSTEM}, {"role": "user", "content": (user_full if FULL_DUMP_ON_RETRY and user_full else user)}],
                            bud, deadline)
            raw = re.sub(r"^```\w*|```$", "", text.strip()).strip().splitlines()[-1] if text.strip() else None
            rec.update(answer=postprocess(raw, query, labels, fmt), path="direct")
    except _Done:
        pass
    except Exception as e:
        rec["err"] = f"{type(e).__name__}: {e}"
        log(f"[{qid}] LOI: {rec['err']}")
    rec.update(tok_in=bud.tin, tok_out=bud.tout, cost=round(bud.cost, 6), calls=bud.calls, secs=round(time.time() - t0, 1))
    log(f"[{qid}] {rec['path']:8s} {rec['secs']:5.1f}s tok={bud.tin}/{bud.tout} ${bud.cost:.5f} votes={rec['votes']} -> {rec['answer']!r}")
    return rec

def write_submission(path, ids, answers):
    tmp = path + ".tmp"
    with open(tmp, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["id", "answer"])
        for i in ids:
            w.writerow([i, answers.get(i, "")])
    os.replace(tmp, path)

def find_queries(data_dir):
    p = os.path.join(data_dir, "queries.json")
    if os.path.exists(p):
        return p
    for root, _, files in os.walk(data_dir):
        if "queries.json" in files:
            return os.path.join(root, "queries.json")
    raise FileNotFoundError("khong tim thay queries.json")

def run(data_dir, output, queries_path=None):
    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)
    qpath = queries_path if queries_path else find_queries(data_dir)
    queries = json.load(open(qpath, encoding="utf-8"))
    ids = [q["id"] for q in queries]
    answers, records = {}, []
    write_submission(output, ids, answers)
    deadline = START + TIME_BUDGET
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = [ex.submit(solve_one, q, data_dir, deadline) for q in queries]
        for f in futs:
            try:
                rec = f.result()
                answers[rec["id"]] = rec["answer"]
                records.append(rec)
            except Exception as e:
                log(f"loi khong mong doi: {e}\n{traceback.format_exc()}")
            write_submission(output, ids, answers)
    write_submission(output, ids, answers)
    try:
        if DEBUG_LOG:
            with open(DEBUG_LOG, "w", encoding="utf-8") as f:
                for r in records:
                    f.write(json.dumps(r, ensure_ascii=False) + "\n")
    except Exception:
        pass
    tin = sum(r.get("tok_in", 0) for r in records)
    tout = sum(r.get("tok_out", 0) for r in records)
    n = max(len(records), 1)
    log(f"Xong {len(answers)}/{len(ids)} cau trong {time.time() - START:.1f}s | token TB/cau: vao {tin // n}, ra {tout // n} | tong ${sum(r.get('cost', 0) for r in records):.4f}")

def main():
    global MODEL, API_URL
    ap = argparse.ArgumentParser()
    # 6 tham số mà hệ thống chấm truyền vào
    ap.add_argument("--xlsx-dir", help="thư mục chứa các file .xlsx")
    ap.add_argument("--queries", help="đường dẫn tới queries.json")
    ap.add_argument("--output", "-o", default="submission.csv", help="đường dẫn file submission.csv")
    ap.add_argument("--openrouter-key", help="OpenRouter API key")
    ap.add_argument("--openrouter-model", help="tên model, ví dụ qwen/qwen3.7-flash")
    ap.add_argument("--openrouter-base-url", help="base URL, ví dụ https://openrouter.ai/api/v1")
    ap.add_argument("--input", "-i", help="(cũ) thư mục chứa queries.json và xlsx")
    ap.add_argument("--api-key", help="(cũ) OpenRouter API key")
    ap.add_argument("--dump", help="in ra dump của một file xlsx rồi thoát")
    a, unknown = ap.parse_known_args()
    if unknown:
        log(f"Bo qua tham so khong nhan dien: {unknown}")
    key = a.openrouter_key or a.api_key
    if key:
        os.environ["OPENROUTER_API_KEY"] = key
    if a.openrouter_model:
        MODEL = a.openrouter_model
    if a.openrouter_base_url:
        base = a.openrouter_base_url.rstrip("/")
        API_URL = base if base.endswith("/chat/completions") else base + "/chat/completions"
    if a.dump:
        print(load_workbook_sheets(a.dump)[1])
        return
    data_dir = a.xlsx_dir or a.input
    if not data_dir:
        log("LOI: thieu --xlsx-dir (hoac --input)")
        sys.exit(1)
    try:
        run(data_dir, a.output, a.queries)
    except (RuntimeError, FileNotFoundError) as e:
        log(f"LOI: {e}")
        sys.exit(1)
    except Exception as e:
        log(f"LOI khong mong doi: {e}\n{traceback.format_exc()}")
        sys.exit(1)
if __name__ == "__main__":
    main()