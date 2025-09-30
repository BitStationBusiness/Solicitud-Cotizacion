# telele.py
# Solicitud de Cotización Proveedores — GUI PyQt6 + Bot Telegram (PTB v21) con Gemini y PDF robusto
# -----------------------------------------------------------------------------

import os, sys, re, json, tempfile, asyncio, smtplib, traceback, requests, ctypes, subprocess, logging
from dataclasses import dataclass
from typing import Optional, Dict, List, Tuple
from datetime import datetime, timedelta  # <-- añadido timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, ROUND_DOWN
import unicodedata
import base64

# --- Logging a consola --------------------------------------------------------
def setup_logging():
    handler = logging.StreamHandler(sys.stdout)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
    handler.setFormatter(fmt)
    root = logging.getLogger()
    root.handlers.clear()
    root.addHandler(handler)
    root.setLevel(logging.DEBUG)
    logging.getLogger("httpx").setLevel(logging.WARNING)
    logging.getLogger("telegram").setLevel(logging.INFO)
    logging.getLogger("google").setLevel(logging.INFO)
    logging.getLogger("asyncio").setLevel(logging.INFO)

setup_logging()
log = logging.getLogger(__name__)

# --- Qt
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QMarginsF, QTimer
from PyQt6.QtGui import QPdfWriter, QTextDocument, QPageSize, QPageLayout, QIcon
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QTabWidget, QGroupBox, QFormLayout, QTextBrowser,
    QSpinBox, QCheckBox, QFileDialog, QMessageBox, QListWidget, QInputDialog,
    QSizePolicy, QPushButton
)

# --- Telegram (PTB v21)
from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import (
    Application, ApplicationBuilder, ContextTypes,
    CommandHandler, MessageHandler, filters
)

# --- IA (Gemini)
import google.generativeai as genai
from google.generativeai.types import GenerationConfig

# --- PDF (PyMuPDF)
import fitz

APP_ORG = "BitStation"
APP_NAME = "Solicitud de Cotización Proveedores"
SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_settings.json")

# --- Auto-update / Autostart
GITHUB_USER = "BitStationBusiness"
GITHUB_REPO = "Solicitud-Cotizacion"
AUTOSTART_REG_NAME = "BitStation_TB_Solicitud-Cotizacion"
AUTOSTART_VBS_NAME = "BitStation_TB_Solicitud-Cotizacion.vbs"

def _resource_path(*parts) -> str:
    return os.path.join(os.getcwd(), *parts)

def _version_file() -> str:
    return _resource_path("version.txt")

def get_current_version() -> str:
    vf = _version_file()
    if not os.path.exists(vf):
        with open(vf, "w", encoding="utf-8") as f:
            f.write("0.1")
        return "0.1"
    try:
        with open(vf, "r", encoding="utf-8") as f:
            v = (f.read() or "").strip()
            return v or "0.1"
    except Exception:
        return "0.1"

APP_VERSION = get_current_version()

# ------------------------- Settings JSON -------------------------
class JsonSettings:
    def __init__(self, path: str):
        self.path = path
        self._data: Dict = {}
        self.ensure()

    def ensure(self):
        if not os.path.exists(self.path):
            defaults = {
                "telegram_token": "",
                "owner_name": "",
                "google_api_key": "",
                "smtp": {"server": "smtp.gmail.com", "port": 587, "email": "", "password": ""},
                "pdf": {"margin_mm": 15},
                "pdf_dir": "",
                "pricing_detection": "auto",
                # --- NUEVO: datos de identidad/empresa + validez + ID secuencial
                "company": {
                    "address": "",
                    "cif": "",
                    "phone": "",
                    "email": "",
                    "web": ""
                },
                "validity_days": 15,
                "doc": {"last_id": 0},
                # --- control de acceso
                "access": {"mode": "off", "whitelist": [], "blacklist": []}
            }
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(defaults, f, indent=2, ensure_ascii=False)
        self.load()

    def load(self):
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                self._data = json.load(f)
        except Exception:
            self._data = {}
        self._data.setdefault("access", {"mode": "off", "whitelist": [], "blacklist": []})
        self._data["access"].setdefault("mode", "off")
        self._data["access"].setdefault("whitelist", [])
        self._data["access"].setdefault("blacklist", [])
        self._data.setdefault("pdf", {"margin_mm": 15})
        self._data.setdefault("pdf_dir", "")
        self._data.setdefault("pricing_detection", "auto")
        # --- NUEVO: asegurar claves de identidad/validez/ID
        self._data.setdefault("company", {"address": "", "cif": "", "phone": "", "email": "", "web": ""})
        self._data.setdefault("validity_days", 15)
        self._data.setdefault("doc", {"last_id": 0})

    def save(self):
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(self._data, f, indent=2, ensure_ascii=False)

    def get(self, k, d=None):
        return self._data.get(k, d)

    def set(self, k, v):
        self._data[k] = v
        self.save()

    def get_smtp(self) -> Dict:
        return self._data.get("smtp", {})

    def set_smtp(self, k, v):
        smtp = self._data.setdefault("smtp", {})
        smtp[k] = v
        self.save()

    def pdf_margin(self) -> int:
        return int(self._data.get("pdf", {}).get("margin_mm", 15))

    def set_pdf_margin(self, mm: int):
        self._data.setdefault("pdf", {})["margin_mm"] = int(mm)
        self.save()

    def pdf_dir(self) -> str:
        return self._data.get("pdf_dir") or os.path.join(_project_root(), "PDF")

    def set_pdf_dir(self, d: str):
        self._data["pdf_dir"] = d
        self.save()

    # ---- helpers de acceso ----
    def access_mode(self) -> str:
        return (self._data.get("access", {}).get("mode") or "off").lower()

    def set_access_mode(self, mode: str):
        self._data.setdefault("access", {})["mode"] = mode
        self.save()

    def get_whitelist(self) -> List[int]:
        lst = self._data.get("access", {}).get("whitelist", [])
        out: List[int] = []
        for x in lst:
            s = str(x).strip()
            if re.fullmatch(r"-?\d+", s):
                out.append(int(s))
        return out

    def get_blacklist(self) -> List[int]:
        lst = self._data.get("access", {}).get("blacklist", [])
        out: List[int] = []
        for x in lst:
            s = str(x).strip()
            if re.fullmatch(r"-?\d+", s):
                out.append(int(s))
        return out

    def add_to_list(self, which: str, user_id: int) -> bool:
        acc = self._data.setdefault("access", {})
        lst = acc.setdefault(which, [])
        normalized: List[int] = []
        for x in lst:
            s = str(x).strip()
            if re.fullmatch(r"-?\d+", s):
                v = int(s)
                if v not in normalized:
                    normalized.append(v)
        uid = int(user_id)
        if uid in normalized:
            acc[which] = normalized
            self.save()
            return False
        normalized.append(uid)
        acc[which] = normalized
        self.save()
        return True

    def remove_from_list(self, which: str, user_id: int):
        acc = self._data.setdefault("access", {})
        lst = acc.setdefault(which, [])
        normalized: List[int] = []
        for x in lst:
            s = str(x).strip()
            if re.fullmatch(r"-?\d+", s):
                v = int(s)
                if v not in normalized:
                    normalized.append(v)
        try:
            normalized.remove(int(user_id))
        except ValueError:
            pass
        acc[which] = normalized
        self.save()

    # ---- NUEVO: helpers identidad / validez / IDs ----
    def get_company(self) -> Dict[str, str]:
        return self._data.get("company", {})

    def set_company(self, k: str, v: str):
        comp = self._data.setdefault("company", {})
        comp[k] = v
        self.save()

    def validity_days(self) -> int:
        try:
            return int(self._data.get("validity_days", 15))
        except Exception:
            return 15

    def set_validity_days(self, n: int):
        self._data["validity_days"] = int(n)
        self.save()

    def next_doc_id(self) -> int:
        doc = self._data.setdefault("doc", {})
        last = int(doc.get("last_id", 0))
        new = last + 1
        doc["last_id"] = new
        self.save()
        return new

SET = JsonSettings(SETTINGS_FILE)

# ------------------------- Utilidades -------------------------
def human_ex(e: Exception) -> str:
    return f"{type(e).__name__}: {e}"

def html_escape(s: str) -> str:
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def _project_root() -> str:
    return os.path.dirname(os.path.abspath(__file__))

def _tmp_dir() -> str:
    d = os.path.join(tempfile.gettempdir(), "telequote_tmp")
    os.makedirs(d, exist_ok=True)
    return d

def _safe_unlink(path: Optional[str]) -> None:
    try:
        if path and os.path.exists(path):
            os.remove(path)
    except Exception:
        pass

# --- NUEVO: suma de días hábiles (lunes-viernes) ---
def _add_business_days(start_dt: datetime, days: int) -> datetime:
    d = start_dt
    remaining = max(0, int(days))
    while remaining > 0:
        d += timedelta(days=1)
        if d.weekday() < 5:  # 0=Lunes … 4=Viernes
            remaining -= 1
    return d

# ---------- Helpers de normalización/regex ----------
def _normalize_number(num_text: str) -> Optional[Decimal]:
    t = (num_text or "").strip()
    t = re.sub(r"[ €\t\r\n]", "", t, flags=re.I)
    t = re.sub(r"(eur|euro|euros)$", "", t, flags=re.I)
    if not t:
        return None
    if "." in t and "," in t:
        if t.rfind(".") > t.rfind(","):
            t = t.replace(",", "")
        else:
            t = t.replace(".", "").replace(",", ".")
    else:
        t = t.replace(",", ".")
    try:
        return Decimal(t)
    except InvalidOperation:
        return None

def _format_eur(dec: Decimal) -> str:
    v = dec.quantize(Decimal("1")) if dec == dec.to_integral_value() else dec.normalize()
    txt = format(v, "f").rstrip("0").rstrip(".")
    return f"{txt}€"

# ---- Detección y formateo que IMITA el estilo del usuario ----
def _analyze_number_style(example: str) -> dict:
    s = example or ""
    m = re.search(r'([+-]?\d[\d.,]*)', s)
    if not m:
        return {"prefix": "", "suffix": "€", "dec_sep": None, "thou_sep": "", "decimals": 0}
    num = m.group(1)
    prefix = s[:m.start(1)]
    suffix = s[m.end(1):]

    has_dot = "." in num
    has_com = "," in num
    dec_sep = None
    decimals = 0
    thou_sep = ""

    if has_dot and has_com:
        dec_sep = "." if num.rfind(".") > num.rfind(",") else ","
        other = "," if dec_sep == "." else "."
        thou_sep = other if other in num else ""
        decimals = len(num.split(dec_sep)[-1])
    elif has_com:
        parts = num.split(",")
        tail = parts[-1]
        if len(parts) == 2 and len(tail) == 3 and parts[0].isdigit():
            dec_sep = None
            thou_sep = ","
            decimals = 0
        else:
            dec_sep = ","
            decimals = len(tail)
    elif has_dot:
        parts = num.split(".")
        tail = parts[-1]
        if len(parts) == 2 and len(tail) == 3 and parts[0].isdigit():
            dec_sep = None
            thou_sep = "."
            decimals = 0
        else:
            dec_sep = "."
            decimals = len(tail)

    return {"prefix": prefix, "suffix": suffix, "dec_sep": dec_sep, "thou_sep": thou_sep, "decimals": decimals}

def _group_thousands(int_str: str, sep: str) -> str:
    if not sep:
        return int_str
    s = int_str
    out = []
    while s:
        out.append(s[-3:])
        s = s[:-3]
    return sep.join(reversed(out))

def _format_number_like(val: Decimal, example_text: Optional[str]) -> str:
    style = _analyze_number_style(example_text or "")
    sign = "-" if val < 0 else ""
    v = abs(val)
    if style["dec_sep"] is None:
        q = v.quantize(Decimal("1"))
        int_str = str(q)
        num = _group_thousands(int_str, style["thou_sep"])
    else:
        decimals = max(0, int(style["decimals"]))
        q = v.quantize(Decimal("1").scaleb(-decimals))
        f = format(q, "f")
        int_part, frac = f.split(".")
        int_part = _group_thousands(int_part, style["thou_sep"])
        num = int_part + style["dec_sep"] + frac[:decimals].ljust(decimals, "0")
    return f"{style['prefix']}{sign}{num}{style['suffix'] or '€'}"

def _strip_qty_prefix(detail: str, qty: Decimal) -> str:
    if not detail or qty is None:
        return detail
    if qty == qty.to_integral_value():
        base = str(int(qty))
        num_pat = re.escape(base)
    else:
        base = format(qty.normalize(), "f").rstrip("0").rstrip(".")
        num_pat = re.escape(base).replace(r"\.", r"[.,]")
    pre_words = r"(?:unid(?:ad(?:es)?)?|cant(?:idad)?|ud?s?\.?|unds?\.?)?"
    post_unit = r"(?:\s*[a-záéíóúüñ]{1,20}\.?)?"
    pat = rf"^\s*(?:{pre_words}\s*)?(?:x\s*)?{num_pat}{post_unit}(?:\s*(?:de|:|-))?\s+"
    new = re.sub(pat, "", detail, flags=re.I)
    return new.strip() or detail.strip()

def _strip_qty_and_capture_unit(detail: str, qty: Decimal) -> Tuple[str, Optional[str]]:
    if not detail or qty is None:
        return detail, None
    if qty == qty.to_integral_value():
        base = str(int(qty))
        num_pat = re.escape(base)
    else:
        base = format(qty.normalize(), "f").rstrip("0").rstrip(".")
        num_pat = re.escape(base).replace(r"\.", r"[.,]")
    pre_words = r"(?:unid(?:ad(?:es)?)?|cant(?:idad)?|ud?s?\.?|unds?\.?)?"
    unit_group = r"(?P<unit>[a-záéíóúüñ]{1,20}\.?)"
    pat = rf"^\s*(?:{pre_words}\s*)?(?:x\s*)?{num_pat}(?:\s*{unit_group})?(?:\s*(?:de|:|-))?\s+"
    m = re.match(pat, detail, flags=re.I)
    if not m:
        return _strip_qty_prefix(detail, qty), None
    unit = (m.group("unit") or "").strip()
    new_detail = detail[m.end():].strip()
    return (new_detail, unit if unit else None)

_UNIT_WORDS = [
    "resma", "resmas", "caja", "cajas", "rollo", "rollos", "paquete", "paquetes",
    "bulto", "bultos", "bolsa", "bolsas", "pieza", "piezas", "pza", "pzas", "pz",
    "pzs", "unidad", "unidades", "ud", "uds", "unid", "unids", "pack", "packs"
]
_UNIT_RE = re.compile(rf"^\s*(?:(?:de)\s+)?(?P<unit>{'|'.join(_UNIT_WORDS)})\b[ \t]*[;,:•\-–—]*\s*", re.I)

def _extract_leading_unit_if_any(detail: str) -> Tuple[str, Optional[str]]:
    if not detail:
        return detail, None
    m = _UNIT_RE.match(detail)
    if not m:
        return detail, None
    unit = (m.group("unit") or "").strip()
    rest = detail[m.end():].strip()
    return rest, unit

def _smart_title_unit(unit: str) -> str:
    u = (unit or "").strip()
    if re.fullmatch(r"[a-záéíóúüñ]+\.?", u, re.I):
        return u[:1].upper() + u[1:].lower()
    return u

# ---- Mover color del final del Ítem a Detalle ----
_COLOR_CANON = {
    "blanco": "Blanco", "blanca": "Blanco", "blancos": "Blanco", "blancas": "Blanco",
    "negro": "Negro", "negra": "Negro", "negros": "Negro", "negras": "Negro",
    "rojo": "Rojo", "roja": "Rojo", "rojos": "Rojo", "rojas": "Rojo",
    "amarillo": "Amarillo", "amarilla": "Amarillo", "amarillos": "Amarillo", "amarillas": "Amarillo",
    "morado": "Morado", "morada": "Morado", "morados": "Morado", "moradas": "Morado",
    "dorado": "Dorado", "dorada": "Dorado", "dorados": "Dorado", "doradas": "Dorado",
    "azul": "Azul", "azules": "Azul",
    "verde": "Verde", "verdes": "Verde",
    "marrón": "Marrón", "marron": "Marrón", "marrones": "Marrón",
    "gris": "Gris", "grises": "Gris",
    "rosa": "Rosa", "rosas": "Rosa",
    "lila": "Lila", "lilas": "Lila",
    "violeta": "Violeta", "violetas": "Violeta",
    "plata": "Plata",
    "beige": "Beige",
    "crema": "Crema",
    "natural": "Natural", "naturales": "Natural",
}
_COLOR_RE = re.compile(r"(?:\bcolor\s+)?\b(" + "|".join([re.escape(k) for k in _COLOR_CANON.keys()]) + r")\b\.?$", re.I)

def _move_color_to_detail(item_text: str, detalle_text: str) -> Tuple[str, str]:
    t = (item_text or "").strip()
    m = _COLOR_RE.search(t)
    if not m:
        return item_text, detalle_text
    raw = (m.group(1) or "").lower()
    color_canon = _COLOR_CANON.get(raw, raw.capitalize())
    new_item = t[:m.start()].strip(" -·,")
    if not re.search(r"\bcolor\b", detalle_text or "", re.I) and not _COLOR_RE.search(detalle_text or ""):
        det = ("; " if detalle_text else "") + f"Color {color_canon}"
        return new_item or item_text, (detalle_text + det).strip("; ").strip()
    return new_item or item_text, (detalle_text or "").strip()

def _strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

def _same_token(a: str, b: str) -> bool:
    na = re.sub(r"[^a-z0-9]", "", _strip_accents(a.lower()))
    nb = re.sub(r"[^a-z0-9]", "", _strip_accents(b.lower()))
    return na == nb or na.rstrip("s") == nb or nb.rstrip("s") == na

def _split_item_detail(item_text: str) -> Tuple[str, Optional[str]]:
    t = item_text.strip()
    m = re.search(r"\b((?:\d+(?:[.,]\d+)?\s*(?:x|×)\s*){1,2}\d+(?:[.,]\d+)?\s*(?:mm|cm|m)\b[^\s|]*)\s*$", t, flags=re.I)
    if not m:
        m = re.search(r"\(([^)]*?\d[^)]*?)\)\s*$", t)
    if m:
        spec = m.group(1).strip()
        name = t[:m.start()].strip(" -·,")
        if name:
            return name, spec
    return t, None

def _normalize_inline_tables(md: str) -> str:
    if not md:
        return md
    import re
    text = md.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(
        r"(:)\s*(?=(?:[ÍI]tem)\s*\|\s*Cantidad\s*\|\s*Detalle\s*\|\s*Precio\s*unitario\b)",
        r":\n\n",
        text,
        flags=re.I,
    )
    def _rebuild_table_from_inline(inline: str) -> tuple[str, int]:
        import unicodedata
        def canon(s: str) -> str:
            s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
            return re.sub(r"\s+", " ", s).strip().lower()
        m_hdr = re.search(r"([ÍI]tem\s*\|\s*Cantidad\s*\|\s*Detalle\s*\|\s*Precio\s*unitario)(?:\s*\|\s*Total)?", inline, flags=re.I)
        if not m_hdr:
            return ("", 0)
        post = inline[m_hdr.end():]
        m_end = re.search(r"(?:\n\s*\n|\b(?:Agradecemos|Atentamente|Esperamos|Saludos|Cordialmente)\b)", post, flags=re.I)
        consumed = m_hdr.end() + (m_end.start() if m_end else len(post))
        seg = inline[:consumed]
        tokens = [t.strip() for t in seg.split("|")]
        tokens = [t for t in tokens if t != ""]
        tokens = [t for t in tokens if not re.fullmatch(r":?-{3,}:?", t)]
        idx_item = None
        for i in range(len(tokens) - 3):
            c0, c1, c2, c3 = map(canon, tokens[i:i+4])
            if (c0 in ("ítem", "item") and c1 == "cantidad" and c2 == "detalle" and c3 == "precio unitario"):
                idx_item = i
                break
        if idx_item is None:
            return ("", 0)
        five_cols = False
        if idx_item + 4 < len(tokens):
            c4 = canon(tokens[idx_item + 4])
            five_cols = (c4 == "total")
        col_count = 5 if five_cols else 4
        data = tokens[idx_item + col_count:]
        out_lines = [
            "| Ítem | Cantidad | Detalle | Precio unitario |",
            "|---|---|---|---|",
        ]
        END = re.compile(r"^(?:Agradecemos|Atentamente|Esperamos|Saludos|Cordialmente)\b", re.I)
        step = col_count
        for j in range(0, len(data), step):
            row = [c.strip() for c in data[j:j+step]]
            if len(row) < step:
                break
            if any(END.match(c) for c in row):
                break
            if step == 5:
                row = row[:4]
            if len(row) > 4:
                row = row[:3] + [" ".join(row[3:])]
            out_lines.append("| " + " | ".join(row) + " |")
        if len(out_lines) <= 2:
            return ("", 0)
        return ("\n".join(out_lines), consumed)

    header_pat = re.compile(r"[ÍI]tem\s*\|\s*Cantidad\s*\|\s*Detalle\s*\|\s*Precio\s*unitario", flags=re.I)
    pos = 0
    while True:
        m = header_pat.search(text, pos)
        if not m:
            break
        lookahead = text[m.end(): m.end()+80]
        if "\n" in lookahead and re.search(r"^\s*\|?\s*:?-{3,}:?\s*\|", lookahead, flags=re.M):
            pos = m.end()
            continue
        rebuilt, consumed = _rebuild_table_from_inline(text[m.start():])
        if rebuilt:
            pre = text[:m.start()].rstrip()
            post = text[m.start()+consumed:].lstrip()
            text = f"{pre}\n\n{rebuilt}\n\n{post}"
            pos = len(pre) + len(rebuilt) + 2
        else:
            pos = m.end()
    text = re.sub(r":\s*(\|)", r":\n\1", text)
    text = re.sub(r"\s*\|\|\s*", "\n| ", text)
    return text

def _format_sum_max2(val: Decimal, sample_texts: List[str]) -> str:
    base = _analyze_number_style(sample_texts[0] if sample_texts else "")
    if base["dec_sep"] is None:
        for s in sample_texts or []:
            st = _analyze_number_style(s or "")
            if st["dec_sep"] is not None:
                base["dec_sep"] = st["dec_sep"]
                base["thou_sep"] = st["thou_sep"]
                break
    q2 = val.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    txt = format(q2, "f").rstrip("0").rstrip(".")
    if "." in txt:
        int_part, frac = txt.split(".", 1)
    else:
        int_part, frac = txt, ""
    int_part = _group_thousands(int_part, base["thou_sep"])
    if frac:
        dec_sep = base["dec_sep"] or "."
        num = int_part + dec_sep + frac
    else:
        num = int_part
    return f"{base['prefix']}{num}{base['suffix'] or '€'}"

# --------- REGLAS DE NEGOCIO ----------
RESMA_EQUIV_UNIDADES = 500
PRICING_EXPLICIT_WORDS_UNIT = {"precio unitario", "€/ud", "€/unidad", "eur/ud", "por unidad", "pu", "unit price"}
PRICING_EXPLICIT_WORDS_TOTAL = {"precio total", "total línea", "total del ítem", "precio final"}

def _detect_decimal_style(sample: str) -> str:
    return "comma" if re.search(r"\d+,\d+\s*€", sample) else "dot"

def _parse_money(text: str) -> Optional[Decimal]:
    m = re.search(r"([0-9]{1,3}(?:[.\s][0-9]{3})*|[0-9]+)(?:[.,]([0-9]{1,3}))?\s*€?", text or "")
    if not m:
        return None
    entero = re.sub(r"[.\s]", "", m.group(1))
    dec = m.group(2) or "00"
    return Decimal(f"{entero}.{dec}")

def _format_money(value: Decimal, style: str, with_symbol: bool = True) -> str:
    q = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    s = f"{q:.2f}"
    int_part, frac = s.split(".")
    if frac == "00":
        out = int_part
    else:
        sep = "," if style == "comma" else "."
        out = f"{int_part}{sep}{frac}"
    return f"{out}€" if with_symbol else out

def _format_money_unit(value: Decimal, style: str, with_symbol: bool = True) -> str:
    q = value.quantize(Decimal("0.001"), rounding=ROUND_DOWN)
    s = format(q, "f").rstrip("0").rstrip(".")
    if "." in s:
        int_part, frac = s.split(".", 1)
        sep = "," if style == "comma" else "."
        out = f"{int_part}{sep}{frac}" if frac else int_part
    else:
        out = s
    return f"{out}€" if with_symbol else out

def _qty_to_units(cant_str: str) -> Optional[int]:
    if not cant_str:
        return None
    s = unicodedata.normalize("NFKC", cant_str).strip().lower()
    m = re.search(r"(\d+)\s*resma[s]?", s)
    if m:
        units = int(m.group(1)) * RESMA_EQUIV_UNIDADES
        m2 = re.search(r"y\s*(\d+)\s*unidade?s?", s)
        if m2:
            units += int(m2.group(1))
        return units
    m = re.search(r"(\d[\d\.\,\s\u00A0\u202F]*)", s)
    if m:
        raw = m.group(1)
        raw = re.sub(r"[\.\,\s\u00A0\u202F]", "", raw)
        if raw.isdigit():
            return int(raw)
    return None

def _is_service_item(item_name: str) -> bool:
    s = (item_name or "").strip().lower()
    return any(k in s for k in [
        "transporte", "envío", "porte", "portes", "montaje", "instalación",
        "mano de obra", "servicio", "mozo", "entrega especial"
    ])

def _classify_price_kind(qty_units: Optional[int], price: Decimal,
                         row_text: str, detection_mode: str = "auto") -> str:
    row_low = (row_text or "").lower()
    if any(w in row_low for w in PRICING_EXPLICIT_WORDS_UNIT):
        return "unit"
    if any(w in row_low for w in PRICING_EXPLICIT_WORDS_TOTAL):
        return "total"
    if detection_mode == "auto" and qty_units is not None:
        try:
            if price < Decimal(qty_units):
                return "unit"
        except Exception:
            pass
    return "total"

# --- Procesado de tablas Markdown ---
def process_markdown_tables(body_md: str, is_cotizacion: bool = False, pricing_mode: str = "auto") -> str:
    body_md = _normalize_inline_tables(body_md or "")
    lines = body_md.splitlines()
    out_lines: List[str] = []
    i = 0

    def split_row(row): return [c.strip() for c in row.strip().strip("|").split("|")]

    while i < len(lines):
        line = lines[i]
        if "|" in line and re.search(r"\|\s*[-:]+\s*\|", lines[i + 1] if i + 1 < len(lines) else ""):
            tbl = []
            while i < len(lines) and "|" in lines[i]:
                tbl.append(lines[i]); i += 1
            orig_header = split_row(tbl[0])
            joined_table_text = "\n".join(tbl)
            style = _detect_decimal_style(joined_table_text)
            if is_cotizacion:
                header = ["Ítem", "Cantidad", "Detalle", "Precio unitario"]
            else:
                header = split_row(tbl[0])
                if len(header) < 4:
                    header = ["Ítem", "Cantidad", "Detalle", "Total"]

            def idx_from_orig(col_name: str) -> Optional[int]:
                for k, h in enumerate(orig_header):
                    if unicodedata.normalize("NFKD", h).casefold() == unicodedata.normalize("NFKD", col_name).casefold():
                        return k
                return None

            i_item_o = idx_from_orig("Ítem") or 0
            i_cant_o = idx_from_orig("Cantidad") or 1
            i_det_o  = idx_from_orig("Detalle") or 2
            i_pu_o   = idx_from_orig("Precio unitario")
            i_tot_o  = idx_from_orig("Total")

            data_rows = [split_row(r) for r in tbl[2:]]
            norm_rows: List[List[str]] = []

            for raw in data_rows:
                first_cell = (raw[0] if raw else "").strip().lower()
                if any(k in first_cell for k in (
                    "precio total sin iva", "total a pagar",
                    "total general", "importe total", "gran total"
                )):
                    continue

                item = raw[i_item_o] if i_item_o is not None and i_item_o < len(raw) else (raw[0] if raw else "")
                cant = raw[i_cant_o] if i_cant_o is not None and i_cant_o < len(raw) else (raw[1] if len(raw) > 1 else "")
                det  = raw[i_det_o]  if i_det_o  is not None and i_det_o  < len(raw) else (raw[2] if len(raw) > 2 else "")

                name, spec = _split_item_detail(item)
                if spec and (not det or _same_token(det, name)):
                    det = spec
                    item = name

                qnum = _normalize_number(cant)
                if qnum is not None:
                    new_det, unit = _strip_qty_and_capture_unit(det, qnum)
                    det = new_det
                    if not unit:
                        new_det2, unit2 = _extract_leading_unit_if_any(det)
                        det = new_det2; unit = unit2
                    if unit and not re.search(rf"\b{re.escape(unit)}\b", cant, re.I):
                        cant = (cant.strip() + " " + _smart_title_unit(unit)).strip()

                item, det = _move_color_to_detail(item, det)

                if not (cant or "").strip() and _is_service_item(item):
                    cant = "1"

                pu_text_candidates: List[str] = []
                if i_pu_o is not None and i_pu_o < len(raw):
                    pu_text_candidates.append(raw[i_pu_o])
                if i_tot_o is not None and i_tot_o < len(raw):
                    pu_text_candidates.append(raw[i_tot_o])
                pu_text_candidates.append(det)
                pu_text_candidates.append(" ".join(raw))

                pu_val = None
                for txt in pu_text_candidates:
                    pu_val = _parse_money(txt)
                    if pu_val is not None:
                        break
                pu_str = _format_money(pu_val, style) if pu_val is not None else ""

                if is_cotizacion:
                    norm_rows.append([item, cant, det, pu_str])
                else:
                    tot_str = ""
                    if i_tot_o is not None and i_tot_o < len(raw):
                        d = _parse_money(raw[i_tot_o])
                        if d is not None:
                            tot_str = _format_money(d, style)
                    norm_rows.append([item, cant, det, tot_str])

            out_lines.append("| " + " | ".join(header) + " |")
            out_lines.append("| " + " | ".join(["---"] * len(header)) + " |")
            for r in norm_rows:
                out_lines.append("| " + " | ".join(r) + " |")
            continue

        out_lines.append(line)
        i += 1

    return "\n".join(out_lines)


def markdown_table_to_html(text: str) -> str:
    text = _normalize_inline_tables(text)
    lines = (text or "").splitlines()
    i = 0
    out: List[str] = []

    def par(ps: List[str]):
        for ln in ps:
            if not ln.strip():
                out.append("<p style='margin:0 0 12px'>&nbsp;</p>")
            else:
                out.append(f"<p style='margin:0 0 10px'>{html_escape(ln)}</p>")

    while i < len(lines):
        if lines[i].strip().startswith("|"):
            block = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                block.append(lines[i].strip())
                i += 1

            cleaned: List[List[str]] = []
            for raw in block:
                parts = [c.strip() for c in raw.split("|")]
                if parts and parts[0] == "": parts = parts[1:]
                if parts and parts[-1] == "": parts = parts[:-1]
                if parts and set("".join(parts)) <= set("-: "):
                    continue
                if parts and any(p.strip() for p in parts):
                    cleaned.append(parts)

            if cleaned:
                header = cleaned[0]
                ncols = len(header)
                spillovers: List[str] = []

                def normalize_row(row: List[str]) -> List[str]:
                    nonlocal spillovers
                    if len(row) > ncols:
                        extra = " ".join(cell for cell in row[ncols:] if cell.strip()).strip()
                        if extra:
                            spillovers.append(extra)
                        row = row[:ncols]
                    elif len(row) < ncols:
                        row = row + [""] * (ncols - len(row))
                    return row

                header = normalize_row(header)
                body_rows = [normalize_row(r) for r in cleaned[1:]]

                thead = "".join(
                    f"<th style='background:#f6f8fa;border:1px solid #dfe3e8;padding:6px 8px;text-align:left;font-weight:600'>{html_escape(c)}</th>"
                    for c in header
                )
                tbody = []
                for row in body_rows:
                    tds = "".join(
                        f"<td style='border:1px solid #e6e9ef;padding:6px 8px;vertical-align:top'>{html_escape(c)}</td>"
                        for c in row
                    )
                    tbody.append(f"<tr style='page-break-inside:avoid'>{tds}</tr>")

                out.append(
                    "<table style='border-collapse:collapse;width:100%;table-layout:fixed;margin:8px 0 12px;"
                    "font-family:Arial;font-size:11pt;color:#222'>"
                    f"<thead><tr>{thead}</tr></thead><tbody>{''.join(tbody)}</tbody></table>"
                )

                if spillovers:
                    par(spillovers)
            else:
                par(block)
        else:
            chunk: List[str] = []
            while i < len(lines) and (not lines[i].strip().startswith("|")):
                chunk.append(lines[i]); i += 1
            par(chunk)

    return "".join(out)

def _logo_data_uri() -> str:
    try:
        logo_path = os.path.join(os.getcwd(), "multimedia", "empresa.png")
        if not os.path.exists(logo_path):
            logging.getLogger("pdf").warning("Logo no encontrado: %s", logo_path)
            return ""
        with open(logo_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("ascii")
        return f"data:image/png;base64,{b64}"
    except Exception:
        logging.getLogger("pdf").exception("No se pudo leer el logo")
        return ""

# --- MODIFICADO: acepta doc_info para pintar "Rojo/Azul" ---
def compose_html(subject: str, body_md: str, is_cotizacion: bool = False, doc_info: Optional[Dict[str, str]] = None) -> str:
    body_html = markdown_table_to_html(body_md)

    logo_uri = _logo_data_uri()
    logo_html = f"<img src='{logo_uri}' alt='Logo' style='width:5cm;height:auto;display:block'/>" if logo_uri else ""

    # Datos empresa desde configuración
    comp = SET.get_company()
    def line(lbl, key):
        val = (comp.get(key) or "").strip()
        return f"<div><strong>{lbl}:</strong> {html_escape(val)}</div>" if val else ""
    company_block = "".join([
        line("Dirección", "address"),
        line("CIF", "cif"),
        line("Teléfono", "phone"),
        line("eMail", "email"),
        line("Web", "web"),
    ])

    # Datos de documento (ID/Hora/Fecha/Validez)
    _id = html_escape((doc_info or {}).get("id", ""))
    _time = html_escape((doc_info or {}).get("time", ""))
    _date = html_escape((doc_info or {}).get("date", ""))
    _valid = html_escape((doc_info or {}).get("validity", ""))

    footer_left = ""
    if is_cotizacion:
        footer_left = (
            "<div class='doc-footer-left'>"
            "Portes no incluidos salvo pactado previamente.<br/>"
            "En caso de necesitar ayuda mozo en la entrega especial en almacén o piso, habrá un cargo extra de 100€ + IVA."
            "</div>"
        )
    footer_right = "<div class='doc-footer-right'>Generado con IA</div>"

    html = f"""
    <html>
    <head><meta charset="utf-8">
    <style>
      body {{
        font-family: Arial, Helvetica, sans-serif;
        font-size: 12pt; color: #222; line-height: 1.6;
      }}
      h2 {{ margin: 8px 0 12px; font-size: 16pt; }}
      hr {{ border: none; border-top: 1px solid #ddd; margin: 10px 0 14px; }}
      table {{ border-collapse: collapse; width: 100%; table-layout: fixed; font-size: 11pt; page-break-inside: avoid; }}
      thead th {{ background:#f6f8fa; border:1px solid #dfe3e8; padding:6px 8px; text-align:left; font-weight:600; }}
      tbody td {{ border:1px solid #e6e9ef; padding:6px 8px; vertical-align: top; }}
      tr, td, th {{ page-break-inside: avoid; }}
      p {{ margin: 0 0 10px; }}

      /* --- Encabezado con 3 columnas: logo | empresa | datos --- */
      .doc-header {{
        display: grid;
        grid-template-columns: auto 1fr 220px;
        align-items: start;
        gap: 16px;
        margin: 0 0 6px;
      }}
      .company-info {{ font-size: 10.5pt; line-height: 1.45; }}
      .doc-meta {{ font-size: 10.5pt; line-height: 1.45; }}
      .doc-meta .title {{ font-weight: 700; letter-spacing: .5px; margin-bottom: 6px; }}

      .doc-footer-left {{
        position: fixed; left: 2mm; bottom: 2mm; max-width: 70%;
        font-size: 6px; color: #777; text-decoration: underline; line-height: 1.3;
      }}
      .doc-footer-right {{
        position: fixed; right: 2mm; bottom: 2mm; font-size: 8px;
        color: #777; text-decoration: underline; white-space: nowrap; text-align: right;
      }}
      .doc-content {{ padding-bottom: 18mm; }}
    </style>
    </head>
    <body>
      <div class="doc-header">
        <div class="logo">{logo_html}</div>
        <div class="company-info">
          {company_block}
        </div>
        <div class="doc-meta">
          <div class="title">DATOS</div>
          {"<div><strong>ID:</strong> " + _id + "</div>" if _id else ""}
          {"<div><strong>Hora:</strong> " + _time + "</div>" if _time else ""}
          {"<div><strong>Fecha:</strong> " + _date + "</div>" if _date else ""}
          {"<div><strong>Validez:</strong> " + _valid + "</div>" if _valid else ""}
        </div>
      </div>

      <h2>{html_escape(subject)}</h2>
      <hr/>
      <div class="doc-content">
        {body_html}
      </div>
      {footer_left}
      {footer_right}
    </body>
    </html>
    """.strip()
    return html


def _validate_pdf(path: str) -> bool:
    try:
        ok = os.path.exists(path) and os.path.getsize(path) >= 800
        if not ok:
            logging.getLogger("pdf").warning("PDF inválido (no existe o tamaño < 800B): %s", path)
            return False
        with open(path, "rb") as f:
            head = f.read(5)
        valid = (head == b"%PDF-")
        if not valid:
            logging.getLogger("pdf").warning("Cabecera PDF inválida: %s", path)
        return valid
    except Exception:
        logging.getLogger("pdf").exception("Error validando PDF")
        return False

def _pdf_output_dir() -> str:
    """
    Siempre guarda los PDFs en <raiz_del_proyecto>/PDF.
    Si la carpeta no existe, la crea; si falla, intenta cwd/PDF y en último extremo cwd.
    """
    pdf_logger = logging.getLogger("pdf")
    base = os.path.join(_project_root(), "PDF")
    try:
        os.makedirs(base, exist_ok=True)
        pdf_logger.debug("Usando carpeta PDF: %s", base)
        return base
    except Exception:
        pdf_logger.exception("No se pudo crear <proyecto>/PDF; intentando con cwd…")
    try:
        alt = os.path.join(os.getcwd(), "PDF")
        os.makedirs(alt, exist_ok=True)
        pdf_logger.debug("Usando carpeta PDF (fallback): %s", alt)
        return alt
    except Exception:
        pdf_logger.exception("No se pudo crear carpeta PDF en cwd; usando cwd sin subcarpeta")
    return os.getcwd()

def html_to_pdf_robust(html: str, margin_mm: int = 15) -> str:
    out_dir = _pdf_output_dir()
    pdf_logger = logging.getLogger("pdf")
    eff_mm = max(0.0, min(float(margin_mm), 2.0))
    fd1, path1 = tempfile.mkstemp(prefix="telequote_", suffix=".pdf", dir=out_dir)
    os.close(fd1)
    try:
        doc = fitz.open()
        W, H = 595, 842
        page = doc.new_page(width=W, height=H)
        m = eff_mm * 72 / 25.4
        rect = fitz.Rect(m, m, W - m, H - m)
        page.insert_htmlbox(rect, html)
        doc.save(path1)
        doc.close()
        pdf_logger.debug("PDF PyMuPDF generado: %s (%d bytes)", path1, os.path.getsize(path1))
        if _validate_pdf(path1):
            return path1
    except Exception:
        pdf_logger.exception("Fallo generando PDF con PyMuPDF")

    fd2, path2 = tempfile.mkstemp(prefix="telequote_", suffix=".pdf", dir=out_dir)
    os.close(fd2)
    try:
        writer = QPdfWriter(path2)
        writer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
        writer.setPageMargins(QMarginsF(eff_mm, eff_mm, eff_mm, eff_mm),
                              QPageLayout.Unit.Millimeter)
        doc2 = QTextDocument()
        doc2.setHtml(html)
        doc2.print(writer)
        pdf_logger.debug("PDF QPdfWriter generado: %s (%d bytes)", path2, os.path.getsize(path2))
        if _validate_pdf(path2):
            try:
                if os.path.exists(path1):
                    os.remove(path1)
            except Exception:
                pass
            return path2
    except Exception:
        pdf_logger.exception("Fallo generando PDF con QPdfWriter")

    raise RuntimeError("No se pudo generar un PDF válido.")


def send_email_with_pdf(server: str, port: int, user: str, password: str,
                        to_addr: str, subject: str, body_md: str, pdf_path: str):
    logging.getLogger("smtp").info("Enviando email a %s via %s:%s (adjunto=%s)",
                                   to_addr, server, port, os.path.basename(pdf_path))
    body_html = markdown_table_to_html(body_md)
    from email.message import EmailMessage
    msg = EmailMessage()
    msg["From"] = user
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.set_content(re.sub("<[^>]+>", "", body_html))
    msg.add_alternative(f"<html><body>{body_html}</body></html>", subtype="html")

    with open(pdf_path, "rb") as f:
        data = f.read()
    msg.add_attachment(data, maintype="application", subtype="pdf",
                       filename=os.path.basename(pdf_path))

    if port == 465:
        s = smtplib.SMTP_SSL(server, port, timeout=10)
    else:
        s = smtplib.SMTP(server, port, timeout=10)
        s.ehlo()
        if port == 587:
            s.starttls(); s.ehlo()
    s.login(user, password)
    s.send_message(msg)
    s.quit()
    logging.getLogger("smtp").info("Email enviado correctamente a %s", to_addr)

# ---------- Limpieza y firma obligatoria ----------
def _strip_placeholders(text: str) -> str:
    t = re.sub(r"\[[^\[\]\n]{1,80}\]", "", text or "")
    t = re.sub(r"\{[^\{\}\n]{1,80}\}", "", t)
    return t

def sanitize_subject_body(subject: str, body: str, owner: str) -> tuple[str, str]:
    s = _strip_placeholders(subject).strip(" -–—")
    b = _strip_placeholders(body)

    _TOTAL_KWS_RE = re.compile(
        r"\b("
        r"importe\s+total|total\s+general|gran\s+total|total\s+sin\s+iva|precio\s+total\s+sin\s+iva|"
        r"total\s+a\s+pagar|monto\s+total|total\s+de\s+la\s+cotizaci[oó]n|importe\s+final|total\s+final"
        r")\b",
        re.I,
    )

    cleaned_lines = []
    for ln in (b.splitlines() if b else []):
        if re.search(r"direcci[óo]n\s+de\s+env[ií]o|shipping address|direccion de envio", ln, re.I):
            continue
        if re.search(r"rellenar|por definir|pendiente|por confirmar", ln, re.I):
            continue
        if not ln.lstrip().startswith("|"):
            if _TOTAL_KWS_RE.search(ln) and (_parse_money(ln) is not None or "€" in ln):
                continue
        cleaned_lines.append(ln)
    b = "\n".join(cleaned_lines)

    fixed = []
    for ln in b.splitlines():
        if ln.lstrip().startswith("|"):
            fixed.append(ln)
            continue
        ln = re.sub(r",(?=\S)", ", ", ln)

        def _lower_after_greeting(m): return m.group(1) + m.group(2).lower()
        ln = re.sub(r"^(Estimad[oa]s?(?:\s+[^\s,]+){0,6},\s*)([A-ZÁÉÍÓÚÑ])", _lower_after_greeting, ln, flags=re.I)
        ln = re.sub(r"([.!?])(?!\s)(?=[A-ZÁÉÍÓÚÑ\"“«])", r"\1 ", ln)
        ln = re.sub(r"^(Estimad[oa]s?(?:\s+[^\s,]+){0,6},\s*)(Adjunto\b)", lambda m: m.group(1) + m.group(2).lower(), ln, flags=re.I)
        ln = re.sub(r"\.(?=(?:\s*)?(Atentamente|Saludos|Cordialmente)\b)", r". ", ln, flags=re.I)

        if re.search(r"\badjunto\b", ln, re.I):
            ln = re.sub(r"\b([Aa]djunto)\s+encontrar\w*\b", r"\1", ln, flags=re.I)
            ln = re.sub(r"\b(cotizaci[oó]n|presupuesto|propuesta)\s+formal\b", r"\1", ln, flags=re.I)
            ln = re.sub(r"(\b(?:cotizaci[oó]n|presupuesto|propuesta)\b)\s+para\s+(los|las|el|la)\b", r"\1 con \2", ln, flags=re.I)

        fixed.append(ln)

    b = "\n".join(fixed)
    b = re.sub(r"[ \t]+", " ", b)
    b = re.sub(r"\n{3,}", "\n\n", b).strip()

    courtesy = ("Esperamos que esta propuesta sea de su agrado. "
                "No dude en contactarnos si tiene alguna pregunta o necesita alguna modificación.")
    has_courtesy = re.search(r"(agrad[oa].*propuesta|no\s+dude\s+en\s+contactar)", b, re.I)

    if owner:
        if re.search(r"(atentamente|saludos|cordialmente)", b, re.I):
            repl = ""
            if not has_courtesy:
                repl += "\n\n" + courtesy
            repl += f"\n\nAtentamente,\n{owner}"
            b = re.sub(r"(?:\s*\.)?\s*(atentamente|saludos|cordialmente)[\s,:-]*.*$", repl, b, flags=re.I | re.S)
        elif owner not in b:
            if not has_courtesy:
                b = b.rstrip() + "\n\n" + courtesy
            b = b.rstrip() + f"\n\nAtentamente,\n{owner}"

    return (s.strip() or "Solicitud de cotización", b)


# ------------------------- Estado por chat -------------------------
@dataclass
class Session:
    phase: str = "idle"
    raw_text: str = ""
    image_path: Optional[str] = None
    subject: str = ""
    body: str = ""
    pdf_path: Optional[str] = None
    mode: str = ""  # "" | "solicitar_precios" | "cotizacion"

# ------------------------- Workers de validación -------------------------
class ApiKeyChecker(QThread):
    result = pyqtSignal(bool, str)
    def __init__(self, api_key: str):
        super().__init__(); self.api_key = api_key.strip()
    def run(self):
        logger = logging.getLogger("check.ApiKey")
        if not self.api_key:
            self.result.emit(False, "Sin API Key"); return
        try:
            genai.configure(api_key=self.api_key)
            ok = any(("gemini-2.5" in getattr(m, "name", "")) or
                    ("gemini-1.5" in getattr(m, "name", "")) or
                    (getattr(m, "name", "").startswith("gemini-"))
                    for m in genai.list_models())
            logger.info("API Key OK=%s", bool(ok))
            self.result.emit(bool(ok), "OK" if ok else "Modelos no disponibles")
        except Exception as e:
            logger.exception("Error comprobando API Key")
            self.result.emit(False, human_ex(e))

class SmtpChecker(QThread):
    result = pyqtSignal(bool, str)
    def __init__(self, server: str, port: int, email: str, pwd: str):
        super().__init__(); self.server = server; self.port = int(port); self.email=email; self.pwd=pwd
    def run(self):
        logger = logging.getLogger("check.SMTP")
        try:
            if not (self.server and self.port and self.email and self.pwd):
                self.result.emit(False, "Campos incompletos"); return
            if self.port == 465:
                s = smtplib.SMTP_SSL(self.server, self.port, timeout=5)
            else:
                s = smtplib.SMTP(self.server, self.port, timeout=5); s.ehlo()
                if self.port == 587: s.starttls(); s.ehlo()
            s.login(self.email, self.pwd); s.quit()
            logger.info("SMTP OK en %s:%s", self.server, self.port)
            self.result.emit(True, "OK")
        except Exception as e:
            logger.exception("SMTP NOK")
            self.result.emit(False, human_ex(e))

class BotTokenChecker(QThread):
    result = pyqtSignal(bool, str)
    def __init__(self, token: str):
        super().__init__(); self.token = token.strip()
    def run(self):
        logger = logging.getLogger("check.Token")
        try:
            if not self.token:
                self.result.emit(False, "Sin token"); return
            r = requests.get(f"https://api.telegram.org/bot{self.token}/getMe", timeout=6)
            ok = r.ok and r.json().get("ok") is True
            logger.info("Token válido=%s", ok)
            self.result.emit(ok, "OK" if ok else "Token inválido")
        except Exception as e:
            logger.exception("Error validando token")
            self.result.emit(False, human_ex(e))

# ------------------------- Bot en QThread -------------------------
class BotWorker(QThread):
    status = pyqtSignal(str)
    error = pyqtSignal(str)
    log = pyqtSignal(str)

    def __init__(self, token: str):
        super().__init__()
        self.token = token.strip()
        self._loop: Optional[asyncio.AbstractEventLoop] = None
        self._stop_event: Optional[asyncio.Event] = None
        self._app: Optional[Application] = None
        self.sessions: Dict[int, Session] = {}
        self.logger = logging.getLogger("BotWorker")

    def stop_bot(self):
        if self._loop and self._stop_event:
            self._loop.call_soon_threadsafe(self._stop_event.set)

    # check de permisos
    def _allowed_user(self, user_id: int) -> bool:
        mode = SET.access_mode()
        allowed = True
        if mode == "whitelist":
            allowed = user_id in SET.get_whitelist()
        elif mode == "blacklist":
            allowed = user_id not in SET.get_blacklist()
        self.logger.debug("Auth user %s -> %s (mode=%s)", user_id, allowed, mode)
        return allowed

    def run(self):
        try:
            self._loop = asyncio.new_event_loop()
            asyncio.set_event_loop(self._loop)
            self.logger.info("Arrancando bucle asyncio…")
            self._loop.run_until_complete(self._amain())
        except Exception as e:
            self.logger.exception("Fallo en hilo del bot")
            self.error.emit(human_ex(e))
            self.error.emit(traceback.format_exc())
        finally:
            try:
                if self._loop and not self._loop.is_closed():
                    self._loop.close()
            except Exception:
                pass

    async def _amain(self):
        if not self.token:
            self.status.emit("Sin token. Cárgalo en la pestaña Telegram.")
            return

        self._stop_event = asyncio.Event()
        self.sessions = {}

        self._app = ApplicationBuilder().token(self.token).concurrent_updates(True).build()

        self._app.add_handler(CommandHandler("start", self.h_start))
        self._app.add_handler(CommandHandler("cancelar", self.h_cancel))
        self._app.add_handler(CommandHandler("print", self.h_print))
        self._app.add_handler(CommandHandler("mail", self.h_mail))
        self._app.add_handler(CommandHandler("solicitar_precios", self.h_mode_solicitar))
        self._app.add_handler(CommandHandler("cotizacion", self.h_mode_cotizacion))

        base_filter = (filters.PHOTO | filters.Document.ALL | (filters.TEXT & ~filters.COMMAND))
        self._app.add_handler(MessageHandler(base_filter, self.h_content))
        self._app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, self.h_text_fallback))

        self.status.emit("Inicializando bot…")
        await self._app.initialize()
        await self._app.start()
        await self._app.updater.start_polling(drop_pending_updates=True)
        self.status.emit("Bot en ejecución.")
        self.log.emit("Listo. Envía un texto o imagen para comenzar.")
        self.logger.info("Bot en ejecución y escuchando actualizaciones")
        try:
            await self._stop_event.wait()
        finally:
            self.status.emit("Deteniendo bot…")
            await self._app.updater.stop()
            await self._app.stop()
            await self._app.shutdown()
            self.status.emit("Bot detenido.")
            self.logger.info("Bot detenido")

    def _get_session(self, chat_id: int) -> Session:
        s = self.sessions.get(chat_id)
        if not s:
            s = Session()
            self.sessions[chat_id] = s
        return s

    async def _download_file(self, file_obj, suffix: str) -> str:
        fd, path = tempfile.mkstemp(prefix="telequote_in_", suffix=suffix, dir=_tmp_dir())
        os.close(fd)
        await file_obj.download_to_drive(path)
        self.logger.debug("Archivo descargado a %s", path)
        return path

    # -------- Helpers de modo y nombre de archivo
    def _ensure_mode_first_time(self, update: Update) -> bool:
        s = self._get_session(update.effective_chat.id)
        if not s.mode:
            msg = ("Antes de comenzar, elige el modo:\n"
                   "• /solicitar_precios – pedir cotización a proveedores\n"
                   "• /cotizacion – enviar una cotización a tu cliente\n\n"
                   "Este modo quedará guardado hasta que lo cambies.")
            asyncio.create_task(update.message.reply_text(msg))
            self.logger.info("Solicitando selección de modo al usuario %s", update.effective_user.id)
            return False
        return True

    def _friendly_filename(self, text: str) -> str:
        base = re.sub(r"[^\w\s.-]", "_", text.strip())
        base = re.sub(r"\s+", "-", base)
        return base[:60] or "destinatario"

    def _rename_pdf(self, old_path: str, mode: str, recipient: str) -> str:
        try:
            date_str = datetime.now().strftime("%Y%m%d")
            prefix = "Cotizacion" if mode == "cotizacion" else "SolicitudPrecios"
            name = f"{prefix}_{self._friendly_filename(recipient)}_{date_str}.pdf"
            dst = os.path.join(_pdf_output_dir(), name)
            if os.path.exists(dst):
                i = 2
                root, ext = os.path.splitext(dst)
                while os.path.exists(f"{root}_{i}{ext}"):
                    i += 1
                dst = f"{root}_{i}{ext}"
            os.replace(old_path, dst)
            self.logger.debug("PDF renombrado a %s", dst)
            return dst
        except Exception:
            self.logger.exception("No se pudo renombrar el PDF")
            return old_path

    # -------- Handlers de comandos
    async def h_start(self, update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        uid = update.effective_user.id
        if not self._allowed_user(uid):
            await update.message.reply_text("⛔ No estás autorizado para usar este bot.")
            return
        await update.message.reply_text(
            "Hola 👋\n\n"
            "Primero elige un modo y quedará guardado hasta que lo cambies:\n"
            "• /solicitar_precios – pedir cotización a proveedores\n"
            "• /cotizacion – enviar una cotización a tu cliente\n\n"
            "Luego, envíame un *texto o una imagen* con la información.\n"
            "Después te preguntaré *¿A quién va dirigido?* y generaré el asunto, el cuerpo y el PDF.\n\n"
            "Comandos finales: /print, /mail y /cancelar.",
            parse_mode=ParseMode.MARKDOWN
        )

    async def h_mode_solicitar(self, update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        s = self._get_session(update.effective_chat.id)
        s.mode = "solicitar_precios"
        s.phase = "idle"; s.raw_text = ""; s.image_path = None; s.pdf_path = None; s.subject = ""; s.body = ""
        self.logger.info("Modo cambiado a solicitar_precios para chat %s", update.effective_chat.id)
        await update.message.reply_text("🟦 Modo seleccionado: *Solicitud de precios a proveedores*.\n"
                                        "Envíame un *texto o imagen* con el MATERIAL.",
                                        parse_mode=ParseMode.MARKDOWN)

    async def h_mode_cotizacion(self, update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        s = self._get_session(update.effective_chat.id)
        s.mode = "cotizacion"
        s.phase = "idle"; s.raw_text = ""; s.image_path = None; s.pdf_path = None; s.subject = ""; s.body = ""
        self.logger.info("Modo cambiado a cotizacion para chat %s", update.effective_chat.id)
        await update.message.reply_text("🟩 Modo seleccionado: *Cotización para cliente*.\n"
                                        "Envíame un *texto o imagen* con el MATERIAL.",
                                        parse_mode=ParseMode.MARKDOWN)

    async def h_cancel(self, update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        uid = update.effective_user.id
        if not self._allowed_user(uid):
            await update.message.reply_text("⛔ No estás autorizado para usar este bot.")
            return
        cur = self._get_session(update.effective_chat.id)
        if cur.image_path:
            _safe_unlink(cur.image_path)
        keep_mode = cur.mode
        self.sessions[update.effective_chat.id] = Session(mode=keep_mode)
        self.logger.info("Conversación reiniciada (modo conservado=%s)", keep_mode)
        await update.message.reply_text("✅ Conversación reiniciada. Envíame un *texto o imagen* para comenzar.",
                                        parse_mode=ParseMode.MARKDOWN)

    async def h_print(self, update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        uid = update.effective_user.id
        if not self._allowed_user(uid):
            await update.message.reply_text("⛔ No estás autorizado para usar este bot.")
            return
        s = self._get_session(update.effective_chat.id)
        if not s.pdf_path:
            await update.message.reply_text("Aún no hay PDF. Envía primero un texto o imagen.")
            return
        try:
            os.startfile(s.pdf_path, "print")
            self.logger.info("Impresión solicitada: %s", s.pdf_path)
            await update.message.reply_text("🖨️ Enviado a la impresora predeterminada.")
        except Exception as e:
            self.logger.exception("Error imprimiendo")
            await update.message.reply_text(f"No se pudo imprimir: {human_ex(e)}")

    async def h_mail(self, update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        uid = update.effective_user.id
        if not self._allowed_user(uid):
            await update.message.reply_text("⛔ No estás autorizado para usar este bot.")
            return
        s = self._get_session(update.effective_chat.id)
        if not s.pdf_path:
            await update.message.reply_text("Aún no hay PDF. Envía primero un texto o imagen.")
            return
        s.phase = "waiting_email"
        self.logger.debug("Esperando email del destinatario…")
        await update.message.reply_text("📧 Indica el *correo destinatario*:",
                                        parse_mode=ParseMode.MARKDOWN)

    async def h_content(self, update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        uid = update.effective_user.id
        if not self._allowed_user(uid):
            await update.message.reply_text("⛔ No estás autorizado para usar este bot.")
            return

        if not self._ensure_mode_first_time(update):
            return

        chat_id = update.effective_chat.id
        s = self._get_session(chat_id)

        if update.message.text and s.phase in ("waiting_recipient", "waiting_email"):
            await self.h_text_fallback(update, ctx)
            return

        if s.image_path:
            _safe_unlink(s.image_path)

        keep_mode = s.mode
        self.sessions[chat_id] = s = Session(mode=keep_mode)

        text = ""
        img_path = None

        if update.message.text and not update.message.text.startswith("/"):
            text = update.message.text.strip()
        elif update.message.photo:
            photo = update.message.photo[-1]
            file = await photo.get_file()
            img_path = await self._download_file(file, ".jpg")
        elif update.message.document:
            doc = update.message.document
            file = await doc.get_file()
            mt = (doc.mime_type or "").lower()
            name = (doc.file_name or "").lower()
            if mt.startswith("image/") or name.endswith((".png", ".jpg", ".jpeg", ".bmp")):
                suffix = ".png"
                if name.endswith(".jpg") or name.endswith(".jpeg"):
                    suffix = ".jpg"
                elif name.endswith(".bmp"):
                    suffix = ".bmp"
                img_path = await self._download_file(file, suffix)
            elif mt == "application/pdf" or name.endswith(".pdf"):
                img_path = await self._download_file(file, ".pdf")
            else:
                await update.message.reply_text("Documento no soportado. Usa imagen o PDF.")
                return

        s.raw_text = text
        s.image_path = img_path
        self.logger.info("Recibido MATERIAL (texto_len=%d, imagen=%s) | modo=%s",
                         len(text), bool(img_path), s.mode)
        s.phase = "waiting_recipient"
        await update.message.reply_text("¿Hacia *quién* va dirigido? (nombre/empresa)",
                                        parse_mode=ParseMode.MARKDOWN)

    async def h_text_fallback(self, update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        uid = update.effective_user.id
        if not self._allowed_user(uid):
            await update.message.reply_text("⛔ No estás autorizado para usar este bot.")
            return

        chat_id = update.effective_chat.id
        s = self._get_session(chat_id)

        if s.phase == "waiting_email":
            email = (update.message.text or "").strip()
            if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email):
                await update.message.reply_text("El email no parece válido. Intenta de nuevo o /cancelar.")
                return
            smtp = SET.get_smtp()
            try:
                send_email_with_pdf(
                    smtp.get("server", ""), int(smtp.get("port", 587)),
                    smtp.get("email", ""), smtp.get("password", ""),
                    email, s.subject, s.body, s.pdf_path or ""
                )
                await update.message.reply_text(f"📨 Enviado a {email}.")
            except Exception as e:
                self.logger.exception("Error enviando correo")
                await update.message.reply_text(f"No se pudo enviar el correo: {human_ex(e)}")
            finally:
                s.phase = "ready"
            return

        if s.phase != "waiting_recipient":
            return

        recipient = (update.message.text or "").strip()
        if not recipient:
            await update.message.reply_text("Escribe el nombre del destinatario, por favor.")
            return

        try:
            api_key = SET.get("google_api_key", "")
            if not api_key:
                await update.message.reply_text("⚠️ Falta la Google API Key (pestaña *Cuentas / API*).")
                s.phase = "idle"
                return

            owner = (SET.get("owner_name", "") or "").strip()
            if not owner:
                await update.message.reply_text("⚠️ Falta tu *Nombre o Empresa* en la pestaña Cuentas / API.")
                s.phase = "idle"
                return

            genai.configure(api_key=api_key)
            model = genai.GenerativeModel("gemini-2.5-flash")

            if s.mode == "cotizacion":
                prompt_lines = [
                    "Genera un CORREO FORMAL DE COTIZACIÓN para enviar a un CLIENTE usando el MATERIAL del usuario.",
                    "El MATERIAL puede ser texto y/o imagen (si hay imagen, léela y extrae ítems/valores relevantes).",
                    f"El correo va dirigido a: {recipient}.",
                    f"Firma/remitente OBLIGATORIO (úsalo literalmente): {owner}.",
                    "Idioma: español de España.",
                    f"Comienza con 'Estimado {recipient},' o 'Estimada {recipient},' (si no sabes el género usa 'Estimado {recipient},').",
                    "Tras el saludo usa una frase natural como 'Adjunto la cotización detallada para su solicitud:' (evita 'adjunto encontrará').",
                    "Moneda: Euro (€). Respeta el formato del material.",
                    "TABLA OBLIGATORIA: usa EXACTAMENTE 4 columnas en este ORDEN y con estos encabezados literales:",
                    "| Ítem | Cantidad | Detalle | Precio unitario |",
                    "Requisitos estrictos:",
                    "- La tabla debe ir en **bloque** Markdown con separador '---'.",
                    "- No incluyas 'Total'.",
                    "- 'Cantidad' puede ser texto; en servicios sin cantidad explícita, usa '1'.",
                    "- Si hay color/material, colócalo en 'Detalle'.",
                    "- Si el usuario da un precio, úsalo tal cual como precio unitario.",
                    "- No inventes datos ni uses placeholders.",
                    'Devuelve SOLO JSON: { "subject": "...", "body": "..." }'
                ]
            else:
                prompt_lines = [
                    "Genera un CORREO FORMAL DE SOLICITUD DE COTIZACIÓN usando el MATERIAL del usuario.",
                    "El MATERIAL puede ser texto y/o imagen.",
                    f"El correo va dirigido a: {recipient}.",
                    f"Firma/remitente OBLIGATORIO (úsalo literalmente): {owner}.",
                    f"Saludo: 'Estimado {recipient},' o 'Estimada {recipient},'.",
                    "Evita 'adjunto encontrará'.",
                    "Si presentas ítems en tabla, usa | Ítem | Cantidad | Detalle |.",
                    'Devuelve SOLO JSON: { "subject": "...", "body": "..." }'
                ]

            parts: List = ["\n".join(prompt_lines)]
            if s.raw_text:
                parts.append(s.raw_text)
            if s.image_path:
                p = s.image_path
                if p.lower().endswith(".pdf"):
                    with fitz.open(p) as d:
                        pg = d.load_page(0)
                        pix = d.get_pixmap(dpi=200)
                        img_bytes = pix.tobytes("png")
                    parts.append({"mime_type": "image/png", "data": img_bytes})
                else:
                    with open(p, "rb") as f:
                        img_bytes = f.read()
                    mime = "image/png"
                    if p.lower().endswith((".jpg", ".jpeg")): mime = "image/jpeg"
                    elif p.lower().endswith(".bmp"): mime = "image/bmp"
                    parts.append({"mime_type": mime, "data": img_bytes})

            schema = {"type": "OBJECT", "properties": {"subject": {"type": "STRING"}, "body": {"type": "STRING"}}, "required": ["subject", "body"]}
            cfg = GenerationConfig(response_mime_type="application/json", response_schema=schema, temperature=0.2)
            resp = model.generate_content(parts, generation_config=cfg)
            raw = getattr(resp, "text", "") or "{}"
            self.logger.debug("Gemini raw JSON: %s", raw[:1000])
            try:
                data = json.loads(raw)
            except Exception:
                self.logger.warning("Gemini no devolvió JSON válido; usando valores por defecto")
                data = {}

            subject = (data.get("subject") or ("Cotización" if s.mode == "cotizacion" else "Solicitud de cotización")).strip()
            body = (data.get("body") or "").strip()

            subject, body = sanitize_subject_body(subject, body, owner)
            is_cot = (s.mode == "cotizacion")
            body = process_markdown_tables(body, is_cotizacion=is_cot, pricing_mode=SET.get("pricing_detection", "auto"))

            s.subject = subject
            s.body = body

            # ---- NUEVO: construir doc_info (ID/Hora/Fecha/Validez) ----
            doc_id = SET.next_doc_id()
            now = datetime.now()
            time_str = now.strftime("%H:%M")
            date_str = now.strftime("%d/%m/%Y")
            x = SET.validity_days()
            valid_dt = _add_business_days(now, x)
            valid_str = f"{valid_dt.strftime('%d/%m/%Y')} ({x} días hábiles)"
            doc_info = {
                "id": f"{doc_id:06d}",
                "time": time_str,
                "date": date_str,
                "validity": valid_str
            }

            html = compose_html(subject, body, is_cotizacion=is_cot, doc_info=doc_info)
            pdf = html_to_pdf_robust(html, SET.pdf_margin())
            pdf_named = self._rename_pdf(pdf, s.mode or "solicitar_precios", recipient)
            s.pdf_path = pdf_named
            s.phase = "ready"
            self.logger.info("PDF listo: %s (%d bytes)", pdf_named, os.path.getsize(pdf_named))

            with open(pdf_named, "rb") as fh:
                await update.message.reply_document(
                    document=fh,
                    filename=os.path.basename(pdf_named),
                    caption="PDF generado automáticamente ✅"
                )
            await update.message.reply_text(
                "Ahora puedes usar:\n• /print — imprimir\n• /mail — enviar por correo\n• /cancelar — reiniciar\n"
                "Para cambiar el modo usa /solicitar_precios o /cotizacion."
            )

            if s.image_path:
                _safe_unlink(s.image_path)
                s.image_path = None

        except Exception as e:
            self.logger.exception("Error generando correo/PDF")
            self.log.emit(traceback.format_exc())
            await update.message.reply_text(f"Error generando el correo/PDF: {human_ex(e)}")
            if s.image_path:
                _safe_unlink(s.image_path)
                s.image_path = None
            s.phase = "idle"


# ------------------------- Auto-update helpers -------------------------
class UpdateCheckerWorker(QThread):
    finished = pyqtSignal(dict)
    def run(self):
        url = f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/releases/latest"
        result = {"ok": False, "error": "", "tag": "", "assets": []}
        try:
            r = requests.get(url, timeout=12)
            r.raise_for_status()
            data = r.json()
            tag = (data.get("tag_name") or "").lstrip("v")
            result.update({"ok": True, "tag": tag, "assets": data.get("assets", [])})
            logging.getLogger("update").info("Última versión remota: %s", tag)
        except Exception as e:
            logging.getLogger("update").exception("Error consultando updates")
            result["error"] = human_ex(e)
        self.finished.emit(result)

def _choose_zip_asset(assets: list) -> Optional[dict]:
    for a in assets or []:
        name = str(a.get("name", "")).lower()
        if name.endswith(".zip"):
            return a
    return None

# ------------------------- Autostart (VBS en Startup) -------------------------
def _startup_vbs_path() -> str:
    appdata = os.environ.get("APPDATA", "")
    return os.path.join(appdata, "Microsoft", "Windows", "Start Menu", "Programs", "Startup", AUTOSTART_VBS_NAME)

def is_windows_autostart_enabled() -> bool:
    if sys.platform != "win32":
        return False
    try:
        return os.path.isfile(_startup_vbs_path())
    except Exception:
        return False

def set_windows_autostart(enable: bool) -> None:
    if sys.platform != "win32":
        return
    vbs_path = _startup_vbs_path()
    if not enable:
        try:
            if os.path.exists(vbs_path):
                os.remove(vbs_path)
                print(f"[AUTOSTART] Eliminado {vbs_path}")
        except Exception as e:
            print(f"[AUTOSTART] No se pudo eliminar: {e}")
        return

    try:
        pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
        if not os.path.exists(pythonw):
            pythonw = sys.executable
        script_path = os.path.realpath(sys.argv[0])
        workdir = os.path.dirname(script_path)
        def q(s: str) -> str: return '"' + s.replace('"','""') + '"'
        vbs = (
            "Option Explicit\r\n"
            "On Error Resume Next\r\n"
            "Dim sh: Set sh = CreateObject(\"WScript.Shell\")\r\n"
            "Dim dq: dq = Chr(34)\r\n"
            "Dim i, rc\r\n"
            "For i = 1 To 60\r\n"
            "  rc = sh.Run(\"cmd /c ping -n 1 -w 1000 1.1.1.1 >nul\", 0, True)\r\n"
            "  If rc = 0 Then Exit For\r\n"
            "  WScript.Sleep 5000\r\n"
            "Next\r\n"
            f"sh.CurrentDirectory = {q(workdir)}\r\n"
            f"Dim cmd: cmd = dq & {q(pythonw)} & dq & \" \" & dq & {q(script_path)} & dq & \" --autostart-bot\"\r\n"
            "sh.Run cmd, 0, False\r\n"
        )
        os.makedirs(os.path.dirname(vbs_path), exist_ok=True)
        with open(vbs_path, "w", encoding="utf-8") as f:
            f.write(vbs)
        print(f"[AUTOSTART] Creado {vbs_path}")
    except Exception as e:
        print(f"[AUTOSTART] Error creando VBS: {e}")

# ------------------------- Interfaz -------------------------
class MainWindow(QMainWindow):
    def __init__(self, autostart_mode: bool = False):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} · {APP_ORG}")
        self.resize(980, 680)
        if os.path.exists("BitStation.ico"):
            self.setWindowIcon(QIcon("BitStation.ico"))
        self.bot_thread: Optional[BotWorker] = None
        self._update_assets = []
        self._update_tag = ""

        tabs = QTabWidget(); self.setCentralWidget(tabs)

        # Telegram
        w_tg = QWidget(); ltg = QVBoxLayout(w_tg)
        row = QHBoxLayout()
        row.addWidget(QLabel("Token del bot:"))
        self.ed_token = QLineEdit(SET.get("telegram_token", "")); self.ed_token.setEchoMode(QLineEdit.EchoMode.Password)
        self.led_bot = QLabel(); self._init_led(self.led_bot, "unknown")
        row.addWidget(self.ed_token, 1); row.addWidget(self.led_bot)
        ltg.addLayout(row)

        self.cb_autostart = QCheckBox("Iniciar bot al iniciar Windows (segundo plano)")
        self.cb_autostart.setChecked(is_windows_autostart_enabled())
        ltg.addWidget(self.cb_autostart)

        # Control de acceso
        access_group = QGroupBox("Control de acceso")
        la = QVBoxLayout(access_group)

        switches = QHBoxLayout()
        self.cb_whitelist = QCheckBox("Habilitar Whitelist")
        self.cb_blacklist = QCheckBox("Habilitar Blacklist")
        switches.addWidget(self.cb_whitelist)
        switches.addWidget(self.cb_blacklist)
        la.addLayout(switches)

        lists_row = QHBoxLayout()
        wl_box = QVBoxLayout()
        wl_hdr = QHBoxLayout()
        wl_hdr.addWidget(QLabel("Whitelist — User ID"))
        self.btn_wl_add = QPushButton("+"); self.btn_wl_del = QPushButton("−")
        self.btn_wl_add.setFixedWidth(28); self.btn_wl_del.setFixedWidth(28)
        wl_hdr.addStretch(1); wl_hdr.addWidget(self.btn_wl_add); wl_hdr.addWidget(self.btn_wl_del)
        self.list_wl = QListWidget()
        wl_box.addLayout(wl_hdr); wl_box.addWidget(self.list_wl)
        lists_row.addLayout(wl_box)

        bl_box = QVBoxLayout()
        bl_hdr = QHBoxLayout()
        bl_hdr.addWidget(QLabel("Blacklist — User ID"))
        self.btn_bl_add = QPushButton("+"); self.btn_bl_del = QPushButton("−")
        self.btn_bl_add.setFixedWidth(28); self.btn_bl_del.setFixedWidth(28)
        bl_hdr.addStretch(1); bl_hdr.addWidget(self.btn_bl_add); bl_hdr.addWidget(self.btn_bl_del)
        self.list_bl = QListWidget()
        bl_box.addLayout(bl_hdr); bl_box.addWidget(self.list_bl)
        lists_row.addLayout(bl_box)

        la.addLayout(lists_row)
        ltg.addWidget(access_group)

        self.lbl_status = QLabel("Estado del bot: —"); ltg.addWidget(self.lbl_status)
        self.log = QTextBrowser(); ltg.addWidget(self.log, 1)
        tabs.addTab(w_tg, "Telegram")

        # Cuentas / API
        w_acc = QWidget(); lacc = QVBoxLayout(w_acc)

        grp_owner = QGroupBox("Identidad (firma)")
        fo = QFormLayout(grp_owner)
        # --- existentes + NUEVOS campos ---
        self.ed_owner = QLineEdit(SET.get("owner_name", ""))
        fo.addRow("Nombre o Empresa:", self.ed_owner)

        comp = SET.get_company()
        self.ed_addr = QLineEdit(comp.get("address", ""))
        self.ed_cif  = QLineEdit(comp.get("cif", ""))
        self.ed_tel  = QLineEdit(comp.get("phone", ""))
        self.ed_mail = QLineEdit(comp.get("email", ""))
        self.ed_web  = QLineEdit(comp.get("web", ""))

        fo.addRow("Dirección:", self.ed_addr)
        fo.addRow("CIF:", self.ed_cif)
        fo.addRow("Teléfono:", self.ed_tel)
        fo.addRow("eMail:", self.ed_mail)
        fo.addRow("Web:", self.ed_web)

        self.sb_valid = QSpinBox()
        self.sb_valid.setRange(1, 365)
        self.sb_valid.setValue(SET.validity_days())
        fo.addRow("Días de validez:", self.sb_valid)

        grp_owner.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        lacc.addWidget(grp_owner)

        grp_g = QGroupBox("Google API (Gemini)")
        fg = QFormLayout(grp_g)
        row_api = QHBoxLayout()
        self.ed_api = QLineEdit(SET.get("google_api_key", "")); self.ed_api.setEchoMode(QLineEdit.EchoMode.Password)
        self.led_api = QLabel(); self._init_led(self.led_api, "unknown")
        row_api.addWidget(self.ed_api, 1); row_api.addWidget(self.led_api)
        fg.addRow("API Key:", row_api)
        grp_g.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        lacc.addWidget(grp_g)

        grp_s = QGroupBox("Correo saliente (SMTP)")
        fs = QFormLayout(grp_s)
        smtp = SET.get_smtp()
        self.ed_smtp_server = QLineEdit(smtp.get("server", "smtp.gmail.com"))
        self.sb_smtp_port = QSpinBox(); self.sb_smtp_port.setRange(1, 65535); self.sb_smtp_port.setValue(int(smtp.get("port", 587)))
        self.ed_smtp_email = QLineEdit(smtp.get("email", ""))
        self.ed_smtp_pass = QLineEdit(smtp.get("password", "")); self.ed_smtp_pass.setEchoMode(QLineEdit.EchoMode.Password)
        self.led_smtp = QLabel(); self._init_led(self.led_smtp, "unknown")
        fs.addRow("Servidor:", self.ed_smtp_server)
        fs.addRow("Puerto:", self.sb_smtp_port)
        fs.addRow("Correo:", self.ed_smtp_email)
        fs.addRow("Contraseña:", self.ed_smtp_pass)
        fs.addRow("Estado:", self.led_smtp)
        lacc.addWidget(grp_s)

        # Update
        self.lbl_version = QLabel(f"Versión instalada: v{APP_VERSION}")
        self.btn_update = QPushButton("Comprobando actualizaciones…")
        self.btn_update.setEnabled(False)
        lacc.addSpacing(6); lacc.addWidget(self.lbl_version); lacc.addWidget(self.btn_update)

        tabs.addTab(w_acc, "Cuentas / API")

        # Autosave
        self.ed_token.textChanged.connect(lambda s: (SET.set("telegram_token", s), self._debounce_bot_check()))
        self.ed_owner.textChanged.connect(lambda s: SET.set("owner_name", s))

        self.ed_addr.textChanged.connect(lambda s: SET.set_company("address", s))
        self.ed_cif.textChanged.connect(lambda s: SET.set_company("cif", s))
        self.ed_tel.textChanged.connect(lambda s: SET.set_company("phone", s))
        self.ed_mail.textChanged.connect(lambda s: SET.set_company("email", s))
        self.ed_web.textChanged.connect(lambda s: SET.set_company("web", s))
        self.sb_valid.valueChanged.connect(lambda v: SET.set_validity_days(int(v)))

        self.ed_api.textChanged.connect(lambda s: (SET.set("google_api_key", s), self._debounce_api_check()))
        self.ed_smtp_server.textChanged.connect(lambda s: (SET.set_smtp("server", s), self._debounce_smtp_check()))
        self.sb_smtp_port.valueChanged.connect(lambda v: (SET.set_smtp("port", int(v)), self._debounce_smtp_check()))
        self.ed_smtp_email.textChanged.connect(lambda s: (SET.set_smtp("email", s), self._debounce_smtp_check()))
        self.ed_smtp_pass.textChanged.connect(lambda s: (SET.set_smtp("password", s), self._debounce_smtp_check()))

        # Otros controles
        self.cb_autostart.toggled.connect(lambda v: set_windows_autostart(bool(v)))
        self.btn_update.clicked.connect(self._start_update)

        # Access connections
        self.cb_whitelist.toggled.connect(self._toggle_whitelist)
        self.cb_blacklist.toggled.connect(self._toggle_blacklist)
        self.btn_wl_add.clicked.connect(lambda: self._add_id("whitelist"))
        self.btn_wl_del.clicked.connect(lambda: self._remove_selected("whitelist"))
        self.btn_bl_add.clicked.connect(lambda: self._add_id("blacklist"))
        self.btn_bl_del.clicked.connect(lambda: self._remove_selected("blacklist"))

        self._refresh_access_ui()

        # Timers de debounce
        self._t_api = QTimer(self); self._t_api.setSingleShot(True); self._t_api.timeout.connect(self._check_api)
        self._t_smtp = QTimer(self); self._t_smtp.setSingleShot(True); self._t_smtp.timeout.connect(self._check_smtp)
        self._t_bot = QTimer(self); self._t_bot.setSingleShot(True); self._t_bot.timeout.connect(self._check_bot)

        # Chequeos iniciales
        self._debounce_api_check(10)
        self._debounce_smtp_check(10)
        self._debounce_bot_check(10)

        # Chequeo de updates
        self._upd = UpdateCheckerWorker()
        self._upd.finished.connect(self._on_update_info)
        self._upd.start()

        if autostart_mode:
            QTimer.singleShot(300, self.start_bot)

    # ----- LED helpers -----
    def _init_led(self, lbl: QLabel, state: str):
        lbl.setFixedSize(14, 14)
        lbl.setToolTip("Desconocido")
        self._set_led(lbl, state)

    def _set_led(self, lbl: QLabel, state: str):
        color = {"ok": "#2ecc71", "bad": "#e74c3c", "unknown": "#bdc3c7"}.get(state, "#bdc3c7")
        lbl.setStyleSheet(f"background:{color}; border-radius:7px; border:1px solid #888;")
        tt = {"ok": "Correcto", "bad": "Error", "unknown": "Desconocido"}.get(state, "Desconocido")
        lbl.setToolTip(tt)

    # ----- Access UI helpers -----
    def _refresh_access_ui(self):
        mode = SET.access_mode()
        self.cb_whitelist.blockSignals(True)
        self.cb_blacklist.blockSignals(True)
        self.cb_whitelist.setChecked(mode == "whitelist")
        self.cb_blacklist.setChecked(mode == "blacklist")
        self.cb_whitelist.blockSignals(False)
        self.cb_blacklist.blockSignals(False)

        self.list_wl.clear()
        self.list_bl.clear()
        for uid in SET.get_whitelist():
            self.list_wl.addItem(str(uid))
        for uid in SET.get_blacklist():
            self.list_bl.addItem(str(uid))

    def _toggle_whitelist(self, checked: bool):
        if checked:
            SET.set_access_mode("whitelist")
            if self.cb_blacklist.isChecked():
                self.cb_blacklist.setChecked(False)
        else:
            if SET.access_mode() == "whitelist":
                SET.set_access_mode("off")

    def _toggle_blacklist(self, checked: bool):
        if checked:
            SET.set_access_mode("blacklist")
            if self.cb_whitelist.isChecked():
                self.cb_whitelist.setChecked(False)
        else:
            if SET.access_mode() == "blacklist":
                SET.set_access_mode("off")

    def _add_id(self, which: str):
        text, ok = QInputDialog.getText(self, "Añadir User ID", "Ingresa el User ID numérico:")
        if not ok or not text.strip():
            return
        s = text.strip()
        if not re.fullmatch(r"-?\d+", s):
            QMessageBox.warning(self, "ID inválido",
                                "El User ID debe ser un número entero (se aceptan IDs negativos de grupos).")
            return
        uid = int(s)
        inserted = SET.add_to_list(which, uid)
        if not inserted:
            QMessageBox.information(self, "Duplicado", f"El ID {uid} ya estaba en la {which}.")
        target = self.list_wl if which == "whitelist" else self.list_bl
        existing = {target.item(i).text() for i in range(target.count())}
        if str(uid) not in existing:
            target.addItem(str(uid))

    def _remove_selected(self, which: str):
        lst = self.list_wl if which == "whitelist" else self.list_bl
        items = lst.selectedItems()
        if not items:
            QMessageBox.information(self, "Quitar", "Selecciona un ID de la lista.")
            return
        for it in items:
            try:
                uid = int(it.text())
                SET.remove_from_list(which, uid)
                row = lst.row(it)
                lst.takeItem(row)
            except Exception:
                pass

    # ----- Debounce + checks -----
    def _debounce_api_check(self, ms: int = 500):
        self._t_api.start(ms)

    def _debounce_smtp_check(self, ms: int = 500):
        self._t_smtp.start(ms)

    def _debounce_bot_check(self, ms: int = 500):
        self._t_bot.start(ms)

    def _check_api(self):
        self._set_led(self.led_api, "unknown")
        worker = ApiKeyChecker(self.ed_api.text())
        worker.result.connect(self._on_api_checked)
        worker.start()
        self._api_worker = worker

    def _on_api_checked(self, ok: bool, msg: str):
        self._set_led(self.led_api, "ok" if ok else "bad")
        if ok:
            self.log.append("✅ Google API Key OK.")
        else:
            self.log.append(f"❌ Google API Key: {html_escape(msg)}")

    def _check_smtp(self):
        self._set_led(self.led_smtp, "unknown")
        worker = SmtpChecker(self.ed_smtp_server.text().strip(),
                             int(self.sb_smtp_port.value()),
                             self.ed_smtp_email.text().strip(),
                             self.ed_smtp_pass.text())
        worker.result.connect(self._on_smtp_checked)
        worker.start()
        self._smtp_worker = worker

    def _on_smtp_checked(self, ok: bool, msg: str):
        self._set_led(self.led_smtp, "ok" if ok else "bad")
        if ok:
            self.log.append("✅ SMTP OK.")
        else:
            self.log.append(f"❌ SMTP error: {html_escape(msg)}")

    def _check_bot(self):
        self._set_led(self.led_bot, "unknown")
        worker = BotTokenChecker(self.ed_token.text())
        worker.result.connect(self._on_bot_checked)
        worker.start()
        self._bot_checker = worker

    def _on_bot_checked(self, ok: bool, msg: str):
        self._set_led(self.led_bot, "ok" if ok else "bad")
        if ok:
            self.log.append("✅ Token de Telegram válido.")
            if not (self.bot_thread and self.bot_thread.isRunning()):
                self.start_bot()
        else:
            self.log.append(f"❌ Token Telegram: {html_escape(msg)}")

    # ----- UI helpers -----
    def choose_pdf_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Selecciona carpeta para guardar PDF", SET.pdf_dir())
        if d:
            SET.set_pdf_dir(d)

    def start_bot(self):
        token = self.ed_token.text().strip()
        if not token:
            self.log.append("⚠️ Ingresa el Token del bot.")
            return
        if self.bot_thread and self.bot_thread.isRunning():
            self.log.append("El bot ya está corriendo.")
            return
        self.bot_thread = BotWorker(token)
        self.bot_thread.status.connect(lambda s: (print(f"[STATUS] {s}"), self.lbl_status.setText(f"Estado del bot: {s}")))
        self.bot_thread.error.connect(lambda e: (print("[ERROR]", e), self.log.append(f"<span style='color:#c0392b'>{html_escape(e)}</span>")))
        self.bot_thread.log.connect(lambda t: (print(t), self.log.append(html_escape(t))))
        self.bot_thread.start()
        self.log.append("⏳ Iniciando…")

    def stop_bot(self):
        if self.bot_thread and self.bot_thread.isRunning():
            self.bot_thread.stop_bot()
            self.log.append("⏹ Solicitando parada…")

    def _on_update_info(self, info: dict):
        if not info.get("ok"):
            self.btn_update.setText("No se pudo comprobar updates")
            self.btn_update.setEnabled(False)
            return
        tag = (info.get("tag") or "").lstrip("v")
        if tag and tag != APP_VERSION:
            self.btn_update.setText(f"Actualizar a v{tag}")
            self.btn_update.setEnabled(True)
            self._update_assets = info.get("assets", [])
            self._update_tag = tag
        else:
            self.btn_update.setText("Estás en la última versión")
            self.btn_update.setEnabled(False)

    def _start_update(self):
        assets = self._update_assets
        tag = self._update_tag
        if not assets or not tag:
            QMessageBox.warning(self, "Update", "No hay actualización disponible.")
            return
        asset = _choose_zip_asset(assets)
        if not asset:
            QMessageBox.warning(self, "Update", "El release no contiene un archivo .zip.")
            return
        url = asset.get("browser_download_url")
        name = asset.get("name", "update.zip")
        upd_dir = _resource_path("update_temp")
        os.makedirs(upd_dir, exist_ok=True)
        dst = os.path.join(upd_dir, name)

        try:
            with requests.get(url, stream=True) as r:
                r.raise_for_status()
                with open(dst, "wb") as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)
            updater = _resource_path("updater.py")
            if not os.path.exists(updater):
                QMessageBox.critical(self, "Update", "No se encontró updater.py en la carpeta.")
                return
            pid = os.getpid()
            creationflags = subprocess.DETACHED_PROCESS if sys.platform == "win32" else 0
            subprocess.Popen([sys.executable, updater, str(pid), tag], creationflags=creationflags)
            self.close()
        except Exception as e:
            QMessageBox.critical(self, "Update", f"Fallo al actualizar:\n{human_ex(e)}")

    def closeEvent(self, ev):
        try:
            if self.bot_thread and self.bot_thread.isRunning():
                self.bot_thread.stop_bot()
                self.bot_thread.wait(2000)
        except Exception:
            pass
        super().closeEvent(ev)

# ------------------------- main -------------------------
def main():
    if sys.platform == "win32":
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("BitStation.TeleQuote")
        except Exception:
            pass

    autostart_only = ("--autostart-bot" in sys.argv)

    app = QApplication(sys.argv)
    win = MainWindow(autostart_mode=autostart_only)
    if not autostart_only:
        win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
