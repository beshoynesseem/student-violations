# -*- coding: utf-8 -*-
"""
برنامج سجل المخالفات والعقوبات (طبقاً للقرار رقم 150 لسنة 2024)
================================================================

طريقة التشغيل:
  1) ثبّت بايثون (من python.org) ثم افتح CMD واكتب:   pip install openpyxl
  2) ضع هذا الملف في مجلد، وبجانبه ملفات:
        شيت_مثالي.xlsx   ع_ع.xlsx   ع_ل.xlsx
  3) شغّل البرنامج:   python student_violations.py

عند أول تشغيل يستورد البرنامج الأسماء من ملفي القوائم تلقائياً.
كل البيانات تُحفظ في ملف violations.db بجانب البرنامج (مع نسخة احتياطية يومية في مجلد backup).
الشيت الجاهز للطباعة يُحفظ في مجلد exports.
"""
import datetime as dt
import hashlib
import hmac
import math
import os
import re
import sqlite3
import subprocess
import shutil
import sys
import traceback
from collections import Counter
from copy import copy
from pathlib import Path

import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, simpledialog, ttk

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment
except ImportError:  # المكتبة غير مثبتة
    _r = tk.Tk()
    _r.withdraw()
    messagebox.showerror(
        "مكتبة ناقصة",
        "لازم تثبّت مكتبة openpyxl أولاً.\n\nافتح موجّه الأوامر (CMD) واكتب:\n\npip install openpyxl",
    )
    sys.exit(1)

# ----------------------------------------------------------------------------
# الإعدادات والمسارات
# ----------------------------------------------------------------------------
APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
DB_PATH = APP_DIR / "violations.db"
BACKUP_DIR = APP_DIR / "backup"
EXPORT_DIR = APP_DIR / "exports"
ERROR_LOG = APP_DIR / "error_log.txt"

TEMPLATE_NAME = "شيت_مثالي.xlsx"            # الشيت المثالي
LIST_FILES = ["ع_ع.xlsx", "ع_ل.xlsx"]       # قوائم الأسماء

# أماكن الخانات داخل الشيت المثالي
T_FIRST_ROW, T_LAST_ROW = 6, 24             # صفوف المخالفات (19 صف)
T_CAPACITY = T_LAST_ROW - T_FIRST_ROW + 1
COL_DATE, COL_VIOL, COL_VERBAL, COL_WRITTEN, COL_REPORT = 6, 7, 8, 9, 10   # F G H I J(:K)
CELL_NAME, CELL_CLASS = "G3", "I3"
CHECK = "√"                                 # علامة الصح

# قائمة المخالفات الافتراضية (لو لم يُعثر على القائمة المنسدلة داخل الشيت المثالي)
DEFAULT_VIOLATIONS = [
    "حضور متاخر للطابور", "زي مخالف", "تنمر على زميله", "ضرب زميله",
    "شتم زميله", "مشاغبة في الفصل", "عدم سماع كلام المستر",
]
# مخالفات تُضاف دائماً لآخر القائمة المنسدلة (حتى لو القائمة داخل الشيت المثالي قديمة)
EXTRA_VIOLATIONS = [
    "مكياج",
    "مشاجرة",
    "حضور متاخر عن الحصة",
    "خناقة شديدة وتم الفصل 3 ايام",
    "خناقة شديدة وتم الفصل 7 ايام",
]
OTHER = "أخرى (اكتب المخالفة)…"
ALL_CLASSES = "كل الفصول"

WEEKDAYS = ["الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "الأحد"]


# ----------------------------------------------------------------------------
# دوال مساعدة للنصوص والتواريخ
# ----------------------------------------------------------------------------
_DIACRITICS = re.compile("[\u064B-\u0652\u0670\u0640]")   # التشكيل والتطويل
_AR_FIX = str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ى": "ي", "ة": "ه", "ؤ": "و", "ئ": "ي"})
_AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")


def norm(text):
    """توحيد النص العربي للبحث (أ/إ/آ=ا ، ى=ي ، ة=ه ، حذف التشكيل)."""
    t = _DIACRITICS.sub("", str(text)).translate(_AR_FIX)
    return " ".join(t.split()).lower()


def clean_name(text):
    return " ".join(str(text).replace("\u0640", "").split())


def parse_date(text):
    """يقبل 2026/10/01 أو 01/10/2026 (وأرقام عربية). يرجع None لو التنسيق غلط."""
    t = str(text).translate(_AR_DIGITS).strip()
    m = re.fullmatch(r"(\d{1,4})\D+(\d{1,2})\D+(\d{1,4})", t)
    if not m:
        return None
    a, b, c = (int(x) for x in m.groups())
    try:
        if len(m.group(1)) == 4:
            return dt.date(a, b, c)
        if len(m.group(3)) == 4:
            return dt.date(c, b, a)
    except ValueError:
        pass
    return None


def fmt_date(d):
    return f"{WEEKDAYS[d.weekday()]}  {d:%Y/%m/%d}"


def safe_filename(text):
    return re.sub(r'[\\/:*?"<>|]+', " ", text).strip()[:80] or "student"


def open_file(path):
    if sys.platform.startswith("win"):
        os.startfile(str(path))
    elif sys.platform == "darwin":
        subprocess.run(["open", str(path)], check=False)
    else:
        subprocess.run(["xdg-open", str(path)], check=False)


# ----------------------------------------------------------------------------
# قاعدة البيانات (SQLite)
# ----------------------------------------------------------------------------
SCHEMA = """
CREATE TABLE IF NOT EXISTS students (
    id   INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    cls  TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS records (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
    rdate      TEXT NOT NULL,
    violation  TEXT NOT NULL,
    verbal     INTEGER NOT NULL DEFAULT 0,
    written    INTEGER NOT NULL DEFAULT 0,
    report     INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_records_student ON records(student_id);
"""


def backup_db():
    """نسخة احتياطية يومية من قاعدة البيانات (آخر 30 نسخة)."""
    if not DB_PATH.exists():
        return
    BACKUP_DIR.mkdir(exist_ok=True)
    target = BACKUP_DIR / f"violations_{dt.date.today():%Y%m%d}.db"
    if target.exists():
        return
    src, dst = sqlite3.connect(DB_PATH), sqlite3.connect(target)
    with dst:
        src.backup(dst)
    src.close()
    dst.close()
    for old in sorted(BACKUP_DIR.glob("violations_*.db"))[:-30]:
        old.unlink()


class DB:
    def __init__(self, path):
        self.con = sqlite3.connect(path)
        self.con.row_factory = sqlite3.Row
        self.con.execute("PRAGMA foreign_keys = ON")
        self.con.executescript(SCHEMA)

    def close(self):
        self.con.close()

    # ---- الطلاب
    def student_count(self):
        return self.con.execute("SELECT COUNT(*) FROM students").fetchone()[0]

    def students_with_counts(self):
        return self.con.execute(
            """SELECT s.id, s.name, s.cls, COUNT(r.id) AS n
               FROM students s LEFT JOIN records r ON r.student_id = s.id
               GROUP BY s.id ORDER BY s.id"""
        ).fetchall()

    def student(self, sid):
        return self.con.execute("SELECT * FROM students WHERE id = ?", (sid,)).fetchone()

    def classes(self):
        rows = self.con.execute("SELECT cls FROM students GROUP BY cls ORDER BY MIN(id)").fetchall()
        return [r["cls"] for r in rows]

    def find_student(self, name, cls):
        return self.con.execute("SELECT id FROM students WHERE name = ? AND cls = ?", (name, cls)).fetchone()

    def add_student(self, name, cls):
        with self.con:
            cur = self.con.execute("INSERT INTO students(name, cls) VALUES (?, ?)", (name, cls))
        return cur.lastrowid

    def delete_student(self, sid):
        with self.con:
            self.con.execute("DELETE FROM students WHERE id = ?", (sid,))

    def import_students(self, entries):
        """يضيف الأسماء الجديدة فقط (لا يكرر الاسم الموجود في نفس الفصل). يرجع (المضاف, المتخطّى)."""
        have = Counter((r["name"], r["cls"]) for r in self.con.execute("SELECT name, cls FROM students"))
        added = skipped = 0
        with self.con:
            for name, cls in entries:
                if have[(name, cls)] > 0:
                    have[(name, cls)] -= 1
                    skipped += 1
                else:
                    self.con.execute("INSERT INTO students(name, cls) VALUES (?, ?)", (name, cls))
                    added += 1
        return added, skipped

    # ---- المخالفات
    def records(self, sid):
        return self.con.execute(
            "SELECT * FROM records WHERE student_id = ? ORDER BY rdate, id", (sid,)
        ).fetchall()

    def record(self, rid):
        return self.con.execute("SELECT * FROM records WHERE id = ?", (rid,)).fetchone()

    def add_record(self, sid, rdate, violation):
        with self.con:
            cur = self.con.execute(
                "INSERT INTO records(student_id, rdate, violation) VALUES (?, ?, ?)", (sid, rdate, violation)
            )
        return cur.lastrowid

    def update_record(self, rid, **fields):
        allowed = {"rdate", "violation", "verbal", "written", "report"}
        fields = {k: v for k, v in fields.items() if k in allowed}
        if not fields:
            return
        sets = ", ".join(f"{k} = ?" for k in fields)
        with self.con:
            self.con.execute(f"UPDATE records SET {sets} WHERE id = ?", (*fields.values(), rid))

    def delete_record(self, rid):
        with self.con:
            self.con.execute("DELETE FROM records WHERE id = ?", (rid,))


# ----------------------------------------------------------------------------
# استيراد قوائم الأسماء من ملفات Excel
# ----------------------------------------------------------------------------
_ORD = {"الأول": "الأول", "الاول": "الأول", "الثانى": "الثاني", "الثاني": "الثاني", "الثالث": "الثالث",
        "الرابع": "الرابع", "الخامس": "الخامس", "السادس": "السادس"}
_ORD_BY_NUM = {1: "الأول", 2: "الثاني", 3: "الثالث", 4: "الرابع", 5: "الخامس", 6: "السادس"}
_HEADER_WORDS = {norm("إسم التلميذ"), norm("اسم الطالب"), norm("اسم التلميذ")}


def _parse_dept(heading):
    m = re.search(r"قسم\s+(.+?)\s+للعام", heading)
    if not m:
        return ""
    d = m.group(1).strip().replace("ى", "ي")
    return d[2:] if d.startswith("ال") and len(d) > 4 else d          # العربى -> عربي


def _parse_class(cell, dept):
    c = clean_name(cell)
    m = re.fullmatch(r"(\d)\s*P\s*/\s*(\d+)", c, flags=re.I)           # مثال: 1 P / 2  (P = إعدادي)
    if m and int(m.group(1)) in _ORD_BY_NUM:
        return f"{_ORD_BY_NUM[int(m.group(1))]} الإعدادي ({dept} {m.group(2)})".replace("( ", "(")
    m = re.match(r"(\S+)\s+ال[اإ]عدا", c)                              # مثال: الأول الاعداى
    if m:
        grade = _ORD.get(m.group(1), m.group(1))
        return f"{grade} الإعدادي ({dept})" if dept else f"{grade} الإعدادي"
    return f"{c} ({dept})" if dept else c


def read_student_lists(path):
    """يقرأ ملف قوائم (عنوان القائمة ثم عمودين أسماء) ويرجع [(الاسم, الفصل), ...]."""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    entries, cur_cls, cols = [], None, {}

    def flush():
        if cur_cls:
            for c in sorted(cols):
                entries.extend((n, cur_cls) for n in cols[c])

    for row in ws.iter_rows(values_only=True):
        cells = [(i, clean_name(v)) for i, v in enumerate(row) if isinstance(v, str) and clean_name(v)]
        if not cells:
            continue
        heading = next((t for _, t in cells if "قوائم" in t), None)
        if heading:                                                    # بداية قائمة جديدة
            flush()
            dept = _parse_dept(heading)
            other = next((t for _, t in cells if "قوائم" not in t), "")
            cur_cls, cols = _parse_class(other, dept), {}
            continue
        for i, t in cells:
            if cur_cls and norm(t) not in _HEADER_WORDS:
                cols.setdefault(i, []).append(t)
    flush()
    wb.close()
    return entries


# ----------------------------------------------------------------------------
# الشيت المثالي: قراءة قائمة المخالفات + التصدير
# ----------------------------------------------------------------------------
def merge_violations(items):
    """يضيف المخالفات الجديدة (EXTRA_VIOLATIONS) لآخر القائمة لو مش موجودة فيها."""
    items = ["مشاغبة في الفصل" if norm(x) == norm("مشغابة في الفصل") else x for x in items]   # تصحيح الإملاء لو القائمة القديمة داخل الشيت
    have = {norm(x) for x in items}
    items.extend(x for x in EXTRA_VIOLATIONS if norm(x) not in have)
    return items


def load_violation_choices(template_path=None):
    """القائمة المنسدلة تُقرأ من الشيت المثالي نفسه (العمود G) فلو عدّلتها هناك يلتقطها البرنامج.
    والمخالفات الجديدة (مكياج / خناقة شديدة...) بتتضاف لآخرها تلقائياً."""
    items = None
    if template_path:
        try:
            ws = load_workbook(template_path).worksheets[0]
            first = f"G{T_FIRST_ROW}"
            for dv in ws.data_validations.dataValidation:
                if dv.type == "list" and dv.formula1 and dv.formula1.startswith('"') and first in dv.sqref:
                    found = [x.strip() for x in dv.formula1.strip('"').split(",") if x.strip()]
                    if found:
                        items = found
                        break
        except Exception:
            pass
    return merge_violations(items or DEFAULT_VIOLATIONS)


def _col_width(ws, idx):
    for cd in ws.column_dimensions.values():
        if cd.min and cd.max and cd.min <= idx <= cd.max and cd.width:
            return cd.width
    return ws.sheet_format.defaultColWidth or 8.43


def _fit_pt(text, width_chars, max_pt=20, min_pt=9):
    """حجم خط تقريبي يخلّي النص يتسع في الخلية."""
    need = len(text) * 0.55 * 1.333
    if not need:
        return max_pt
    return max(min_pt, min(max_pt, int((width_chars * 7 + 5) / need)))


def _set_font_size(cell, size):
    f = copy(cell.font)
    f.sz = size
    cell.font = f
    a = copy(cell.alignment)
    a.shrink_to_fit = True
    a.horizontal = "center"
    cell.alignment = a


def _refresh_dropdown(ws, items):
    """يحدّث القائمة المنسدلة للمخالفات داخل ملف Excel المصدَّر لتشمل المخالفات الجديدة.
    (Excel يقبل 255 حرف فقط في القائمة المكتوبة داخل الخلية، فلو زادت نسيبها زي ما هي.)"""
    text = ",".join(items)
    if len(text) > 255:
        return
    first = f"G{T_FIRST_ROW}"
    for dv in ws.data_validations.dataValidation:
        if dv.type == "list" and dv.formula1 and dv.formula1.startswith('"') and first in dv.sqref:
            dv.formula1 = f'"{text}"'


def export_student(template_path, name, cls, records, out_path, violations=None):
    """يملأ نسخة من الشيت المثالي باسم الطالب وفصله ومخالفاته ويحفظها في out_path.
    لو المخالفات أكتر من 19 صف يُضاف ورقة تكملة بنفس التنسيق."""
    wb = load_workbook(template_path)
    base = wb.worksheets[0]
    if violations:
        _refresh_dropdown(base, violations)

    wb._external_links = []                      # فصل الرابط الخارجي (قوائم_عربي_مكتملة.xlsx)
    base["A1"].hyperlink = None                  # زر "رجوع لقائمة الأسماء" الأصلي لا معنى له في ملف مستقل
    base["A1"].value = None
    base._hyperlinks = []

    pages = max(1, math.ceil(len(records) / T_CAPACITY))
    sheets = [base]
    for p in range(2, pages + 1):                # نسخ الشيت وهو فارغ قبل التعبئة
        ws = wb.copy_worksheet(base)
        ws.title = str(p)
        ws.sheet_view.rightToLeft = base.sheet_view.rightToLeft
        ws.print_area = f"E1:K{T_LAST_ROW + 2}"
        for dv in base.data_validations.dataValidation:
            ws.add_data_validation(copy(dv))
        sheets.append(ws)

    center = Alignment(horizontal="center", vertical="center")
    for page, ws in enumerate(sheets):
        # الاسم والفصل (قيم ثابتة بدل الروابط الخارجية)
        ws.merge_cells("I3:K3")                  # نوسّع خانة الفصل على الخانات الفارغة جنبها
        ws[CELL_NAME].value = name
        ws[CELL_CLASS].value = cls
        _set_font_size(ws[CELL_NAME], _fit_pt(name, _col_width(ws, 7)))
        _set_font_size(ws[CELL_CLASS], _fit_pt(cls, sum(_col_width(ws, c) for c in (9, 10, 11))))

        chunk = records[page * T_CAPACITY:(page + 1) * T_CAPACITY]
        for i in range(T_CAPACITY):
            r = T_FIRST_ROW + i
            for c in (COL_DATE, COL_VIOL, COL_VERBAL, COL_WRITTEN, COL_REPORT):
                ws.cell(r, c).value = None       # يمسح معادلة التاريخ الدائرية الأصلية
            if i < len(chunk):
                rec = chunk[i]
                ws.cell(r, COL_DATE).value = dt.date.fromisoformat(rec["rdate"])
                ws.cell(r, COL_VIOL).value = rec["violation"]
                for col, key in ((COL_VERBAL, "verbal"), (COL_WRITTEN, "written"), (COL_REPORT, "report")):
                    if rec[key]:
                        ws.cell(r, col).value = CHECK
                        ws.cell(r, col).alignment = center

    wb.save(out_path)


# ----------------------------------------------------------------------------
# الواجهة الرسومية
# ----------------------------------------------------------------------------
FONT = "Tahoma" if sys.platform.startswith("win") else "Arial"
BG, PRIMARY, DARK = "#f3f5f9", "#1f4e79", "#17375e"
GREEN, BLUE, RED, GRAY, PURPLE = "#2e7d32", "#1565c0", "#c62828", "#546e7a", "#6a1b9a"
PINK = "#f8d7da"

# أعمدة جدول الشيت (الترتيب من اليمين لليسار: رقم أكبر = أقرب لليمين)
C_DATE, C_VIOL, C_VERBAL, C_WRITTEN, C_REPORT, C_DEL = 5, 4, 3, 2, 1, 0
MARKS = ((C_VERBAL, "verbal"), (C_WRITTEN, "written"), (C_REPORT, "report"))


# ----------------------------------------------------------------------------
# خانة كتابة عربية (يمين ← يسار) بتحديد صحيح بالماوس
# ----------------------------------------------------------------------------
USE_CUSTOM_ENTRY = True      # لو حصلت أي مشكلة غريبة في خانات الكتابة غيّرها لـ False يرجع الشكل القديم
_RTL_CHARS = re.compile("[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]")
SEL_BG = "#bbdefb"


class RTLEntry(tk.Canvas):
    """بديل لخانة tk.Entry للنص العربي.

    المشكلة: Entry العادية في Tk بتحسب مكان الماوس والتحديد من الشمال لليمين حتى لو النص عربي،
    فلما تضغط على كلمة «محمد» في «احمد محمد» بيتحدد «احمد» (والعكس).
    هنا كل الحسابات (الماوس، التحديد، المؤشر، الأسهم) بتتعمل من اليمين لليسار للنص العربي.
    """
    PAD = 7

    def __init__(self, parent, textvariable=None, font=None, width=30, fg="black", bg="white"):
        self._font = tkfont.Font(root=parent, font=font) if font else tkfont.nametofont("TkDefaultFont")
        self._width = max(80, self._font.measure("0") * width + 2 * self.PAD)
        self._height = self._font.metrics("linespace") + 14
        super().__init__(parent, width=self._width, height=self._height, bg=bg, bd=0, highlightthickness=0,
                         cursor="xterm", takefocus=1)
        self._fg = fg
        self._text, self._pos, self._anchor = "", 0, None       # النص ومكان المؤشر وبداية التحديد
        self._off, self._xs = 0, [0]
        self._focused = self._dragging = False
        self._caret_on, self._blink_job = True, None
        self._var, self._busy, self._trace_id = textvariable, False, None
        if self._var is not None:
            self._text = self._var.get()
            self._pos = len(self._text)
            self._trace_id = self._var.trace_add("write", self._var_changed)

        self.bind("<Button-1>", self._on_press)
        self.bind("<B1-Motion>", self._on_drag)
        self.bind("<ButtonRelease-1>", lambda e: setattr(self, "_dragging", False))
        self.bind("<Double-Button-1>", self._on_double)
        self.bind("<Triple-Button-1>", lambda e: self._select_all())
        self.bind("<Key>", self._on_key)
        self.bind("<FocusIn>", lambda e: self._set_focus(True))
        self.bind("<FocusOut>", lambda e: self._set_focus(False))
        self.bind("<Configure>", self._on_configure)
        self.bind("<Destroy>", self._on_destroy)
        self._redraw()

    # ---- واجهة شبيهة بـ Entry
    def get(self):
        return self._text

    def select_range(self, first, last):
        n = len(self._text)
        a = n if first == "end" else int(first)
        b = n if last == "end" else int(last)
        self._anchor, self._pos = max(0, min(a, n)), max(0, min(b, n))
        self._touch()

    # ---- حسابات الاتجاه
    def _is_rtl(self):
        return not self._text or bool(_RTL_CHARS.search(self._text))

    def _selection(self):
        if self._anchor is None or self._anchor == self._pos:
            return None
        return min(self._anchor, self._pos), max(self._anchor, self._pos)

    def _index_at(self, x):
        """أقرب فاصل بين حرفين لموضع x (بالبكسل)."""
        return min(range(len(self._xs)), key=lambda k: abs(self._xs[k] - x))

    def _char_at(self, x):
        n = len(self._text)
        for k in range(n):
            lo, hi = sorted((self._xs[k], self._xs[k + 1]))
            if lo <= x <= hi:
                return k
        return min(self._index_at(x), n - 1)

    # ---- الرسم
    def _redraw(self):
        self.delete("all")
        w, h, pad, t = self._width, self._height, self.PAD, self._text
        n, measure = len(t), self._font.measure
        total = measure(t) if t else 0
        right = w - pad
        if self._is_rtl():                          # أول حرف عند اليمين
            xs = [right - measure(t[:k]) for k in range(n + 1)]
        else:                                       # نص إنجليزي/أرقام: من الشمال لكن محاذي لليمين
            xs = [right - total + measure(t[:k]) for k in range(n + 1)]

        if total <= w - 2 * pad:                    # النص داخل الخانة
            self._off = 0
        else:                                       # النص أطول من الخانة: نحرّكه عشان المؤشر يفضل ظاهر
            cx = xs[self._pos] + self._off
            if cx < pad:
                self._off += pad - cx
            elif cx > right:
                self._off -= cx - right
        self._xs = [x + self._off for x in xs]

        focused = self._focused
        self.create_rectangle(1, 1, w - 2, h - 2, outline=PRIMARY if focused else "#9aa5b1",
                              width=2 if focused else 1)
        sel = self._selection()
        if sel:
            x1, x2 = sorted((self._xs[sel[0]], self._xs[sel[1]]))
            self.create_rectangle(x1, 5, x2, h - 5, fill=SEL_BG, outline="")
        if t:
            self.create_text(right + self._off, h // 2, text=t, anchor="e", font=self._font, fill=self._fg)
        cx = self._xs[self._pos]
        self.create_line(cx, 5, cx, h - 5, fill="black", tags="caret",
                         state="normal" if (focused and self._caret_on) else "hidden")

    def _touch(self):
        """إعادة رسم + إظهار المؤشر (بعد أي حركة أو كتابة)."""
        self._caret_on = True
        self._restart_blink()
        self._redraw()

    def _restart_blink(self):
        if self._blink_job:
            self.after_cancel(self._blink_job)
            self._blink_job = None
        if self._focused:
            self._blink_job = self.after(530, self._blink)

    def _blink(self):
        self._blink_job = None
        if not self._focused or not self.winfo_exists():
            return
        self._caret_on = not self._caret_on
        self.itemconfigure("caret", state="normal" if self._caret_on else "hidden")
        self._blink_job = self.after(530, self._blink)

    def _set_focus(self, on):
        self._focused = on
        if on:
            self._touch()
        else:
            self._restart_blink()
            self._redraw()

    def _on_configure(self, e):
        self._width, self._height = e.width, e.height
        self._redraw()

    def _on_destroy(self, e):
        if e.widget is not self:
            return
        if self._blink_job:
            self.after_cancel(self._blink_job)
        if self._var is not None and self._trace_id:
            try:
                self._var.trace_remove("write", self._trace_id)
            except tk.TclError:
                pass

    # ---- تعديل النص
    def _set_text(self, text, pos):
        self._text, self._pos, self._anchor = text, pos, None
        if self._var is not None:
            self._busy = True
            try:
                self._var.set(text)
            finally:
                self._busy = False
        self._touch()

    def _var_changed(self, *_):
        if self._busy or not self.winfo_exists():
            return
        self._text = self._var.get()
        self._pos = min(self._pos, len(self._text))
        self._anchor = None
        self._touch()

    def _replace(self, a, b, s):
        self._set_text(self._text[:a] + s + self._text[b:], a + len(s))

    def _insert(self, s):
        a, b = self._selection() or (self._pos, self._pos)
        self._replace(a, b, s)

    def _delete_selection(self):
        sel = self._selection()
        if sel:
            self._replace(sel[0], sel[1], "")
        return bool(sel)

    def _select_all(self):
        self._anchor, self._pos = 0, len(self._text)
        self._touch()

    def _select_word(self, c):
        t = self._text
        if not t.strip():
            return
        c = max(0, min(c, len(t) - 1))
        while c > 0 and t[c].isspace():              # الضغطة وقعت على مسافة: نروح لأقرب كلمة
            c -= 1
        a, b = c, c + 1
        while a > 0 and not t[a - 1].isspace():
            a -= 1
        while b < len(t) and not t[b].isspace():
            b += 1
        self._anchor, self._pos = a, b
        self._touch()

    def _move(self, new_pos, extend):
        if extend:
            if self._anchor is None:
                self._anchor = self._pos
        else:
            self._anchor = None
        self._pos = max(0, min(new_pos, len(self._text)))
        self._touch()

    # ---- الماوس
    def _on_press(self, e):
        self.focus_set()
        i = self._index_at(e.x)
        if e.state & 0x1 and self._anchor is None:   # Shift + ضغطة: يمدّ التحديد من مكان المؤشر
            self._anchor = self._pos
        elif not e.state & 0x1:
            self._anchor = i
        self._pos = i
        self._dragging = True
        self._touch()

    def _on_drag(self, e):
        if self._dragging:
            self._pos = self._index_at(e.x)
            self._touch()

    def _on_double(self, e):
        self.focus_set()
        self._dragging = False
        self._select_word(self._char_at(e.x))

    # ---- لوحة المفاتيح
    def _copy(self):
        sel = self._selection()
        if sel:
            self.clipboard_clear()
            self.clipboard_append(self._text[sel[0]:sel[1]])
        return bool(sel)

    def _paste(self):
        try:
            s = self.clipboard_get()
        except tk.TclError:
            return
        s = re.sub(r"[\r\n\t]+", " ", s)
        if s:
            self._insert(s)

    def _on_key(self, e):
        ks = e.keysym
        ctrl, shift = bool(e.state & 0x4), bool(e.state & 0x1)
        rtl, n = self._is_rtl(), len(self._text)

        if ctrl:
            key = ks.lower() if len(ks) == 1 else ""
            if sys.platform.startswith("win") and key not in ("a", "c", "v", "x"):
                key = {65: "a", 67: "c", 86: "v", 88: "x"}.get(e.keycode, key)   # شغّال حتى والكيبورد عربي
            if key == "a":
                self._select_all()
            elif key == "c" or ks == "Insert":
                self._copy()
            elif key == "x":
                if self._copy():
                    self._delete_selection()
            elif key == "v":
                self._paste()
            else:
                return None
            return "break"
        if shift and ks == "Insert":
            self._paste()
            return "break"

        if ks == "BackSpace":
            if not self._delete_selection() and self._pos > 0:
                self._replace(self._pos - 1, self._pos, "")
        elif ks == "Delete":
            if not self._delete_selection() and self._pos < n:
                self._replace(self._pos, self._pos + 1, "")
        elif ks in ("Left", "Right"):
            # في النص العربي «يسار» على الشاشة = الحرف التالي في النص، و«يمين» = الحرف السابق
            step = 1 if rtl else -1
            if ks == "Right":
                step = -step
            sel = self._selection()
            if sel and not shift:
                a, b = sel
                left_edge, right_edge = (b, a) if rtl else (a, b)
                self._move(left_edge if ks == "Left" else right_edge, False)
            else:
                self._move(self._pos + step, shift)
        elif ks == "Home":
            self._move(0, shift)
        elif ks == "End":
            self._move(n, shift)
        elif e.char and all(ch.isprintable() for ch in e.char):
            self._insert(e.char)
        else:
            return None                              # Return / Escape / Tab... تكمل لباقي البرنامج
        return "break"


def arabic_entry(parent, textvariable, width):
    """خانة كتابة للأسماء العربية (RTLEntry، أو Entry العادية لو USE_CUSTOM_ENTRY = False)."""
    if USE_CUSTOM_ENTRY:
        return RTLEntry(parent, textvariable=textvariable, font=(FONT, 13), width=width)
    return tk.Entry(parent, textvariable=textvariable, justify="right", font=(FONT, 13), width=width)


class App:
    def __init__(self, root):
        self.root = root
        self.db = DB(DB_PATH)
        self.template = self.find_template()
        self.violations = load_violation_choices(self.template)
        self.current = None          # رقم الطالب المفتوح حالياً (None = شاشة الأسماء)
        self.last_opened = None
        self.delete_mode = False
        self._render_job = None

        root.title("سجل المخالفات والعقوبات — قرار 150 لسنة 2024")
        root.configure(bg=BG)
        root.report_callback_exception = self.on_error
        self.center(1120, 740)
        root.minsize(940, 600)

        self.build_style()
        self.build_menu()
        self.build_toolbar()
        self.body = tk.Frame(root, bg=BG)
        self.body.pack(fill="both", expand=True)
        self.build_names_view()
        self.build_sheet_view()

        root.bind("<Escape>", lambda e: self.show_names() if self.current is not None else None)
        root.bind_all("<MouseWheel>", self.on_wheel)
        root.bind_all("<Button-4>", self.on_wheel)       # لينكس
        root.bind_all("<Button-5>", self.on_wheel)
        root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.show_names()
        root.after(300, self.first_run_import)

    # ------------------------------------------------------------- عام
    def center(self, w, h):
        self.root.update_idletasks()
        x = max(0, (self.root.winfo_screenwidth() - w) // 2)
        y = max(0, (self.root.winfo_screenheight() - h) // 2 - 20)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def on_error(self, exc, val, tb):
        msg = "".join(traceback.format_exception(exc, val, tb))
        try:
            with open(ERROR_LOG, "a", encoding="utf-8") as f:
                f.write(f"\n[{dt.datetime.now():%Y-%m-%d %H:%M:%S}]\n{msg}")
        except Exception:
            pass
        messagebox.showerror("حدث خطأ", f"حدث خطأ غير متوقع:\n{val}\n\n(التفاصيل محفوظة في error_log.txt)")

    def on_close(self):
        try:
            self.db.close()
        finally:
            self.root.destroy()

    def on_wheel(self, event):
        if self.current is None:
            return
        num = getattr(event, "num", None)
        step = -1 if num == 4 else 1 if num == 5 else (-1 if event.delta > 0 else 1)
        self.canvas.yview_scroll(step * 2, "units")

    def make_button(self, parent, text, cmd, color):
        return tk.Button(parent, text=text, command=cmd, bg=color, fg="white", activebackground=color,
                         activeforeground="white", disabledforeground="#9fb3c8", relief="flat", bd=0,
                         font=(FONT, 11, "bold"), padx=16, pady=7, cursor="hand2")

    def build_style(self):
        st = ttk.Style()
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("Treeview", rowheight=30, font=(FONT, 12), background="white", fieldbackground="white")
        st.configure("Treeview.Heading", font=(FONT, 11, "bold"), background=PRIMARY, foreground="white")
        st.map("Treeview.Heading", background=[("active", DARK)])
        st.map("Treeview", background=[("selected", "#bbdefb")], foreground=[("selected", "black")])
        self.root.option_add("*TCombobox*Listbox.font", f"{FONT} 12")

    def build_menu(self):
        bar = tk.Menu(self.root)
        m = tk.Menu(bar, tearoff=0)
        m.add_command(label="استيراد قوائم أسماء من Excel…", command=self.import_dialog)
        m.add_command(label="فتح مجلد الشيتات المصدّرة", command=self.open_exports)
        m.add_separator()
        m.add_command(label="خروج", command=self.on_close)
        bar.add_cascade(label="ملف", menu=m)
        self.root.config(menu=bar)

    def build_toolbar(self):
        bar = tk.Frame(self.root, bg=DARK)
        bar.pack(fill="x")
        kw = dict(side="right", padx=(0, 8), pady=8)
        self.btn_search = self.make_button(bar, "بحث", self.focus_search, BLUE)
        self.btn_back = self.make_button(bar, "رجوع للأسماء", self.show_names, GRAY)
        self.btn_add = self.make_button(bar, "إضافة اسم جديد", self.add_student_dialog, GREEN)
        self.btn_del = self.make_button(bar, "حذف الاسم", self.delete_action, RED)
        self.btn_search.pack(side="right", padx=(12, 8), pady=8)
        for b in (self.btn_back, self.btn_add, self.btn_del):
            b.pack(**kw)
        self.btn_export = self.make_button(bar, "فتح الشيت في Excel", self.export_current, PURPLE)
        self.btn_export.pack(side="left", padx=12, pady=8)

    def update_buttons(self):
        in_sheet = self.current is not None
        self.btn_back.config(state="normal" if in_sheet else "disabled")
        self.btn_export.config(state="normal" if in_sheet else "disabled")
        self.btn_del.config(text="إلغاء وضع الحذف" if (self.delete_mode and not in_sheet) else "حذف الاسم")

    # ------------------------------------------------------------- شاشة الأسماء
    def build_names_view(self):
        f = self.names_frame = tk.Frame(self.body, bg=BG)

        top = tk.Frame(f, bg=BG)
        top.pack(fill="x", padx=12, pady=(12, 6))
        tk.Label(top, text="بحث عن اسم:", font=(FONT, 12, "bold"), bg=BG).pack(side="right")
        self.search_var = tk.StringVar()
        self.search_entry = arabic_entry(top, self.search_var, 30)
        self.search_entry.pack(side="right", padx=6)
        self.search_entry.bind("<KeyRelease>", lambda e: self.refresh_names())
        self.search_entry.bind("<Return>", lambda e: self.refresh_names())
        tk.Button(top, text="مسح", command=self.clear_search, relief="flat", bg="#dfe5ec", font=(FONT, 10),
                  cursor="hand2").pack(side="right", padx=2)
        tk.Label(top, text="الفصل:", font=(FONT, 12, "bold"), bg=BG).pack(side="right", padx=(22, 0))
        self.class_var = tk.StringVar(value=ALL_CLASSES)
        self.class_cb = ttk.Combobox(top, textvariable=self.class_var, state="readonly", justify="right",
                                     font=(FONT, 12), width=26)
        self.class_cb.pack(side="right", padx=6)
        self.class_cb.bind("<<ComboboxSelected>>", lambda e: self.refresh_names())
        self.class_cb.bind("<MouseWheel>", lambda e: "break")        # يمنع تغيير الفصل بالغلط بعجلة الماوس
        self.count_lbl = tk.Label(top, text="", font=(FONT, 11), bg=BG, fg=GRAY)
        self.count_lbl.pack(side="left")

        self.banner = tk.Label(f, text="وضع الحذف: اضغط على الاسم الذي تريد حذفه (أو اضغط «إلغاء وضع الحذف»)",
                               bg="#ffebee", fg=RED, font=(FONT, 12, "bold"), pady=6)

        self.tree_frame = tf = tk.Frame(f, bg=BG)
        tf.pack(fill="both", expand=True, padx=12, pady=(0, 4))
        cols = ("count", "cls", "name", "num")                      # بترتيب معكوس عشان يظهر من اليمين
        self.tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse")
        heads = {"num": "م", "name": "اسم الطالب", "cls": "الفصل", "count": "عدد المخالفات"}
        sizes = {"num": 60, "name": 420, "cls": 230, "count": 130}
        for c in cols:
            self.tree.heading(c, text=heads[c], anchor="e" if c in ("name", "cls") else "center")
            self.tree.column(c, width=sizes[c], minwidth=50, stretch=(c == "name"),
                             anchor="e" if c in ("name", "cls") else "center")
        self.tree.tag_configure("has", foreground="#b71c1c")
        sb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="left", fill="y")
        self.tree.pack(side="right", fill="both", expand=True)
        self.tree.bind("<ButtonRelease-1>", self.on_tree_click)
        self.tree.bind("<Return>", self.on_tree_enter)

        tk.Label(f, text="اضغط على اسم الطالب لفتح سجل مخالفاته", font=(FONT, 10), bg=BG, fg=GRAY).pack(pady=(0, 8))

    def refresh_class_filter(self):
        vals = [ALL_CLASSES] + self.db.classes()
        self.class_cb.config(values=vals)
        if self.class_var.get() not in vals:
            self.class_var.set(ALL_CLASSES)

    def refresh_names(self):
        tokens = norm(self.search_var.get()).split()
        cf = self.class_var.get()
        self.tree.delete(*self.tree.get_children())
        shown = 0
        for s in self.db.students_with_counts():
            if cf != ALL_CLASSES and s["cls"] != cf:
                continue
            if tokens:
                n = norm(s["name"])
                if not all(t in n for t in tokens):
                    continue
            shown += 1
            self.tree.insert("", "end", iid=str(s["id"]), values=(s["n"], s["cls"], s["name"], shown),
                             tags=("has",) if s["n"] else ())
        self.count_lbl.config(text=f"المعروض: {shown} من {self.db.student_count()}")
        if self.last_opened is not None and self.tree.exists(str(self.last_opened)):
            self.tree.selection_set(str(self.last_opened))
            self.tree.see(str(self.last_opened))

    def clear_search(self):
        self.search_var.set("")
        self.refresh_names()
        self.search_entry.focus_set()

    def focus_search(self):
        if self.current is not None:
            self.show_names()
        self.search_entry.focus_set()
        self.search_entry.select_range(0, "end")
        self.refresh_names()

    def on_tree_click(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        iid = self.tree.identify_row(event.y)
        if iid:
            self.activate_student(int(iid))

    def on_tree_enter(self, event):
        sel = self.tree.selection()
        if sel:
            self.activate_student(int(sel[0]))

    def activate_student(self, sid):
        """ضغطة على اسم: تفتح الشيت (أو تحذف الاسم لو وضع الحذف شغّال)."""
        if self.delete_mode:
            self.confirm_delete(sid)
        else:
            self.open_student(sid)

    def show_names(self):
        if self.current is not None:
            self.last_opened = self.current
        self.current = None
        self.sheet_frame.pack_forget()
        self.names_frame.pack(fill="both", expand=True)
        self.set_delete_mode(False)
        self.refresh_class_filter()
        self.refresh_names()
        self.update_buttons()

    # ------------------------------------------------------------- الحذف
    def delete_action(self):
        if self.current is not None:
            self.confirm_delete(self.current)
        else:
            self.set_delete_mode(not self.delete_mode)

    def set_delete_mode(self, on):
        self.delete_mode = on
        if on:
            self.banner.pack(fill="x", padx=12, pady=(0, 6), before=self.tree_frame)
        else:
            self.banner.pack_forget()
        self.update_buttons()

    def confirm_delete(self, sid):
        s = self.db.student(sid)
        if not s:
            return
        n = len(self.db.records(sid))
        extra = f"\nومعه {n} مخالفة مسجّلة سيتم حذفها أيضاً." if n else ""
        ok = messagebox.askyesno(
            "تأكيد الحذف", f"هل تريد حذف الطالب:\n«{s['name']}»؟{extra}\n\nلا يمكن التراجع عن الحذف.",
            icon="warning", parent=self.root)
        if ok:
            self.db.delete_student(sid)
            if self.last_opened == sid:
                self.last_opened = None
            self.show_names()

    # ------------------------------------------------------------- إضافة اسم
    def add_student_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("إضافة اسم جديد")
        dlg.configure(bg=BG)
        dlg.transient(self.root)
        dlg.resizable(False, False)
        pad = dict(padx=14, pady=8)

        tk.Label(dlg, text="اسم الطالب:", bg=BG, font=(FONT, 12, "bold")).grid(row=0, column=1, sticky="e", **pad)
        name_var = tk.StringVar()
        entry = arabic_entry(dlg, name_var, 34)
        entry.grid(row=0, column=0, **pad)

        tk.Label(dlg, text="الفصل:", bg=BG, font=(FONT, 12, "bold")).grid(row=1, column=1, sticky="e", **pad)
        if self.current is not None:
            default_cls = self.db.student(self.current)["cls"]
        else:
            default_cls = "" if self.class_var.get() == ALL_CLASSES else self.class_var.get()
        cls_var = tk.StringVar(value=default_cls)
        ttk.Combobox(dlg, textvariable=cls_var, values=self.db.classes(), justify="right",
                     font=(FONT, 12), width=32).grid(row=1, column=0, **pad)
        tk.Label(dlg, text="اختر فصلاً من القائمة أو اكتب اسم فصل جديد", bg=BG, fg=GRAY,
                 font=(FONT, 10)).grid(row=2, column=0, columnspan=2)

        def save(event=None):
            name, cls = clean_name(name_var.get()), clean_name(cls_var.get())
            if not name:
                messagebox.showwarning("بيانات ناقصة", "اكتب اسم الطالب.", parent=dlg)
                return
            if not cls:
                messagebox.showwarning("بيانات ناقصة", "اختر الفصل أو اكتبه.", parent=dlg)
                return
            if self.db.find_student(name, cls) and not messagebox.askyesno(
                    "اسم مكرر", "هذا الاسم موجود بالفعل في نفس الفصل.\nهل تريد إضافته على أي حال؟", parent=dlg):
                return
            sid = self.db.add_student(name, cls)
            dlg.destroy()
            self.last_opened = sid
            self.search_var.set("")
            self.show_names()
            self.class_var.set(cls)               # يعرض فصل الطالب الجديد ويحدده في القائمة
            self.refresh_names()

        row = tk.Frame(dlg, bg=BG)
        row.grid(row=3, column=0, columnspan=2, pady=12)
        self.make_button(row, "حفظ", save, GREEN).pack(side="right", padx=6)
        self.make_button(row, "إلغاء", dlg.destroy, GRAY).pack(side="right", padx=6)
        dlg.bind("<Return>", save)
        dlg.bind("<Escape>", lambda e: dlg.destroy())
        entry.focus_set()
        try:
            dlg.grab_set()                     # يجعل النافذة مودال
        except tk.TclError:
            pass

    # ------------------------------------------------------------- استيراد
    def first_run_import(self):
        if self.db.student_count() > 0:
            return
        paths = [str(APP_DIR / n) for n in LIST_FILES if (APP_DIR / n).exists()]
        if not paths:
            messagebox.showinfo("قوائم الأسماء", "لا توجد أسماء في البرنامج بعد.\nاختر ملفات قوائم الأسماء (Excel) ليتم استيرادها.",
                                parent=self.root)
            paths = filedialog.askopenfilenames(title="اختر ملفات قوائم الأسماء", filetypes=[("Excel", "*.xlsx")])
        if paths:
            self.import_files(paths)

    def import_dialog(self):
        paths = filedialog.askopenfilenames(title="اختر ملفات قوائم الأسماء", filetypes=[("Excel", "*.xlsx")])
        if paths:
            self.import_files(paths)

    def import_files(self, paths):
        added = skipped = 0
        errors = []
        for p in paths:
            try:
                a, s = self.db.import_students(read_student_lists(p))
                added += a
                skipped += s
            except Exception as e:                                    # ملف بتنسيق غير متوقع
                errors.append(f"{Path(p).name}: {e}")
        self.show_names()
        msg = f"تم استيراد {added} اسم جديد."
        if skipped:
            msg += f"\n(تم تخطّي {skipped} اسم موجود بالفعل)"
        if errors:
            msg += "\n\nتعذّر قراءة:\n" + "\n".join(errors)
        messagebox.showinfo("الاستيراد", msg, parent=self.root)

    def open_exports(self):
        EXPORT_DIR.mkdir(exist_ok=True)
        open_file(EXPORT_DIR)

    # ------------------------------------------------------------- شاشة الشيت
    def build_sheet_view(self):
        f = self.sheet_frame = tk.Frame(self.body, bg=BG)
        tk.Label(f, text="سجل المخالفات والعقوبات طبقا للقرار رقم 150 لسنة 2024", font=(FONT, 18, "bold"),
                 bg=BG, fg=DARK).pack(pady=(10, 2))

        info = tk.Frame(f, bg=BG)
        info.pack(fill="x", padx=16, pady=4)
        tk.Label(info, text="اسم الطالب:", font=(FONT, 13, "bold"), bg=BG).pack(side="right")
        self.lbl_name = tk.Label(info, text="", font=(FONT, 16, "bold"), bg=BG, fg=PRIMARY)
        self.lbl_name.pack(side="right", padx=(6, 30))
        tk.Label(info, text="الفصل:", font=(FONT, 13, "bold"), bg=BG).pack(side="right")
        self.lbl_class = tk.Label(info, text="", font=(FONT, 16, "bold"), bg=BG, fg=PRIMARY)
        self.lbl_class.pack(side="right", padx=6)
        self.lbl_count = tk.Label(info, text="", font=(FONT, 11), bg=BG, fg=GRAY)
        self.lbl_count.pack(side="left")

        tk.Label(f, text="اختر المخالفة من القائمة في أول خانة فارغة ← يُسجَّل التاريخ تلقائياً.  "
                         "اضغط على التاريخ لتعديله، وعلى خانات العقوبة لوضع √ أو إزالتها.",
                 font=(FONT, 10), bg=BG, fg=GRAY).pack()

        foot = tk.Frame(f, bg=BG)                                      # خانات التوقيعات (مثل الشيت)
        foot.pack(side="bottom", fill="x", padx=16, pady=(6, 10))
        for t in ("معلم الفصل", "الاخصائي", "وكيل المدرسة", "مدير المدرسة"):
            tk.Label(foot, text=t, font=(FONT, 11, "bold"), bg=BG).pack(side="right", expand=True)

        wrap = tk.Frame(f, bg=BG)
        wrap.pack(fill="both", expand=True, padx=16, pady=(6, 0))
        self.canvas = tk.Canvas(wrap, bg=BG, highlightthickness=0)
        vs = ttk.Scrollbar(wrap, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vs.set)
        vs.pack(side="left", fill="y")
        self.canvas.pack(side="right", fill="both", expand=True)
        self.grid_frame = tk.Frame(self.canvas, bg=BG)
        self._win = self.canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")
        self.grid_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self._win, width=e.width))

    def open_student(self, sid):
        s = self.db.student(sid)
        if not s:
            return
        self.current = self.last_opened = sid
        self.lbl_name.config(text=s["name"])
        self.lbl_class.config(text=s["cls"])
        self.names_frame.pack_forget()
        self.sheet_frame.pack(fill="both", expand=True)
        self.canvas.yview_moveto(0)
        self.render_sheet()
        self.update_buttons()

    def schedule_render(self):
        """إعادة رسم الجدول بعد انتهاء الحدث الحالي (عشان ما نمسحش عنصر وهو بينفّذ حدثه)."""
        if self._render_job is None:
            self._render_job = self.root.after(10, self.render_sheet)

    def render_sheet(self):
        self._render_job = None
        if self.current is None:
            return
        g = self.grid_frame
        for w in g.winfo_children():
            w.destroy()
        recs = self.db.records(self.current)
        n_rows = max(T_CAPACITY, len(recs) + 1)
        choices = list(self.violations) + [OTHER]

        for c, width in {C_DATE: 170, C_VIOL: 330, C_VERBAL: 110, C_WRITTEN: 110, C_REPORT: 190, C_DEL: 44}.items():
            g.grid_columnconfigure(c, minsize=width, weight=1 if c == C_VIOL else 0)

        def head(text, r, c, rs=1, cs=1, wrap=0, bg=PINK, fg="#8b1a1a"):
            tk.Label(g, text=text, bg=bg, fg=fg, font=(FONT, 10, "bold"), relief="solid", bd=1,
                     wraplength=wrap, justify="center").grid(row=r, column=c, rowspan=rs, columnspan=cs, sticky="nsew")

        head("اليوم والتاريخ", 0, C_DATE, rs=2, bg="white", fg="black")
        head("وصف المخالفة", 0, C_VIOL, rs=2)
        head("العقوبات المتدرجة لمخالفات الدرجة الاولى", 0, C_WRITTEN, cs=2, wrap=220)
        head("التنبيه شفوى", 1, C_VERBAL)
        head("التنبيه كتابى", 1, C_WRITTEN)
        head("رفع تقارير للمدير للمخالفات المتوسطة او الخطيرة", 0, C_REPORT, rs=2, wrap=180)

        for i in range(n_rows):
            r = i + 2
            rec = recs[i] if i < len(recs) else None
            g.grid_rowconfigure(r, minsize=36)

            # التاريخ (اضغط عليه للتعديل)
            d = tk.Label(g, text=fmt_date(dt.date.fromisoformat(rec["rdate"])) if rec else "", bg="white",
                         font=(FONT, 11), relief="solid", bd=1, cursor="hand2" if rec else "")
            d.grid(row=r, column=C_DATE, sticky="nsew")
            if rec:
                d.bind("<Button-1>", lambda e, rid=rec["id"]: self.edit_date(rid))

            # وصف المخالفة (قائمة منسدلة) — الصف الفارغ الأول فقط يكون مفعّلاً
            cell = tk.Frame(g, bg="white", relief="solid", bd=1)
            cell.grid(row=r, column=C_VIOL, sticky="nsew")
            active = rec is not None or i == len(recs)
            cb = ttk.Combobox(cell, values=choices, state="readonly" if active else "disabled",
                              justify="right", font=(FONT, 12))
            if rec:
                cb.set(rec["violation"])
            cb.pack(fill="x", padx=2, pady=4)
            cb.bind("<<ComboboxSelected>>",
                    lambda e, rid=(rec["id"] if rec else None), w=cb: self.apply_violation(rid, w.get()))
            cb.bind("<MouseWheel>", lambda e: (self.on_wheel(e), "break")[1])   # العجلة تحرّك الصفحة ولا تغيّر المخالفة

            # خانات العقوبة
            for col, key in MARKS:
                m = tk.Label(g, text=CHECK if (rec and rec[key]) else "", bg="white", fg=GREEN,
                             font=(FONT, 14, "bold"), relief="solid", bd=1, cursor="hand2" if rec else "")
                m.grid(row=r, column=col, sticky="nsew")
                if rec:
                    m.bind("<Button-1>", lambda e, rid=rec["id"], k=key: self.toggle_mark(rid, k))

            # حذف السطر
            if rec:
                tk.Button(g, text="×", fg=RED, relief="flat", bg=BG, font=(FONT, 13, "bold"), cursor="hand2",
                          command=lambda rid=rec["id"]: self.remove_record(rid)).grid(row=r, column=C_DEL, sticky="nsew")

        self.lbl_count.config(text=f"عدد المخالفات المسجّلة: {len(recs)}")

    # ------------------------------------------------------------- أفعال الجدول
    def apply_violation(self, rid, value):
        """اختيار مخالفة من القائمة: صف جديد ← يسجّل بتاريخ اليوم، صف موجود ← يغيّر الوصف فقط."""
        if value == OTHER:
            current = self.db.record(rid)["violation"] if rid is not None else ""
            value = self.ask_text("مخالفة أخرى", "اكتب وصف المخالفة (أو عدّل الحالي):", initial=current)
            value = clean_name(value) if value else ""
        if value:
            if rid is None:
                self.db.add_record(self.current, dt.date.today().isoformat(), value)
            else:
                self.db.update_record(rid, violation=value)
        self.schedule_render()

    def ask_text(self, title, prompt, initial="", width=44):
        """نافذة إدخال بخانة عربية (يمين ← يسار) تدعم اللصق (Ctrl+V) والتحديد الصحيح بالماوس.
        لو فيه نص حالي بيظهر جاهز للتعديل (ومحدد بالكامل عشان لو كتبت فوقه يتبدّل). يرجع None عند الإلغاء."""
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.configure(bg=BG)
        dlg.transient(self.root)
        dlg.resizable(False, False)
        result = {"value": None}
        var = tk.StringVar(value=initial)

        tk.Label(dlg, text=prompt, bg=BG, font=(FONT, 12, "bold")).pack(anchor="e", padx=16, pady=(14, 4))
        entry = arabic_entry(dlg, var, width)
        entry.pack(padx=16, pady=6)

        def ok(event=None):
            result["value"] = var.get()
            dlg.destroy()

        row = tk.Frame(dlg, bg=BG)
        row.pack(pady=12)
        self.make_button(row, "حفظ", ok, GREEN).pack(side="right", padx=6)
        self.make_button(row, "إلغاء", dlg.destroy, GRAY).pack(side="right", padx=6)
        dlg.bind("<Return>", ok)
        dlg.bind("<Escape>", lambda e: dlg.destroy())

        dlg.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dlg.winfo_reqwidth()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dlg.winfo_reqheight()) // 3
        dlg.geometry(f"+{max(0, x)}+{max(0, y)}")
        entry.focus_set()
        if initial and hasattr(entry, "select_range"):
            entry.select_range(0, "end")
        try:
            dlg.grab_set()
        except tk.TclError:
            pass
        self.root.wait_window(dlg)
        return result["value"]

    def toggle_mark(self, rid, field):
        rec = self.db.record(rid)
        if rec:
            self.db.update_record(rid, **{field: 0 if rec[field] else 1})
            self.schedule_render()

    def edit_date(self, rid):
        rec = self.db.record(rid)
        if not rec:
            return
        cur = dt.date.fromisoformat(rec["rdate"])
        text = self.ask_text("تعديل التاريخ", "اكتب التاريخ بصيغة 2026/10/01", initial=f"{cur:%Y/%m/%d}", width=24)
        if text is None:
            return
        d = parse_date(text)
        if d is None:
            messagebox.showerror("تاريخ غير صحيح", "اكتب التاريخ بصيغة 2026/10/01 أو 01/10/2026.", parent=self.root)
            return
        self.db.update_record(rid, rdate=d.isoformat())
        self.schedule_render()

    def remove_record(self, rid):
        if messagebox.askyesno("حذف المخالفة", "هل تريد حذف هذا السطر من السجل؟", parent=self.root):
            self.db.delete_record(rid)
            self.schedule_render()

    # ------------------------------------------------------------- التصدير للشيت المثالي
    def find_template(self):
        p = APP_DIR / TEMPLATE_NAME
        if p.exists():
            return p
        found = sorted(APP_DIR.glob("*مثالي*.xlsx"))
        return found[0] if found else None

    def ensure_template(self):
        if self.template and Path(self.template).exists():
            return self.template
        messagebox.showinfo("الشيت المثالي", "لم أجد ملف «شيت_مثالي.xlsx» بجانب البرنامج.\nاختره من فضلك.",
                            parent=self.root)
        p = filedialog.askopenfilename(title="اختر الشيت المثالي", filetypes=[("Excel", "*.xlsx")])
        if not p:
            return None
        try:
            shutil.copy(p, APP_DIR / TEMPLATE_NAME)
            p = APP_DIR / TEMPLATE_NAME
        except Exception:
            pass
        self.template = Path(p)
        self.violations = load_violation_choices(self.template)
        self.schedule_render()
        return self.template

    def export_current(self):
        if self.current is None:
            return
        tpl = self.ensure_template()
        if not tpl:
            return
        s = self.db.student(self.current)
        recs = self.db.records(self.current)
        EXPORT_DIR.mkdir(exist_ok=True)
        base = safe_filename(f"{s['name']} - {s['cls']}")
        out = EXPORT_DIR / f"{base}.xlsx"
        try:
            export_student(tpl, s["name"], s["cls"], recs, out, self.violations)
        except PermissionError:                                         # الملف القديم مفتوح في Excel
            out = EXPORT_DIR / f"{base}_{dt.datetime.now():%H%M%S}.xlsx"
            export_student(tpl, s["name"], s["cls"], recs, out, self.violations)
        try:
            open_file(out)
        except Exception:
            messagebox.showinfo("تم الحفظ", f"تم حفظ الشيت في:\n{out}", parent=self.root)


# ----------------------------------------------------------------------------
# قفل البرنامج بكلمة سر (تعمل أيضاً بعد التحويل إلى exe)
# كلمة السر غير مكتوبة في الكود، المحفوظ فقط بصمتها المشفّرة (PBKDF2).
# ----------------------------------------------------------------------------
PW_SALT = bytes.fromhex("a3f19c7e5b2d4488")
PW_HASH = "301ecafc008a19dfca3bea5e75480301078c3f444a5a5522fca5ebcd8966a7eb"
PW_ROUNDS = 200_000
PW_MAX_TRIES = 3


def check_password(text):
    h = hashlib.pbkdf2_hmac("sha256", text.strip().translate(_AR_DIGITS).encode("utf-8"), PW_SALT, PW_ROUNDS)
    return hmac.compare_digest(h.hex(), PW_HASH)


def ask_password(root):
    """يعرض شاشة كلمة السر. يرجع True لو صحيحة، وFalse لو أُلغيت أو انتهت المحاولات."""
    result = {"ok": False, "tries": 0}
    dlg = tk.Toplevel(root)
    dlg.title("كلمة السر")
    dlg.configure(bg=BG)
    dlg.resizable(False, False)
    dlg.attributes("-topmost", True)

    tk.Label(dlg, text="🔒 سجل المخالفات والعقوبات", font=(FONT, 15, "bold"), bg=BG, fg=DARK).pack(padx=40, pady=(22, 4))
    tk.Label(dlg, text="اكتب كلمة السر لفتح البرنامج", font=(FONT, 11), bg=BG, fg=GRAY).pack()
    var = tk.StringVar()
    entry = tk.Entry(dlg, textvariable=var, show="•", justify="center", font=(FONT, 15), width=22)
    entry.pack(padx=40, pady=12)
    msg = tk.Label(dlg, text="", font=(FONT, 10, "bold"), bg=BG, fg=RED)
    msg.pack()

    def submit(event=None):
        if check_password(var.get()):
            result["ok"] = True
            dlg.destroy()
            return
        result["tries"] += 1
        var.set("")
        left = PW_MAX_TRIES - result["tries"]
        if left <= 0:
            dlg.destroy()
        else:
            msg.config(text=f"كلمة السر غير صحيحة — متبقي {left} محاولة")

    row = tk.Frame(dlg, bg=BG)
    row.pack(pady=(6, 20))
    tk.Button(row, text="دخول", command=submit, bg=GREEN, fg="white", activebackground=GREEN,
              activeforeground="white", relief="flat", font=(FONT, 11, "bold"), padx=22, pady=5,
              cursor="hand2").pack(side="right", padx=6)
    tk.Button(row, text="خروج", command=dlg.destroy, bg=GRAY, fg="white", activebackground=GRAY,
              activeforeground="white", relief="flat", font=(FONT, 11, "bold"), padx=22, pady=5,
              cursor="hand2").pack(side="right", padx=6)
    dlg.bind("<Return>", submit)
    dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)

    dlg.update_idletasks()
    w, h = dlg.winfo_reqwidth(), dlg.winfo_reqheight()
    x = (dlg.winfo_screenwidth() - w) // 2
    y = (dlg.winfo_screenheight() - h) // 3
    dlg.geometry(f"+{x}+{y}")
    entry.focus_force()
    dlg.grab_set()
    root.wait_window(dlg)
    return result["ok"]


def main():
    try:                                                                # وضوح أفضل للخطوط على ويندوز
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    root.withdraw()                                                     # نخفي النافذة الرئيسية لحد ما الباسورد يتأكد
    if not ask_password(root):
        root.destroy()
        return
    backup_db()
    root.deiconify()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
