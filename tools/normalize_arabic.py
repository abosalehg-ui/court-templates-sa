# -*- coding: utf-8 -*-
"""
تطبيع النصوص العربية لقوالب المحاكم — court-templates-sa

يُستخدم في موضعين:
  1) tools/wordtotemplates.py  عند التوليد من ملفات Word (المصدر الحقيقي للمشكلة)
  2) tools/clean_templates.py  لتنظيف data/templates.js الموجود حاليًا

قاعدة واحدة في مكان واحد، فلا تتكرر المشكلة عند إعادة التوليد.
"""

import re
import unicodedata

TATWEEL = "\u0640"          # ـ  علامة التطويل (الكشيدة)
NBSP = "\u00a0"             # مسافة غير قابلة للكسر
EM_DASH = "\u2014"          # —

# محارف تنسيق خفية تُحذف نهائيًا
INVISIBLE = {
    "\u200b",  # ZERO WIDTH SPACE
    "\u200c",  # ZERO WIDTH NON-JOINER
    "\u200d",  # ZERO WIDTH JOINER
    "\u200e",  # LEFT-TO-RIGHT MARK
    "\u200f",  # RIGHT-TO-LEFT MARK
    "\u061c",  # ARABIC LETTER MARK
    "\u00ad",  # SOFT HYPHEN
    "\ufeff",  # BOM / ZERO WIDTH NO-BREAK SPACE
    "\u2060",  # WORD JOINER
    "\u202a", "\u202b", "\u202c", "\u202d", "\u202e",   # embeddings / overrides / pop
    "\u2066", "\u2067", "\u2068", "\u2069",             # isolates
}

# مسافات غير قياسية تتحول إلى مسافة عادية
ODD_SPACES = {
    NBSP, "\u2000", "\u2001", "\u2002", "\u2003", "\u2004", "\u2005",
    "\u2006", "\u2007", "\u2008", "\u2009", "\u200a", "\u202f",
    "\u205f", "\u3000", "\u1680",
}

_INVISIBLE_RE = re.compile("[" + "".join(INVISIBLE) + "]")
_ODD_SPACE_RE = re.compile("[" + "".join(ODD_SPACES) + "]")

# تطويل مستقل بين مسافتين → شرطة حقيقية (كان يُستخدم كفاصل)
_STANDALONE_TATWEEL_RE = re.compile(
    r"(?<!\S)" + TATWEEL + r"+(?!\S)"
)
# هاء متبوعة بتطويل في آخر الكلمة → هـ  (اختصار «انتهى» و«هجري»)
_HE_ABBREV_RE = re.compile(r"\u0647" + TATWEEL + r"+(?=$|[^\u0621-\u064a])")

_LINE_SEP_RE = re.compile("[\u2028\u2029]")


def strip_invisible(text: str) -> str:
    """يحذف محارف التنسيق الخفية ويحوّل المسافات الشاذة إلى مسافة عادية."""
    text = _INVISIBLE_RE.sub("", text)
    text = _ODD_SPACE_RE.sub(" ", text)
    text = _LINE_SEP_RE.sub("\n", text)
    # محارف تحكّم متبقية عدا السطر الجديد والجدولة
    text = "".join(
        ch for ch in text
        if ch in ("\n", "\t") or unicodedata.category(ch) != "Cc"
    )
    return text


def strip_tatweel(text: str) -> str:
    """
    يزيل تطويل الكشيدة الزخرفي مع الحفاظ على حالتين:
      • التطويل المستقل المستخدم كفاصل  ـــ  →  —
      • اختصار الهاء في آخر الكلمة  هــــ  →  هـ   (أ.هـ / 1441هـ)
    """
    text = _STANDALONE_TATWEEL_RE.sub(EM_DASH, text)
    # حارس مؤقت من نطاق الاستخدام الخاص، حتى لا يمسح الكنسُ الأخير الهاءَ المحمية
    text = _HE_ABBREV_RE.sub("\u0647\uE000", text)
    # ما تبقى داخل الكلمات زخرفي بحت
    text = text.replace(TATWEEL, "")
    return text.replace("\uE000", TATWEEL)


def tidy_spacing(text: str) -> str:
    """يوحّد المسافات دون المساس بفواصل الأسطر أو بنقاط التعبئة (.....)"""
    text = re.sub(r"[ \t]{2,}", " ", text)
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def normalize(text: str, tidy: bool = True) -> str:
    """التطبيع الكامل — هذه هي الدالة التي تُستدعى من الخارج."""
    if not text:
        return text
    text = strip_invisible(text)
    text = strip_tatweel(text)
    if tidy:
        text = tidy_spacing(text)
    return text


def audit(text: str) -> dict:
    """يرجع إحصاء بما سيتغيّر — مفيد للتحقق قبل الكتابة."""
    return {
        "invisible": len(_INVISIBLE_RE.findall(text)),
        "odd_spaces": len(_ODD_SPACE_RE.findall(text)),
        "tatweel": text.count(TATWEEL),
    }
