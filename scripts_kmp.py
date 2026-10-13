# -*- coding: utf-8 -*-
#| Экспорт развёрток листовых тел из активной сборки/детали в DXF
#| Совместим с КОМПАС-3D v21 и v25.
#| Универсальный автопоиск конвертера DXF (реестр + все диски).
#| Обозначение размещается ПАРАЛЛЕЛЬНО самой длинной грани контура.
#| УПРАВЛЕНИЕ линиями переходов через BreakLinesVisible (API 7).
#| Удаление рамки/штампа/масштаба 1:1 из экспортированного DXF.
#| Всего 2 слоя на выходе: "0" (геометрия + зелёные насечки) и "Mark".
#|
#| Однолинейный векторный шрифт (порт GetStrokeFont из VBA SolidWorks).
#| Выходная папка: <папка модели>\DXF_kps
#| Версия под упаковку в EXE (PyInstaller) — БЕЗ matplotlib.

import os
import sys
import io
import math
import re
import glob
import string
import traceback
from collections import Counter

# === Адаптация под EXE ===
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', line_buffering=True)
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace', line_buffering=True)
except Exception:
    pass

try:
    import multiprocessing
    multiprocessing.freeze_support()
except Exception:
    pass

import win32com.client
import pythoncom
from win32com.client import Dispatch, gencache
import ezdxf
from ezdxf.enums import TextEntityAlignment
from shapely.geometry import LineString, Point, Polygon
from shapely.ops import unary_union


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# ============================================================
# НАСТРОЙКИ
# ============================================================
# Выходная папка создаётся РЯДОМ С МОДЕЛЬЮ, в подпапке DXF_kps.
# Если модель не сохранена / путь не определён — fallback на FALLBACK_OUTPUT_DIR.
OUTPUT_SUBDIR_NAME   = "DXF_kps"
FALLBACK_OUTPUT_DIR  = r"C:\DXF"
OUTPUT_DIR           = None   # вычисляется после открытия активного документа

DXF_CONVERTER_PATH = None  # None = искать автоматически

# === СЛОИ ===
GEOMETRY_LAYER = "0"          # слой для всей геометрии и насечек
DEFPOINTS_LAYER = "Mark"      # слой для обозначения

# === РАЗМЕР ОБОЗНАЧЕНИЯ ===
# DEFPOINT_TEXT_HEIGHT — ЧИСТАЯ высота цифры (штриха) в мм.
# GLYPH_SPAN — доля ячейки глифа, которую занимает штрих.
#   В STROKE_FONT все символы описаны в y ∈ [0.05 … 0.95] → штрих = 0.9 ячейки.
#   Поэтому высота ЯЧЕЙКИ = DEFPOINT_TEXT_HEIGHT / GLYPH_SPAN,
#   и итоговая высота цифры на чертеже будет ровно DEFPOINT_TEXT_HEIGHT мм.
DEFPOINT_TEXT_HEIGHT = 8.0    # ЧИСТАЯ высота цифры, мм
GLYPH_SPAN           = 0.9    # доля ячейки, занимаемая штрихом глифа

# === ЦВЕТА (ACI) ===
COLOR_LINES = 0   # цвет контура/отверстий (0=ByBlock, поставьте 7 для чёрного)
COLOR_BEND  = 3   # цвет насечек сгиба (3 = зелёный)
COLOR_RED   = 1   # цвет обозначения (1 = красный)

# === ЛИНИИ ПЕРЕХОДОВ ===
TRANSITION_LINES_VISIBLE = False

LAYERS_TO_DELETE = [
    "ТТ", "Технические требования",
    "Шероховатость", "Неуказанная шероховатость",
]

TYPES_TO_REMOVE_FROM_ZERO = ("TEXT", "MTEXT", "INSERT", "ATTRIB")
SCALE_TEXT_REGEX = re.compile(r"^\s*1\s*:\s*1\s*$")

TARGET_BEND_LINETYPE = "CONTINUOUS"
BEND_LENGTH = 0.8
OUTER_EDGE_TOL = 1.0

BEND_LINETYPE_KEYWORDS = ("K5LT_CHAIN_LINE_WITH_2_POINTS",)

# Системные блоки КОМПАСа для рамки/штампа/масштаба (T25, T26, T27, ...)
SYS_BLOCK_REGEX = re.compile(r"^T\d+$", re.IGNORECASE)

# Дополнительный фильтр удаления линий переходов из DXF (страховка)
TRANSITION_LAYERS = [
    "Тонкие линии",
    "Линии переходов",
    "Переходы",
    "THIN_LINES",
    "TRANSITION_LINES",
    "HIDDEN",
    "Скрытые линии",
]
TRANSITION_LINETYPE_KEYWORDS = (
    "K5LT_THIN_LINE",
    "K5LT_TRANSITION",
    "THIN",
    "ТОНКАЯ",
    "TRANSITION",
    "ПЕРЕХОД",
)

MIN_HOLE_AREA = 1.0
GLUE_BUFFER = 0.1
INSIDE_TOLERANCE = 0.5

PROP_NAME_DESIGNATION = "Обозначение"
PROP_NAME_NAME        = "Наименование"


# ============================================================
# ОДНОЛИНЕЙНЫЙ ВЕКТОРНЫЙ ШРИФТ (порт GetStrokeFont из VBA SolidWorks)
# ============================================================
STROKE_FONT = {
    "0": "0.10,0.15 0.15,0.05 0.45,0.05 0.50,0.15 0.50,0.85 0.45,0.95 0.15,0.95 0.10,0.85 0.10,0.15",
    "1": "0.10,0.80 0.30,0.95 0.30,0.05|0.10,0.05 0.50,0.05",
    "2": "0.10,0.85 0.15,0.95 0.45,0.95 0.50,0.85 0.50,0.65 0.10,0.10 0.10,0.05 0.50,0.05",
    "3": "0.10,0.85 0.15,0.95 0.45,0.95 0.50,0.85 0.50,0.65 0.45,0.55 0.20,0.50|0.45,0.55 0.50,0.45 0.50,0.15 0.45,0.05 0.15,0.05 0.10,0.15",
    "4": "0.40,0.05 0.40,0.95 0.10,0.30 0.55,0.30",
    "5": "0.50,0.95 0.10,0.95 0.10,0.55 0.40,0.55 0.50,0.45 0.50,0.15 0.40,0.05 0.10,0.05 0.05,0.15",
    "6": "0.50,0.85 0.45,0.95 0.15,0.95 0.10,0.85 0.10,0.15 0.15,0.05 0.40,0.05 0.50,0.15 0.50,0.40 0.40,0.50 0.15,0.50 0.10,0.40",
    "7": "0.05,0.95 0.50,0.95 0.25,0.05",
    "8": "0.15,0.50 0.10,0.60 0.10,0.85 0.15,0.95 0.45,0.95 0.50,0.85 0.50,0.60 0.45,0.50 0.15,0.50|0.15,0.50 0.10,0.40 0.10,0.15 0.15,0.05 0.45,0.05 0.50,0.15 0.50,0.40 0.45,0.50",
    "9": "0.50,0.60 0.45,0.50 0.15,0.50 0.10,0.60 0.10,0.85 0.15,0.95 0.45,0.95 0.50,0.85 0.50,0.15 0.45,0.05 0.15,0.05 0.10,0.15",

    "A": "0.05,0.05 0.30,0.95 0.55,0.05|0.15,0.35 0.45,0.35",
    "B": "0.10,0.95 0.10,0.05 0.40,0.05 0.50,0.15 0.50,0.40 0.40,0.50 0.10,0.50|0.40,0.50 0.50,0.60 0.50,0.85 0.40,0.95 0.10,0.95",
    "C": "0.50,0.85 0.45,0.95 0.15,0.95 0.10,0.85 0.10,0.15 0.15,0.05 0.45,0.05 0.50,0.15",
    "D": "0.10,0.95 0.10,0.05 0.40,0.05 0.50,0.15 0.50,0.85 0.40,0.95 0.10,0.95",
    "E": "0.50,0.95 0.10,0.95 0.10,0.05 0.50,0.05|0.10,0.50 0.40,0.50",
    "F": "0.50,0.95 0.10,0.95 0.10,0.05|0.10,0.50 0.40,0.50",
    "G": "0.50,0.85 0.45,0.95 0.15,0.95 0.10,0.85 0.10,0.15 0.15,0.05 0.40,0.05 0.50,0.15 0.50,0.45 0.30,0.45",
    "H": "0.10,0.95 0.10,0.05|0.50,0.95 0.50,0.05|0.10,0.50 0.50,0.50",
    "I": "0.15,0.95 0.45,0.95|0.30,0.95 0.30,0.05|0.15,0.05 0.45,0.05",
    "J": "0.15,0.95 0.45,0.95|0.40,0.95 0.40,0.10 0.30,0.05 0.15,0.05 0.05,0.15",
    "K": "0.10,0.95 0.10,0.05|0.50,0.95 0.10,0.45|0.25,0.60 0.50,0.05",
    "L": "0.10,0.95 0.10,0.05 0.50,0.05",
    "M": "0.05,0.05 0.05,0.95 0.30,0.50 0.55,0.95 0.55,0.05",
    "N": "0.10,0.05 0.10,0.95 0.50,0.05 0.50,0.95",
    "O": "0.10,0.15 0.15,0.05 0.45,0.05 0.50,0.15 0.50,0.85 0.45,0.95 0.15,0.95 0.10,0.85 0.10,0.15",
    "P": "0.10,0.05 0.10,0.95 0.40,0.95 0.50,0.85 0.50,0.60 0.40,0.50 0.10,0.50",
    "Q": "0.10,0.15 0.15,0.05 0.45,0.05 0.50,0.15 0.50,0.85 0.45,0.95 0.15,0.95 0.10,0.85 0.10,0.15|0.30,0.25 0.55,0.00",
    "R": "0.10,0.05 0.10,0.95 0.40,0.95 0.50,0.85 0.50,0.60 0.40,0.50 0.10,0.50|0.30,0.50 0.50,0.05",
    "S": "0.50,0.85 0.45,0.95 0.15,0.95 0.10,0.85 0.10,0.60 0.15,0.50 0.45,0.50 0.50,0.40 0.50,0.15 0.45,0.05 0.15,0.05 0.10,0.15",
    "T": "0.05,0.95 0.55,0.95|0.30,0.95 0.30,0.05",
    "U": "0.10,0.95 0.10,0.15 0.15,0.05 0.45,0.05 0.50,0.15 0.50,0.95",
    "V": "0.05,0.95 0.30,0.05 0.55,0.95",
    "W": "0.05,0.95 0.05,0.05 0.30,0.50 0.55,0.05 0.55,0.95",
    "X": "0.05,0.05 0.55,0.95|0.05,0.95 0.55,0.05",
    "Y": "0.05,0.95 0.30,0.50 0.55,0.95|0.30,0.50 0.30,0.05",
    "Z": "0.05,0.95 0.55,0.95 0.05,0.05 0.55,0.05",

    "-": "0.10,0.50 0.50,0.50",
    ".": "0.30,0.05 0.30,0.20",
    ",": "0.30,0.10 0.25,-0.05",
    "/": "0.05,0.00 0.55,0.95",
    "\\": "0.05,0.95 0.55,0.00",
    ":": "0.25,0.15 0.35,0.15 0.35,0.25 0.25,0.25 0.25,0.15|0.25,0.55 0.35,0.55 0.35,0.65 0.25,0.65 0.25,0.55",
    "_": "0.00,-0.15 0.60,-0.15",
    "+": "0.10,0.50 0.50,0.50|0.30,0.30 0.30,0.70",
    "=": "0.10,0.40 0.50,0.40|0.10,0.60 0.50,0.60",
    " ": "",

    "А": "0.05,0.05 0.30,0.95 0.55,0.05|0.15,0.35 0.45,0.35",
    "Б": "0.10,0.95 0.50,0.95|0.10,0.95 0.10,0.05 0.50,0.05 0.50,0.55 0.10,0.55",
    "В": "0.10,0.95 0.10,0.05 0.40,0.05 0.50,0.15 0.50,0.40 0.40,0.50 0.10,0.50|0.40,0.50 0.50,0.60 0.50,0.85 0.40,0.95 0.10,0.95",
    "Г": "0.10,0.95 0.50,0.95|0.10,0.95 0.10,0.05",
    "Д": "0.05,0.05 0.15,0.05 0.15,0.20 0.10,0.20 0.20,0.95 0.40,0.95 0.50,0.20 0.45,0.20 0.45,0.05 0.55,0.05|0.20,0.20 0.40,0.20",
    "Е": "0.50,0.95 0.10,0.95 0.10,0.05 0.50,0.05|0.10,0.50 0.40,0.50",
    "Ж": "0.05,0.95 0.30,0.50 0.05,0.05|0.55,0.95 0.30,0.50 0.55,0.05|0.30,0.95 0.30,0.05",
    "З": "0.05,0.85 0.15,0.95 0.45,0.95 0.50,0.85 0.50,0.60 0.45,0.50 0.25,0.50|0.45,0.50 0.50,0.40 0.50,0.15 0.45,0.05 0.15,0.05 0.05,0.15",
    "И": "0.10,0.05 0.10,0.95|0.50,0.05 0.50,0.95|0.10,0.05 0.50,0.95",
    "Й": "0.10,0.05 0.10,0.95|0.50,0.05 0.50,0.95|0.10,0.05 0.50,0.95|0.20,1.10 0.30,1.20 0.40,1.10",
    "К": "0.10,0.95 0.10,0.05|0.50,0.95 0.10,0.45|0.25,0.60 0.50,0.05",
    "Л": "0.10,0.05 0.20,0.95 0.50,0.95 0.50,0.05",
    "М": "0.05,0.05 0.05,0.95 0.30,0.50 0.55,0.95 0.55,0.05",
    "Н": "0.10,0.95 0.10,0.05|0.50,0.95 0.50,0.05|0.10,0.50 0.50,0.50",
    "О": "0.10,0.15 0.15,0.05 0.45,0.05 0.50,0.15 0.50,0.85 0.45,0.95 0.15,0.95 0.10,0.85 0.10,0.15",
    "П": "0.10,0.05 0.10,0.95 0.50,0.95 0.50,0.05",
    "Р": "0.10,0.05 0.10,0.95 0.40,0.95 0.50,0.85 0.50,0.60 0.40,0.50 0.10,0.50",
    "С": "0.50,0.85 0.45,0.95 0.15,0.95 0.10,0.85 0.10,0.15 0.15,0.05 0.45,0.05 0.50,0.15",
    "Т": "0.05,0.95 0.55,0.95|0.30,0.95 0.30,0.05",
    "У": "0.05,0.95 0.30,0.50|0.55,0.95 0.10,0.05",
    "Ф": "0.30,0.95 0.30,0.05|0.10,0.50 0.15,0.65 0.25,0.70 0.35,0.70 0.45,0.65 0.50,0.50 0.45,0.35 0.35,0.30 0.25,0.30 0.15,0.35 0.10,0.50",
    "Х": "0.05,0.05 0.55,0.95|0.05,0.95 0.55,0.05",
    "Ц": "0.10,0.95 0.10,0.05 0.45,0.05 0.45,0.95|0.45,0.05 0.55,0.05 0.55,-0.10",
    "Ч": "0.10,0.95 0.10,0.55 0.40,0.55 0.50,0.45|0.50,0.95 0.50,0.05",
    "Ш": "0.05,0.95 0.05,0.05 0.30,0.05 0.30,0.95|0.30,0.95 0.30,0.05 0.55,0.05 0.55,0.95",
    "Щ": "0.05,0.95 0.05,0.05 0.30,0.05 0.30,0.95|0.30,0.95 0.30,0.05 0.55,0.05 0.55,0.95|0.55,0.05 0.65,0.05 0.65,-0.10",
    "Ъ": "0.05,0.95 0.25,0.95 0.25,0.05 0.50,0.05 0.50,0.55 0.25,0.55",
    "Ы": "0.10,0.95 0.10,0.05 0.40,0.05 0.40,0.55 0.10,0.55|0.55,0.95 0.55,0.05",
    "Ь": "0.10,0.95 0.10,0.05 0.50,0.05 0.50,0.55 0.10,0.55",
    "Э": "0.05,0.85 0.15,0.95 0.45,0.95 0.50,0.85 0.50,0.15 0.45,0.05 0.15,0.05 0.05,0.15|0.05,0.50 0.40,0.50",
    "Ю": "0.10,0.95 0.10,0.05|0.10,0.50 0.30,0.50|0.30,0.65 0.40,0.70 0.50,0.65 0.55,0.50 0.50,0.35 0.40,0.30 0.30,0.35 0.30,0.65",
    "Я": "0.50,0.95 0.50,0.05|0.15,0.95 0.50,0.95 0.50,0.50 0.15,0.50 0.15,0.95|0.30,0.50 0.15,0.05",
}


# ============================================================
# УНИВЕРСАЛЬНЫЙ ПОИСК КОНВЕРТЕРА DXF
# ============================================================
def find_dxf_converter():
    if DXF_CONVERTER_PATH and os.path.exists(DXF_CONVERTER_PATH):
        print(f"  Конвертер (задан вручную): {DXF_CONVERTER_PATH}")
        return DXF_CONVERTER_PATH

    try:
        app5 = Dispatch("Kompas.Application.5")
        try:
            libs_path = app5.ksSystemPath(1)
        except Exception:
            libs_path = None
        if libs_path:
            p = os.path.join(libs_path, "ImpExp", "dwgdxfexp.rtw")
            if os.path.exists(p):
                print(f"  Конвертер найден через API 5: {p}")
                return p
    except Exception as e:
        print(f"  API 5 не сработал: {e}")

    try:
        import winreg
        keys_to_try = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\ASCON\KOMPAS-3D"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\ASCON\KOMPAS-3D"),
            (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\ASCON\KOMPAS-3D"),
        ]
        for hive, path in keys_to_try:
            try:
                key = winreg.OpenKey(hive, path, 0, winreg.KEY_READ)
                i = 0
                while True:
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        i += 1
                        try:
                            subkey = winreg.OpenKey(key, subkey_name, 0, winreg.KEY_READ)
                            for value_name in ("Path", "InstallPath", "InstallDir", ""):
                                try:
                                    if value_name:
                                        install_path, _ = winreg.QueryValueEx(subkey, value_name)
                                    else:
                                        install_path, _ = winreg.QueryValueEx(subkey, None)
                                    if install_path:
                                        p = os.path.join(install_path, "Libs", "ImpExp", "dwgdxfexp.rtw")
                                        if os.path.exists(p):
                                            winreg.CloseKey(subkey)
                                            winreg.CloseKey(key)
                                            print(f"  Конвертер найден через реестр: {p}")
                                            return p
                                except Exception:
                                    continue
                            winreg.CloseKey(subkey)
                        except Exception:
                            continue
                    except OSError:
                        break
                winreg.CloseKey(key)
            except Exception:
                continue
    except Exception as e:
        print(f"  Поиск в реестре не удался: {e}")

    print("  Поиск по всем доступным дискам...")
    sub_paths = [
        r"Libs\ImpExp",
        r"Program Files\ASCON\KOMPAS-3D v*\Libs\ImpExp",
        r"Program Files (x86)\ASCON\KOMPAS-3D v*\Libs\ImpExp",
        r"Program Files\ASCON\KOMPAS-3D*\Libs\ImpExp",
        r"ASCON\KOMPAS-3D v*\Libs\ImpExp",
        r"Компас 3D *\Libs\ImpExp",
        r"KOMPAS-3D *\Libs\ImpExp",
        r"Program Files\ASCON\KOMPAS*\Libs\ImpExp",
    ]
    for drive_letter in string.ascii_uppercase:
        drive = f"{drive_letter}:\\"
        if not os.path.exists(drive):
            continue
        for sub in sub_paths:
            pattern = os.path.join(drive, sub, "dwgdxfexp.rtw")
            try:
                found = glob.glob(pattern)
                if found:
                    found.sort(reverse=True)
                    print(f"  Конвертер найден при переборе дисков: {found[0]}")
                    return found[0]
            except Exception:
                continue

    print("  ВНИМАНИЕ: не удалось найти dwgdxfexp.rtw.")
    return None


# ============================================================
# ЧТЕНИЕ СВОЙСТВ ЧЕРЕЗ API 7
# ============================================================
def _flatten_value(v, out):
    if v is None:
        return
    if isinstance(v, bytes):
        try:
            v = v.decode("utf-8", "replace")
        except Exception:
            v = str(v)
    if isinstance(v, str):
        s = v.strip()
        if s:
            out.append(s)
        return
    if isinstance(v, bool):
        out.append(str(v)); return
    if isinstance(v, (int, float)):
        out.append(str(int(v)) if isinstance(v, float) and v.is_integer() else str(v))
        return
    if isinstance(v, (list, tuple)):
        for it in v:
            _flatten_value(it, out)
        return
    try:
        it = iter(v)
    except TypeError:
        s = str(v).strip()
        if s:
            out.append(s)
        return
    for x in it:
        _flatten_value(x, out)


def _extract_code_prefix(s):
    if not s:
        return ""
    s = str(s).strip()
    m = re.match(r"^(\d[\d\.\-\/\\]*\d|\d+)", s)
    if m:
        return m.group(1).rstrip(".-/\\")
    parts = s.split()
    if not parts:
        return ""
    tok = parts[0]
    tok = re.sub(r"[^\w\.\-\/]", "", tok, flags=re.UNICODE)
    return tok.rstrip(".-/\\")


def read_property_value(doc, part, prop_name, label=""):
    print(f"  [PROPS] ===== '{prop_name}' ({label}) =====")
    try:
        prop_mng = api7_module.IPropertyMng(iApplication7)
    except Exception as e:
        print(f"  [PROPS] IPropertyMng: {e}")
        return ""

    prop = None
    try:
        prop = prop_mng.GetProperty(doc, prop_name)
    except Exception as e:
        print(f"  [PROPS] GetProperty('{prop_name}'): {e}")

    if prop is None:
        print(f"  [PROPS] Свойство '{prop_name}' не найдено")
        return ""

    try:
        keeper = api7_module.IPropertyKeeper(part)
    except Exception as e:
        print(f"  [PROPS] IPropertyKeeper: {e}")
        return ""

    try:
        res, value, from_source = keeper.GetPropertyValue(prop, "", True, True)
        if res and value:
            vals = []
            _flatten_value(value, vals)
            if vals:
                print(f"  [PROPS] Значение: '{vals[0]}'")
                return vals[0]
    except Exception as e:
        print(f"  [PROPS] GetPropertyValue: {e}")
    return ""


# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ DXF
# ============================================================
def iter_layouts(doc):
    res = [("MODEL", doc.modelspace())]
    for blk in doc.blocks:
        nm = blk.dxf.name
        if nm.upper().startswith("*"):
            continue
        res.append((nm, blk))
    return res


def color_stats(container):
    cc, tt = Counter(), Counter()
    for e in container:
        try:
            cc[str(e.dxf.get("color", "?"))] += 1
        except Exception:
            cc["ERR"] += 1
        try:
            tt[e.dxftype()] += 1
        except Exception:
            tt["?"] += 1
    return cc, tt


def dump_color_diag(doc, title):
    print(f"  [POST] ===== {title} =====")
    for name, cont in iter_layouts(doc):
        cc, tt = color_stats(cont)
        total = sum(tt.values())
        print(f"  [POST]   [{name}] сущностей={total} "
              f"цвета(ACI)={dict(cc)} типы={dict(tt)}")


def get_effective_linetype(doc, entity):
    lt = (entity.dxf.get("linetype", "") or "").upper()
    if not lt or lt == "BYLAYER":
        layer_name = entity.dxf.get("layer", "")
        try:
            lt = (doc.layers.get(layer_name).dxf.linetype or "").upper()
        except Exception:
            lt = ""
    return lt


def is_bend_line(entity, doc):
    eff_lt = get_effective_linetype(doc, entity)
    if not eff_lt:
        return False
    return any(kw in eff_lt for kw in BEND_LINETYPE_KEYWORDS)


def is_transition_line(entity, doc):
    layer_name = entity.dxf.get("layer", "") or ""
    if layer_name in TRANSITION_LAYERS:
        return True
    eff_lt = get_effective_linetype(doc, entity)
    if eff_lt and any(kw in eff_lt for kw in TRANSITION_LINETYPE_KEYWORDS):
        return True
    return False


def entity_to_shapely(entity):
    t = entity.dxftype()
    try:
        if t == "LINE":
            s = entity.dxf.start; e = entity.dxf.end
            return LineString([(s.x, s.y), (e.x, e.y)])
        if t == "CIRCLE":
            c = entity.dxf.center; r = entity.dxf.radius; n = 48
            pts = [(c.x + r * math.cos(2 * math.pi * i / n),
                    c.y + r * math.sin(2 * math.pi * i / n)) for i in range(n + 1)]
            return LineString(pts)
        if t == "ARC":
            c = entity.dxf.center; r = entity.dxf.radius
            a1 = math.radians(entity.dxf.start_angle); a2 = math.radians(entity.dxf.end_angle)
            if a2 < a1: a2 += 2 * math.pi
            n = 32; pts = []
            for i in range(n + 1):
                a = a1 + (a2 - a1) * i / n
                pts.append((c.x + r * math.cos(a), c.y + r * math.sin(a)))
            return LineString(pts)
        if t == "ELLIPSE":
            pts = list(entity.flattening(0.05))
            if len(pts) >= 2:
                return LineString([(p.x, p.y) for p in pts])
        if t == "LWPOLYLINE":
            try:
                from ezdxf.path import make_path
                path = make_path(entity)
                pts = list(path.flattening(0.05))
                if len(pts) >= 2:
                    return LineString([(p.x, p.y) for p in pts])
            except Exception:
                pts = [(p[0], p[1]) for p in entity.get_points()]
                if len(pts) >= 2:
                    return LineString(pts)
        if t == "POLYLINE":
            pts = [(v.dxf.location.x, v.dxf.location.y) for v in entity.vertices]
            if len(pts) >= 2:
                return LineString(pts)
        if t == "SPLINE":
            pts = list(entity.flattening(0.05))
            if len(pts) >= 2:
                return LineString([(p.x, p.y) for p in pts])
    except Exception:
        pass
    return None


def find_outer_and_holes(line_geoms):
    """Возвращает (outer_filled_polygon, holes_list).

    Если после буферизации контур получился 'рамкой' (плотность < 0.3),
    превращаем его в залитый полигон по внешней границе."""
    if not line_geoms:
        return None, []

    try:
        buffered = [g.buffer(GLUE_BUFFER) for g in line_geoms]
        merged = unary_union(buffered)
    except Exception as e:
        print(f"  [POST] Ошибка буферизации: {e}")
        return None, []

    polys = []
    if merged.geom_type == "Polygon":
        polys = [merged]
    elif merged.geom_type == "MultiPolygon":
        polys = list(merged.geoms)
    else:
        return None, []

    polys.sort(key=lambda p: p.area, reverse=True)
    outer_ring = polys[0]

    try:
        outer_filled = Polygon(outer_ring.exterior.coords)
    except Exception:
        outer_filled = outer_ring

    minx, miny, maxx, maxy = outer_filled.bounds
    bbox_area = max(1e-9, (maxx - minx) * (maxy - miny))
    density = outer_filled.area / bbox_area
    if density < 0.3:
        print(f"  [POST] Обнаружена рамка (плотность {density:.3f}) -> "
              f"заполняю внешнюю границу")

    holes = []
    for p in polys[1:]:
        try:
            if p.interiors:
                hp = Polygon(p.interiors[0].coords)
            else:
                hp = p
            if hp.area < MIN_HOLE_AREA:
                continue
            rp = hp.representative_point()
            if outer_filled.contains(rp):
                holes.append(hp)
        except Exception:
            continue

    return outer_filled, holes


def entity_inside_holes(entity, holes):
    geom = entity_to_shapely(entity)
    if geom is None:
        return False
    for h in holes:
        try:
            if geom.within(h.buffer(INSIDE_TOLERANCE)):
                return True
        except Exception:
            continue
    return False


def remove_scale_text(msp):
    removed = 0
    for e in list(msp):
        t = e.dxftype()
        text = ""
        if t == "TEXT":
            try: text = e.dxf.text or ""
            except Exception: continue
        elif t == "MTEXT":
            try: text = e.text or ""
            except Exception: continue
        else:
            continue
        if not text:
            continue
        clean = text
        if ";" in clean: clean = clean.split(";")[-1]
        clean = clean.replace("{", "").replace("}", "").strip()
        if SCALE_TEXT_REGEX.match(clean):
            try:
                msp.delete_entity(e); removed += 1
            except Exception:
                pass
    print(f"  [POST] Удалено масштабных надписей: {removed}")
    return removed


# ============================================================
# РАСЧЁТ РАЗМЕРА И ПОЗИЦИИ ТЕКСТА ОБОЗНАЧЕНИЯ
# ============================================================
def compute_designation_layout(designation, outer_geom):
    """Возвращает (cx, cy, cell_h, angle).

    ВАЖНО: возвращается ВЫСОТА ЯЧЕЙКИ (cell_h), а не чистая высота цифры.
    draw_designation_stroke ожидает именно высоту ячейки.
    Чистая высота цифры = cell_h * GLYPH_SPAN = DEFPOINT_TEXT_HEIGHT (в норме)."""
    if outer_geom is None:
        return 0.0, 0.0, DEFPOINT_TEXT_HEIGHT / GLYPH_SPAN, 0.0

    minx, miny, maxx, maxy = outer_geom.bounds
    part_w = maxx - minx
    part_h = maxy - miny
    if part_w <= 0 or part_h <= 0:
        return ((minx + maxx) / 2.0, (miny + maxy) / 2.0,
                DEFPOINT_TEXT_HEIGHT / GLYPH_SPAN, 0.0)

    n = max(1, len(designation))

    try:
        exterior = outer_geom.exterior
        coords = list(exterior.coords)
    except Exception:
        return ((minx + maxx) / 2.0, (miny + maxy) / 2.0,
                DEFPOINT_TEXT_HEIGHT / GLYPH_SPAN, 0.0)

    segments = []
    for i in range(len(coords) - 1):
        x1, y1 = coords[i]
        x2, y2 = coords[i + 1]
        length = math.hypot(x2 - x1, y2 - y1)
        if length < 1.0:
            continue
        mx = (x1 + x2) / 2.0
        my = (y1 + y2) / 2.0
        segments.append((length, (x1, y1), (x2, y2), (mx, my)))

    if not segments:
        rp = outer_geom.representative_point()
        return rp.x, rp.y, DEFPOINT_TEXT_HEIGHT / GLYPH_SPAN, 0.0

    segments.sort(key=lambda s: (-s[0], -s[3][1]))

    best_len, p1, p2, mid = segments[0]
    x1, y1 = p1
    x2, y2 = p2
    dx = x2 - x1
    dy = y2 - y1
    L = math.hypot(dx, dy)
    if L < 1e-6:
        rp = outer_geom.representative_point()
        return rp.x, rp.y, DEFPOINT_TEXT_HEIGHT / GLYPH_SPAN, 0.0

    angle = math.degrees(math.atan2(dy, dx))
    while angle > 90.0:
        angle -= 180.0
    while angle <= -90.0:
        angle += 180.0

    perp1 = (-dy / L, dx / L)
    perp2 = (dy / L, -dx / L)

    rp = outer_geom.representative_point()
    to_inside = (rp.x - mid[0], rp.y - mid[1])
    dot1 = to_inside[0] * perp1[0] + to_inside[1] * perp1[1]
    inward = perp1 if dot1 >= 0 else perp2

    # ---- ЧИСТАЯ высота цифры (штриха) ----
    glyph_h = DEFPOINT_TEXT_HEIGHT

    # Ограничение по меньшей стороне детали (в чистых мм)
    min_x, min_y, max_x, max_y = outer_geom.bounds
    part_min = min(max_x - min_x, max_y - min_y)
    if glyph_h > part_min * 0.18:
        glyph_h = part_min * 0.18

    # Ограничение по длине самой длинной грани.
    # Длина строки ≈ n * cell_h * 0.7,  где cell_h = glyph_h / GLYPH_SPAN
    cell_h_est = glyph_h / GLYPH_SPAN
    text_len_est = n * cell_h_est * 0.7
    if text_len_est > best_len * 0.85:
        cell_h_est = (best_len * 0.85) / (n * 0.7)
        glyph_h = cell_h_est * GLYPH_SPAN

    if glyph_h < 2.7:
        glyph_h = 2.7

    # Переводим ЧИСТУЮ высоту в ВЫСОТУ ЯЧЕЙКИ (её ждёт draw_designation_stroke)
    text_h = glyph_h / GLYPH_SPAN

    margin = glyph_h * 1.3
    cx = mid[0] + inward[0] * margin
    cy = mid[1] + inward[1] * margin

    try:
        if not outer_geom.contains(Point(cx, cy)):
            margin = glyph_h * 0.8
            cx = mid[0] + inward[0] * margin
            cy = mid[1] + inward[1] * margin
    except Exception:
        pass

    print(f"  [POST]   Грань: L={best_len:.1f}  угол={angle:.1f}°  "
          f"mid=({mid[0]:.1f},{mid[1]:.1f})  inward=({inward[0]:.2f},{inward[1]:.2f})")
    print(f"  [POST]   Итог: cx={cx:.2f}  cy={cy:.2f}  "
          f"glyph_h={glyph_h:.2f}  cell_h={text_h:.2f}  rot={angle:.1f}°")

    return cx, cy, text_h, angle


# ============================================================
# ПАРСЕР И ОТРИСОВКА ОДНОЛИНЕЙНОГО ШРИФТА
# ============================================================
def _parse_glyph(glyph_str):
    """'x,y x,y ... | x,y ...' -> [[(x,y), ...], [...]]"""
    result = []
    if not glyph_str:
        return result
    for part in glyph_str.split("|"):
        pts = []
        for token in part.split():
            token = token.strip()
            if not token:
                continue
            try:
                xs, ys = token.split(",")
                pts.append((float(xs), float(ys)))
            except Exception:
                continue
        if len(pts) >= 2:
            result.append(pts)
    return result


def draw_designation_stroke(msp, designation, cx, cy, height, rotation_deg,
                            layer, color):
    """Рисует обозначение однолинейным векторным шрифтом.
    height — ВЫСОТА ЯЧЕЙКИ (не чистая высота цифры).
    Чистая высота цифры = height * GLYPH_SPAN (при GLYPH_SPAN=0.9 → 0.9*height).
    Возвращает число созданных LWPOLYLINE."""
    if not designation or height <= 0:
        return 0

    norm = designation.replace("\r\n", "\n").replace("\r", "\n")
    lines = norm.split("\n")
    line_count = len(lines)

    LINE_SPACING = 1.4
    line_step = height * LINE_SPACING
    total_block_h = (line_count - 1) * line_step

    CHAR_ADV = 0.7
    CHAR_W   = 1.0

    ang = math.radians(rotation_deg)
    cos_r = math.cos(ang)
    sin_r = math.sin(ang)

    created = 0

    for li, line_text in enumerate(lines):
        n = len(line_text)
        if n == 0:
            continue

        y_off = total_block_h / 2.0 - li * line_step
        line_cx = cx - y_off * sin_r
        line_cy = cy + y_off * cos_r

        total_w = ((n - 1) * CHAR_ADV + CHAR_W) * height

        for i, ch in enumerate(line_text):
            glyph_str = STROKE_FONT.get(ch)
            if glyph_str is None:
                glyph_str = STROKE_FONT.get(ch.upper())
            if glyph_str is None:
                continue

            polylines = _parse_glyph(glyph_str)
            if not polylines:
                continue

            x_base = i * CHAR_ADV * height - total_w / 2.0

            for poly in polylines:
                world_pts = []
                for (px, py) in poly:
                    lx = x_base + px * CHAR_W * height
                    ly = (py - 0.5) * height
                    wx = line_cx + lx * cos_r - ly * sin_r
                    wy = line_cy + lx * sin_r + ly * cos_r
                    world_pts.append((wx, wy))

                try:
                    msp.add_lwpolyline(
                        world_pts, close=False,
                        dxfattribs={"layer": layer, "color": color},
                    )
                    created += 1
                except Exception as e:
                    print(f"  [POST]   Ошибка LWPOLYLINE глифа: {e}")

    return created


# ============================================================
# ПОСТОБРАБОТКА DXF
# ============================================================
def postprocess_dxf(dxf_path, designation=""):
    print(f"  [POST] Постобработка: {dxf_path}")
    print(f"  [POST] Обозначение для Mark: '{designation}'")

    if not os.path.exists(dxf_path):
        print(f"  [POST] Файл не найден: {dxf_path}")
        return
    try:
        doc = ezdxf.readfile(dxf_path)
    except Exception as e:
        print(f"  [POST] Не удалось открыть DXF: {e}")
        return

    msp = doc.modelspace()
    print(f"  [POST] Слои в DXF: {[l.dxf.name for l in doc.layers]}")
    print(f"  [POST] Типы линий в DXF: {[lt.dxf.name for lt in doc.linetypes]}")
    dump_color_diag(doc, "ЦВЕТА ДО ПЕРЕКРАСКИ")

    total_in_model = len(list(msp))
    print(f"  [POST] Всего сущностей в MODEL до обработки: {total_in_model}")

    # ---------- Шаг A1: ТТ и шероховатость ----------
    to_delete = []
    for e in msp.query('*[layer=="0"]'):
        if e.dxftype() in TYPES_TO_REMOVE_FROM_ZERO:
            to_delete.append(e)
    for e in msp:
        if e.dxf.get("layer", "") in LAYERS_TO_DELETE:
            to_delete.append(e)
    deleted = 0
    for e in set(to_delete):
        try:
            msp.delete_entity(e); deleted += 1
        except Exception:
            pass
    print(f"  [POST] Удалено объектов (ТТ/шероховатость): {deleted}")

    # ---------- Шаг A1.1: Удаление рамки/штампа КОМПАСа ----------
    frame_deleted = 0
    text_deleted = 0

    for name, cont in iter_layouts(doc):
        for e in list(cont):
            try:
                if e.dxftype() == "INSERT":
                    block_name = (e.dxf.name or "").strip()
                    if SYS_BLOCK_REGEX.match(block_name):
                        print(f"  [POST]   Удаляю INSERT '{block_name}' "
                              f"в [{name}] (рамка/штамп)")
                        cont.delete_entity(e)
                        frame_deleted += 1
                        continue

                if e.dxftype() in ("TEXT", "MTEXT"):
                    if e.dxftype() == "TEXT":
                        text = e.dxf.get("text", "") or ""
                    else:
                        text = e.text or ""
                    clean_text = re.sub(r"\\[A-Za-z][^;]*;", "", text)
                    clean_text = clean_text.replace(" ", "").replace("{", "").replace("}", "")
                    stripped = clean_text.strip()
                    if "1:1" in clean_text or stripped in ("", "(", ")", "()"):
                        print(f"  [POST]   Удаляю TEXT в [{name}]: "
                              f"'{text[:60]}'")
                        cont.delete_entity(e)
                        text_deleted += 1
            except Exception as ex:
                print(f"  [POST]   Ошибка в A1.1: {ex}")

    print(f"  [POST] A1.1: удалено INSERT рамок={frame_deleted}, "
          f"текстов '1:1'={text_deleted}")

    remaining = len(list(msp))
    print(f"  [POST] Сущностей в MODEL после A1.1: {remaining}")
    if remaining == 0:
        print("  [POST] !!! ВНИМАНИЕ: MODEL пуста после A1.1. "
              "Возможно, весь вид не отрисован КОМПАСом.")

    # ---------- Шаг A2: масштабный текст ----------
    remove_scale_text(msp)

    # ---------- Шаг A3: удаление линий переходов (страховка) ----------
    transition_deleted = 0
    for e in list(msp):
        if e.dxftype() not in ("LINE", "LWPOLYLINE", "POLYLINE", "ARC", "CIRCLE", "SPLINE"):
            continue
        layer_name = e.dxf.get("layer", "") or ""
        eff_lt = get_effective_linetype(doc, e)
        by_layer = layer_name in TRANSITION_LAYERS
        by_lt = bool(eff_lt and any(kw in eff_lt for kw in TRANSITION_LINETYPE_KEYWORDS))
        if by_layer or by_lt:
            try:
                msp.delete_entity(e)
                transition_deleted += 1
            except Exception:
                pass
    print(f"  [POST] Удалено линий переходов (страховка): {transition_deleted}")

    # ---------- Шаг B: контур (для насечек и позиции обозначения) ----------
    line_geoms = []
    for e in msp:
        if e.dxftype() in ("TEXT", "MTEXT"):
            continue
        if e.dxf.get("layer", "").lower() == DEFPOINTS_LAYER.lower():
            continue
        if e.dxftype() == "LINE" and is_bend_line(e, doc):
            continue
        g = entity_to_shapely(e)
        if g is not None and not g.is_empty:
            line_geoms.append(g)
    print(f"  [POST] Геометрий для анализа: {len(line_geoms)}")
    outer_geom, holes_geoms = find_outer_and_holes(line_geoms)
    if outer_geom is not None:
        print(f"  [POST] Контур детали (залитый): площадь={outer_geom.area:.1f}, "
              f"bbox={outer_geom.bounds}")
        print(f"  [POST] Отверстий найдено: {len(holes_geoms)}")
    else:
        print(f"  [POST] Контур не построен.")

    # ---------- Шаг D1: СОБИРАЕМ позиции линий сгиба ----------
    bend_positions = []
    bend_entities = []
    for e in msp:
        if e.dxftype() != "LINE" or not is_bend_line(e, doc):
            continue
        s, en = e.dxf.start, e.dxf.end
        bend_positions.append(((s.x, s.y), (en.x, en.y)))
        bend_entities.append(e)
    print(f"  [POST] Найдено линий сгиба: {len(bend_positions)}")

    # ---------- Шаг E: удаляем штриховые линии сгиба из DXF ----------
    removed = 0
    for ent in bend_entities:
        try:
            msp.delete_entity(ent); removed += 1
        except Exception:
            pass
    print(f"  [POST] Удалено штриховых линий сгиба: {removed}")

    # ---------- Шаг F: переносим ВСЁ (кроме Mark) на слой "0" и перекрашиваем ----------
    moved_to_geom = 0
    repainted_lines = 0
    for name, cont in iter_layouts(doc):
        for e in cont:
            try:
                cur_layer = e.dxf.get("layer", "") or ""
                if cur_layer.lower() == DEFPOINTS_LAYER.lower():
                    continue
                if cur_layer != GEOMETRY_LAYER:
                    e.dxf.layer = GEOMETRY_LAYER
                    moved_to_geom += 1
                e.dxf.color = COLOR_LINES
                repainted_lines += 1
            except Exception:
                pass
    print(f"  [POST] Перенесено на слой '{GEOMETRY_LAYER}': {moved_to_geom}, "
          f"перекрашено в основной цвет: {repainted_lines}")

    # ---------- Шаг D2: обозначение на Mark ----------
    if designation:
        if DEFPOINTS_LAYER not in doc.layers:
            doc.layers.add(name=DEFPOINTS_LAYER, color=COLOR_RED)
        else:
            try:
                doc.layers.get(DEFPOINTS_LAYER).dxf.color = COLOR_RED
            except Exception:
                pass

        cx, cy, text_h, rotation = compute_designation_layout(designation, outer_geom)
        print(f"  [POST]   Обозначение: {designation!r}  "
              f"cell_h={text_h:.2f} (glyph={text_h * GLYPH_SPAN:.2f} мм)  "
              f"({cx:.1f},{cy:.1f})  rot={rotation:.1f}°")

        try:
            n_polys = draw_designation_stroke(
                msp, designation, cx, cy, text_h, rotation,
                DEFPOINTS_LAYER, COLOR_RED,
            )
            if n_polys > 0:
                print(f"  [POST]   OK: {n_polys} LWPOLYLINE на слое "
                      f"'{DEFPOINTS_LAYER}' (stroke font)")
            else:
                print(f"  [POST]   ВНИМАНИЕ: не создано ни одной полилинии "
                      f"для {designation!r}")
        except Exception as ex:
            print(f"  [POST]   Ошибка отрисовки обозначения: {ex}")
            traceback.print_exc()
    else:
        print("  [POST] Обозначение пустое -> слой Mark не заполняется.")

    # ---------- Шаг D3: СОЗДАЁМ насечки сгиба (ЗЕЛЁНЫЕ, слой "0") ----------
    created, skipped_inner = 0, 0
    if outer_geom is None:
        print("  [POST] ВНИМАНИЕ: контур не построен -> насечки не создаются.")
    else:
        outer_ring = outer_geom.exterior
        for s, en in bend_positions:
            dx, dy = en[0] - s[0], en[1] - s[1]
            length = math.hypot(dx, dy)
            if length < 1e-6:
                continue
            ux, uy = dx / length, dy / length
            for pt, ix, iy in [(s, ux, uy), (en, -ux, -uy)]:
                dist = outer_ring.distance(Point(pt))
                if dist > OUTER_EDGE_TOL:
                    skipped_inner += 1
                    continue
                end_mark = (pt[0] + ix * BEND_LENGTH, pt[1] + iy * BEND_LENGTH)
                try:
                    msp.add_line(
                        (pt[0], pt[1], 0), end_mark,
                        dxfattribs={"layer": GEOMETRY_LAYER,
                                    "linetype": TARGET_BEND_LINETYPE,
                                    "color": COLOR_BEND},
                    )
                    created += 1
                except Exception as ex:
                    print(f"  [POST]   Ошибка насечки: {ex}")
    print(f"  [POST] Создано ЗЕЛЁНЫХ насечек: {created}, "
          f"пропущено (внутренних): {skipped_inner}")

    # ---------- Диагностика: что реально лежит на слое Mark ----------
    try:
        mark_ents = [e for e in msp
                     if (e.dxf.get("layer", "") or "").lower()
                        == DEFPOINTS_LAYER.lower()]
        types_on_mark = Counter(e.dxftype() for e in mark_ents)
        print(f"  [POST] На слое '{DEFPOINTS_LAYER}': "
              f"{len(mark_ents)} сущностей, типы={dict(types_on_mark)}")
    except Exception as e:
        print(f"  [POST] Диагностика слоя '{DEFPOINTS_LAYER}': {e}")

    # ---------- Шаг H: удалить все неиспользуемые слои ----------
    ALLOWED_LAYERS = {GEOMETRY_LAYER.lower(), DEFPOINTS_LAYER.lower()}
    removed_layers = []
    for layer in list(doc.layers):
        try:
            lname = layer.dxf.name
            if lname.lower() in ALLOWED_LAYERS:
                continue
            if lname.startswith("*"):
                continue
            doc.layers.remove(lname)
            removed_layers.append(lname)
        except Exception as e:
            print(f"  [POST]   Не удалось удалить слой '{layer.dxf.name}': {e}")

    print(f"  [POST] Удалено лишних слоёв: {len(removed_layers)} "
          f"-> {removed_layers}")
    print(f"  [POST] Остались слои: {[l.dxf.name for l in doc.layers]}")

    dump_color_diag(doc, "ЦВЕТА ПОСЛЕ ОБРАБОТКИ")

    try:
        doc.saveas(dxf_path)
        print(f"  [POST] Сохранено: {dxf_path}")
    except Exception as e:
        print(f"  [POST] Ошибка сохранения DXF: {e}")


# ============================================================
# 1. Подключение к КОМПАС-3D (API 7)
# ============================================================
try:
    api7_module = gencache.EnsureModule(
        "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0
    )
    kompas_app7 = Dispatch("Kompas.Application.7")
    iApplication7 = api7_module.IApplication(
        kompas_app7._oleobj_.QueryInterface(
            api7_module.IApplication.CLSID,
            win32com.client.pythoncom.IID_IDispatch,
        )
    )
except Exception as e:
    print("Не удалось подключиться к КОМПАС-3D. Убедитесь, что он запущен.")
    print("Ошибка:", e)
    sys.exit(1)

iApplication7.Visible = True

const7 = None
try:
    const7 = gencache.EnsureModule(
        "{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0
    ).constants
except Exception as e:
    print(f"  Не удалось получить константы API 7: {e}")

if const7 is not None:
    try:
        iApplication7.HideMessage = const7.ksHideMessageNo
        print("  Диалоги КОМПАС подавлены (HideMessage = ksHideMessageNo)")
    except Exception as e:
        print(f"  Не удалось установить HideMessage: {e}")

print("  Поиск конвертера DXF...")
RESOLVED_CONVERTER = find_dxf_converter()
if RESOLVED_CONVERTER:
    print(f"  Используется конвертер: {RESOLVED_CONVERTER}")

# ============================================================
# 2. Активный 3D-документ
# ============================================================
doc = iApplication7.ActiveDocument
if doc is None:
    print("Нет активного документа. Откройте сборку или деталь в КОМПАС-3D.")
    sys.exit(1)
try:
    doc3D = api7_module.IKompasDocument3D(doc)
    topPart = doc3D.TopPart
except Exception as e:
    print("Активный документ не является 3D-документом.")
    print("Ошибка:", e)
    sys.exit(1)

# ---- Определяем выходную папку: рядом с файлом активной модели -> DXF_kps ----
def _resolve_output_dir(active_doc, root_part):
    """Возвращает путь к подпапке DXF_kps рядом с файлом модели.
    Если модель не сохранена или путь недоступен — возвращает fallback."""
    candidates = []
    # 1) сам файл активного документа
    try:
        fn = active_doc.FileName
        if fn:
            candidates.append(fn)
    except Exception:
        pass
    # 2) путь верхней детали/сборки
    try:
        fn = root_part.FileName
        if fn:
            candidates.append(fn)
    except Exception:
        pass
    # 3) PathName у документа
    try:
        fn = active_doc.PathName
        if fn:
            candidates.append(fn)
    except Exception:
        pass

    for path in candidates:
        try:
            if not path:
                continue
            path = str(path)
            if not os.path.exists(path):
                continue
            base_dir = os.path.dirname(path)
            if not base_dir:
                continue
            return os.path.join(base_dir, OUTPUT_SUBDIR_NAME)
        except Exception:
            continue

    print("  Не удалось определить путь к модели — использую fallback:", FALLBACK_OUTPUT_DIR)
    return FALLBACK_OUTPUT_DIR

OUTPUT_DIR = _resolve_output_dir(doc, topPart)
try:
    os.makedirs(OUTPUT_DIR, exist_ok=True)
except Exception as e:
    print(f"  Не удалось создать папку '{OUTPUT_DIR}': {e}")
    print(f"  Переключаюсь на fallback: {FALLBACK_OUTPUT_DIR}")
    OUTPUT_DIR = FALLBACK_OUTPUT_DIR
    os.makedirs(OUTPUT_DIR, exist_ok=True)

print(f"Папка для DXF: {OUTPUT_DIR}")


# ============================================================
# 3. Рекурсивный обход компонентов
# ============================================================
def traverse_part(part, level=0):
    indent = "  " * level
    try:
        name = part.Name
    except Exception:
        name = "<без имени>"
    print(f"{indent}Обработка: {name}")
    process_sheet_metal(part, level)
    try:
        parts = part.Parts
    except Exception:
        return
    if parts is None:
        return
    try:
        count = parts.Count
    except Exception:
        return
    if count == 0:
        return
    for i in range(count):
        try:
            child = parts.Item(i)
        except Exception:
            continue
        if child is None:
            continue
        try:
            child_part = win32com.client.CastTo(child, "IPart7")
        except Exception:
            continue
        traverse_part(child_part, level + 1)


# ============================================================
# 4. Обработка листовой детали
# ============================================================
def _restore_straighten(states, indent):
    """Возвращает SheetMetalBody.Straighten в исходное состояние.
    states — список [(body, old_value), ...]"""
    if not states:
        return
    ok = 0
    for body, old_val in states:
        try:
            body.Straighten = old_val
            body.Update()
            ok += 1
        except Exception as e:
            print(f"{indent}  Не удалось восстановить Straighten: {e}")
    print(f"{indent}  Состояние развёртки восстановлено ({ok}/{len(states)} тел).")


def process_sheet_metal(part, level=0):
    indent = "  " * level
    try:
        sheet_container = api7_module.ISheetMetalContainer(part)
    except Exception:
        return
    try:
        sheet_bodies = sheet_container.SheetMetalBodies
    except Exception:
        return
    if sheet_bodies is None:
        return
    try:
        bodies_count = sheet_bodies.Count
    except Exception:
        return
    if bodies_count == 0:
        return
    print(f"{indent}  >>> Найдена листовая деталь: {part.Name} "
          f"(листовых тел: {bodies_count})")
    file_name = part.FileName
    if not file_name or not os.path.exists(file_name):
        print(f"{indent}  Файл не найден или виртуальный компонент.")
        return
    doc_part = iApplication7.Documents.Open(file_name, False, False)
    if doc_part is None:
        print(f"{indent}  Не удалось открыть: {file_name}")
        return

    try:
        doc3D_part = api7_module.IKompasDocument3D(doc_part)
        part7 = doc3D_part.TopPart
    except Exception as e:
        print(f"{indent}  Ошибка получения TopPart: {e}")
        return

    des_value = read_property_value(doc_part, part7, PROP_NAME_DESIGNATION, "МОДЕЛЬ")
    if not des_value:
        des_value = read_property_value(doc_part, part7, PROP_NAME_NAME, "МОДЕЛЬ")

    code_from_model = _extract_code_prefix(des_value)
    print(f"{indent}  Обозначение из модели: '{code_from_model}'")

    # --- Список (body, old_straighten) для отката ---
    straighten_states = []

    try:
        container = api7_module.ISheetMetalContainer(part7)
        bodies = container.SheetMetalBodies

        # 1) Запоминаем исходное состояние и разворачиваем
        for j in range(bodies.Count):
            body = bodies.SheetMetalBody(j)
            if body is None:
                continue
            try:
                old_val = bool(body.Straighten)
            except Exception:
                old_val = False
            straighten_states.append((body, old_val))
            body.Straighten = True
            body.Update()
        print(f"{indent}  Развёртка построена (тело разогнуто). "
              f"Запомнено состояний: {len(straighten_states)}")

        # 2) Включаем отображение линий сгиба
        bended_straightens = container.SheetMetalBendedStraightens
        if bended_straightens is not None:
            for k in range(bended_straightens.Count):
                bs = bended_straightens.SheetMetalBendedStraighten(k)
                if bs is not None:
                    bs.FoldLinesEnabled = True
                    bs.Update()
            print(f"{indent}  Отображение линий сгиба включено.")
    except Exception as e:
        print(f"{indent}  Ошибка при разгибании: {e}")
        traceback.print_exc()
        _restore_straighten(straighten_states, indent)
        return

    # --- Основная работа + ГАРАНТИРОВАННЫЙ ОТКАТ ---
    try:
        create_drawing_and_export(file_name, part.Name, code_from_model, level, part7)
    finally:
        # Восстанавливаем Straighten. Документ НЕ закрываем.
        _restore_straighten(straighten_states, indent)
# ============================================================
# 5. Чертеж «Развертка» -> DXF -> постобработка
# ============================================================
def create_drawing_and_export(model_file, part_name, design_from_model="",
                              level=0, model_part7=None):
    indent = "  " * level
    safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in part_name)
    dxf_path = os.path.join(OUTPUT_DIR, safe_name + ".dxf")
    if os.path.exists(dxf_path):
        try:
            os.remove(dxf_path)
        except Exception as e:
            print(f"{indent}  Не удалось удалить старый DXF: {e}")
            return
    drawing_doc = None
    try:
        drawing_doc = iApplication7.Documents.Add(1, True)
        if drawing_doc is None:
            print(f"{indent}  Не удалось создать документ чертежа.")
            return
        try:
            layout_sheets = drawing_doc.LayoutSheets
            if layout_sheets is not None and layout_sheets.Count > 0:
                sheet = layout_sheets.Item(0)
                sheet.LayoutLibraryFileName = ""
                sheet.LayoutStyleNumber = 0
                sheet.Update()
        except Exception as e:
            print(f"{indent}  Не удалось отключить оформление: {e}")
        try:
            drawing_doc_iface = drawing_doc._oleobj_.QueryInterface(
                api7_module.NamesToIIDMap['IDrawingDocument'],
                win32com.client.pythoncom.IID_IDispatch,
            )
            iDrawingDocument = api7_module.IDrawingDocument(drawing_doc_iface)
            try:
                td = iDrawingDocument.TechnicalDemand
                if td is not None: td.Delete()
            except Exception: pass
            try:
                sr = iDrawingDocument.SpecRough
                if sr is not None: sr.Delete()
            except Exception: pass
        except Exception as e:
            print(f"{indent}  Не удалось получить IDrawingDocument: {e}")
        doc2D = api7_module.IKompasDocument2D(drawing_doc)
        views = doc2D.ViewsAndLayersManager.Views
        view = views.Add(2)
        if view is None:
            print(f"{indent}  Не удалось создать вид.")
            return
        assoc_view = api7_module.IAssociationView(view)
        if assoc_view is None:
            print(f"{indent}  Вид не является ассоциативным.")
            return
        assoc_view.SourceFileName = model_file

        try:
            assoc_elements = api7_module.IAssociationViewElements(assoc_view)
            assoc_elements.CreateCentresMarkers = False
            assoc_elements.CreateCircularCentres = False
            assoc_elements.CreateLinearCentres = False
            assoc_elements.ProjectSpecRough = False
        except Exception as e:
            print(f"{indent}  Не удалось настроить элементы вида: {e}")

        try: assoc_view.ProjectionName = "#Развертка"
        except Exception: pass
        try: assoc_view.Unfold = True
        except Exception: pass
        try: assoc_view.Scale = 1.0
        except Exception: pass

        try:
            assoc_view.BendLinesVisible = True
            assoc_view.BendLinesStyle = 8
            print(f"{indent}  Линии сгиба ВКЛЮЧЕНЫ (стиль: пунктир 2)")
        except Exception as e:
            print(f"{indent}  Не удалось включить линии сгиба: {e}")

        try:
            assoc_view.BreakLinesVisible = TRANSITION_LINES_VISIBLE
            state = "ВКЛЮЧЕНЫ" if TRANSITION_LINES_VISIBLE else "ВЫКЛЮЧЕНЫ"
            print(f"{indent}  Линии переходов (BreakLinesVisible) {state}")
        except Exception as e:
            print(f"{indent}  Не удалось установить BreakLinesVisible: {e}")

        try:
            if TRANSITION_LINES_VISIBLE:
                assoc_view.BreakLinesStyle = 1
                print(f"{indent}  Стиль линий переходов: тонкая")
        except Exception:
            pass

        assoc_view.Update()
        view.Update()

        des_draw = ""
        if model_part7 is not None:
            des_draw_val = read_property_value(
                drawing_doc, model_part7, PROP_NAME_DESIGNATION, "ЧЕРТЁЖ")
            if not des_draw_val:
                des_draw_val = read_property_value(
                    drawing_doc, model_part7, PROP_NAME_NAME, "ЧЕРТЁЖ")
            des_draw = _extract_code_prefix(des_draw_val)
        print(f"{indent}  Обозначение из чертежа: '{des_draw}'")

        base = os.path.splitext(os.path.basename(model_file))[0]
        base_code = _extract_code_prefix(base) or base
        if des_draw:
            designation, src = des_draw, "свойства чертежа"
        elif design_from_model:
            designation, src = design_from_model, "свойства модели"
        else:
            designation, src = base_code, "имя файла"
        print(f"{indent}  Итоговое обозначение '{designation}' (источник: {src})")

        try:
            frames = drawing_doc.DocumentFrames
            if frames is not None and frames.Count > 0:
                frame = frames.Item(0)
                try:
                    frame.Activate()
                except AttributeError:
                    try:
                        frame.Active = True
                    except Exception as e:
                        print(f"{indent}  Не удалось активировать рамку: {e}")
                print(f"{indent}  Рамка активирована (для рендеринга вида)")
        except Exception as e:
            print(f"{indent}  Ошибка активации рамки: {e}")

        try:
            import time
            time.sleep(0.5)
        except Exception:
            pass

        saved = False
        if RESOLVED_CONVERTER:
            try:
                converter = iApplication7.Converter(RESOLVED_CONVERTER)
                if converter is not None:
                    converter.Convert("", dxf_path, 1, False)
                    saved = True
                else:
                    print(f"{indent}  Converter вернул None")
            except Exception as e:
                print(f"{indent}  Ошибка конвертера DXF: {e}")
        if not saved:
            try:
                if drawing_doc.SaveAs(dxf_path): saved = True
            except Exception as e:
                print(f"{indent}  Ошибка SaveAs: {e}")
        if saved and os.path.exists(dxf_path):
            postprocess_dxf(dxf_path, designation)
    except Exception as e:
        print(f"{indent}  Ошибка при создании чертежа: {e}")
        traceback.print_exc()
    finally:
        if drawing_doc is not None:
            try: drawing_doc.Close(False)
            except Exception: pass


# ============================================================
# 6. Запуск
# ============================================================
if __name__ == "__main__":
    try:
        print("=" * 60)
        traverse_part(topPart)
        print("=" * 60)
        print("\nГотово. Проверьте папку:", OUTPUT_DIR)
    except Exception as ex:
        print("\nОшибка выполнения:", ex)
        traceback.print_exc()

    try:
        if const7 is not None:
            iApplication7.HideMessage = const7.ksShowMessage
            print("  Диалоги КОМПАС возвращены в обычный режим")
    except Exception:
        pass

    try:
        input("\nНажмите Enter для выхода...")
    except Exception:
        pass