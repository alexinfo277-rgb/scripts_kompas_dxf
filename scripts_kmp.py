# -*- coding: utf-8 -*-
#| Экспорт развёрток листовых тел из активной сборки/детали в DXF
#| Совместим с КОМПАС-3D v21 и v25.
#| Универсальный автопоиск конвертера DXF (реестр + все диски).
#| Обозначение размещается ПАРАЛЛЕЛЬНО самой длинной грани контура.
#| УПРАВЛЕНИЕ линиями переходов через BreakLinesVisible (API 7).
#| Удаление рамки/штампа/масштаба 1:1 из экспортированного DXF.
#| Всего 2 слоя на выходе: "0" (геометрия + зелёные насечки) и "Defpoints".
#|
#| Версия под упаковку в EXE (PyInstaller).

import osgi
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
OUTPUT_DIR = r"C:\DXF"

DXF_CONVERTER_PATH = None  # None = искать автоматически

# === СЛОИ ===
GEOMETRY_LAYER = "0"          # слой для всей геометрии и насечек
DEFPOINTS_LAYER = "Mark" # слой для обозначения
DEFPOINT_TEXT_HEIGHT = 8.0

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
    outer = polys[0]
    holes = []
    for interior in outer.interiors:
        hp = Polygon(interior)
        if hp.area > MIN_HOLE_AREA:
            holes.append(hp)
    for p in polys[1:]:
        if MIN_HOLE_AREA < p.area < outer.area * 0.5:
            holes.append(p)
    return outer, holes


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
    if outer_geom is None:
        return 0.0, 0.0, DEFPOINT_TEXT_HEIGHT, 0.0

    minx, miny, maxx, maxy = outer_geom.bounds
    part_w = maxx - minx
    part_h = maxy - miny
    if part_w <= 0 or part_h <= 0:
        return (minx + maxx) / 2.0, (miny + maxy) / 2.0, DEFPOINT_TEXT_HEIGHT, 0.0

    n = max(1, len(designation))

    try:
        exterior = outer_geom.exterior
        coords = list(exterior.coords)
    except Exception:
        return (minx + maxx) / 2.0, (miny + maxy) / 2.0, DEFPOINT_TEXT_HEIGHT, 0.0

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
        return rp.x, rp.y, DEFPOINT_TEXT_HEIGHT, 0.0

    segments.sort(key=lambda s: (-s[0], -s[3][1]))

    best_len, p1, p2, mid = segments[0]
    x1, y1 = p1
    x2, y2 = p2
    dx = x2 - x1
    dy = y2 - y1
    L = math.hypot(dx, dy)
    if L < 1e-6:
        rp = outer_geom.representative_point()
        return rp.x, rp.y, DEFPOINT_TEXT_HEIGHT, 0.0

    angle = math.degrees(math.atan2(dy, dx))
    while angle > 90.0:
        angle -= 180.0
    while angle <= -90.0:
        angle += 180.0

    perp1 = (-dy / L, dx / L)
    perp2 = (dy / L, -dx / L)
    test_d = 0.5
    try:
        if outer_geom.contains(Point(mid[0] + perp1[0] * test_d,
                                     mid[1] + perp1[1] * test_d)):
            inward = perp1
        else:
            inward = perp2
    except Exception:
        inward = perp1

    h_by_seg = (best_len * 0.85) / (n * 0.6) if n > 0 else DEFPOINT_TEXT_HEIGHT
    h_by_size = min(part_w, part_h) * 0.06
    text_h = min(DEFPOINT_TEXT_HEIGHT, h_by_seg, h_by_size)
    if text_h < 2.5:
        text_h = 2.5
    if text_h > min(part_w, part_h) * 0.15:
        text_h = min(part_w, part_h) * 0.15

    margin = text_h * 1.3
    cx = mid[0] + inward[0] * margin
    cy = mid[1] + inward[1] * margin

    try:
        if not outer_geom.contains(Point(cx, cy)):
            margin = text_h * 0.8
            cx = mid[0] + inward[0] * margin
            cy = mid[1] + inward[1] * margin
    except Exception:
        pass

    print(f"  [POST]   Грань: L={best_len:.1f}  угол={angle:.1f}°  "
          f"mid=({mid[0]:.1f},{mid[1]:.1f})  inward=({inward[0]:.2f},{inward[1]:.2f})")
    print(f"  [POST]   h_by_seg={h_by_seg:.2f}  h_by_size={h_by_size:.2f}  "
          f"→ text_h={text_h:.2f}  margin={margin:.2f}")
    print(f"  [POST]   Итог: cx={cx:.2f}  cy={cy:.2f}  rotation={angle:.1f}°")

    return cx, cy, text_h, angle

def text_entity_to_polylines(msp, text_entity, target_layer, target_color,
                             font_name="Arial"):
    """Преобразует TEXT/MTEXT в набор LWPOLYLINE (обводка глифов).
    Возвращает количество созданных полилиний."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        from matplotlib.textpath import TextPath
        from matplotlib.font_manager import FontProperties
    except ImportError as e:
        print(f"  [POST] matplotlib не установлен — текст останется TEXT: {e}")
        return 0

    if text_entity.dxftype() == "TEXT":
        content  = text_entity.dxf.text or ""
        height   = float(text_entity.dxf.height)
        rotation = float(text_entity.dxf.get("rotation", 0.0))
        insert   = text_entity.dxf.insert
    elif text_entity.dxftype() == "MTEXT":
        content  = text_entity.text or ""
        height   = float(text_entity.dxf.char_height)
        rotation = float(text_entity.dxf.get("rotation", 0.0))
        insert   = text_entity.dxf.insert
    else:
        return 0

    if not content.strip():
        return 0

    fp = FontProperties(family=font_name)
    try:
        tp = TextPath((0, 0), content, size=height, prop=fp)
    except Exception as e:
        print(f"  [POST] TextPath не построен: {e}")
        return 0

    # Центрирование по bbox = MIDDLE_CENTER
    bbox = tp.get_extents()
    cx_t = (bbox.x0 + bbox.x1) / 2.0
    cy_t = (bbox.y0 + bbox.y1) / 2.0

    ang = math.radians(rotation)
    ca, sa = math.cos(ang), math.sin(ang)
    ix, iy = insert.x, insert.y

    # to_polygons уже разворачивает безье в ломаную
    polys = tp.to_polygons(closed_only=False)
    created = 0
    for poly in polys:
        pts = []
        for (x, y) in poly:
            x -= cx_t
            y -= cy_t
            xr = x * ca - y * sa
            yr = x * sa + y * ca
            pts.append((ix + xr, iy + yr))
        if len(pts) < 2:
            continue
        closed = (len(pts) >= 3
                  and abs(pts[0][0] - pts[-1][0]) < 1e-6
                  and abs(pts[0][1] - pts[-1][1]) < 1e-6)
        try:
            msp.add_lwpolyline(
                pts, close=closed,
                dxfattribs={"layer": target_layer, "color": target_color},
            )
            created += 1
        except Exception as e:
            print(f"  [POST] Ошибка LWPOLYLINE: {e}")
    return created

# ============================================================
# ПОСТОБРАБОТКА DXF
# ============================================================
def postprocess_dxf(dxf_path, designation=""):
    print(f"  [POST] Постобработка: {dxf_path}")
    print(f"  [POST] Обозначение для Defpoints: '{designation}'")

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
        print(f"  [POST] Контур детали: площадь={outer_geom.area:.1f}, "
              f"bbox={outer_geom.bounds}")
        print(f"  [POST] Отверстий найдено: {len(holes_geoms)}")
    else:
        print(f"  [POST] Контур не построен.")

    # ---------- Шаг D1: СОБИРАЕМ позиции линий сгиба (пока не создаём) ----------
    bend_positions = []  # список (start, end)
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

    # ---------- Шаг F: переносим ВСЁ (кроме Defpoints) на слой "0" и перекрашиваем ----------
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
                # Красим в основной цвет
                e.dxf.color = COLOR_LINES
                repainted_lines += 1
            except Exception:
                pass
    print(f"  [POST] Перенесено на слой '{GEOMETRY_LAYER}': {moved_to_geom}, "
          f"перекрашено в основной цвет: {repainted_lines}")

    # ---------- Шаг D2: обозначение на Defpoints ----------
    if designation:
        if DEFPOINTS_LAYER not in doc.layers:
            doc.layers.add(name=DEFPOINTS_LAYER, color=COLOR_RED)
        else:
            try: doc.layers.get(DEFPOINTS_LAYER).dxf.color = COLOR_RED
            except Exception: pass

        cx, cy, text_h, rotation = compute_designation_layout(designation, outer_geom)
        print(f"  [POST]   Размер/позиция: h={text_h:.2f}  "
              f"({cx:.1f},{cy:.1f})  rot={rotation:.1f}°")

        try:
            txt = msp.add_text(
                designation,
                dxfattribs={"layer": DEFPOINTS_LAYER,
                            "color": COLOR_RED,
                            "height": text_h,
                            "rotation": rotation,
                            "insert": (cx, cy, 0)},
            )
            txt.dxf.insert = (cx, cy, 0)

            n_polys = text_entity_to_polylines(
                msp, txt, DEFPOINTS_LAYER, COLOR_RED, font_name="Arial"
            )
            msp.delete_entity(txt)  # сам TEXT больше не нужен
            print(f"  [POST]   Defpoints '{designation}' в ({cx:.1f},{cy:.1f}) "
                  f"rot={rotation:.1f}° -> {n_polys} LWPOLYLINE")
        except Exception as ex:
            print(f"  [POST]   Ошибка записи Defpoints: {ex}")
    else:
        print("  [POST] Обозначение пустое -> Defpoints не заполняется.")

    # ---------- Шаг D3: СОЗДАЁМ насечки сгиба (ЗЕЛЁНЫЕ, слой "0") ----------
    # Делаем это ПОСЛЕ перекраски, чтобы наш зелёный цвет никто не стёр.
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

    try:
        container = api7_module.ISheetMetalContainer(part7)
        bodies = container.SheetMetalBodies
        for j in range(bodies.Count):
            body = bodies.SheetMetalBody(j)
            if body is not None:
                body.Straighten = True
                body.Update()
        print(f"{indent}  Развёртка построена (тело разогнуто).")
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
        return

    create_drawing_and_export(file_name, part.Name, code_from_model, level, part7)


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