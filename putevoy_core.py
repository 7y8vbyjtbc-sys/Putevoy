"""
Ядро генератора путевых листов: расчёт расстояний офлайн, справочник
магазинов, генерация поездок и запись в шаблон .xlsm.

Расстояния считаются по населённому пункту из адреса: встроенная таблица
расстояний по автодорогам между населёнными пунктами Мурманской области
(«Мурманавтодор», опубликована на kolamap.ru), от центра до центра.
"""
import calendar
import datetime as dt
import json
import os
import random
import re
import warnings
import zipfile
from collections import Counter
from decimal import Decimal, ROUND_HALF_UP
from xml.sax.saxutils import escape

import openpyxl

warnings.filterwarnings("ignore", module="openpyxl")


class UserError(Exception):
    """Понятная пользователю ошибка (показывается в интерфейсе как есть)."""


# ============================================================ настройки
MAX_ADJUST_KM = 20          # максимальная подгонка одного расстояния, км
MORNING_DEPART = (9, 0)     # выезд из офиса
EVENING_DEPART = (18, 0)    # выезд обратно в офис
AVG_SPEED_KMH = 80          # для расчёта времени заезда

RU_HOLIDAYS = {(1, d) for d in range(1, 9)} | {(2, 23), (3, 8), (5, 1), (5, 9), (6, 12), (11, 4)}

MONTHS_GEN = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля",
              "августа", "сентября", "октября", "ноября", "декабря"]
MONTHS_NOM = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль",
              "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]

FRONT = "Лицевая сторона"
BACK1 = "Оборотная сторона"
BACK2 = "Оборотная сторона2"
COLS_RIGHT = dict(date="J", tout="K", tin="L", src="M", dst="N", odo="O", km="Q", kind="R")
COLS_LEFT = dict(date="A", tout="B", tin="C", src="D", dst="E", odo="F", km="H", kind="I")
ROWS = range(5, 52, 2)
SLOTS = ([(FRONT, r, COLS_RIGHT) for r in ROWS]
         + [(BACK1, r, COLS_LEFT) for r in ROWS] + [(BACK1, r, COLS_RIGHT) for r in ROWS]
         + [(BACK2, r, COLS_LEFT) for r in ROWS] + [(BACK2, r, COLS_RIGHT) for r in ROWS])

FIELDS = [
    ("B8", "Организация", "str"),
    ("C9", "Марка автомобиля", "str"),
    ("C10", "Госномер", "str"),
    ("C11", "Водитель (ФИО полностью)", "str"),
    ("B12", "СНИЛС", "str"),
    ("B13", "Удостоверение, серия", "num"),
    ("D13", "Удостоверение, номер", "num"),
    ("F13", "Удостоверение, дата выдачи (дд.мм.гггг)", "date"),
    ("B19", "В распоряжение", "str"),
    ("B21", "Адрес подачи", "str"),
    ("E50", "Тариф, руб. за 1 км", "num"),
    ("F47", "Основание (договор аренды)", "str"),
    ("B53", "Арендатор, должность", "str"),
    ("G53", "Арендатор, ФИО", "str"),
    ("B56", "Арендодатель, должность", "str"),
]
DEFAULT_KIND = "Междугородное/Перевозка для нужд компании"


# ============================================================ таблица расстояний
# Нижний треугольник: каждая строка — расстояния до всех пунктов выше неё.
_MATRIX = """
Мурманск
Кировск 208
Абрам-Мыс 27 220
Апатиты 192 16 199
Верхнетуломский 79 273 77 254
Видяево 84 277 60 260 138
Выходной 21 200 21 183 75 81
Заозерск 118 313 95 294 170 94 114
Заполярный 169 363 145 346 220 145 165 95
Зеленоборский 295 181 305 160 358 365 287 398 448
Кандалакша 248 128 251 111 306 311 235 345 395 56
Кильдинстрой 23 194 29 176 83 89 11 123 173 230 226
Килпъявр 64 258 39 240 116 41 59 75 126 343 288 67
Кола 14 207 13 188 67 72 7 106 157 291 236 15 52
Ковдор 294 177 301 160 355 359 280 394 444 206 152 276 338 286
Лавна 34 227 9 210 95 60 29 96 144 312 260 36 39 20 307
Лиинахамари 157 350 134 332 208 132 152 82 39 434 381 160 113 143 431 132
Ловозеро 185 189 190 171 245 250 172 284 334 274 231 166 230 170 270 198 321
Минькино 31 224 6 206 81 59 25 92 143 309 256 34 38 17 305 7 131 195
Мишуково 43 235 18 217 93 58 36 92 143 320 267 45 37 29 316 16 130 206 16
Молочный 17 201 16 183 70 77 4 111 161 288 233 11 57 3 282 24 148 173 21 32
Мончегорск 136 85 143 67 197 201 124 235 286 175 118 118 182 129 167 151 273 112 146 158 123
Снежногорск 75 268 50 250 126 71 69 126 176 354 302 79 70 51 351 49 163 221 46 36 64 191
Мурмаши 32 214 30 196 34 89 16 123 174 300 246 23 69 17 295 37 163 171 34 46 17 137 73
Никель 204 397 179 378 254 179 195 129 24 481 428 207 151 182 465 179 63 339 169 167 183 310 200 197
Оленегорск 110 113 117 95 170 175 96 208 259 193 145 69 154 102 193 122 246 85 119 131 98 35 163 109 284
Алакуртти 348 234 351 216 403 406 334 441 494 140 111 331 388 338 257 358 482 323 356 366 336 216 398 349 522 247
Печенга 147 339 122 321 196 121 140 65 27 423 370 150 102 132 420 121 12 292 120 119 129 260 152 149 51 235 471
Полярный 70 262 45 244 119 82 65 115 172 346 294 73 64 55 344 43 157 217 42 32 58 185 14 72 196 158 393 146
Причальное 29 221 26 203 53 84 22 119 168 306 256 31 65 15 302 32 155 176 30 42 18 143 73 32 194 118 352 144 67
Пушной 62 156 70 139 123 128 50 162 213 242 188 45 108 56 237 77 200 112 73 65 51 79 118 63 234 51 292 190 112 71
Росляково 13 227 47 208 96 101 34 137 187 309 257 34 83 30 304 53 174 181 48 57 33 147 92 46 215 120 359 164 36 45 71
Сафоново 17 229 50 212 102 105 38 140 192 314 262 39 88 35 309 55 179 184 52 61 37 151 96 50 219 124 363 168 90 49 75 4
Североморск 28 239 61 223 114 114 43 153 203 325 273 50 99 46 318 64 188 193 61 70 46 160 105 59 228 133 372 177 99 58 84 13 9
Ревда 155 177 162 160 214 217 143 250 306 262 209 135 198 149 258 169 292 27 166 176 144 100 209 156 330 73 310 281 194 162 99 167 171 180
Спутник 139 332 119 315 190 114 134 54 32 417 365 144 94 125 414 113 19 285 113 111 128 256 144 142 57 228 461 8 138 137 182 158 162 171 272
Туманный 137 321 143 303 195 200 130 234 284 406 355 132 180 128 402 146 271 268 146 156 130 244 139 143 308 216 452 260 183 142 171 129 133 142 255 253
Териберка 136 320 142 302 194 199 129 233 283 405 354 131 179 127 401 145 270 267 145 155 129 243 188 142 307 215 451 259 182 141 170 127 132 141 254 252 84
Тайбола 68 149 75 151 128 207 56 168 218 234 181 50 114 62 230 82 205 105 78 86 56 72 123 69 241 45 284 194 117 76 6 79 83 92 92 187 176 175
Титовка 112 306 87 288 162 86 107 34 59 390 333 117 67 99 387 86 47 260 86 85 101 229 118 115 93 201 437 35 108 110 155 130 134 143 247 27 226 225 160
Умба 359 242 366 226 419 426 347 459 509 172 115 341 405 353 278 375 496 335 370 381 347 231 414 360 532 258 222 482 408 367 302 370 375 383 322 475 465 464 293 450
Ура-Губа 79 263 55 255 128 7 74 88 138 356 303 82 34 66 352 42 125 224 51 50 67 194 35 81 162 168 403 113 45 78 121 95 99 108 211 107 192 191 126 80 415
Шонгуй 34 185 42 167 94 100 21 134 185 267 220 17 80 25 266 48 174 131 45 55 22 108 90 35 205 82 324 157 82 42 35 45 49 58 128 153 142 141 40 126 331 92
Кица 60 161 51 143 104 110 32 144 194 245 198 42 90 38 238 76 196 119 73 83 50 86 108 63 233 60 296 185 110 70 13 67 71 80 106 175 164 163 18 154 303 114 34
Тулома 32 224 29 206 48 86 25 122 172 313 257 37 68 18 305 37 160 179 33 43 22 147 77 35 197 120 360 146 72 4 74 49 53 62 166 140 145 144 78 114 378 33 45 72
Полярные Зори 221 106 227 88 279 282 206 314 370 87 28 201 261 211 128 237 357 194 234 244 208 91 273 221 394 120 143 345 269 227 164 232 236 245 181 337 328 327 156 312 141 280 193 171 231
"""

# Пункты, которых нет в таблице: (ближайший пункт таблицы, добавочные км)
_DERIVED = {
    "Гаджиево": ("Снежногорск", 5),
    "Высокий": ("Оленегорск", 5),
    "Титан": ("Апатиты", 8),
}


def _parse_matrix():
    names, dist = [], {}
    for line in _MATRIX.strip().splitlines():
        parts = line.split()
        nums = [p for p in parts if p.isdigit()]
        name = " ".join(p for p in parts if not p.isdigit())
        if len(nums) != len(names):
            raise RuntimeError(f"Ошибка во встроенной таблице расстояний: {name}")
        for other, km in zip(names, nums):
            dist[(name, other)] = dist[(other, name)] = int(km)
        names.append(name)
    return names, dist


TOWNS, _DIST = _parse_matrix()
ALL_TOWNS = TOWNS + list(_DERIVED)


def town_distance(a, b):
    """Расстояние по дорогам между населёнными пунктами, км."""
    if a == b:
        return 0
    ba, oa = _DERIVED.get(a, (a, 0))
    bb, ob = _DERIVED.get(b, (b, 0))
    base = 0 if ba == bb else _DIST[(ba, bb)]
    return base + oa + ob


def _norm(s):
    return str(s or "").lower().replace("ё", "е")


_MARK = r"(?:г|гор|город|пгт|нп|с|село|п|пос|поселок|рп|д|мкр|р-н)"
_L = r"[а-яa-z0-9]"
_TOWN_PATTERNS = []
for _t in ALL_TOWNS:
    _n = re.escape(_norm(_t)).replace(r"\ ", r"\s+")
    _TOWN_PATTERNS.append((_t,
                           re.compile(rf"(?<!{_L}){_n}\s*{_MARK}\.?(?!{_L})"),
                           re.compile(rf"(?<!{_L}){_MARK}\.?\s*{_n}(?!{_L})"),
                           re.compile(rf"(?<!{_L}){_n}(?!{_L})")))


def find_town(address):
    """Населённый пункт из адреса. Берётся самый «внутренний» (последний),
    чтобы «Североморск г, Сафоново пгт» дал Сафоново. Названия улиц
    («Полярные Зори ул») не путаются с городами: нужен признак г/пгт/нп/с/п."""
    a = _norm(address)
    if "•" in a:                                    # формат «Россия•…•Апатиты»
        for token in reversed(a.split("•")):
            for t, *_ in _TOWN_PATTERNS:
                if _norm(t) == token.strip():
                    return t
    best = None
    for t, p1, p2, _ in _TOWN_PATTERNS:
        for p in (p1, p2):
            for m in p.finditer(a):
                if best is None or m.start() > best[0]:
                    best = (m.start(), t)
    if best:
        return best[1]
    best = None                                     # запасной вариант: просто название
    for t, _, _, p3 in _TOWN_PATTERNS:
        for m in p3.finditer(a):
            if best is None or m.start() > best[0]:
                best = (m.start(), t)
    return best[1] if best else None


# ============================================================ числа
def D(x):
    return Decimal(0) if x is None else Decimal(str(x))


def to_tenths(x):
    return int((D(x) * 10).quantize(Decimal(1), ROUND_HALF_UP))


def tenths_str(t):
    return f"{t // 10}" if t % 10 == 0 else f"{t // 10}.{t % 10}"


def excel_round2(x):
    return x.quantize(Decimal("0.01"), ROUND_HALF_UP)


def fact_liters(km_tenths, norm):
    return excel_round2(Decimal(km_tenths) / 10 * norm / 100)


def parse_decimal(text, what):
    try:
        v = Decimal(str(text).strip().replace(",", ".").replace(" ", ""))
    except Exception:
        raise UserError(f"{what}: введите число, например 295 или 7,5")
    if v <= 0:
        raise UserError(f"{what}: значение должно быть больше нуля")
    return v


def target_tenths(liters, norm):
    """Пробег (в десятых км), при котором шаблон покажет ровно `liters`."""
    exact = liters * 100 / norm
    t0 = int((exact * 10).to_integral_value())
    cands = [t for t in range(t0 - 30, t0 + 31) if fact_liters(t, norm) == liters]
    if not cands:
        raise UserError(f"Не удаётся подобрать пробег для {liters} л при норме {norm}.")
    whole = [t for t in cands if t % 10 == 0]
    return min(whole or cands, key=lambda t: abs(Decimal(t) / 10 - exact))


# ============================================================ справочник магазинов
def load_db(path):
    db = {}
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            db = json.load(f)
    db.setdefault("min_km", 60)
    db.setdefault("prefix", {"МД": "ММ", "МК": "МК"})
    db.setdefault("excluded", {})
    db.setdefault("shops", {})
    return db


def save_db(db, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=2)


def shop_display_name(name, typ, prefix):
    name = str(name).strip()
    typ = str(typ or "").strip()
    if typ in prefix:
        return f"{prefix[typ]} {name}"
    if name.lower().startswith("аптека") or not typ:
        return name
    return f"{typ} {name}"


def is_shop_list(path):
    """Похоже ли на список магазинов (есть столбцы «Магазин» и «Полный адрес»)."""
    try:
        return _find_shop_header(openpyxl.load_workbook(path, read_only=True)) is not None
    except Exception:
        return False


def _find_shop_header(wb):
    for ws in wb.worksheets:
        for r_idx, row in enumerate(ws.iter_rows(max_row=10, values_only=True), 1):
            cells = [_norm(c).strip() for c in row]
            if "магазин" in cells and any("адрес" in c for c in cells):
                col = {"name": cells.index("магазин"),
                       "addr": next(i for i, c in enumerate(cells) if "адрес" in c),
                       "type": cells.index("тип") if "тип" in cells else None}
                return ws, r_idx, col
    return None


def import_shops(path, db):
    """Заменяет справочник актуальным списком из Excel. Возвращает отчёт."""
    found = _find_shop_header(openpyxl.load_workbook(path, read_only=True, data_only=True))
    if not found:
        raise UserError("В файле не найдены столбцы «Магазин» и «Полный адрес».")
    ws, header_row, col = found
    new = {}
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        name = row[col["name"]] if col["name"] < len(row) else None
        addr = row[col["addr"]] if col["addr"] < len(row) else None
        if not name or not addr:
            continue
        typ = row[col["type"]] if col["type"] is not None and col["type"] < len(row) else None
        new[shop_display_name(name, typ, db["prefix"])] = str(addr).strip()
    if not new:
        raise UserError("В списке не найдено ни одного магазина с адресом.")
    old = db["shops"]
    report = dict(
        added=sorted(set(new) - set(old)),
        removed=sorted(set(old) - set(new)),
        changed=sorted(k for k in set(new) & set(old) if new[k] != old[k]),
        total=len(new),
    )
    db["shops"] = dict(sorted(new.items()))
    return report


def compute_points(db, office_address):
    """Расстояния от офиса до всех магазинов.
    Возвращает (office_town, points{имя: км}, skipped{имя: причина})."""
    office_town = find_town(office_address)
    if not office_town:
        raise UserError("Не удалось определить населённый пункт в адресе выезда.\n"
                        "Укажите его явно, например: «г. Мурманск, ул. Карла Либкнехта, 28».")
    points, skipped = {}, {}
    for name, addr in db["shops"].items():
        if name in db["excluded"]:
            skipped[name] = db["excluded"][name]
            continue
        town = find_town(addr)
        if not town:
            skipped[name] = "не определён населённый пункт"
            continue
        km = town_distance(office_town, town)
        if km == 0:
            skipped[name] = f"в том же населённом пункте ({town})"
        elif km < db["min_km"]:
            skipped[name] = f"ближе {db['min_km']} км ({town}, {km} км)"
        else:
            points[name] = km
    return office_town, points, skipped


# ============================================================ прошлый путевой лист
def is_trip_sheet(path):
    try:
        return FRONT in openpyxl.load_workbook(path, read_only=True).sheetnames
    except Exception:
        return False


def read_prev(path):
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        raise UserError(f"Не удалось открыть файл: {e}")
    if FRONT not in wb.sheetnames:
        raise UserError("Это не путевой лист: в файле нет листа «Лицевая сторона».")
    fr = wb[FRONT]
    info = {"path": path}

    m = re.search(r"([а-яё]+)\s+(\d{4})", str(fr["C6"].value or "").lower())
    if m and m.group(1) in MONTHS_GEN:
        info["month"] = (int(m.group(2)), MONTHS_GEN.index(m.group(1)) + 1)
    elif isinstance(fr["J5"].value, dt.datetime):
        info["month"] = (fr["J5"].value.year, fr["J5"].value.month)
    else:
        raise UserError("Не удалось определить месяц путевого листа (ячейка C6).")
    y, mo = info["month"]
    info["next_month"] = (y + 1, 1) if mo == 12 else (y, mo + 1)

    info["consts"] = {cell: fr[cell].value for cell, _, _ in FIELDS}
    info["end_odo"] = to_tenths(fr["H39"].value)
    issued = D(fr["H32"].value)
    rest_start = D(fr["H33"].value)
    info["norm"] = D(fr["H35"].value)

    trips = []
    for sheet, row, c in SLOTS:
        ws = wb[sheet]
        if ws[f"{c['date']}{row}"].value is None:
            continue
        trips.append(dict(src=str(ws[f"{c['src']}{row}"].value or "").strip(),
                          odo=to_tenths(ws[f"{c['odo']}{row}"].value),
                          km=to_tenths(ws[f"{c['km']}{row}"].value),
                          kind=ws[f"{c['kind']}{row}"].value))
    total = sum(t["km"] for t in trips)
    fact = fact_liters(total, info["norm"]) if info["norm"] else Decimal(0)
    info["rest_end"] = max(Decimal(0), issued + rest_start - fact)

    kinds = Counter(t["kind"] for t in trips if t["kind"])
    info["kind"] = kinds.most_common(1)[0][0] if kinds else DEFAULT_KIND
    # Адрес выезда: точка выезда первой поездки, иначе «Адрес подачи»
    office = trips[0]["src"] if trips else ""
    if not office:
        office = re.sub(r"^\s*офис\s*,\s*", "", str(fr["B21"].value or ""), flags=re.I)
    info["office"] = office

    info["warning"] = ""
    if trips and trips[-1]["odo"] + trips[-1]["km"] != info["end_odo"]:
        info["warning"] = (f"Одометр на конец месяца ({tenths_str(info['end_odo'])}) не совпадает "
                           f"с последней поездкой ({tenths_str(trips[-1]['odo'] + trips[-1]['km'])}). "
                           f"Взято значение из ячейки H39.")
    return info


def describe_prev(info):
    y, m = info["month"]
    ny, nm = info["next_month"]
    c = info["consts"]
    return (f"{MONTHS_NOM[m - 1]} {y} · {c.get('C11') or ''} · {c.get('C9') or ''} {c.get('C10') or ''}",
            f"Будет создан лист на {MONTHS_NOM[nm - 1].lower()} {ny}")


# ============================================================ генерация
def workdays(year, month, weekends=False):
    days = []
    for d in range(1, calendar.monthrange(year, month)[1] + 1):
        day = dt.date(year, month, d)
        if not weekends and (day.weekday() >= 5 or (month, d) in RU_HOLIDAYS):
            continue
        days.append(day)
    return days


def _pick_sequence(n, names, rng):
    seq, used = [], Counter()
    for _ in range(n):
        choices = [x for x in names if not seq or x != seq[-1]] or names
        weights = [1.0 / (1 + used[x]) ** 2 for x in choices]
        x = rng.choices(choices, weights)[0]
        seq.append(x)
        used[x] += 1
    return seq


def _leg_bounds(base_t):
    adj = MAX_ADJUST_KM * 10
    return max(10, base_t - adj), base_t + adj


def _plan(target, days, points, rng, tries=4000):
    names = list(points)
    if not names:
        raise UserError("Нет ни одного магазина, до которого можно посчитать расстояние.")
    base = {x: points[x] * 10 for x in names}
    mean = sum(base.values()) / len(base)
    n0 = max(1, round(target / (2 * mean)))
    ns = [n for n in range(n0 - 3, n0 + 4) if 1 <= n <= len(days)] or [len(days)]
    best = None
    for n in ns:
        for _ in range(tries // len(ns)):
            seq = _pick_sequence(n, names, rng)
            gap = target - 2 * sum(base[x] for x in seq)
            up = sum(2 * (_leg_bounds(base[x])[1] - base[x]) for x in seq)
            down = sum(2 * (base[x] - _leg_bounds(base[x])[0]) for x in seq)
            if not (-down <= gap <= up):
                continue
            score = abs(gap) / n
            if best is None or score < best[0]:
                best = (score, seq, gap)
    if best is None:
        raise UserError("Не удалось уложить пробег в рабочие дни с подгонкой ±20 км.\n"
                        "Разрешите поездки в выходные или проверьте количество литров.")
    return best[1], best[2]


def _distribute(seq, gap, points, rng):
    legs = [[points[x] * 10, points[x] * 10] for x in seq]
    bounds = [_leg_bounds(points[x] * 10) for x in seq]
    remaining = gap
    order = list(range(len(seq)))
    rng.shuffle(order)
    sign = 1 if remaining > 0 else -1
    while abs(remaining) >= 20:
        moved = False
        for i in order:
            if abs(remaining) < 20:
                break
            lo, hi = bounds[i]
            nv = legs[i][0] + sign * 10
            if lo <= nv <= hi:
                legs[i][0] = legs[i][1] = nv
                remaining -= sign * 20
                moved = True
        if not moved:
            break
    rng.shuffle(order)
    for i in order:
        if remaining == 0:
            break
        for k in (1, 0):
            lo, hi = bounds[i]
            nv = min(hi, max(lo, legs[i][k] + remaining))
            remaining -= nv - legs[i][k]
            legs[i][k] = nv
            if remaining == 0:
                break
    if remaining != 0:
        raise UserError("Внутренняя ошибка распределения пробега.")
    return legs


def _arrival(km_t, depart):
    hours = Decimal(km_t) / 10 / AVG_SPEED_KMH
    half_hours = max(1, int((hours * 2).quantize(Decimal(1), ROUND_HALF_UP)))
    end = min(depart[0] * 60 + depart[1] + half_hours * 30, 23 * 60 + 30)
    return dt.time(end // 60, end % 60)


def generate(info, db, office, liters, norm, weekends=False, seed=None):
    """Строит поездки на следующий месяц. Возвращает словарь с результатом."""
    office = office.strip()
    office_town, points, skipped = compute_points(db, office)
    rest_start = info["rest_end"]
    target_liters = excel_round2(liters + rest_start)
    target = target_tenths(target_liters, norm)
    year, month = info["next_month"]
    days = workdays(year, month, weekends)
    rng = random.Random(seed)
    seq, gap = _plan(target, days, points, rng)
    legs = _distribute(seq, gap, points, rng)
    if len(seq) == len(days):
        chosen = days
    else:
        chosen = [days[i] for i in sorted(rng.sample(range(len(days)), len(seq)))]
    trips, odo = [], info["end_odo"]
    for day, point, (out_t, back_t) in zip(chosen, seq, legs):
        for src, dst, km, dep in ((office, point, out_t, MORNING_DEPART), (point, office, back_t, EVENING_DEPART)):
            trips.append(dict(date=day, tout=dt.time(*dep), tin=_arrival(km, dep),
                              src=src, dst=dst, km=km, odo=odo))
            odo += km
    if len(trips) > len(SLOTS):
        raise UserError(f"Поездок ({len(trips)}) больше, чем строк в шаблоне ({len(SLOTS)}).")
    total = sum(t["km"] for t in trips)
    fact = fact_liters(total, norm)
    return dict(trips=trips, office=office, office_town=office_town, liters=liters, norm=norm,
                rest_start=rest_start, start_odo=info["end_odo"], end_odo=odo, total=total,
                fact=fact, rest_end=max(Decimal(0), liters + rest_start - fact),
                year=year, month=month, kind=info["kind"], points=len(points), skipped=skipped)


def summary_text(res):
    lines = [f"{'Дата':<11}{'Выезд':<7}{'Заезд':<7}{'Куда':<28}{'Км':>7}"]
    for t in res["trips"]:
        where = t["dst"] if t["src"] == res["office"] else "← офис"
        lines.append(f"{t['date']:%d.%m.%Y} {t['tout']:%H:%M}  {t['tin']:%H:%M}  {where[:26]:<28}"
                     f"{tenths_str(t['km']):>7}")
    lines += ["",
              f"Дней с поездками: {len(res['trips']) // 2}",
              f"Одометр: {tenths_str(res['start_odo'])} → {tenths_str(res['end_odo'])}",
              f"Пробег за месяц: {tenths_str(res['total'])} км",
              f"Выдано {res['liters']} л + остаток {res['rest_start']} л → расход {res['fact']} л, "
              f"остаток {res['rest_end']} л"]
    return "\n".join(lines)


def output_name(info, res, changes=None):
    fio = (changes or {}).get("C11") or info["consts"].get("C11")
    parts = str(fio or "").split()
    if len(parts) >= 3:
        stem = f"ПЛ_{parts[0]}_{parts[1][0]}_{parts[2][0]}_"
    elif parts:
        stem = f"ПЛ_{parts[0]}_"
    else:
        stem = "ПЛ"
    return f"{stem}_{MONTHS_NOM[res['month'] - 1]}_{res['year']}.xlsm"


def parse_field(text, kind):
    text = str(text).strip()
    if kind == "date":
        try:
            return dt.datetime.strptime(text, "%d.%m.%Y").date()
        except ValueError:
            raise UserError(f"Дата должна быть в формате дд.мм.гггг: {text}")
    if kind == "num":
        try:
            return int(text)
        except ValueError:
            return parse_decimal(text, "Число")
    return text


def format_field(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%d.%m.%Y")
    return "" if v is None else str(v)


# ============================================================ запись в .xlsm
class XlsmPatcher:
    """Меняет значения ячеек прямо в XML, остальное в файле не трогает
    (макрос, выпадающие списки, условное форматирование сохраняются)."""

    def __init__(self, path):
        with zipfile.ZipFile(path) as z:
            self.infos = z.infolist()
            self.data = {i.filename: z.read(i.filename) for i in self.infos}
        wb = self.data["xl/workbook.xml"].decode("utf-8")
        rels = self.data["xl/_rels/workbook.xml.rels"].decode("utf-8")
        rid_target = {}
        for rel in re.findall(r"<Relationship [^>]*>", rels):
            rid, tgt = re.search(r'Id="([^"]+)"', rel), re.search(r'Target="([^"]+)"', rel)
            if rid and tgt:
                rid_target[rid.group(1)] = tgt.group(1)
        self.sheets = {}
        for tag in re.findall(r"<sheet [^>]*>", wb):
            name = re.search(r'name="([^"]+)"', tag).group(1).replace("&amp;", "&")
            target = rid_target[re.search(r'r:id="([^"]+)"', tag).group(1)].lstrip("/")
            self.sheets[name] = target if target.startswith("xl/") else "xl/" + target
        self.pending = {}

    def set(self, sheet, ref, value):
        self.pending.setdefault(sheet, {})[ref] = value

    @staticmethod
    def _cell_xml(ref, style, value):
        s = f' s="{style}"' if style else ""
        if value is None or value == "":
            return f'<c r="{ref}"{s}/>'
        if isinstance(value, dt.datetime):
            value = value.date()
        if isinstance(value, dt.date):
            num = str((value - dt.date(1899, 12, 30)).days)
        elif isinstance(value, dt.time):
            num = repr((value.hour * 3600 + value.minute * 60) / 86400)
        elif isinstance(value, Decimal):
            num = format(value.normalize(), "f")
        elif isinstance(value, (int, float)):
            num = repr(value)
        else:
            return (f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">'
                    f'{escape(str(value))}</t></is></c>')
        return f'<c r="{ref}"{s}><v>{num}</v></c>'

    @staticmethod
    def _col_idx(ref):
        n = 0
        for ch in re.match(r"[A-Z]+", ref).group(0):
            n = n * 26 + ord(ch) - 64
        return n

    def _patch(self, xml, ref, value):
        m = re.compile(r'<c r="%s"(?P<a>[^>]*?)(?:/>|>(?P<b>.*?)</c>)' % ref, re.S).search(xml)
        if m:
            if m.group("b") and "<f" in m.group("b"):
                raise UserError(f"Ячейка {ref} содержит формулу — шаблон отличается от ожидаемого.")
            st = re.search(r'\ss="(\d+)"', m.group("a"))
            return xml[:m.start()] + self._cell_xml(ref, st.group(1) if st else None, value) + xml[m.end():]
        row = re.search(r"\d+", ref).group(0)
        rm = re.search(r'<row r="%s"[^>]*?(?:/>|>(.*?)</row>)' % row, xml, re.S)
        if not rm:
            raise UserError(f"В шаблоне нет строки {row}.")
        new = self._cell_xml(ref, None, value)
        if rm.group(0).endswith("/>"):
            return xml[:rm.start()] + rm.group(0)[:-2] + ">" + new + "</row>" + xml[rm.end():]
        for cm in re.finditer(r'<c r="([A-Z]+\d+)"', rm.group(1)):
            if self._col_idx(cm.group(1)) > self._col_idx(ref):
                pos = rm.start(1) + cm.start()
                return xml[:pos] + new + xml[pos:]
        return xml[:rm.end(1)] + new + xml[rm.end(1):]

    def save(self, out):
        for sheet, cells in self.pending.items():
            name = self.sheets[sheet]
            xml = self.data[name].decode("utf-8")
            for ref, value in cells.items():
                xml = self._patch(xml, ref, value)
            self.data[name] = xml.encode("utf-8")
        wb = self.data["xl/workbook.xml"].decode("utf-8")
        if "fullCalcOnLoad" not in wb:
            wb = re.sub(r"<calcPr", '<calcPr fullCalcOnLoad="1"', wb, count=1)
        self.data["xl/workbook.xml"] = wb.encode("utf-8")
        tmp = out + ".tmp"
        with zipfile.ZipFile(tmp, "w") as z:
            for i in self.infos:
                z.writestr(i, self.data[i.filename], compress_type=i.compress_type)
        os.replace(tmp, out)


def write_result(info, res, out_path, changes=None):
    p = XlsmPatcher(info["path"])
    year, month = res["year"], res["month"]
    p.set(FRONT, "C6", f"{MONTHS_GEN[month - 1]} {year} г.")
    p.set(FRONT, "A6", 1)
    p.set(FRONT, "B6", calendar.monthrange(year, month)[1])
    p.set(FRONT, "H19", Decimal(tenths_str(res["start_odo"])))
    p.set(FRONT, "H32", res["liters"])
    p.set(FRONT, "H33", res["rest_start"])
    p.set(FRONT, "H35", res["norm"])
    for cell, value in (changes or {}).items():
        p.set(FRONT, cell, value)
    for sheet, row, c in SLOTS:
        for key in c:
            p.set(sheet, f"{c[key]}{row}", None)
    for (sheet, row, c), t in zip(SLOTS, res["trips"]):
        p.set(sheet, f"{c['date']}{row}", t["date"])
        p.set(sheet, f"{c['tout']}{row}", t["tout"])
        p.set(sheet, f"{c['tin']}{row}", t["tin"])
        p.set(sheet, f"{c['src']}{row}", t["src"])
        p.set(sheet, f"{c['dst']}{row}", t["dst"])
        p.set(sheet, f"{c['odo']}{row}", Decimal(tenths_str(t["odo"])))
        p.set(sheet, f"{c['km']}{row}", Decimal(tenths_str(t["km"])))
        p.set(sheet, f"{c['kind']}{row}", res["kind"])
    p.set(FRONT, "H39", Decimal(tenths_str(res["end_odo"])))
    p.set(FRONT, "B39", res["trips"][-1]["date"])
    p.set(FRONT, "D39", res["trips"][-1]["tin"])
    try:
        p.save(out_path)
    except PermissionError:
        raise UserError("Не удалось сохранить файл — возможно, он открыт в Excel. Закройте его и повторите.")
