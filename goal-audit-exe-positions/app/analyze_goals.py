# -*- coding: utf-8 -*-
"""Аудит целей сотрудников по единому каскаду метрик.

Вход: xlsx/csv с целями сотрудников (шапка может быть не в первой строке).
Выход: xlsx с двумя листами — «Аудит целей» (по строке на цель) и «Сводка».

Три критерия:
  К1 — формула цели (действие + измеримый результат, без абстракций)
  К2 — соответствие уровню должности по матрице каскада
  К3 — наличие подобной метрики в едином справочнике (мягкая проверка)

Запуск:
  python analyze_goals.py цели.xlsx --out Аудит_целей.xlsx
  python analyze_goals.py цели.xlsx --header-row 2 --out Аудит.xlsx
"""
import argparse
import csv
import json
import os
import re
import sys
from pathlib import Path
from collections import Counter, defaultdict

HERE = Path(__file__).resolve().parent
BUNDLE_ROOT = Path(getattr(sys, "_MEIPASS", HERE))
REF = BUNDLE_ROOT / "references"
if not REF.is_dir():
    REF = HERE.parent / "references"

# ---------------------------------------------------------------- справочники

def load_catalog():
    """ALLOW: метрика -> {колонка: отметка}; TYPE/DIR: метрика -> вид/направление."""
    path = REF / "cascade-metrics.csv"
    allow, typ, dirn = {}, {}, {}
    with open(path, encoding="utf-8") as f:
        rd = csv.reader(f, delimiter=";")
        head = next(rd)
        cols = head[3:]
        for row in rd:
            if not row or not row[0].strip():
                continue
            m = row[0].strip()
            typ[m] = row[1].strip()
            dirn[m] = row[2].strip()
            allow[m] = {c: row[3 + i].strip() for i, c in enumerate(cols) if len(row) > 3 + i and row[3 + i].strip()}
    return allow, typ, dirn


def load_families():
    with open(REF / "metric-families.json", encoding="utf-8") as f:
        d = json.load(f)
    fams = [(re.compile(x["rx"], re.I), x["family"], x["metric"], x["direction"]) for x in d["families"]]
    return fams, d["candidates"]


ALLOW, TYPE, DIRN = load_catalog()
FAMS, CAND = load_families()

# должность в файле -> (колонка справочника, уровень матрицы, есть ли колонка)
POSMAP = [
    (r"директор по производству|ДПП|ДЗ\b|главный инженер", "ДПП/ ДЗ", "ГД-1/ДЗ", True),
    (r"начальник производства", "НП", "НП/СМП", True),
    (r"старший менеджер", "СМП", "НП/СМП", True),
    (r"начальник отдела|руководител", "СМП", "НП/СМП", False),
    (r"начальник смены", "НС", "НС/Инженеры", True),
    (r"эксперт по качеству", "Инж. по качеству", "НС/Инженеры", True),
    (r"эксперт", "Инж. по качеству", "НС/Инженеры", False),
    (r"ведущий инженер по учету|ведущий инженер по учёту", "ВИП", "НС/Инженеры", False),
    (r"ведущий инженер|старший технолог", "ВИП", "НС/Инженеры", True),
    (r"инженер по подготовке", "ИПП", "НС/Инженеры", True),
    (r"инженер по планированию", "ИПП", "НС/Инженеры", False),
    (r"специалист", "ИПП", "НС/Инженеры", False),
    (r"оператор пульта|пультов", "Операторы пульта", "Рабочие", True),
    (r"оператор|аппаратчик|машинист|слесар|прибориc?т", "Операторы поля", "Рабочие", True),
    (r"инженер", "Инж смены", "НС/Инженеры", True),
]

DZ_ONLY = {"LTIF в ЗО", "ИА-1 в ЗО", "% Комплексный показатель (УМД+ПОФ)/EBITDA",
           "Потери по Качеству в ЗО", "ПОФ, млрд в ЗО"}

ACTION = re.compile(
    r"провед|провер|настро|организ|обеспеч|разработ|актуализ|внедр|контрол|ознаком|устран|оформ|"
    r"внесен|реализ|выполн|поддерж|сформир|подготов|фиксац|анализ|стандартиз|обучен|маркиров|загрузк|"
    r"перенос|пересмотр|оптимизац|прораб|аудит|обход|ритуал|встреч|тренировк|расследован|наставнич|"
    r"разъясн|согласов|синхрониз|замер|инвентариз", re.I)
THRESH = re.compile(r"(не менее|не более|не ниже|не выше|≥|≤|>|<|=)\s*\d|\d+\s*%|"
                    r"\d+\s*(шт|т\.|тн|тонн|дн|мг|м3|кг|ед|коэф|событ|млн|ГДж)", re.I)
QUAL_OK = re.compile(r"в полном объеме|в полном объёме|в установленный срок|согласно график|"
                     r"без нарушений|отсутствие замечаний|отсутствие случаев|100\s*%", re.I)
ABSTRACT = [
    "развитие команды", "подготовка к пуску", "управление травматизмом", "ltif=0", "ltif = 0",
    "мини-т", "реализация проекта мини-т", "выполнение целей стратегии устойчивого развития",
    "достижение целей по ключевым показателям эффективности", "достижение кпэ", "качество планирования",
    "соблюдение требований ср1.01", "повышение производственной эффективности", "кадровая устойчивость",
    "обеспечить приживаемость производственной системы", "команда (развитие + компетенции)",
]

LEVEL_NORM = {
    "ГД-1/ДЗ": "бизнес-цели + 10-20% стратегия, процессные метрики 30-40%",
    "НП/СМП": "бизнес-цели 20-30%, процессные метрики 60-70%, индикаторы 10-20%",
    "НС/Инженеры": "процессные метрики в ЗО 40-50%, индикаторы активности 30-40%, бизнес-цели 10-20%",
    "Рабочие": "индикаторы активности 70-90%, процессные метрики 30-40%",
}

# ---------------------------------------------------------------- утилиты

def resolve_pos(pos):
    p = (pos or "").strip()
    for rx, col, lvl, incat in POSMAP:
        if re.search(rx, p, re.I):
            return col, lvl, incat
    return "НС", "НС/Инженеры", False


def family(text):
    for rx, fam, met, d in FAMS:
        if rx.search(text):
            return fam, met, d
    return "Не классифицировано", None, "—"


def allowed(metric, col):
    return ALLOW.get(metric, {}).get(col, "")


def where(metric):
    d = ALLOW.get(metric, {})
    return ", ".join(f"{k}—{v}" for k, v in d.items()) if d else "нет в справочнике"


def rec_metrics(fam, col, limit=3):
    out = []
    for m in CAND.get(fam, []):
        if allowed(m, col):
            out.append(f"«{m}» ({allowed(m, col)})")
        if len(out) >= limit:
            break
    if not out:
        for m, cols in ALLOW.items():
            if col in cols and DIRN.get(m) == DIRN.get(fam, ""):
                out.append(f"«{m}» ({cols[col]})")
            if len(out) >= limit:
                break
    return out


def num(x):
    try:
        return float(str(x).replace(",", ".").replace("%", "").strip())
    except Exception:
        return 0.0

# ---------------------------------------------------------------- чтение файла

ALIASES = {
    "fio": ["фио", "сотрудник", "работник", "ф.и.о"],
    "tab": ["табельный", "таб. №", "таб.номер", "код сотрудника", "id сотрудника"],
    "pos": ["должность", "наименование должности"],
    "path": ["подразделение", "путь до подразделения", "структурное подразделение", "полный путь"],
    "name": ["наименование цели", "название цели", "цель", "формулировка цели"],
    "desc": ["описание цели", "ожидаемый результат", "описание", "результат"],
    "grp": ["группа целей", "тип группы", "категория цели"],
    "w": ["вес цели", "вес, %", "вес"],
    "gtype": ["тип цели"],
    "thr": ["порог частичного", "порог"],
}


def find_header(rows, forced=None):
    if forced:
        return forced - 1
    best, score = 0, -1
    for i, r in enumerate(rows[:20]):
        cells = [str(c).lower() for c in r if c is not None]
        s = sum(1 for c in cells if any(a in c for al in ALIASES.values() for a in al))
        if s > score:
            best, score = i, s
    return best


def map_columns(header):
    idx = {}
    low = [str(h).lower().strip() if h is not None else "" for h in header]
    for key, aliases in ALIASES.items():
        for j, h in enumerate(low):
            if not h:
                continue
            if any(a in h for a in aliases):
                idx.setdefault(key, j)
    return idx


def read_rows(path, header_row=None):
    if path.lower().endswith(".csv"):
        with open(path, encoding="utf-8-sig") as f:
            raw = [r for r in csv.reader(f, delimiter=";")]
    else:
        from openpyxl import load_workbook
        wb = load_workbook(path, data_only=True)
        ws = wb[wb.sheetnames[0]]
        raw = [[c for c in r] for r in ws.iter_rows(values_only=True)]
        wb.close()
    h = find_header(raw, header_row)
    source_width = max(len(r) for r in raw[h:])
    source_headers = tuple(raw[h]) + (None,) * (source_width - len(raw[h]))
    idx = map_columns(raw[h])
    missing = [k for k in ("pos", "name") if k not in idx]
    if missing:
        sys.exit(f"Не найдены обязательные колонки: {missing}. Укажите --header-row или проверьте шапку файла.")
    out = []
    for n, r in enumerate(raw[h + 1:], start=h + 2):
        get = lambda k: ("" if idx.get(k) is None or idx[k] >= len(r) or r[idx[k]] is None else str(r[idx[k]]).strip())
        if not any(get(k) for k in ("fio", "pos", "name", "desc")):
            continue
        out.append(dict(row=n, fio=get("fio"), tab=get("tab"), pos=get("pos"), path=get("path"),
                        name=get("name"), desc=get("desc"), grp=get("grp"), w=get("w"),
                        gtype=get("gtype"), thr=get("thr"),
                        source_headers=source_headers,
                        source_values=tuple(r) + (None,) * (source_width - len(r))))
    return out

# ---------------------------------------------------------------- проверка

def analyze(rows, progress=None):
    cards, index = [], {}
    has_person = any(r["tab"] or r["fio"] for r in rows)
    if has_person:
        # карточка = сотрудник в конкретной должности
        for r in rows:
            key = (r["tab"] or r["fio"], r["pos"])
            if key not in index:
                index[key] = dict(id=len(cards) + 1, pos=r["pos"], fio=r["fio"], tab=r["tab"], items=[])
                cards.append(index[key])
            index[key]["items"].append(r)
            r["card"] = index[key]["id"]
    else:
        # ФИО в файле нет: новая карточка при смене должности или повторе наименования цели
        cur, seen = None, set()
        for r in rows:
            if cur is None or cur["pos"] != r["pos"] or r["name"] in seen:
                cur = dict(id=len(cards) + 1, pos=r["pos"], fio=r["fio"], tab=r["tab"], items=[])
                cards.append(cur)
                seen = set()
            cur["items"].append(r)
            seen.add(r["name"])
            r["card"] = cur["id"]

    cstat = {}
    for c in cards:
        fams = Counter()
        biz = 0
        for x in c["items"]:
            fam, met, _ = family(x["name"] + " " + x["desc"])
            fams[fam] += 1
            if TYPE.get(met or "", "") == "Бизнес-цели":
                biz += 1
        wtot = sum(num(x["w"]) for x in c["items"])
        wbase = sum(num(x["w"]) for x in c["items"] if "базов" in x["grp"].lower())
        cstat[c["id"]] = dict(n=len(c["items"]), fams=fams, biz=biz, wtot=round(wtot),
                              shb=round(100 * wbase / wtot) if wtot else 0,
                              base=sum(1 for x in c["items"] if "базов" in x["grp"].lower()),
                              focus=sum(1 for x in c["items"] if x["grp"] and "базов" not in x["grp"].lower()))

    res = []
    for r in rows:
        col, lvl, incat = resolve_pos(r["pos"])
        text = (r["name"] + " " + r["desc"]).strip()
        fam, metric, direction = family(text)
        typ = TYPE.get(metric or "", "—")
        cs = cstat[r["card"]]

        # --- K1
        k1p = []
        has_a = bool(ACTION.search(text))
        has_t = bool(THRESH.search(text))
        qual_ok = bool(QUAL_OK.search(r["desc"]))
        low = r["name"].strip().lower()
        is_abs = any(low.startswith(a) or low == a for a in ABSTRACT)
        if not has_a:
            k1p.append("нет конкретного действия ни в наименовании, ни в описании")
        if not has_t and not qual_ok:
            k1p.append("нет измеримого порога/целевого значения")
        note = ""
        if not has_t and qual_ok:
            note = "числового порога нет, результат задан перечнем проверяемых условий"
        if has_a and (has_t or qual_ok) and is_abs and len(r["desc"]) < 90:
            k1p.append("формулировка абстрактна: описание сведено к нормативу без указания управляемых действий")
        if len(text) < 6:
            k1p = ["строка не содержит формулировки цели (пустое/некорректное значение)"]
        k1 = "Соответствует" if not k1p else "Требует корректировки"

        # --- K2
        k2p = []
        if typ == "Бизнес-цели":
            if lvl == "Рабочие":
                k2p.append(f"бизнес-цель «{metric}» на уровне рабочего персонала: допустимы только индикаторы активности (70-90%)")
            elif lvl == "НС/Инженеры" and not (col == "НС" and allowed(metric, col) == "Б"):
                k2p.append(f"бизнес-цель «{metric}» на уровне {col}: по матрице требуется {LEVEL_NORM[lvl]}")
        if metric in DZ_ONLY and lvl != "ГД-1/ДЗ":
            k2p.append(f"«{metric}» — итоговая бизнес-цель уровня ГД-1/ДЗ; для {col} требуется процессная метрика или индикатор активности")
        if metric is None:
            k2p.append("вид метрики не определяется — цель не отнесена ни к бизнес-цели, ни к процессной метрике, ни к индикатору активности")
        if not k2p and lvl == "НС/Инженеры" and cs["biz"] > 2 and typ == "Бизнес-цели":
            k2p.append(f"в карточке {cs['biz']} бизнес-целей при норме «минимум бизнес-целей» для уровня {col}")
        if lvl == "Рабочие" and typ == "Процессные метрики" and not k2p:
            pass
        if len(text) < 6:
            k2p = ["строка не содержит формулировки цели"]
        k2 = "Соответствует" if not k2p else "Требует корректировки"

        # --- K3 (мягкая: достаточно найденного эквивалента в справочнике)
        if len(text) < 6:
            k3, k3p = "Требует корректировки", ["строка не содержит формулировки цели"]
        elif metric is None:
            k3 = "Требует корректировки"
            k3p = [f"метрика по направлению «{fam}» отсутствует в едином справочнике — подобрать эквивалент из каскада или снять цель"]
        else:
            mk = allowed(metric, col)
            base = f"метрика сопоставлена со справочником: «{metric}» ({typ}, {direction})"
            k3 = "Соответствует"
            if mk:
                k3p = [base + f"; закреплена за {col} как «{mk}»"]
            elif not incat:
                k3p = [base + f"; должность «{r['pos']}» отдельной колонки в каскаде не имеет — сопоставлена с ближайшей ролью {col}"]
            else:
                k3p = [base + f"; за {col} прямо не закреплена (закреплена: {where(metric)}) — засчитано как наличие в справочнике"]

        # --- рекомендация
        if k1 == k2 == k3 == "Соответствует":
            rec = "Оставить без изменений; при пересмотре сверять порог с дашбордом и фиксировать частоту сверок."
        else:
            parts = []
            rm = rec_metrics(fam, col)
            if rm:
                parts.append("Переформулировать по схеме «направление + действие + измеримый результат» с метриками справочника: " + "; ".join(rm) + ".")
            else:
                parts.append(f"Заменить метрику на закреплённую за должностью в едином справочнике (направление «{fam}») либо снять цель.")
            if metric in DZ_ONLY and lvl != "ГД-1/ДЗ":
                parts.append(f"Итоговый показатель «{metric}» оставить на уровне ДПП/ДЗ, у {col} — управляемые действия и процессные метрики.")
            rec = " ".join(parts)

        res.append(dict(row=r["row"], card=r["card"], fio=r["fio"], pos=r["pos"], col=col, lvl=lvl,
                        unit=r["path"],
                        grp=r["grp"] or "—", w=r["w"], fam=fam, metric=metric or "—", typ=typ,
                        text=(r["name"] + " | " + r["desc"])[:900],
                        k1=k1, k1r="; ".join(k1p) or ("действие и измеримый порог присутствуют" + (f" ({note})" if note else "")),
                        k2=k2, k2r="; ".join(k2p) or f"вид метрики ({typ}) соответствует уровню {col}",
                        k3=k3, k3r="; ".join(k3p), rec=rec,
                        source_headers=r.get("source_headers", ()),
                        source_values=r.get("source_values", ())))
        if progress:
            progress(len(res), len(rows))
    return res, cstat

# ---------------------------------------------------------------- выгрузка

def write_xlsx(res, cstat, out):
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    wb = Workbook()
    ws = wb.active
    ws.title = f"Аудит целей ({len(res)})"
    H = ["№ стр.", "Карточка", "Сотрудник", "Должность", "Подразделение", "Группа целей", "Вес, %",
         "Направление каскада", "Анализируемый текст (наименование + описание)", "Метрика справочника / вид",
         "К1. Формула цели", "К2. Уровень каскада", "К3. Справочник метрик",
         "Обоснование статусов", "Рекомендация по исправлению"]
    hf = PatternFill("solid", fgColor="1F3864")
    okf = PatternFill("solid", fgColor="E2EFDA")
    badf = PatternFill("solid", fgColor="FCE4E4")
    thin = Side(style="thin", color="BFBFBF")
    bd = Border(thin, thin, thin, thin)
    ws.append(H)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = hf
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = bd
    for r in res:
        met = r["metric"] + (f" [{r['typ']}]" if r["metric"] != "—" and r["typ"] != "—" else "")
        just = f"К1: {r['k1r']}\nК2: {r['k2r']}\nК3: {r['k3r']}"
        ws.append([r["row"], r["card"], r["fio"], r["pos"], r["unit"], r["grp"], r["w"], r["fam"],
                   r["text"], met, r["k1"], r["k2"], r["k3"], just, r["rec"]])
        i = ws.max_row
        for col in ("K", "L", "M"):
            cell = ws[f"{col}{i}"]
            cell.fill = okf if cell.value == "Соответствует" else badf
            cell.font = Font(size=9, bold=True, color="375623" if cell.value == "Соответствует" else "9C0006")
        for j in range(1, 16):
            c = ws.cell(row=i, column=j)
            c.border = bd
            c.alignment = Alignment(wrap_text=True, vertical="top",
                                    horizontal="center" if j in (1, 2, 7, 11, 12, 13) else "left")
            if not c.font.bold:
                c.font = Font(size=9)
    for col, w in zip("ABCDEFGHIJKLMNO", [7, 9, 22, 24, 26, 20, 7, 22, 60, 34, 13, 13, 13, 62, 60]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:O{ws.max_row}"
    ws.row_dimensions[1].height = 46

    # Исходные столбцы сохраняются справа от отчёта, а не удаляются.
    # Связь со строкой источника остаётся корректной после отбора должностей.
    source_headers = next((r["source_headers"] for r in res if r.get("source_headers")), ())
    if len(H) + len(source_headers) > 16384:
        raise ValueError("Исходные столбцы и отчёт превышают лимит Excel: 16384 столбца.")
    for offset, header in enumerate(source_headers, start=len(H) + 1):
        cell = ws.cell(row=1, column=offset, value=header)
        if isinstance(header, str):
            cell.data_type = "s"
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.fill = hf
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = bd
        ws.column_dimensions[get_column_letter(offset)].hidden = True
    for row_number, r in enumerate(res, start=2):
        for offset, value in enumerate(r.get("source_values", ()), start=len(H) + 1):
            cell = ws.cell(row=row_number, column=offset, value=value)
            # Текст CSV, начинающийся с "=", не должен становиться формулой.
            if isinstance(value, str):
                cell.data_type = "s"

    s = wb.create_sheet("Сводка")
    tot = len(res)
    cnt = lambda k: sum(1 for r in res if r[k] == "Требует корректировки")
    s.append(["Показатель", "Значение", "", "", ""])
    s.append(["Всего проверено целей (строк)", tot])
    s.append(["Карточек сотрудников", len(cstat)])
    for t, k in [("К1. Требует корректировки (формула цели)", "k1"),
                 ("К2. Требует корректировки (уровень каскада)", "k2"),
                 ("К3. Требует корректировки (метрика не найдена в справочнике)", "k3")]:
        s.append([t, f"{cnt(k)} ({round(100 * cnt(k) / tot)}%)"])
    s.append(["Целей без замечаний по всем трём критериям",
              sum(1 for r in res if all(r[k] == "Соответствует" for k in ("k1", "k2", "k3")))])
    s.append(["Методика", "Критерий распределения целей (База+Фокус) в проверку не входит. К3 засчитывается при наличии подобной метрики в справочнике."])
    s.append([])
    s.append(["Разрез по должностям", "Целей", "К1 — не соотв.", "К2 — не соотв.", "К3 — не соотв."])
    byp = defaultdict(list)
    for r in res:
        byp[r["pos"]].append(r)
    for p, v in sorted(byp.items(), key=lambda x: -len(x[1])):
        s.append([p, len(v)] + [sum(1 for r in v if r[k] == "Требует корректировки") for k in ("k1", "k2", "k3")])
    s.append([])
    s.append(["Группа целей", "Целей", "Доля", "", ""])
    for g, n in Counter(r["grp"] for r in res).most_common():
        s.append([g, n, f"{round(100 * n / tot)}%"])
    s.append([])
    s.append(["Направления с метриками вне справочника", "Целей", "", "", ""])
    for g, n in Counter(r["fam"] for r in res if r["k3"] == "Требует корректировки").most_common(10):
        s.append([g, n])
    s.append([])
    s.append(["Справочная информация по карточкам", "Значение", "", "", ""])
    s.append(["Среднее число целей в карточке", round(sum(v["n"] for v in cstat.values()) / len(cstat), 1)])
    s.append(["Максимум целей в карточке", max(v["n"] for v in cstat.values())])
    s.append(["Карточек с более чем 7 целями", sum(1 for v in cstat.values() if v["n"] > 7)])
    from openpyxl.styles import Font as F2
    for c in s[1]:
        c.font = F2(bold=True, color="FFFFFF")
        c.fill = hf
    s.column_dimensions["A"].width = 62
    s.column_dimensions["B"].width = 18
    for col in "CDE":
        s.column_dimensions[col].width = 18
    wb.save(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("--out", default="Аудит_целей_по_каскаду_метрик.xlsx")
    ap.add_argument("--header-row", type=int, default=None, help="номер строки шапки (1-based)")
    a = ap.parse_args()
    rows = read_rows(a.input, a.header_row)
    if not rows:
        sys.exit("Не удалось прочитать ни одной строки с целями.")
    res, cstat = analyze(rows)
    write_xlsx(res, cstat, a.out)
    tot = len(res)
    print(f"Проверено целей: {tot}, карточек: {len(cstat)}")
    for k, t in (("k1", "К1 формула цели"), ("k2", "К2 уровень каскада"), ("k3", "К3 справочник метрик")):
        n = sum(1 for r in res if r[k] == "Требует корректировки")
        print(f"  {t}: требует корректировки {n} ({round(100 * n / tot)}%)")
    print("  без замечаний по всем трём:",
          sum(1 for r in res if all(r[k] == "Соответствует" for k in ("k1", "k2", "k3"))))
    unc = sum(1 for r in res if r["fam"] == "Не классифицировано")
    if unc:
        print(f"  ВНИМАНИЕ: не классифицировано направлений: {unc} — дополните references/metric-families.json")
    print("Файл сохранён:", a.out)


if __name__ == "__main__":
    main()
