"""Excel shablonni to'ldirish.

- Ko'k shrift (0000FF) = input. Bot faqat shularni yozadi.
- Qora = formulalar. Dinamik kengayishda formulalar avtomatik tuzatiladi.
- 2-bo'lim B18:I27 (10 qator). 4-Manbalar B49:G54 (6 qator, to'liq tozalanadi).
- 6-bo'lim: yangi xodimlar 70-76 (7 qator), ketganlar 78-82 (5 qator).
  Sig'masa qator INSERT qilinadi + barcha formulalar tuzatiladi.
"""
from __future__ import annotations

import re
from copy import copy
from datetime import date
from pathlib import Path

import openpyxl

SHEET = "Oylik hisobot"

SOURCE_ROWS = ["hh.uz", "OLX", "Telegram kanallari", "Tavsiya (referral)", "LinkedIn", "O'zi murojaat qilgan"]
VACANCY_STAGE_DEFAULT = "E'lon"

HIRES_TOP, HIRES_BASE_END = 70, 76     # B70:B76 (76 - zaxira qator)
LEAVERS_TOP, LEAVERS_BASE_END = 78, 82  # B78:B82 (81-82 zaxira qatorlar)

_REF_RE = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def _set(ws, coord: str, value):
    ws[coord].value = value


def _clear_block(ws, min_row: int, max_row: int, cols: str):
    for r in range(min_row, max_row + 1):
        for col in cols:
            ws[f"{col}{r}"].value = None


def _shift_formula(formula: str, at_row: int, n: int) -> str:
    """Qo'shtirnoq ichidagilarga tegmasdan, at_row dan pastdagi ref'larni +n."""
    parts = formula.split('"')

    def rep(m):
        row = int(m.group(4))
        if row >= at_row:
            return m.group(1) + m.group(2) + m.group(3) + str(row + n)
        return m.group(0)

    for i in range(0, len(parts), 2):
        parts[i] = _REF_RE.sub(rep, parts[i])
    return '"'.join(parts)


def _repair_formulas(ws, at_row: int, n: int):
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str) and v.startswith("="):
                c.value = _shift_formula(v, at_row, n)


def _shift_merges(ws, at_row: int, n: int, saved: list):
    """Merge'larni hujayralarga tegmasdan surish (unmerge_cells/merge_cells
    qiymatlarni o'chirib yuboradi - ulardan foydalanilmaydi)."""
    from openpyxl.worksheet.cell_range import CellRange
    kept = []
    for (c1, r1, c2, r2) in saved:
        try:
            if r1 >= at_row:
                kept.append(CellRange(min_col=c1, min_row=r1 + n,
                                      max_col=c2, max_row=r2 + n))
            else:
                kept.append(CellRange(min_col=c1, min_row=r1,
                                      max_col=c2, max_row=r2 + n))
        except Exception:
            pass
    try:
        ws.merged_cells.ranges = {
            r for r in ws.merged_cells.ranges
            if not (r.min_row >= at_row or (r.min_row < at_row <= r.max_row))
        } | set(kept)
    except Exception:
        pass


def _copy_row_style(ws, src_row: int, dst_row: int, cols: str):
    for col in cols:
        s, d = ws[f"{col}{src_row}"], ws[f"{col}{dst_row}"]
        d.font = copy(s.font)
        d.fill = copy(s.fill)
        d.border = copy(s.border)
        d.alignment = copy(s.alignment)
        d.number_format = s.number_format


def _ensure_rows(ws, top: int, base_end: int, need: int, merge_template_row: int,
                count_cell: str, count_col: str = "B") -> int:
    """Blokda need qator bo'lishini ta'minlash. Yetmasa INSERT + formula repair.
    count_cell (mas M70) ni =COUNTA(B<top>:B<last>) ga yangilaydi.
    Oxirgi data qator raqamini qaytaradi."""
    have = base_end - top + 1
    last = base_end
    if need > have:
        extra = need - have
        at = base_end + 1
        saved = [(r.min_col, r.min_row, r.max_col, r.max_row)
                 for r in ws.merged_cells.ranges
                 if r.min_row >= at or (r.min_row < at <= r.max_row)]
        ws.insert_rows(at, extra)
        _shift_merges(ws, at, extra, saved)
        _repair_formulas(ws, at, extra)
        last = base_end + extra
    # zaxira/yangi qatorlarga merge + style
    for r in range(merge_template_row + 1, last + 1):
        for a, b in (("C", "D"), ("F", "H")):
            rng = f"{a}{r}:{b}{r}"
            try:
                if rng not in ws.merged_cells:
                    ws.merge_cells(rng)
            except Exception:
                pass
        _copy_row_style(ws, merge_template_row, r, "BCEF")
        ws.row_dimensions[r].height = ws.row_dimensions[merge_template_row].height
    ws[count_cell].value = f"=COUNTA({count_col}{top}:{count_col}{last})"
    return last


def fill_vacancies(ws, vacancies: list[dict], report_end: date):
    _clear_block(ws, 18, 27, "BCDEFGHI")
    for i, v in enumerate(vacancies[:10]):
        r = 18 + i
        _set(ws, f"B{r}", v.get("lavozim", ""))
        _set(ws, f"C{r}", v.get("bulim", ""))
        _set(ws, f"D{r}", v.get("buyurtmachi", ""))
        if v.get("ochilgan_sana"):
            _set(ws, f"E{r}", v["ochilgan_sana"])
            ws[f"E{r}"].number_format = "DD.MM.YYYY"
        _set(ws, f"F{r}", v.get("bosqich", VACANCY_STAGE_DEFAULT))
        _set(ws, f"H{r}", v.get("ustuvorlik", "O'rta"))
        _set(ws, f"I{r}", v.get("izoh", ""))
    for r in range(18, 28):
        if isinstance(ws[f"G{r}"].value, (int, float)) or ws[f"G{r}"].value is None:
            ws[f"G{r}"].value = f'=IF(OR(E{r}="",F{r}="Yopilgan"),"",$C$12-E{r})'


def fill_kpi(ws, kpi: dict):
    for coord, val in kpi.items():
        if val is None:
            continue
        _set(ws, coord, val)


def fill_sources(ws, sources: list[dict]):
    """6 qator to'liq tozalanib qayta yoziladi (eski namuna qoldiqlari o'chadi)."""
    _clear_block(ws, 49, 54, "BCDEFG")
    for i, s in enumerate(sources[:6]):
        r = 49 + i
        _set(ws, f"B{r}", s.get("manba", ""))
        for col, key in (("C", "kelgan"), ("D", "suhbat"), ("E", "taklif"),
                         ("F", "qabul"), ("G", "xarajat")):
            if s.get(key) is not None:
                _set(ws, f"{col}{r}", s[key])


def fill_hires(ws, hires: list[dict]) -> int:
    """Qaytadi: insert qilingan qatorlar soni (pastdagi bloklar shuncha suriladi)."""
    _clear_block(ws, HIRES_TOP, LEAVERS_TOP - 2, "BCEF")
    need = max(len(hires), 1)
    inserted = max(0, need - (HIRES_BASE_END - HIRES_TOP + 1))
    last = _ensure_rows(ws, HIRES_TOP, HIRES_BASE_END, need,
                        HIRES_TOP, "M70")
    for i, h in enumerate(hires):
        r = HIRES_TOP + i
        _set(ws, f"B{r}", h.get("fish", ""))
        _set(ws, f"C{r}", h.get("lavozim", ""))
        if h.get("sana"):
            _set(ws, f"E{r}", h["sana"])
            ws[f"E{r}"].number_format = "DD.MM.YYYY"
        _set(ws, f"F{r}", h.get("manba", ""))
    return inserted


def fill_leavers(ws, leavers: list[dict], top: int = LEAVERS_TOP) -> int:
    """Qaytadi: insert qilingan qatorlar soni."""
    base_end = top + (LEAVERS_BASE_END - LEAVERS_TOP)
    need = max(len(leavers), 1)
    inserted = max(0, need - (base_end - top + 1))
    _clear_block(ws, top, base_end, "BCEF")
    last = _ensure_rows(ws, top, base_end, need, top, "M71")
    for i, lv in enumerate(leavers):
        r = top + i
        _set(ws, f"B{r}", lv.get("fish", ""))
        _set(ws, f"C{r}", lv.get("lavozim", ""))
        if lv.get("sana"):
            _set(ws, f"E{r}", lv["sana"])
            ws[f"E{r}"].number_format = "DD.MM.YYYY"
        _set(ws, f"F{r}", lv.get("sabab", ""))
    return inserted


def fill_problems(ws, problems: list[dict], top: int = 96):
    """8-bo'lim: B=muammo, C=ta'sir, F=chora, I=mas'ul, L=muddat."""
    _clear_block(ws, top, top + 2, "BCFIL")
    for i, p in enumerate(problems[:3]):
        r = top + i
        if p.get("muammo"):
            _set(ws, f"B{r}", p["muammo"])
        if p.get("tasir"):
            _set(ws, f"C{r}", p["tasir"])
        if p.get("chora"):
            _set(ws, f"F{r}", p["chora"])
        if p.get("masul"):
            _set(ws, f"I{r}", p["masul"])
        if p.get("muddat"):
            _set(ws, f"L{r}", p["muddat"])


def fill_costs(ws, costs: dict):
    for coord, val in costs.items():
        if val is not None:
            _set(ws, coord, val)


def build_report(template_path: str | Path, output_path: str | Path, data: dict) -> str:
    wb = openpyxl.load_workbook(template_path)
    ws = wb[SHEET] if SHEET in wb.sheetnames else wb.active

    if data.get("davr_boshi"):
        _set(ws, "C11", data["davr_boshi"])
        ws["C11"].number_format = "DD.MM.YYYY"
    if data.get("davr_oxiri"):
        _set(ws, "C12", data["davr_oxiri"])
        ws["C12"].number_format = "DD.MM.YYYY"
    if data.get("tayyorladi"):
        _set(ws, "C13", data["tayyorladi"])
    if data.get("hisobot_sanasi"):
        _set(ws, "C14", data["hisobot_sanasi"])
        ws["C14"].number_format = "DD.MM.YYYY"

    if "vacancies" in data:
        fill_vacancies(ws, data["vacancies"], data.get("davr_oxiri"))
    if "kpi" in data:
        fill_kpi(ws, data["kpi"])
    if "sources" in data:
        fill_sources(ws, data["sources"])
    n1 = 0
    if "hires" in data:
        n1 = fill_hires(ws, data["hires"]) or 0
    n2 = 0
    if "leavers" in data:
        n2 = fill_leavers(ws, data["leavers"], top=LEAVERS_TOP + n1) or 0
    shift = n1 + n2  # pastdagi statik bo'limlar surilishi
    if "costs" in data:
        fill_costs(ws, data["costs"])
    if "problems" in data:
        fill_problems(ws, data["problems"], top=96 + shift)

    def _sh(coord: str) -> str:
        m = re.match(r"([A-Z]+)(\d+)", coord)
        return f"{m.group(1)}{int(m.group(2)) + shift}" if m else coord

    for coord, key in (("B96", "muammo1"), ("C96", "tasir1"), ("F96", "chora1"),
                       ("I96", "masul1"), ("L96", "muddat1")):
        if data.get(key) is not None:
            _set(ws, _sh(coord), data[key])
    if data.get("reja_ustuvor") is not None:
        _set(ws, _sh("C103"), data["reja_ustuvor"])
    if data.get("reja_kanal") is not None:
        _set(ws, _sh("C104"), data["reja_kanal"])
    if data.get("reja_qaror") is not None:
        _set(ws, _sh("C105"), data["reja_qaror"])
    if data.get("xulosa_qoshimcha") is not None:
        _set(ws, _sh("B112"), data["xulosa_qoshimcha"])
    if data.get("turnover_note"):
        _set(ws, "J73", data["turnover_note"])

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return str(output_path)
