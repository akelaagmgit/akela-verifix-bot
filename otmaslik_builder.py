"""Suhbatdan o'tmaslik sabablari hisoboti (template_otmaslik.xlsx).

3 varaq: Hisobot (formulalar), Nomzodlar (bot to'ldiradi), Sabablar (12 sabab).
Bot Nomzodlar 6-500 qatorlarni Verifix'dan to'ldiradi:
- oyda yaratilgan + suhbat bosqichiga yetganlar (interview, internship, offer, accepted)
- E (keldi)=Ha, F (natija): offer/accepted=O'tdi, rejected=O'tmadi, jarayonda=bo'sh
- G (sabab): maosh kutilmasi vilkadan yuqori bo'lsa avtomatik, qolgani qo'lda
- Manba: 2421=Telegram, 2422=HH.uz, else=Boshqa (DV ro'yxatiga mos)
- H (guruh) formulasi tegilmaydi.
"""
from __future__ import annotations

from datetime import date
from pathlib import Path

import openpyxl

SHEET = "Nomzodlar"

CHAN_LABEL = {"2421": "Telegram", "2422": "HH.uz"}
INTERVIEW_STAGES = {"Собеседование", "Стажировка (5 ден)", "Предложение о работе",
                    "Предложение принято"}
PASS_STAGES = {"Предложение о работе", "Предложение принято"}
FAIL_STAGES = {"Отклонена"}
WAGE_REASON = "Maosh kutilmasi mos kelmadi"


def _parse_money(s) -> int | None:
    """Pul formati: minglik ajratkichlar olib tashlanadi, diapazonda yuqori
    chegara olinadi ('3-4 000 000' -> 4000000)."""
    import re
    if not s:
        return None
    t = re.sub(r"[.,\s]", "", str(s))
    try:
        nums = [int(x) for x in re.findall(r"\d+", t)]
        return max(nums) if nums else None
    except ValueError:
        return None


TEMPLATE_REASONS = [
    "Lokatsiya mos kelmadi (uzoq)", "Ish grafigi mos kelmadi",
    "Maosh kutilmasi mos kelmadi", "Bilimi yetarli emas",
    "Tajribasi yo'q / kam", "Sohaga mos emas (boshqa yo'nalish)",
    "Til bilmaydi (rus/ingliz)", "Muloqot / ko'rinish / motivatsiya past",
    "Qiymatlar, jamoaga mos emas", "Nomzodning o'zi rad etdi",
    "Hujjat / tekshiruvdan o'tmadi", "Boshqa (izoh bilan)",
]


def _norm_reason(s: str) -> str:
    return (s or "").strip().lower()


_REASON_INDEX = {_norm_reason(x): x for x in TEMPLATE_REASONS}


def _fetch_reject_map(client) -> dict:
    """Rad etilganlar: candidate_id -> (sabab, yaratilgan sana)."""
    cols = ["name", "candidate_id", "reject_reason_name", "vacancy_id", "created_on"]
    out: dict = {}
    try:
        off = 0
        total = None
        from verifix_client import parse_verifix_date
        while True:
            j = client._ipost(
                "/b/vhr/hrec/candidate/candidate_management:reject_candidates",
                {"p": {"column": cols, "filter": ["candidate_kind", "=", "J"],
                         "sort": ["name"], "offset": off, "limit": 100}})
            if total is None:
                total = j.get("count", 0)
            data = j.get("data", [])
            if not data:
                break
            for row in data:
                try:
                    out[str(row[1])] = (row[2] or "", parse_verifix_date(row[4]))
                except Exception:
                    pass
            off += len(data)
            if off >= (total or 0):
                break
    except Exception:
        pass
    return out


def collect_fail_data(client, y: int, m: int) -> tuple[list[dict], dict]:
    """Oyda yaratilgan + intervyuga yetgan kandidatlar.
    Sabab ustuvorligi: Verifix'dagi rad sababi > maosh mos kelmasligi > bo'sh."""
    cands = client.get_all_candidates()
    rej_map = _fetch_reject_map(client)
    vac_wage = {}
    try:
        for v in client.get_vacancy_table("O") + client.get_vacancy_table("C"):
            vac_wage[str(v.get("vacancy_id"))] = _parse_money(v.get("wage_to"))
    except Exception:
        pass
    rows = []
    n_verifix = 0
    for x in cands:
        if not (x.get("_created") and x["_created"].year == y and x["_created"].month == m):
            continue
        if x.get("_stage") not in INTERVIEW_STAGES | FAIL_STAGES:
            continue
        st = x.get("_stage")
        if st in PASS_STAGES:
            res = "O'tdi"
        elif st in FAIL_STAGES:
            res = "O'tmadi"
        else:
            res = ""
        reason = ""
        if res == "O'tmadi":
            v_reason, _ = rej_map.get(str(x.get("candidate_id", "")), ("", None))
            mapped = _REASON_INDEX.get(_norm_reason(v_reason))
            if mapped:
                reason = mapped
                n_verifix += 1
            else:
                exp = _parse_money(x.get("wage_expactation"))
                lim = vac_wage.get(str(x.get("vacancy_id")))
                if exp and lim and exp > lim:
                    reason = WAGE_REASON
        ctx = [f"Hozir: {st}"]
        if x.get("age"):
            ctx.append(f"{x['age']} yosh")
        exp = _parse_money(x.get("wage_expactation"))
        lim = vac_wage.get(str(x.get("vacancy_id")))
        if exp:
            ctx.append(f"kutilma {exp:,}".replace(",", " "))
        if lim:
            ctx.append(f"vilka ~{lim:,}".replace(",", " "))
        rows.append({
            "sana": x["_created"],
            "fish": x.get("candidate_full_name") or x.get("candidate_name", ""),
            "lavozim": x.get("vacancy_name", "") or x.get("job_name", ""),
            "manba": CHAN_LABEL.get(str(x.get("channel_id")), "Boshqa"),
            "keldi": "Ha", "natija": res, "sabab": reason,
            "izoh": "; ".join(ctx),
        })
    rows.sort(key=lambda r: (r["sana"] is None, r["sana"]))
    info = {"total": len(rows),
            "passed": sum(1 for r in rows if r["natija"] == "O'tdi"),
            "failed": sum(1 for r in rows if r["natija"] == "O'tmadi"),
            "pending": sum(1 for r in rows if not r["natija"]),
            "wage_auto": sum(1 for r in rows if r["sabab"] == WAGE_REASON),
            "verifix_auto": n_verifix}
    return rows, info


def build_fail_report(template_path: str | Path, output_path: str | Path,
                      rows: list[dict]) -> str:
    wb = openpyxl.load_workbook(template_path)
    ws = wb[SHEET] if SHEET in wb.sheetnames else wb.active
    # 5-qator namuna + 6-500 tozalash (H formulalar saqlanadi)
    for r in range(5, 501):
        for col in "BCDEFGI":
            ws[f"{col}{r}"].value = None
    for i, x in enumerate(rows[:495]):
        r = 6 + i
        ws[f"A{r}"].value = x["sana"]
        ws[f"A{r}"].number_format = "DD.MM.YYYY"
        ws[f"B{r}"].value = x["fish"]
        ws[f"C{r}"].value = x["lavozim"]
        ws[f"D{r}"].value = x["manba"]
        ws[f"E{r}"].value = x["keldi"]
        ws[f"F{r}"].value = x["natija"] or None
        ws[f"G{r}"].value = x["sabab"] or None
        ws[f"I{r}"].value = x["izoh"]
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return str(output_path)
