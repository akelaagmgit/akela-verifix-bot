"""Oylik hisobot uchun Verifix ma'lumotlarini yig'ish.

Manbalar (ustuvorlik bo'yicha):
1. Ichki sessiya API: vacancy_list:table (ochilish sanalari), dashboard
   (voronka/time), candidate_management (manbalar/voronka).
2. Public API: hiring/dismissal (qabul/ketganlar - ground truth).
3. Ichki API bloklansa (SessionBlocked) - public-only fallback.

Excel KPI xaritas:
- C32 oy boshida ochiq: opened < oy_boshi AND (status O OR ...)
- C33 yangi ochilgan: opened oy ichida
- C34 yopilgan: shu oy qabul (proxy, shablon namunasida ham C34=C40)
- C36 kelgan: oyda yaratilgan kandidatlar (barcha bosqich)
- C37 skriningdan o'tgan: kelganlardan bosqichi todo dan keyingi
- C38 suhbatga chiqqan: Suhbat+ bosqichiga yetganlar
- C39 taklif berilgan: Taklif+ bosqichidagilar
- C40 qabul: hiring ro'yxati (ground truth)
- C41/C42 time-to-fill/hire: dashboard o'rtachalari (bo'lmasa qo'lda)
"""
from __future__ import annotations

import calendar
from collections import Counter
from datetime import date

from verifix_client import (
    CHANNEL_NAMES, SOURCE_ROWS, STAGE_ORDER, SessionBlocked, VerifixClient,
    clean_verifix_label, parse_verifix_date,
)

INTERVIEW_PLUS = {"Собеседование", "Стажировка (5 ден)", "Предложение о работе",
                  "Предложение принято"}
OFFER_PLUS = {"Предложение о работе", "Предложение принято"}
TODO_STAGES = {"К выполнению"}


def month_range(y: int, m: int) -> tuple[date, date]:
    return date(y, m, 1), date(y, m, calendar.monthrange(y, m)[1])


def _norm_tokens(name: str) -> set:
    import re
    s = (name or "").lower().replace("o'", "o").replace("g'", "g").replace("`", "")
    toks = set(re.findall(r"[a-z'\-]+", s)) | set(re.findall(r"[\w']+", s)) - {""}
    # o'zbek transliteratsiyasi: x ~ h (Xoshim/Hoshim, Xusanova/Husanova)
    toks |= {t.replace("x", "h") for t in toks if "x" in t}
    return toks


def _same_person(a: str, b: str) -> bool:
    ta, tb = _norm_tokens(a), _norm_tokens(b)
    if not ta or not tb:
        return False
    if ta == tb:
        return True
    small, big = (ta, tb) if len(ta) <= len(tb) else (tb, ta)
    if len(small) >= 2 and small <= big:
        return True
    # initsial: "R Sug'diyona" vs "Rajabova Sug'diyona"
    if len(small) >= 2:
        initials = {t for t in small if len(t) == 1}
        rest = small - initials
        if initials and rest and rest <= big:
            for ini in initials:
                if any(w.startswith(ini) for w in big - rest):
                    return True
    return False


# Qabul ta'rifi: "started" = ish boshlaganlar (jurnallar + kartochkalar).
# Birlashtirish YO'Q - faqat aynan bir xil yozuvlar (ism+sanasi bir xil).
HIRES_MODE = "started"


def _dedupe_people(records: list[dict], key: str = "fish") -> list[dict]:
    """Faqat aynan bir xil yozuvlar birlashadi (bir xil ism + bir xil sana -
    masalan jurnal va kartochkadagi bir xil hodisa). Qolgan hamma alohida."""
    out: list[dict] = []
    for r in records:
        hit = None
        for o in out:
            if o.get(key, "") and _norm_tokens(o.get(key, "")) == _norm_tokens(r.get(key, "")) \
                    and o.get("sana") and o.get("sana") == r.get("sana"):
                hit = o
                break
        if hit is None:
            out.append(dict(r))
            continue
        if len(r.get(key, "")) > len(hit.get(key, "")):
            hit[key] = r[key]
        for f in ("lavozim", "manba", "sabab"):
            cur, new = (hit.get(f) or "").strip(), (r.get(f) or "").strip()
            if new and new.lower() not in cur.lower():
                hit[f] = (cur + " / " + new).strip(" / ")
    return out


def _fmtd(d: date) -> str:
    return d.strftime("%d.%m.%Y")


def _accepted_hires(client: VerifixClient, y: int, m: int) -> list[dict]:
    """Taklifni qabul qilganlar: oyda yaratilgan + offer_response='A'.
    Sana/lavozim xodim kartochkasidan (hiring_date) olinadi."""
    from verifix_client import parse_position_name as _ppn
    acc = [x for x in client.get_all_candidates()
            if x.get("_created") and client.in_month(x["_created"], y, m)
            and (x.get("offer_response") or "").upper() == "A"]
    emps = client.get_employees()
    recs: list[dict] = []
    for x in acc:
        name = x.get("candidate_full_name") or x.get("candidate_name", "")
        match = None
        for e in emps:
            if _same_person(e.get("employee_name", ""), name):
                match = e
                break
        if match is not None:
            job_c, div_c = _ppn(match.get("position_name"))
            recs.append({"fish": match.get("employee_name", name),
                         "lavozim": job_c or clean_verifix_label(match.get("job_name", "")),
                         "sana": parse_verifix_date(match.get("hiring_date")) or x["_created"],
                         "manba": div_c or clean_verifix_label(match.get("division_name", "")),
                         "_eid": str(match.get("employee_id", ""))})
        else:
            recs.append({"fish": name, "lavozim": clean_verifix_label(x.get("job_name", "")),
                         "sana": x["_created"], "manba": clean_verifix_label(x.get("division_name", "")),
                         "_eid": ""})
    return sorted(_dedupe_people(recs), key=lambda x: (x["sana"] is None, x["sana"]))


def _started_hires(client: VerifixClient, y: int, m: int) -> list[dict]:
    """Ish boshlaganlar: hiring jurnallari + kartochkalar (hiring_date)."""
    from verifix_client import parse_position_name as _ppn
    hire_recs: list[dict] = []
    for h in client.hirings_in_month(y, m):
        hire_recs.append({"fish": h.get("employee_name", ""), "lavozim": h.get("_job_name", ""),
                           "sana": h.get("_parsed_date"), "manba": h.get("_division_name", ""),
                           "_eid": str(h.get("employee_id", ""))})
    for e in client.get_employees():
        d = parse_verifix_date(e.get("hiring_date"))
        if client.in_month(d, y, m):
            job_c, div_c = _ppn(e.get("position_name"))
            hire_recs.append({"fish": e.get("employee_name", ""),
                               "lavozim": job_c or clean_verifix_label(e.get("job_name", "")),
                               "sana": d, "manba": div_c or clean_verifix_label(e.get("division_name", "")),
                               "_eid": str(e.get("employee_id", ""))})
    return sorted(_dedupe_people(hire_recs), key=lambda x: (x["sana"] is None, x["sana"]))


def find_transfers(client: VerifixClient) -> tuple[set, set, list]:
    """Transferlar: ishdan bo'shab (jurnal/kartochka), 45 kun ichida qayta
    ishga olinganlar (jurnal/kartochka) - employee_id bo'yicha."""
    out_le, out_hi, notes = set(), set(), []
    try:
        hires_all, dis_all = [], []
        for h in client.get_hirings():
            d = parse_verifix_date(h.get("hiring_date"))
            if d and h.get("employee_id"):
                hires_all.append((str(h["employee_id"]), h.get("employee_name", ""), d))
        for e in client.get_employees():
            d = parse_verifix_date(e.get("hiring_date"))
            if d and e.get("employee_id"):
                hires_all.append((str(e["employee_id"]), e.get("employee_name", ""), d))
        for d in client.get_dismissals():
            dd = parse_verifix_date(d.get("dismissal_date"))
            if dd and d.get("employee_id"):
                dis_all.append((str(d["employee_id"]), d.get("employee_name", ""), dd))
        for e in client.get_employees():
            dd = parse_verifix_date(e.get("dismissal_date"))
            if dd and e.get("employee_id"):
                dis_all.append((str(e["employee_id"]), e.get("employee_name", ""), dd))
    except Exception:
        return out_le, out_hi, notes
    seen = set()
    for eid, dname, dd in dis_all:
        for eid2, hname, hd in hires_all:
            if eid and eid == eid2 and 0 < (hd - dd).days <= 45 and (eid, dd, hd) not in seen:
                seen.add((eid, dd, hd))
                out_le.add((eid, dd.year, dd.month))
                out_hi.add((eid, hd.year, hd.month))
                notes.append((dname or hname, dd, hd))
    return out_le, out_hi, notes


def get_month_people(client: VerifixClient, y: int, m: int) -> tuple[list[dict], list[dict]]:
    """Qabul HIRES_MODE bo'yicha; ketganlar = jurnallar + kartochkalar."""
    from verifix_client import parse_position_name as _ppn
    hires = _accepted_hires(client, y, m) if HIRES_MODE == "accepted" else _started_hires(client, y, m)
    leave_recs: list[dict] = []
    for dd in client.dismissals_in_month(y, m):
        leave_recs.append({"fish": dd.get("employee_name", ""), "lavozim": "",
                            "sana": dd.get("_parsed_date"), "sabab": dd.get("_reason", ""),
                            "_eid": str(dd.get("employee_id", ""))})
    for e in client.get_employees():
        d = parse_verifix_date(e.get("dismissal_date"))
        if client.in_month(d, y, m):
            job_c, _ = _ppn(e.get("position_name"))
            leave_recs.append({"fish": e.get("employee_name", ""),
                                "lavozim": job_c or clean_verifix_label(e.get("job_name", "")),
                                "sana": d, "sabab": "",
                                "_eid": str(e.get("employee_id", ""))})
    leavers = sorted(_dedupe_people(leave_recs), key=lambda x: (x["sana"] is None, x["sana"]))
    # transferlar chiqarib tashlanadi (xodim ishdan bo'shamay, ko'chirilgan)
    tr_le, tr_hi, tr_notes = find_transfers(client)
    hires = [h for h in hires if (h.get("_eid"), y, m) not in tr_hi]
    leavers = [x for x in leavers if (x.get("_eid"), y, m) not in tr_le]
    tnote = "; ".join(f"{n} ({d.strftime('%d.%m')} → {h.strftime('%d.%m')}, transfer)"
                        for n, d, h in tr_notes
                        if (d.year, d.month) == (y, m) or (h.year, h.month) == (y, m))
    return hires, leavers, tnote


def _parse_money(s) -> int | None:
    import re
    if not s:
        return None
    d = re.sub(r"\D", "", str(s))
    try:
        return int(d) if d else None
    except ValueError:
        return None


def collect_month_data(client: VerifixClient, y: int, m: int, tayyorladi: str = "") -> tuple[dict, dict]:
    start, end = month_range(y, m)
    auto: list[str] = []   # avtomatik to'lgan bo'limlar
    manual: list[str] = []  # qo'lda to'ldiriladiganlar
    data: dict = {
        "davr_boshi": start, "davr_oxiri": end,
        "tayyorladi": tayyorladi or "Recruiting menejer",
        "hisobot_sanasi": date.today(),
    }

    # ---- 1) Vakansiyalar jadvali (ichki API) ----
    vac_open: list[dict] = []
    vac_closed: list[dict] = []
    internal = True
    try:
        for v in client.get_vacancy_table("O"):
            vac_open.append({
                "id": v.get("vacancy_id", ""), "lavozim": clean_verifix_label(v.get("name", "")),
                "bulim": clean_verifix_label(v.get("division_name", "")),
                "buyurtmachi": "", "ochilgan_sana": v.get("opened"),
                "bosqich": "E'lon", "ustuvorlik": "O'rta",
                "izoh": clean_verifix_label(v.get("job_name", "")),
                "_quantity": v.get("quantity", "1"),
            })
        vac_closed = client.get_vacancy_table("C")
    except SessionBlocked:
        internal = False
        for v in client.get_vacancies_simple():
            vac_open.append({
                "id": "", "lavozim": clean_verifix_label(v.get("vacancy_name", "")),
                "bulim": clean_verifix_label(v.get("division_name", "")),
                "buyurtmachi": "", "ochilgan_sana": None,
                "bosqich": "E'lon", "ustuvorlik": "O'rta",
                "izoh": clean_verifix_label(v.get("job_name", "")),
                "_quantity": "1",
            })

    # Vakansiya bosqichini kandidatlar bo'yicha aniqlash (ichki API bo'lsa)
    cand_all: list[dict] = []
    if internal:
        try:
            cand_all = client.get_all_candidates()
        except SessionBlocked:
            internal = False
    if internal and cand_all:
        best_stage: dict[str, int] = {}
        order = {name: i for i, name in enumerate(STAGE_ORDER)}
        for cd in cand_all:
            vid = str(cd.get("vacancy_id", ""))
            lvl = order.get(cd.get("_stage", ""), -1)
            if lvl > best_stage.get(vid, -2):
                best_stage[vid] = lvl
        stage_map = {0: "E'lon", 1: "Skrining", 2: "Suhbat", 3: "5 kunlik sinov",
                       4: "Taklif", 5: "Taklif"}
        for v in vac_open:
            if v.get("id") and str(v["id"]) in best_stage:
                v["bosqich"] = stage_map.get(best_stage[str(v["id"])], "E'lon")

    # 2-bo'lim: eng eski 10 ochiq vakansiya (muddati o'tganlar birinchi)
    data["vacancies"] = sorted(
        vac_open, key=lambda v: (v["ochilgan_sana"] is None, v["ochilgan_sana"] or date.max))[:10]
    auto.append("2-Vakansiyalar")

    # ---- 2) Kandidatlar: voronka + manbalar (oy kesimi) ----
    kpi: dict = {}
    sources: list[dict] = []
    by_chan: dict[str, list] = {}
    month_cand: list[dict] = []
    funnel_note = ""
    if internal and cand_all:
        month_cand = [c for c in cand_all if c.get("_created") and client.in_month(c["_created"], y, m)]
        kelgan = len(month_cand)
        skrining = sum(1 for c in month_cand if c.get("_stage") not in TODO_STAGES)
        suhbat = sum(1 for c in month_cand if c.get("_stage") in INTERVIEW_PLUS)
        taklif = sum(1 for c in month_cand if c.get("_stage") in OFFER_PLUS)
        kpi.update({"C36": kelgan, "C37": skrining, "C38": suhbat, "C39": taklif})
        # Manbalar: kanal bo'yicha
        for c in month_cand:
            by_chan.setdefault(str(c.get("channel_id")), []).append(c)
        for cid, lst in sorted(by_chan.items(), key=lambda kv: -len(kv[1])):
            name = CHANNEL_NAMES.get(cid, "Boshqa (kanalsiz)" if cid in ("None", "") else f"Boshqa ({cid})")
            sources.append({
                "manba": name, "kelgan": len(lst),
                "suhbat": sum(1 for c in lst if c.get("_stage") in INTERVIEW_PLUS),
                "taklif": sum(1 for c in lst if c.get("_stage") in OFFER_PLUS),
                "qabul": 0, "xarajat": 0,
            })
        auto.extend(["3-Voronka (C36-C39)", "4-Manbalar"])
        funnel_note = f"oyda kelgan {kelgan} (Telegram va boshqalar)"
    else:
        manual.extend(["3-Voronka (C36-C39)", "4-Manbalar"])

    # ---- 3) Vakansiya hisoblari C32/C33 ----
    if internal and (vac_open or vac_closed):
        month_start = start
        c32 = sum(1 for v in vac_open if v["ochilgan_sana"] and v["ochilgan_sana"] < month_start)
        c32 += sum(1 for v in vac_closed if v.get("opened") and v["opened"] < month_start)
        # ochilish sanasi yo'qlar - joriy ochiqlar sifatida qo'shilmaydi (aniq emas)
        c33 = sum(1 for v in vac_open if v["ochilgan_sana"] and month_start <= v["ochilgan_sana"] <= end)
        c33 += sum(1 for v in vac_closed if v.get("opened") and month_start <= v["opened"] <= end)
        kpi["C32"] = c32
        kpi["C33"] = c33
        auto.append("3-C32/C33 (ochiq/yangi)")
    else:
        kpi["C32"] = len(vac_open)
        manual.append("3-C33 (yangi ochilgan)")

    # ---- 4) Dashboard: time-to-fill/hire, konversiya ----
    dash_hint = ""
    if internal:
        try:
            d = client.get_dashboard(_fmtd(start), _fmtd(end))
            conv = d.get("rec_funnel_conversion") or ""
            over = d.get("rec_overdue_vacancies") or ""
            dash_hint = f"konversiya {conv}%, muddati o'tgan {over}"
            auto.append("dashboard")
        except SessionBlocked:
            internal = False
            manual.append("3-C41/C42 (time)")
    else:
        manual.append("3-C41/C42 (time)")

    # ---- 5) Qabul / ketganlar (ground truth) ----
    hires, leavers, turnover_note = get_month_people(client, y, m)
    if turnover_note:
        data["turnover_note"] = "Transfer (ketgan/qabul emas): " + turnover_note
    # manba->qabul + time-to-hire: ishga olinganlarni kandidat yozuvi bilan bog'lash
    name_chan: dict[str, str] = {}
    for x in cand_all:
        nm = x.get("candidate_full_name") or x.get("candidate_name", "")
        if nm and nm not in name_chan:
            name_chan[nm] = str(x.get("channel_id"))
    hire_chan: dict[str, int] = {}
    th_days: list[int] = []
    for h in hires:
        for x in cand_all:
            nm = x.get("candidate_full_name") or x.get("candidate_name", "")
            if _same_person(nm, h["fish"]):
                cid = str(x.get("channel_id"))
                hire_chan[cid] = hire_chan.get(cid, 0) + 1
                if x.get("_created") and h.get("sana"):
                    try:
                        dd = (h["sana"] - x["_created"]).days
                        if 0 <= dd <= 365:
                            th_days.append(dd)
                    except Exception:
                        pass
                break
    # sources dagi qabul ustunini to'ldirish
    _src_by_cid = {}
    for cid, lst in (by_chan or {}).items():
        nm = CHANNEL_NAMES.get(cid, "Boshqa (kanalsiz)" if cid in ("None", "") else f"Boshqa ({cid})")
        _src_by_cid[nm] = cid
    for s in sources:
        cid = None
        for nm, c2 in _src_by_cid.items():
            if nm == s["manba"]:
                cid = c2
                break
        s["qabul"] = hire_chan.get(cid, 0) if cid else 0
    if th_days:
        kpi["C42"] = round(sum(th_days) / len(th_days))
        auto.append("3-C42 (time-to-hire)")
    else:
        manual.append("3-C42 (time-to-hire)")
    # ---- 6) Muammolarni avtomatik aniqlash (8-bo'lim taklifi) ----
    problems: list[dict] = []
    if internal:
        over = sorted(
            [(v, (end - v["ochilgan_sana"]).days) for v in vac_open
             if v["ochilgan_sana"] and (end - v["ochilgan_sana"]).days > 30],
            key=lambda t: -t[1])
        if over:
            v0, d0 = over[0]
            problems.append({
                "muammo": f"{len(over)} ta vakansiya 30+ kun ochiq (eng eskisi: {v0['lavozim']}, {d0} kun)",
                "tasir": "Vakansiya yopilishi cho'zilmoqda",
                "chora": "Ustuvor vakansiyalar bo'yicha alohida yig'ilish",
                "masul": "", "muddat": ""})
        seq = [("Skrining", kpi.get("C36", 0), kpi.get("C37", 0)),
               ("Suhbat", kpi.get("C37", 0), kpi.get("C38", 0)),
               ("Taklif", kpi.get("C38", 0), kpi.get("C39", 0))]
        worst = None
        for name, a, b in seq:
            if a and b is not None:
                rate = b / a
                if worst is None or rate < worst[1]:
                    worst = (name, rate)
        if worst and worst[1] < 0.5:
            problems.append({
                "muammo": f"{worst[0]} bosqichida yo'qotish yuqori ({worst[1]:.0%} o'tmoqda)",
                "tasir": "Voronka toraymoqda",
                "chora": f"{worst[0]} mezonlarini qayta ko'rib chiqish",
                "masul": "", "muddat": ""})
        wto = {}
        try:
            for v in client.get_vacancy_table("O") + client.get_vacancy_table("C"):
                wto[str(v.get("vacancy_id"))] = _parse_money(v.get("wage_to"))
        except SessionBlocked:
            pass
        high_pay = 0
        for x in month_cand:
            exp = _parse_money(x.get("wage_expactation"))
            lim = wto.get(str(x.get("vacancy_id")))
            if exp and lim and exp > lim:
                high_pay += 1
        if high_pay:
            problems.append({
                "muammo": f"{high_pay} ta nomzod maosh kutilmasi vilkadan yuqori",
                "tasir": "Taklifni rad etish ko'payadi",
                "chora": "Maosh vilkasini qayta ko'rib chiqish",
                "masul": "", "muddat": ""})
    data["problems"] = problems[:3]
    if problems:
        auto.append("8-Muammolar (taklif)")
    else:
        manual.append("8-Muammolar")
    qabul = len(hires)
    kpi["C34"] = qabul
    kpi["C40"] = qabul
    data.update({"vacancies": data["vacancies"], "kpi": kpi, "hires": hires,
                 "leavers": leavers})
    if sources:
        data["sources"] = sources
    auto.extend(["3-C34/C40 (qabul)", "6-Yangi/ketganlar"])

    manual.extend(["7-Xarajatlar", "8-Muammolar", "9-Reja",
                   "3-C41/C42 (time)", "3-D ustun (o'tgan oy)", "2-Bosqich/Ustuvorlik/Buyurtmachi"])

    data["_audit"] = {
        "vac_open": [(v.get("id"), v["ochilgan_sana"].isoformat() if v["ochilgan_sana"] else None,
                        v["lavozim"], v["bulim"]) for v in vac_open],
        "vac_closed": [(v.get("vacancy_id"), v["opened"].isoformat() if v.get("opened") else None)
                          for v in vac_closed],
        "month_cand": [(x.get("_stage"), str(x.get("channel_id"))) for x in month_cand],
        "hires": [(h["fish"], h["sana"].isoformat() if h.get("sana") else None) for h in hires],
        "leavers": [(x["fish"], x["sana"].isoformat() if x.get("sana") else None) for x in leavers],
    }
    stats = {"vacancies_open": len(vac_open), "vacancies_closed_total": len(vac_closed),
             "hired_in_month": qabul, "dismissed_in_month": len(leavers),
             "c32": kpi.get("C32"), "c33": kpi.get("C33"),
             "internal": internal, "auto": sorted(set(auto)),
             "manual": sorted(set(manual)), "funnel_note": funnel_note,
             "dash_hint": dash_hint}
    return data, stats
