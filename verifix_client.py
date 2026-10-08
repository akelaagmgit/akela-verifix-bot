"""Verifix API mijozi: public API (Basic) + ichki sessiya API.

Ichki API (dashboard, vacancy_list:table, candidate_management) sessiya
cookies talab qiladi. Server bir vaqtda kam sessiyaga ruxsat beradi, shuning
uchun bitta persistent sessiya ishlatiladi (session.json da saqlanadi).
409/401 da bir marta qayta login qilinadi; yana xato bo'lsa caller
public-API fallback rejimga o'tadi.
"""
from __future__ import annotations

import base64
import hashlib
import json
import re
import time
from datetime import date, datetime
from pathlib import Path
from typing import Any

import httpx

DATE_FMTS = ("%d.%m.%Y", "%d.%m.%Y %H:%M:%S", "%Y-%m-%d", "%Y-%m-%dT%H:%M:%S")
RU_PREFIX_RE = re.compile(r"^(Должности|Подразделение)\s*:\s*", re.IGNORECASE)

# Excel manbalar qatorlari <-> Verifix channel_id
CHANNEL_NAMES = {
    "2421": "Telegram kanallari",
    "2422": "hh.uz",
}
SOURCE_ROWS = ["hh.uz", "OLX", "Telegram kanallari", "Tavsiya (referral)", "LinkedIn", "O'zi murojaat qilgan"]

# Bosqichlar tartibi (voronka)
STAGE_ORDER = [
    "К выполнению",
    "Фильтрация",
    "Собеседование",
    "Стажировка (5 ден)",
    "Предложение о работе",
    "Предложение принято",
]
FINAL_STAGES = {"Предложение принято", "Кадровый резерв", "Отклонена"}

BROWSER_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "Chrome/154.0.0.0 Safari/537.36")


def parse_verifix_date(s: str | None) -> date | None:
    if not s:
        return None
    s = str(s).strip()
    for f in DATE_FMTS:
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            continue
    m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", s)
    if m:
        d, mo, y = map(int, m.groups())
        try:
            return date(y, mo, d)
        except ValueError:
            return None
    return None


def clean_verifix_label(s: str | None) -> str:
    if not s:
        return ""
    return RU_PREFIX_RE.sub("", str(s).strip()).strip()


def parse_position_name(pos: str | None) -> tuple[str, str]:
    if not pos:
        return "", ""
    parts = str(pos).split("/")
    if len(parts) >= 2:
        return clean_verifix_label(parts[0]), clean_verifix_label(parts[1])
    return clean_verifix_label(pos), ""


class SessionBlocked(Exception):
    pass


class VerifixClient:
    def __init__(self, login: str, password: str, filial_id: str = "483143",
                 project_code: str = "vhr", base_url: str = "https://app.verifix.com",
                 session_file: str = "session.json"):
        self.username = login
        self.password = password
        self.filial_id = filial_id
        self.project_code = project_code
        self.base_url = base_url.rstrip("/")
        self.session_file = Path(session_file)
        cred = f"{login}:{password}"
        self._basic = base64.b64encode(cred.encode()).decode()
        self._client: httpx.Client | None = None
        self.user_id: str = ""
        self.internal_ok: bool = False

    # ---------- transport ----------
    def _new_client(self) -> httpx.Client:
        return httpx.Client(timeout=40, headers={
            "User-Agent": BROWSER_UA,
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "ru-RU,ru;q=0.9",
            "Referer": self.base_url + "/",
            "Origin": self.base_url,
            "X-Requested-With": "XMLHttpRequest",
        })

    @property
    def client(self) -> httpx.Client:
        if self._client is None:
            self._client = self._new_client()
            self._load_session()
        return self._client

    def _load_session(self):
        try:
            if self.session_file.exists():
                data = json.loads(self.session_file.read_text(encoding="utf-8"))
                for k, v in data.get("cookies", {}).items():
                    self._client.cookies.set(k, v, domain="app.verifix.com", path="/")
                self.user_id = data.get("user_id", "")
                self.internal_ok = bool(self.user_id)
        except Exception:
            pass

    def _save_session(self):
        try:
            cookies = {k: v for k, v in self.client.cookies.items()}
            self.session_file.write_text(json.dumps({
                "cookies": cookies, "user_id": self.user_id,
            }), encoding="utf-8")
        except Exception:
            pass

    def login(self) -> bool:
        """Yangi sessiya ochish. True = internal API ishlashi mumkin."""
        self._client = self._new_client()
        sha1 = hashlib.sha1(self.password.encode()).hexdigest()
        try:
            r = self._client.post(
                f"{self.base_url}/b/biruni/s$log_in",
                data={"login": self.username, "password": sha1, "lang_code": "ru"})
            if r.status_code != 200:
                return False
            r2 = self._client.post(f"{self.base_url}/b/biruni/m:session", json={})
            if r2.status_code != 200:
                return False
            self.user_id = str(r2.json().get("user", {}).get("user_id", ""))
            self.internal_ok = bool(self.user_id)
            self._save_session()
            return self.internal_ok
        except Exception:
            return False

    def ensure_internal(self) -> bool:
        # avval saqlangan sessiyani yuklash (client property _load_session chaqiradi)
        _ = self.client
        if self.internal_ok and self.user_id:
            return True
        return self.login()

    def _headers(self) -> dict:
        return {"project_code": self.project_code, "filial_id": self.filial_id,
                "user_id": self.user_id, "lang_code": "ru",
                "Content-Type": "application/json;charset=UTF-8"}

    def _ipost(self, path: str, body: dict, retry_login: bool = True) -> Any:
        """Ichki sessiya API. 409/401 da bir marta relogin; ishlamasa eski
        sessiya tiklanadi (yaxshi sessiya ezilmasligi uchun) va SessionBlocked."""
        if not self.ensure_internal():
            raise SessionBlocked("login failed")
        now = time.time()
        dt = now - getattr(self, "_last_call", 0)
        if dt < 0.5:
            time.sleep(0.5 - dt)
        r = self.client.post(f"{self.base_url}{path}", json=body, headers=self._headers())
        self._last_call = time.time()
        if r.status_code in (401, 409) and retry_login:
            backup_cookies = {k: v for k, v in self.client.cookies.items()}
            backup_uid, backup_file = self.user_id, None
            try:
                if self.session_file.exists():
                    backup_file = self.session_file.read_text(encoding="utf-8")
            except Exception:
                pass
            time.sleep(2)
            if self.login():
                r = self.client.post(f"{self.base_url}{path}", json=body, headers=self._headers())
            if r.status_code in (401, 409):
                # relogin yordam bermadi - eski sessiyani tiklash
                try:
                    self._client = self._new_client()
                    for k, v in backup_cookies.items():
                        self._client.cookies.set(k, v, domain="app.verifix.com", path="/")
                    self.user_id = backup_uid
                    if backup_file is not None:
                        self.session_file.write_text(backup_file, encoding="utf-8")
                except Exception:
                    pass
        if r.status_code in (401, 409):
            self.internal_ok = False
            raise SessionBlocked(f"internal API blocked ({r.status_code})")
        r.raise_for_status()
        return r.json()

    def _bpost(self, path: str, body: dict) -> Any:
        """Public API (Basic auth) - sessiya cookiesiz alohida transport."""
        h = {"Authorization": "Basic " + self._basic, "project_code": self.project_code,
             "filial_id": self.filial_id, "Content-Type": "application/json",
             "Accept": "application/json", "User-Agent": BROWSER_UA}
        if getattr(self, "_pub", None) is None:
            self._pub = httpx.Client(timeout=40, headers={
                "User-Agent": BROWSER_UA, "Accept": "application/json"})
        r = self._pub.post(f"{self.base_url}{path}", json=body, headers=h)
        r.raise_for_status()
        j = r.json()
        return j.get("data", j) if isinstance(j, dict) and "data" in j else j

    # ---------- public API ----------
    def get_vacancies_simple(self) -> list[dict]:
        try:
            data = self._bpost("/b/vhr/api/v1/rec/vacancy$vacancies", {"job_id": None, "region_id": None})
            return data if isinstance(data, list) else []
        except Exception:
            return []

    def get_hirings(self) -> list[dict]:
        data = self._bpost("/b/vhr/api/v1/pro/hiring$list", {})
        out: list[dict] = []
        if isinstance(data, list):
            for j in data:
                for h in j.get("hirings", []) or []:
                    h = dict(h)
                    h["_journal_date"] = j.get("journal_date", "")
                    out.append(h)
        return out

    def get_employees(self) -> list[dict]:
        try:
            data = self._bpost("/b/vhr/api/v1/pro/employee$list", {})
        except Exception:
            data = self._bpost("/b/vhr/api/v1/core/employee$list", {})
        return data if isinstance(data, list) else []

    def get_dismissals(self) -> list[dict]:
        data = self._bpost("/b/vhr/api/v1/pro/dismissal$list", {})
        out: list[dict] = []
        if isinstance(data, list):
            for j in data:
                for d in j.get("dismissals", []) or []:
                    d = dict(d)
                    d["_journal_date"] = j.get("journal_date", "")
                    out.append(d)
        return out

    # ---------- ichki API ----------
    VAC_COLUMNS = ["name", "opened_date", "division_name", "job_name", "funnel_name",
                   "quantity", "scope_name", "vacancy_id", "wage_from", "wage_to",
                   "region_id", "status", "vacancy_type", "published_head_hunter",
                   "published_olx", "offer_accepted_candidates_count"]

    def get_vacancy_table(self, status: str = "O", limit: int = 100) -> list[dict]:
        j = self._ipost("/b/vhr/hrec/vacancy_list:table", {"p": {
            "column": self.VAC_COLUMNS, "filter": ["status", "=", status],
            "sort": ["-opened_date", "-created_on"], "offset": 0, "limit": limit}})
        rows = j.get("data", []) if isinstance(j, dict) else []
        out = []
        for r in rows:
            d = dict(zip(self.VAC_COLUMNS, r))
            d["opened"] = parse_verifix_date(d.get("opened_date"))
            out.append(d)
        return out

    def get_dashboard(self, date_from: str | None = None, date_to: str | None = None) -> dict:
        return self._ipost("/b/vhr/hrec/candidate/dashboard:load_dashboard_infos", {
            "date_from": date_from, "date_to": date_to, "vacancy_id": None,
            "division_id": None, "job_id": None, "region_id": None, "channel_id": None,
            "gender_code": None, "recruiter_id": None, "score_from": None, "score_to": None,
            "phone": None, "age_from": None, "age_to": None, "address": None,
            "dynamic_filters": []})

    EXTRA_STAGES = [("22487", "Предложение принято"), ("22488", "Кадровый резерв"),
                    ("22489", "Отклонена")]

    def get_all_candidates(self) -> list[dict]:
        """Barcha bosqichlardagi kandidatlar (aktiv + qabul/rezerv/rad, sahifalash bilan)."""
        url = "/b/vhr/hrec/candidate/candidate_management:load_candidates_by_stage"
        j = self._ipost(url, {"sort_by": "created", "sort_dir": "D", "dynamic_filters": []})
        out: list[dict] = []
        stages = list(j.get("stages", []))
        seen = {s["stage_id"] for s in stages}
        for sid, sname in self.EXTRA_STAGES:
            if sid in seen:
                continue
            try:
                jj = self._ipost(url, {"sort_by": "created", "sort_dir": "D",
                                       "dynamic_filters": [], "vacancy_type_ids": [],
                                       "stage_id": sid, "candidates": []})
                first = jj.get("candidates", []) if isinstance(jj, dict) else []
                if first:
                    stages.append({"stage_id": sid, "name": sname,
                                   "all_candidate_count": str(len(first)), "candidates": first,
                                   "_extra": True})
            except SessionBlocked:
                raise
            except Exception:
                pass
        for s in stages:
            loaded = list(s.get("candidates", []) or [])
            while True:
                sent = [{"candidate_id": x["candidate_id"], "vacancy_id": x.get("vacancy_id")}
                        for x in loaded]
                jj = self._ipost(url, {"sort_by": "created", "sort_dir": "D",
                                       "dynamic_filters": [], "vacancy_type_ids": [],
                                       "stage_id": s["stage_id"], "candidates": sent})
                more = jj.get("candidates", []) if isinstance(jj, dict) else []
                if not more:
                    break
                loaded.extend(more)
                if not s.get("_extra") and len(loaded) >= int(s.get("all_candidate_count", 0)):
                    break
            for cd in loaded:
                cd["_stage"] = s["name"]
                cd["_created"] = parse_verifix_date(cd.get("created_on"))
            out.extend(loaded)
        return out

    # ---------- oy filtrlari ----------
    @staticmethod
    def in_month(d: date | None, y: int, m: int) -> bool:
        return d is not None and d.year == y and d.month == m

    def hirings_in_month(self, y: int, m: int) -> list[dict]:
        out = []
        for h in self.get_hirings():
            d = parse_verifix_date(h.get("hiring_date")) or parse_verifix_date(h.get("_journal_date"))
            if self.in_month(d, y, m):
                h["_parsed_date"] = d
                job_c, div_c = parse_position_name(h.get("position_name"))
                h["_job_name"] = job_c or clean_verifix_label(h.get("job_name", ""))
                h["_division_name"] = div_c
                out.append(h)
        return sorted(out, key=lambda x: str(x.get("hiring_date")))

    def dismissals_in_month(self, y: int, m: int) -> list[dict]:
        out = []
        for d in self.get_dismissals():
            dt = parse_verifix_date(d.get("dismissal_date")) or parse_verifix_date(d.get("_journal_date"))
            if self.in_month(dt, y, m):
                d["_parsed_date"] = dt
                d["_reason"] = d.get("based_on_doc") or d.get("dismissal_note") or d.get("dismissal_reason_name") or ""
                out.append(d)
        return sorted(out, key=lambda x: str(x.get("dismissal_date")))
