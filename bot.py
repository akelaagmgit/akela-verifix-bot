"""Telegram bot - Verifix asosida Recruiting oylik hisobot .xlsx tayyorlaydi.

Buyruqlar:
- /start - Salomlashuv (inline menu)
- /hisobot - Oylik hisobot yaratish (oy tanlash)
- /vakansiyalar - ochiq vakansiyalar ro'yxati (Verifix)
- /qabul - shu oydagi qabul qilinganlar
- /yordam - yordam

Oqim: /hisobot -> inline tugmalar (oylarni tanlash, ko'p oy ham mumkin) -> "Tayyorladi (F.I.Sh.)"
ni so'raydi -> Verifix dan yig'adi -> template.xlsx dan yangi fayl quradi -> yuboradi.

Muhim: voronka (kelgan/skrining/suhbat/taklif), manbalar, xarajatlar Verifix
public API da yo'q (ichki candidate API 409) - ular shablondagi NAMUNA
qiymatlar bilan qoladi, Excelda ko'k kataklarni qo'lda to'g'rilaysiz.
Vakansiyalar, qabul/ketganlar va KPI (C32/C34/C40) avtomatik to'ladi.
"""
from __future__ import annotations

import asyncio
import calendar
from datetime import date
from pathlib import Path

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import CallbackQuery, FSInputFile, InlineKeyboardButton, InlineKeyboardMarkup, Message

from config import settings
from monthly import collect_month_data
from report_builder import build_report
from verifix_client import VerifixClient

MONTHS_UZ = ["", "Yanvar", "Fevral", "Mart", "Aprel", "May", "Iyun",
             "Iyul", "Avgust", "Sentyabr", "Oktyabr", "Noyabr", "Dekabr"]


class ReportFlow(StatesGroup):
    waiting_fio = State()
    resend = State()


async def _send_document(msg: Message, path: str, caption: str, kb, tries: int = 3):
    """Telegram tarmoq uzilishlarida qayta urinish bilan yuborish."""
    last_err = None
    for i in range(tries):
        try:
            await msg.answer_document(FSInputFile(path), caption=caption,
                                      parse_mode="HTML", reply_markup=kb)
            return
        except Exception as e:
            last_err = e
            await asyncio.sleep(3 * (i + 1))
    raise last_err


def resend_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🔁 Qayta yuborish", callback_data="report:resend")],
        [back_button()],
    ])


async def cb_resend(call: CallbackQuery, state: FSMContext):
    d = await state.get_data()
    path, caption = d.get("resend_path", ""), d.get("resend_caption", "")
    if not path or not Path(path).exists():
        await call.answer("Fayl topilmadi, /hisobot dan qayta boshlang", show_alert=True)
        return
    try:
        await _send_document(call.message, path, caption, main_menu_keyboard())
        await state.clear()
    except Exception:
        await call.answer("Hali ham aloqa yo'q, biroz kutib qayta bosing", show_alert=True)
    await call.answer()


def back_button() -> InlineKeyboardButton:
    return InlineKeyboardButton(text="⬅️ Orqaga", callback_data="menu:main")


def main_menu_keyboard() -> InlineKeyboardMarkup:
    """Asosiy menyu - inline tugmalar (doim ko'rinadi)"""
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📊 Hisobot yaratish", callback_data="menu:report")],
        [InlineKeyboardButton(text="🔍 O'tmaslik sabablari", callback_data="menu:failreasons")],
        [InlineKeyboardButton(text="ℹ️ Yordam", callback_data="menu:help")],
    ])


def month_selection_keyboard(prefix: str = "month") -> InlineKeyboardMarkup:
    """Oy tanlash klaviaturasi (bitta oy)."""
    today = date.today()
    rows = []
    # Joriy oydan 6 oy oldingi (yoki ko'proq) oylarni ko'rsatamiz
    for delta in range(6):
        y, m = today.year, today.month - delta
        while m < 1:
            m += 12
            y -= 1
        key = f"{y}-{m:02d}"
        label = f"{MONTHS_UZ[m]} {y}"
        rows.append([InlineKeyboardButton(text=label, callback_data=f"{prefix}:{key}")])
    rows.append([back_button()])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def get_client() -> VerifixClient:
    return VerifixClient(
        login=settings.verifix_login,
        password=settings.verifix_password,
        filial_id=settings.verifix_filial_id,
        project_code=settings.verifix_project_code,
        base_url=settings.verifix_base_url,
    )


async def cmd_start(msg: Message):
    await msg.answer(
        "Salom! Men <b>Akela Recruiting hisobot boti</b>man.\n\n"
        "Verifix dagi vakansiya + qabul/bo'shash ma'lumotlari asosida\n"
        "<b>Recruiting_oylik_hisobot.xlsx</b> ni tayyorlab beraman.\n\n"
        "Quyidagi tugmalardan birini tanlang:",
        parse_mode="HTML",
        reply_markup=main_menu_keyboard(),
    )


async def cmd_help(msg: Message):
    await msg.answer(
        "📖 <b>Yordam</b>\n\n"
        "📊 <b>Hisobot yaratish</b> - 1 oy tanlang, F.I.Sh. yozing, Excel oling\n"
        "🔍 <b>O'tmaslik sabablari</b> - oy tanlang, suhbat natijalari Excel oling\n\n"
        "Buyruqlar:\n"
        "/hisobot - oylik Excel hisobot\n"
        "/yordam - bu xabar",
        parse_mode="HTML",
        reply_markup=main_menu_keyboard(),
    )


async def _edit_msg(call: CallbackQuery, text: str, reply_markup=None):
    """Hisobot (hujjat) xabari o'zgarmaydi - javob har doim alohida xabar.
    Oddiy menyu xabarlari joyida tahrirlanadi."""
    try:
        if getattr(call.message, "document", None):
            await call.message.answer(text, parse_mode="HTML",
                                      reply_markup=reply_markup)
        else:
            await call.message.edit_text(text, parse_mode="HTML",
                                         reply_markup=reply_markup)
    except Exception:
        await call.message.answer(text, parse_mode="HTML", reply_markup=reply_markup)


async def cb_menu(call: CallbackQuery, state: FSMContext):
    action = call.data.split(":")[1]
    if action == "main":
        await _edit_msg(call, "Asosiy menyu:\nQuyidagi tugmalardan birini tanlang:",
                        main_menu_keyboard())
    elif action == "report":
        await _edit_msg(call, "Qaysi oy uchun hisobot tayyorlayman? (1 oy)",
                        month_selection_keyboard())
    elif action == "failreasons":
        await _edit_msg(call, "Qaysi oy uchun o'tmaslik sabablari hisoboti?",
                        month_selection_keyboard(prefix="fmonth"))
    elif action == "help":
        await _edit_msg(call,
            "📖 <b>Yordam</b>\n\n"
            "📊 <b>Hisobot yaratish</b> - 1 oy tanlang, F.I.Sh. yozing, Excel oling\n"
            "🔍 <b>O'tmaslik sabablari</b> - oy tanlang, suhbat natijalari Excel oling\n\n"
            "Buyruqlar:\n"
            "/hisobot - oylik Excel hisobot\n"
            "/yordam - bu xabar",
            main_menu_keyboard())
    await call.answer()


async def cb_month_single(call: CallbackQuery, state: FSMContext):
    key = call.data.split(":")[1]
    y, m = map(int, key.split("-"))
    await state.update_data(year=y, month=m)
    await state.set_state(ReportFlow.waiting_fio)
    await _edit_msg(call,
        f"{MONTHS_UZ[m]} {y} tanlandi.\n"
        "<b>1-bo'lim (Tayyorladi)</b> uchun F.I.Sh. ni yozing:\n"
        "(mas: Rajabova Sug'diyona)")
    await call.answer()


async def on_fio(msg: Message, state: FSMContext):
    fio = msg.text.strip()
    data = await state.get_data()
    await state.clear()
    
    wait = await msg.answer(f"⏳ Verifix dan ma'lumotlar olinmoqda...")
    try:
        client = get_client()
        loop = asyncio.get_running_loop()
        
        # Single month
        if "year" in data and "month" in data:
            y, m = data["year"], data["month"]
            m_data, m_stats = await loop.run_in_executor(None, collect_month_data, client, y, m, fio)
            out_dir = Path(settings.output_dir)
            fname = f"Recruiting_hisobot_{y}_{m:02d}.xlsx"
            out_path = out_dir / fname
            await loop.run_in_executor(None, build_report, settings.template_path, str(out_path), m_data)
            
            last = calendar.monthrange(y, m)[1]
            extra = ""
            if m_stats.get('funnel_note'):
                extra += f"\n• Voronka: {m_stats['funnel_note']}"
            if m_stats.get('dash_hint'):
                extra += f"\n• Dashboard: {m_stats['dash_hint']}"
            if not m_stats.get('internal'):
                extra += "\n• ⚠️ Ichki API bloklandi - faqat public ma'lumot (qayta urinib ko'ring)"
            caption = (
                f"✅ <b>{MONTHS_UZ[m]} {y}</b> (01.{m:02d}.{y} - {last:02d}.{m:02d}.{y})\n"
                f"• Oy boshida ochiq (C32): {m_stats.get('c32', m_stats['vacancies_open'])}, yangi (C33): {m_stats.get('c33', '-')}\n"
                f"• Shu oyda qabul: {m_stats['hired_in_month']}\n"
                f"• Shu oyda ketgan: {m_stats['dismissed_in_month']}{extra}"
            )
            try:
                await _send_document(msg, str(out_path), caption, main_menu_keyboard())
            except Exception:
                await state.update_data(resend_path=str(out_path), resend_caption=caption)
                await state.set_state(ReportFlow.resend)
                await wait.edit_text("⚠️ Internet uzildi, hisobot tayyor. Aloqa tiklangach pastdagi tugmani bosing.", reply_markup=resend_keyboard())
                return
            await wait.delete()
            return
        else:
            await wait.edit_text("Oy tanlanmagan. /hisobot dan qayta tanlang.", reply_markup=main_menu_keyboard())
            return
            
    except Exception:
        await wait.edit_text("❌ Hisobot tayyorlanmadi (internet yoki Verifix aloqasi). Birozdan keyin /hisobot dan qayta urinib ko'ring.", reply_markup=main_menu_keyboard())


async def show_vacancies(message_or_callback):
    try:
        client = get_client()
        loop = asyncio.get_running_loop()
        vacs = await loop.run_in_executor(None, client.get_vacancies)
        lines = [f"🏢 <b>Ochiq vakansiyalar: {len(vacs)}</b>"]
        for v in vacs[:25]:
            from verifix_client import clean_verifix_label
            name = clean_verifix_label(v.get("vacancy_name", ""))
            div = clean_verifix_label(v.get("division_name", ""))
            lines.append(f"• {name} — <i>{div}</i>")
        if len(vacs) > 25:
            lines.append(f"... yana {len(vacs)-25} ta (to'liq ro'yxat Excelda)")
        text = "\n".join(lines)
        kb = main_menu_keyboard()
        if isinstance(message_or_callback, Message):
            await message_or_callback.answer(text, parse_mode="HTML", reply_markup=kb)
        else:
            await message_or_callback.edit_text(text, parse_mode="HTML", reply_markup=kb)
    except Exception:
        text = "❌ Aloqa uzildi, birozdan keyin qayta urinib ko'ring"
        kb = main_menu_keyboard()
        if isinstance(message_or_callback, Message):
            await message_or_callback.answer(text, reply_markup=kb)
        else:
            await message_or_callback.edit_text(text, reply_markup=kb)


async def show_hired_current_month(message_or_callback):
    today = date.today()
    y, m = today.year, today.month
    try:
        client = get_client()
        loop = asyncio.get_running_loop()
        from monthly import get_month_people
        hires, dismissals, _tn = await loop.run_in_executor(None, get_month_people, client, y, m)
        lines = [f"✅ <b>{MONTHS_UZ[m]} {y}: qabul {len(hires)} ta, ketgan {len(dismissals)} ta</b>"]
        if hires:
            lines.append("\n<b>Qabul qilinganlar:</b>")
            for h in hires[:20]:
                dt = h.get('sana').strftime('%d.%m.%Y') if h.get('sana') else ''
                lines.append(f"• {h.get('fish','')} — {dt} ({h.get('lavozim','')[:50]})")
        if dismissals:
            lines.append("\n<b>Ketganlar:</b>")
            for d in dismissals[:20]:
                dt = d.get('sana').strftime('%d.%m.%Y') if d.get('sana') else ''
                reason = d.get('sabab', '')[:50]
                lines.append(f"• {d.get('fish','')} — {dt} ({reason})")
        text = "\n".join(lines)
        kb = main_menu_keyboard()
        if isinstance(message_or_callback, Message):
            await message_or_callback.answer(text, parse_mode="HTML", reply_markup=kb)
        else:
            await message_or_callback.edit_text(text, parse_mode="HTML", reply_markup=kb)
    except Exception:
        text = "❌ Aloqa uzildi, birozdan keyin qayta urinib ko'ring"
        kb = main_menu_keyboard()
        if isinstance(message_or_callback, Message):
            await message_or_callback.answer(text, reply_markup=kb)
        else:
            await message_or_callback.edit_text(text, reply_markup=kb)


async def cmd_vacancies(msg: Message):
    await show_vacancies(msg)


async def cmd_hired(msg: Message):
    parts = msg.text.split()
    today = date.today()
    y, m = today.year, today.month
    if len(parts) > 1:
        try:
            y, m = map(int, parts[1].split("-"))
        except ValueError:
            await msg.answer("Format: /qabul YYYY-MM (mas: /qabul 2026-09)", reply_markup=main_menu_keyboard())
            return
    try:
        client = get_client()
        loop = asyncio.get_running_loop()
        from monthly import get_month_people
        hires, dismissals, _tn = await loop.run_in_executor(None, get_month_people, client, y, m)
        lines = [f"✅ <b>{MONTHS_UZ[m]} {y}: qabul {len(hires)} ta, ketgan {len(dismissals)} ta</b>"]
        for h in hires[:30]:
            dt = h.get('sana').strftime('%d.%m.%Y') if h.get('sana') else ''
            lines.append(f"• {h.get('fish','')} — {dt} ({h.get('lavozim','')[:50]})")
        for d in dismissals[:30]:
            dt = d.get('sana').strftime('%d.%m.%Y') if d.get('sana') else ''
            reason = d.get('sabab', '')[:50]
            lines.append(f"• {d.get('fish','')} — {dt} ({reason})")
        await msg.answer("\n".join(lines), parse_mode="HTML", reply_markup=main_menu_keyboard())
    except Exception:
        await msg.answer("❌ Aloqa uzildi, birozdan keyin qayta urinib ko'ring", reply_markup=main_menu_keyboard())


async def cmd_report(msg: Message):
    await msg.answer("Qaysi oy uchun hisobot tayyorlayman? (1 oy)", reply_markup=month_selection_keyboard())


async def cb_fail_month(call: CallbackQuery, state: FSMContext):
    key = call.data.split(":")[1]
    y, m = map(int, key.split("-"))
    await call.answer()
    wait = await call.message.answer(f"⏳ {MONTHS_UZ[m]} {y} uchun o'tmaslik sabablari yig'ilmoqda...")
    try:
        from otmaslik_builder import collect_fail_data, build_fail_report
        client = get_client()
        loop = asyncio.get_running_loop()
        rows, info = await loop.run_in_executor(None, collect_fail_data, client, y, m)
        out_dir = Path(settings.output_dir)
        fname = f"Otmaslik_sabablari_{y}_{m:02d}.xlsx"
        out_path = out_dir / fname
        await loop.run_in_executor(
            None, build_fail_report, "template_otmaslik.xlsx", str(out_path), rows)
        caption = (
            f"🔍 <b>{MONTHS_UZ[m]} {y} — suhbatdan o'tmaslik sabablari</b>\n"
            f"• Suhbatga kelgan: {info['total']}\n"
            f"• O'tgan: {info['passed']}, o'tmagan: {info['failed']}, jarayonda: {info['pending']}\n"
            f"• Maosh sababi avtomatik: {info['wage_auto']} ta\n"
            f"• Sabab belgilanmagan: {info['failed'] - info['wage_auto']} ta — Excel'da G ustundan tanlang."
        )
        try:
            await _send_document(call.message, str(out_path), caption, main_menu_keyboard())
        except Exception:
            await state.update_data(resend_path=str(out_path), resend_caption=caption)
            await state.set_state(ReportFlow.resend)
            await wait.edit_text("⚠️ Internet uzildi, hisobot tayyor. Aloqa tiklangach pastdagi tugmani bosing.", reply_markup=resend_keyboard())
            return
        await wait.delete()
    except Exception:
        await wait.edit_text("❌ Hisobot tayyorlanmadi (internet yoki Verifix aloqasi). Birozdan keyin qayta urinib ko'ring.", reply_markup=main_menu_keyboard())


def _start_health_server():
    """Render free: PORT ga /health endpoint (uxlab qolmaslik + UptimeRobot ping uchun)."""
    import os
    import threading
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

    port = int(os.getenv("PORT", "0"))
    if not port:
        return

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            if self.path == "/health":
                body = b"ok"
                self.send_response(200)
                self.send_header("Content-Type", "text/plain")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            else:
                self.send_response(404)
                self.end_headers()

        def log_message(self, *args):
            pass

    srv = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    print(f"Health server on :{port}/health")


async def _session_keepalive():
    """Verifix sessiyasi o'lmasligi uchun har 7 daqiqada yangilash."""
    import time
    while True:
        await asyncio.sleep(7 * 60)
        try:
            loop = asyncio.get_running_loop()
            await loop.run_in_executor(None, _touch_session)
        except Exception:
            pass


def _touch_session():
    try:
        c = get_client()
        cli = c.client
        if c.user_id:
            cli.post(f"{c.base_url}/b/biruni/m:session", json={},
                      headers=c._headers(), timeout=20)
    except Exception:
        pass


def main():
    if not settings.bot_token:
        raise SystemExit("BOT_TOKEN .env da yo'q! .env.example dan nusxa oling.")
    bot = Bot(token=settings.bot_token)
    dp = Dispatcher()
    async def _on_startup(dispatcher):
        asyncio.create_task(_session_keepalive())
    dp.startup.register(_on_startup)
    dp.message.register(cmd_start, Command("start"))
    dp.message.register(cmd_help, Command("yordam", "help"))
    dp.message.register(cmd_report, Command("hisobot", "report"))
    dp.message.register(cmd_vacancies, Command("vakansiyalar", "vacancies"))
    dp.message.register(cmd_hired, Command("qabul", "hired"))
    dp.callback_query.register(cb_menu, F.data.startswith("menu:"))
    dp.callback_query.register(cb_month_single, F.data.startswith("month:"))
    dp.callback_query.register(cb_fail_month, F.data.startswith("fmonth:"))
    dp.callback_query.register(cb_resend, F.data == "report:resend")
    dp.message.register(on_fio, ReportFlow.waiting_fio)

    print("Bot ishga tushdi...")
    _start_health_server()
    dp.run_polling(bot)


if __name__ == "__main__":
    main()