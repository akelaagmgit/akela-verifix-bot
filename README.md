# Akela Verifix Recruiting Oylik Hisobot Boti

Telegram bot Verifix dagi real ma'lumotlar asosida
`Recruiting_oylik_hisobot.xlsx` (sheet `Oylik hisobot`) ni avtomatik tayyorlaydi.

## Excel tahlili (qisqa)

Shablon `template.xlsx`:
- **Ko'k shrift (0000FF)** = input. Bot faqat shularni yozadi.
- **Qora** = formulalar, tegilmaydi (`C35`, `E**`, `C43`, voronka, jami).
- `C11` davr boshi, `C12` davr oxiri, `C13` tayyorladi, `C14` hisobot sanasi.
- `B18:I27` (10 qator) vakansiyalar, `G` formula (`C12-E`).
- KPI: `C32` oy boshida ochiq, `C33` yangi, `C34` yopilgan, `C36-C40` voronka.
- Manbalar `B49:G54`, xarajatlar `C85/D85..`, muammolar, reja, xulosa.

## Verifix manbalari (barchasi API dan)

Login: `admin@akela` (sessiya `session.json` da, keep-alive har 15 daq).

| Excel | Manba |
|---|---|
| Vakansiyalar (nom, bo'lim, **ochilish sanasi**, id) | `vacancy_list:table` (O/C status) |
| Vakansiya bosqichi (E'lon..Taklif) | kandidat bosqichlaridan avtomatik |
| C32 oy boshida ochiq / C33 yangi | ochilish sanasidan hisoblanadi |
| C36-C39 voronka | `candidate_management` (aktiv+rezerv+rad, sahifalash) |
| 4-Manbalar (kanal bo'yicha) | `channel_id` (2421=Telegram, 2422=hh.uz) |
| C34/C40, 6-bo'lim (yangi/ketgan) | `hiring$list` / `dismissal$list` (ground truth) |
| Dashboard hint (konversiya, muddati o'tgan) | `dashboard:load_dashboard_infos` (oy filtri) |

**Qo'lda to'ldiriladi** (Verifixda yo'q): 7-Xarajatlar, C41/C42 time,
D ustun (o'tgan oy), 8-Muammolar, 9-Reja, 2-Bosqich/Ustuvorlik/Buyurtmachi.
Hisobot caption'ida Avtomatik/Qo'lda ro'yxati chiqadi.

Muhim: ichki API yangi sessiyalarda 409 beradi (server sessiya limiti),
shuning uchun bitta persistent sessiya ishlatiladi + uni ezmaslik himoyasi bor.
Ichki API bloklansa bot public-only fallback'ga o'tadi.

## Ishga tushirish

```powershell
cd C:\forwork\akela_verifix_bot
python -m pip install -r requirements.txt
python bot.py
```

Telegram: `/start` → tugmalar → `/hisobot` (1 oy) / `/multihisobot` (2-3 oy bir faylda),
`/vakansiyalar`, `/qabul 2026-09`, `/yordam`.

## Fayllar

- `bot.py` - aiogram v3 bot (inline menu, multi-oy, keep-alive)
- `verifix_client.py` - sessiya API + public API
- `monthly.py` - oy kesimida yig'ish va KPI xaritası
- `report_builder.py` - shablonni to'ldirish (formulalarni saqlaydi)
- `template.xlsx`, `session.json`, `config.py`, `.env`, `requirements.txt`
