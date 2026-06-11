# -*- coding: utf-8 -*-
"""
חילוץ שבוע אימון מ-JSON של Google Calendar אל קובץ Markdown מובנה.
שלב ביניים לפני הפקת הפק"א ל-Word.

כללי חילוץ (לפי דרישות המפתח):
1. פורמט שעה: "HH:MM-HH:MM".
2. אירוע שמשתרע על יותר ממשבצת זמן אחת מסומן כ"מאוחד" (rowspan ב-Word בהמשך).
3. שום טקסט לא נזרק. כשהשדה לא מזוהה — נכתב הטקסט הגולמי כפי שהוא ("זה מה שנכתב").
4. טכניקות/תרגולות — נשמר כל הטקסט מהכותרת/התיאור כולל סוגריים ורשימות.
5. לו"ז מקביל (ורוד) ולו"ז פלס"ם (אפור) — מופרדים לטבלה נפרדת בסוף.
   ** הערה: ה-MCP אינו מחזיר colorId. הסיווג כאן מבוסס טקסט בלבד.
      באתר (משיכה ישירה מה-API) יש colorId ואז הסיווג לפי צבע. **

שימוש:
    python calendar_to_md.py <events.json> <out.md> [--start 2026-06-14] [--days 5]
"""
import json, re, sys, argparse
from datetime import datetime, date, timedelta

HE_DAYS = ["א'", "ב'", "ג'", "ד'", "ה'", "ו'", "ש'"]

# כללי #5 — סיווג לו"ז מקביל / פלס"ם לפי טקסט (עד שיהיה colorId)
PALSAM_RE   = re.compile(r'פלס[\"׳\']?ם|פלסם')
PARALLEL_KW = ["חובשים", "תאג\"ד", "תאגד", "סמב\"צ", "סמבצ", "קשרים",
               "רענון", "הסמכ", "פלס\"ם", "פלסם", "מאבטח", "מגנ\"ט"]

# תחילית זמן משובצת בכותרת: "10:15-11:00| ..." → מוסרת (יש לנו את הזמן האמיתי)
TIME_PREFIX = re.compile(r'^\s*\d{1,2}:\d{2}\s*[-–]\s*\d{1,2}:\d{2}\s*\|\s*')
PLATOON_RE  = re.compile(r'פלוגה\s*([א-ד])')
TECH_RE     = re.compile(r'טכניק|תרגול')


def parse_iso(s):
    return datetime.fromisoformat(s)


def hhmm(dt):
    return dt.strftime('%H:%M')


def clean_title(summary):
    """מסיר רק תחילית-זמן כפולה; שאר הטקסט נשמר ככתבו (כלל #3, #4)."""
    return TIME_PREFIX.sub('', (summary or '')).strip()


def classify(title, desc, all_day, dur_hours):
    """מחזיר 'palsam' / 'parallel' / 'main'."""
    blob = (title or '') + ' ' + (desc or '')
    if PALSAM_RE.search(blob):
        return 'palsam'
    if any(k in blob for k in PARALLEL_KW):
        return 'parallel'
    # אירוע שמשתרע כמעט על כל היום ורץ במקביל → מקביל
    if all_day or dur_hours >= 6:
        return 'parallel'
    return 'main'


def load_events(path, start, days):
    raw = json.load(open(path, encoding='utf-8'))['events']
    last = start + timedelta(days=days)
    out = []
    for e in raw:
        s, en = e['start'], e['end']
        all_day = 'date' in s
        if all_day:
            d = date.fromisoformat(s['date'][:10])
            sd = datetime.combine(d, datetime.min.time())
            ed = sd
            dur = 24.0
        else:
            sd = parse_iso(s['dateTime'])
            ed = parse_iso(en['dateTime'])
            dur = (ed - sd).total_seconds() / 3600.0
        if not (start <= sd.date() < last):
            continue
        title = clean_title(e.get('summary', ''))
        desc = (e.get('description') or '').strip()
        m = PLATOON_RE.search(title)
        out.append({
            'day': sd.date(),
            'start': '' if all_day else hhmm(sd),
            'end': '' if all_day else hhmm(ed),
            'time': 'כל היום' if all_day else f"{hhmm(sd)}-{hhmm(ed)}",
            'title': title,
            'desc': desc,
            'loc': (e.get('location') or '').strip(),   # ריק דרך MCP
            'platoon': m.group(1) if m else None,
            'all_day': all_day,
            'dur': dur,
            'kind': classify(title, desc, all_day, dur),
            'tech': bool(TECH_RE.search(title + ' ' + desc)),
            'multi_slot': (not all_day) and dur >= 2.0,   # כלל #2: מועמד למיזוג משבצות
        })
    out.sort(key=lambda r: (r['day'], r['start'] or '99'))
    return out


def md_escape(s):
    return (s or '').replace('|', '\\|').replace('\n', '<br>')


def build_md(events, start, days, calendar_name):
    L = []
    L.append(f"# שליפת לו\"ז — {calendar_name}")
    L.append(f"\nשבוע אימון: {start.strftime('%d/%m/%Y')} ({HE_DAYS[(start.weekday()+1)%7]}) — "
             f"{days} ימים · סה\"כ {len(events)} אירועים\n")
    L.append("> ⚠️ ה-MCP לא מחזיר צבע (colorId) ולא מיקום. סיווג מקביל/פלס\"ם כאן לפי טקסט בלבד; "
             "באתר עצמו הסיווג לפי צבע (ורוד=מקביל, אפור=פלס\"ם).\n")

    for i in range(days):
        d = start + timedelta(days=i)
        day_ev = [e for e in events if e['day'] == d]
        main = [e for e in day_ev if e['kind'] == 'main']
        para = [e for e in day_ev if e['kind'] != 'main']
        plats = sorted({e['platoon'] for e in main if e['platoon']})
        is_split = len(plats) >= 2

        L.append(f"\n## יום {HE_DAYS[(d.weekday()+1)%7]} — {d.strftime('%d/%m')}"
                 + ("  ⇆ *(יום מפוצל)*" if is_split else ""))

        if not day_ev:
            L.append("\n_אין אירועים._")
            continue

        # ---- טבלת הלו"ז הראשי ----
        if is_split:
            L.append("\n**לו\"ז פלוגתי (מפוצל):**\n")
            header = "| שעה | " + " | ".join(f"פלוגה {p}" for p in plats) + " | כלל הגדוד |"
            sep = "|" + "---|" * (len(plats) + 2)
            L.append(header); L.append(sep)
            times = sorted({e['time'] for e in main})
            for t in times:
                row = [t]
                for p in plats:
                    cell = " · ".join(md_escape(e['title']) for e in main
                                      if e['time'] == t and e['platoon'] == p)
                    row.append(cell or "—")
                allrow = " · ".join(md_escape(e['title']) for e in main
                                    if e['time'] == t and not e['platoon'])
                row.append(allrow or "—")
                L.append("| " + " | ".join(row) + " |")
        else:
            L.append("\n**לו\"ז יומי:**\n")
            L.append("| שעה | פעילות | מיקום | מאוחד\\* | טכ' |")
            L.append("|---|---|---|---|---|")
            for e in main:
                L.append("| {t} | {a} | {loc} | {m} | {tech} |".format(
                    t=e['time'],
                    a=md_escape(e['title']) + (f"<br>_{md_escape(e['desc'])}_" if e['desc'] else ""),
                    loc=md_escape(e['loc']) or "—",
                    m="◧" if e['multi_slot'] else "",
                    tech="✓" if e['tech'] else ""))

        # ---- כלל #5: לו"ז מקביל + פלס"ם בנפרד ----
        if para:
            L.append("\n**לו\"ז מקביל / פלס\"ם (כלל #5):**\n")
            L.append("| שעה | פעילות | סוג |")
            L.append("|---|---|---|")
            for e in sorted(para, key=lambda x: x['start'] or ''):
                kind = "פלס\"ם" if e['kind'] == 'palsam' else "מקביל"
                L.append(f"| {e['time']} | {md_escape(e['title'])}"
                         + (f"<br>_{md_escape(e['desc'])}_" if e['desc'] else "")
                         + f" | {kind} |")

    L.append("\n\n---\n\\* ◧ = אירוע ≥שעתיים — מועמד למיזוג שתי משבצות בלו\"ז ה-Word (כלל #2).")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('json')
    ap.add_argument('out')
    ap.add_argument('--start', default='2026-06-14')
    ap.add_argument('--days', type=int, default=5)
    ap.add_argument('--name', default='מפקץ דרומי')
    a = ap.parse_args()
    start = date.fromisoformat(a.start)
    ev = load_events(a.json, start, a.days)
    md = build_md(ev, start, a.days, a.name)
    open(a.out, 'w', encoding='utf-8').write(md)
    print(f"OK: {len(ev)} events -> {a.out}")


if __name__ == '__main__':
    main()
