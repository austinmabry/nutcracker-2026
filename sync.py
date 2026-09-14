#!/usr/bin/env python3
"""
Sync the Nutcracker "My Schedule" site from the Ballet Arkansas master Google Sheet.

  python sync.py                # download the master sheet and rebuild docs/
  python sync.py --file x.xlsx  # rebuild from a local copy instead (for testing)
  python sync.py --check        # parse + validate only, write nothing

Exit code 0 = ok (outputs written, or nothing changed).
Exit code 1 = the sheet could not be parsed with confidence; nothing was written,
              so the published site keeps its last good version.
"""
import argparse, datetime, hashlib, io, json, os, re, sys
from openpyxl import load_workbook

SHEET_ID = "1xWFCPfY5m0aK9d5dnxvALNre-6-N7Cl0kkw24rvm6cE"
EXPORT_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
YEAR = 2026
TZ = "America/Chicago"
HERE = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(HERE, "docs")
THANKSGIVING = datetime.date(YEAR, 11, 26)          # BA: rehearsals after Thanksgiving are mandatory

ROLES = ["Infantry", "Commandants", "Artillery", "Brigade", "Lieutenants", "Mice", "Rats", "Clara", "Fritz",
         "Party Children", "Prince", "Snow", "Cherubs", "Seraphs", "Trumpeter", "Archangels", "Hot Chocolate",
         "Tea", "Coffee", "Candy Cane", "Bon Bons", "Flowers"]
SOLDIERS = ["Infantry", "Commandants", "Artillery", "Brigade", "Lieutenants"]
BATTLE = SOLDIERS + ["Mice", "Rats", "Clara"]
PARTY = ["Party Children", "Clara", "Fritz"]
ANGELS = ["Cherubs", "Seraphs", "Trumpeter", "Archangels"]
ACT1 = PARTY + BATTLE + ["Snow", "Prince"]
ACT2 = ["Hot Chocolate", "Tea", "Coffee", "Candy Cane", "Bon Bons", "Flowers", "Clara", "Prince"] + ANGELS
# text that legitimately has no community-cast role attached
NO_ROLE_OK = re.compile(r"Marzipan|Grand Pas|Notes\b|Understud", re.I)

MONTHS = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "jul": 7, "aug": 8, "sep": 9, "sept": 9,
          "oct": 10, "nov": 11, "dec": 12}
DAY_RE = re.compile(r"^(Mon|Tue|Tues|Wed|Thu|Thur|Thurs|Fri|Sat|Sun)\.?\s+([A-Za-z]+)\.?\s+(\d{1,2})", re.I)
WEEK_RE = re.compile(r"^(WEEK\s*\d+|OFF\s*WEEK)", re.I)
TIME = r"(\d{1,2}):(\d{2})\s*([AaPp][Mm])?"
GRID_ENTRY_RE = re.compile(rf"^{TIME}\s*[-–]\s*{TIME}\s+(.*)$")
PW_CALL_RE = re.compile(rf"^{TIME}\s*[-–]?\s*Call\s*Time\s*[-–]\s*(.+)$", re.I)
PW_RANGE_RE = re.compile(rf"^{TIME}\s*[-–]\s*{TIME}\s*[-–]?\s*(.+)$")


class ParseError(Exception):
    pass


# ----------------------------------------------------------------------------- helpers
def has(word, text):
    return re.search(r"\b" + word + r"\b", text, re.I) is not None


def tag(text):
    """Map a schedule entry's text to the community-cast roles it applies to."""
    t = text
    s = set()
    if has("Backup Fitting", t):
        return list(ROLES)
    for w in ["Infantry", "Commandants", "Artillery", "Brigade", "Lieutenants", "Mice", "Clara", "Fritz", "Snow",
              "Flowers", "Tea", "Hot Chocolate", "Coffee", "Candy Cane", "Cherubs", "Seraphs", "Trumpeter"]:
        if has(w, t):
            s.add(w)
    if has("Rats?", t):
        s.add("Rats")
    if has("Soldiers", t):
        s.update(SOLDIERS)
    if has("Battle", t) and not re.search(r"Battle\s*\(Infantry", t):   # "Battle (Infantry/...)" names its groups
        s.update(BATTLE)
    if has("Party", t):
        s.update(PARTY)
    if has("Prince", t) or has("Nutcracker", t):
        s.add("Prince")
    if has("Bon Bons?", t):
        s.add("Bon Bons")
    if has("Archangels?", t):
        s.add("Archangels")
    if has("Angels", t):
        s.update(ANGELS)
    if has("ACT I", t) and not has("ACT II", t) and not re.search(r"ACT I\s+(Party|Battle|Snow)", t):
        s.update(ACT1)          # "ACT I Party Cast B" is a Party-only call; keyword rules handle those
    if has("ACT II", t):
        if "(" in t:                         # explicit list, e.g. school show "(Hot Chocolate, Tea, Candy Cane, Flowers)"
            s.update(["Clara", "Prince"])
        else:
            s.update(ACT2)
    if has("Snow Scene", t):
        s.update(["Clara", "Prince"])
    if has("Battle Scene", t) or has("Party Scene", t):
        s.add("Prince")
    return [r for r in ROLES if r in s]


def cast_of(text):
    m = re.search(r"\bCast\s+([AB])\b", text)
    return m.group(1) if m else "All"


def to24(h, mn, ampm, default_pm=True):
    h = int(h)
    if ampm:
        ampm = ampm.upper()
        if ampm == "PM" and h != 12:
            h += 12
        if ampm == "AM" and h == 12:
            h = 0
    elif default_pm and h < 8:
        h += 12
    return f"{h:02d}:{mn}"


def fmt12(hhmm):
    h, m = map(int, hhmm.split(":"))
    return f"{(h % 12) or 12}:{m:02d} {'AM' if h < 12 else 'PM'}"


def clean(t):
    t = re.sub(r"\s+", " ", t.replace("\n", " ")).strip()
    t = t.replace("Mandatory Cherubs", "Cherubs")
    t = re.sub(r"(\d)\s*-\s*(\d)", r"\1–\2", t)
    return t


# ----------------------------------------------------------------------------- weekend grid
def parse_grid(ws):
    """Rows between the studio header and the production-week marker."""
    hdr = None
    for r in range(1, min(ws.max_row, 40) + 1):
        cells = [str(ws.cell(r, c).value or "") for c in range(1, ws.max_column + 1)]
        if sum("studio" in c.lower() for c in cells) >= 3:
            hdr = r
            cols = [(c, re.sub(r"^BA\s*[-–]", "Ballet Arkansas –", clean(cells[c - 1]).replace("S&BII", "S&B II").replace(" - ", " – ")))
                    for c in range(2, ws.max_column + 1) if cells[c - 1].strip()]
            break
    if hdr is None:
        raise ParseError("could not find the studio header row (a row with 3+ 'Studio' cells)")

    events, seen, week, day = [], set(), None, None
    pw_start = None
    for r in range(hdr + 1, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or "").strip()
        if "PRODUCTION WEEK" in a.upper():
            pw_start = r
            break
        if a and WEEK_RE.match(a):
            week = clean(a).title().replace("Off Week", "Off week")
            continue
        m = DAY_RE.match(a)
        if m:
            mon = MONTHS.get(m.group(2).lower()[:4]) or MONTHS.get(m.group(2).lower()[:3])
            if not mon:
                raise ParseError(f"row {r}: unknown month in {a!r}")
            date = datetime.date(YEAR, mon, int(m.group(3)))
            day = {"date": date, "label": f"{m.group(1).title()[:3]}. {date.strftime('%b')}. {date.day}", "week": week}
        elif a and not re.match(r"^[A-Za-z]{3,4}\.?\s+\d", a):
            pass  # things like "Oct. 24-25" off-week rows fall through to the OFF check below
        if day is None:
            continue
        for c, loc in cols:
            v = ws.cell(r, c).value
            if not v:
                continue
            text = str(v)
            if "No Rehearsal" in text:
                continue
            if re.search(r"\b(SFD|YAGP)\b", text):
                continue
            if re.search(r"Studio|Costume Shop", text) and not re.match(r"\s*\d", text):   # repeated header row
                continue
            lines = [l.strip() for l in text.split("\n") if l.strip()]
            if len(lines) > 1 and not all(re.match(r"\d", l) for l in lines):
                lines = [" ".join(lines)]
            for line in lines:
                line = clean(line)
                m = GRID_ENTRY_RE.match(line.replace("–", "-"))
                if not m:
                    raise ParseError(f"row {r} col {c}: cannot read {line!r} as 'H:MM-H:MM text'")
                h1, m1, ap1, h2, m2, ap2, what = m.groups()
                start, end = to24(h1, m1, ap1 or ap2), to24(h2, m2, ap2 or ap1)
                what = clean(what)
                if "Fitting" in what:
                    typ, note = "Costume fitting – MANDATORY", ""
                    if "Backup" in what:
                        what, note = "Backup fitting date", "Only if you missed your fitting"
                    else:
                        what = re.sub(r"\s*Mandatory Fitting", "", what)
                elif day["date"] > THANKSGIVING or re.search(r"w/\s*Company", what, re.I):
                    typ, note = "Mandatory – w/ Company", ""
                else:
                    typ, note = "Regular rehearsal", ""
                roles = tag(what)
                cast = cast_of(what)
                if typ == "Regular rehearsal":
                    keys = {(ro, cs) for ro in roles for cs in (["A", "B"] if cast == "All" else [cast])}
                    if any(k not in seen for k in keys):
                        typ = "1st rehearsal – MANDATORY"
                    seen |= keys
                if not roles and not NO_ROLE_OK.search(what):
                    raise ParseError(f"row {r} col {c}: no role recognised in {what!r} – add a rule to tag()")
                events.append(dict(date=date_iso(day["date"]), day=day["label"], week=week or "",
                                   time=f"{line.split(' ')[0]}", what=what, cast=cast, type=typ,
                                   loc=loc, note=note, tags=roles, start=start, end=end))
    if pw_start is None:
        raise ParseError("could not find the 'PRODUCTION WEEK' row")
    return events, pw_start


def date_iso(d):
    return d.isoformat()


# ----------------------------------------------------------------------------- production week
def parse_production(ws, start_row):
    loc = "Robinson Center"
    for r in range(start_row, min(start_row + 3, ws.max_row) + 1):
        txt = " ".join(str(ws.cell(r, c).value or "") for c in range(1, 3))
        m = re.search(r"@\s*([A-Za-z ]+Center)", txt)
        if m:
            loc = m.group(1).strip()
    events = []
    for r in range(start_row, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or "").strip()
        m = DAY_RE.match(a)
        if not m:
            continue
        mon = MONTHS.get(m.group(2).lower()[:4]) or MONTHS.get(m.group(2).lower()[:3])
        date = datetime.date(YEAR, mon, int(m.group(3)))
        label = f"{m.group(1).title()[:3]}. {date.strftime('%b')}. {date.day}"
        body = "\n".join(str(ws.cell(r, c).value or "") for c in range(2, ws.max_column + 1))
        section, sec_note, pending = "", "", []
        for raw in body.split("\n"):
            line = raw.strip()
            if not line:
                continue
            mc = PW_CALL_RE.match(line)
            mr = PW_RANGE_RE.match(line)
            if mc:
                pending.append((to24(mc.group(1), mc.group(2), mc.group(3)), clean(mc.group(4))))
                continue
            if mr:
                h1, m1, ap1, h2, m2, ap2, desc = mr.groups()
                s, e = to24(h1, m1, ap1 or ap2), to24(h2, m2, ap2 or ap1)
                desc = clean(desc)
                if desc.startswith("["):
                    who = desc.split(":")[-1].strip(" ]")
                    items = [(None, who, "")]
                elif pending:
                    items = [(call, who, desc) for call, who in pending]
                else:
                    items = [(None, desc, "")]
                pending = []
                for call, who, activity in items:
                    cast = cast_of(who)
                    if cast == "All":
                        cast = cast_of(section)
                    if "Both Casts" in section:
                        cast = "All"
                    roles = tag(who)
                    if not roles and not NO_ROLE_OK.search(who):
                        raise ParseError(f"production week {label}: no role recognised in {who!r}")
                    time = (f"Call {fmt12(call)} · " if call else "") + f"{fmt12(s)}–{fmt12(e)}"
                    note = "; ".join(x for x in [activity if activity and activity != who else "", sec_note] if x)
                    events.append(dict(date=date_iso(date), day=label, week="Production Week", time=time,
                                       what=f"{section} – {who}" if section else who, cast=cast,
                                       type="Mandatory – Production Week", loc=loc, note=note, tags=roles,
                                       start=call or s, end=e))
                continue
            # anything else is a section heading, e.g. "Cast B Dress Rehearsal (In Performance Costumes, ...)"
            sec_note = ""
            mm = re.match(r"^(.*?)\s*\*\*\s*(.+)$", line)
            if mm:
                line, sec_note = mm.group(1), mm.group(2).strip()
            mm = re.match(r"^(.*?)\s*\((.+)\)\s*$", line)
            if mm and not re.search(r"@", mm.group(1)):
                line, sec_note = mm.group(1), mm.group(2)
            section = clean(line)
            pending = []
    return events


# ----------------------------------------------------------------------------- validation
def validate(events):
    problems = []
    if len(events) < 150:
        problems.append(f"only {len(events)} events parsed (expected 150+)")
    dates = sorted({e["date"] for e in events})
    if len(dates) < 25:
        problems.append(f"only {len(dates)} distinct dates (expected 25+)")
    pw = [e for e in events if e["week"] == "Production Week"]
    if len(pw) < 20:
        problems.append(f"only {len(pw)} production-week events (expected 20+)")
    for role in ROLES:
        for cast in "AB":
            n = sum(1 for e in events if role in e["tags"] and e["cast"] in ("All", cast))
            if not 5 <= n <= 80:
                problems.append(f"{role} Cast {cast}: {n} events looks wrong")
    for e in events:
        if not re.match(r"^\d\d:\d\d$", e["start"]) or e["end"] <= e["start"]:
            problems.append(f"{e['day']} {e['what']}: bad times {e['start']}–{e['end']}")
    return problems


# ----------------------------------------------------------------------------- standing notes from the sheet
DEFAULT_POLICY = {
    "mandatory": "Mandatory Rehearsals (1st Rehearsal, Rehearsals after Thanksgiving and Production Week Rehearsals and Performances)",
    "fitting": "Mandatory Fittings",
    "regular": "Regular Rehearsal",
    "casts": "All Casts are called unless a cast is named",
    "calltime": "Call Time is the time dancers should be checked in and backstage - Arrive 15 minutes before Call Time",
}


def parse_policies(ws):
    """The colour-key / instruction text BA keeps at the top of the master and above production week."""
    p = dict(DEFAULT_POLICY)
    for r in range(1, ws.max_row + 1):
        for c in range(1, min(ws.max_column, 3) + 1):
            t = clean(str(ws.cell(r, c).value or ""))
            if not t:
                continue
            if t.lower().startswith("mandatory rehearsal"):
                p["mandatory"] = t
            elif t.lower().startswith("mandatory fitting"):
                p["fitting"] = t
            elif t.lower().startswith("regular rehearsal"):
                p["regular"] = t
            elif t.lower().startswith("all casts"):
                p["casts"] = t
            elif "call time" in t.lower() and "backstage" in t.lower():
                i = t.lower().find("call time")
                p["calltime"] = t[i:].replace("*", "").strip(" (")
    return p


GUIDE_MAP = [  # keyword in the guideline sheet's Role column -> our roles
    ("Clara", ["Clara"]), ("Fritz", ["Fritz"]), ("Party Children", ["Party Children"]), ("Mice", ["Mice"]),
    ("Rats", ["Rats"]), ("Lieutenant", ["Lieutenants"]), ("Leiutenant", ["Lieutenants"]),
    ("Artillery", ["Artillery", "Brigade", "Infantry"]), ("Commandant", ["Commandants"]), ("Snow", ["Snow"]),
    ("Trumpeter", ["Trumpeter"]), ("Cherubs", ["Cherubs"]), ("Seraphs", ["Seraphs"]), ("Archangels", ["Archangels"]),
    ("Hot Chocolate", ["Hot Chocolate"]), ("Tea", ["Tea"]), ("Coffee", ["Coffee"]), ("Candy Cane", ["Candy Cane"]),
    ("Bon Bon", ["Bon Bons"]), ("Flowers", ["Flowers"]),
]
GUIDE_COLS = ["Costume (you provide)", "Shoes", "Makeup", "Hair", "Accessories (BA provides)", "Prop (BA provides)"]


def parse_guidelines(wb):
    """The 'Costume, Hair, Makeup Guidelines' sheet -> per-role text + the makeup/care instruction blocks."""
    ws = next((w for w in wb.worksheets if "costume" in w.title.lower()), None)
    out = {"roles": {}, "blocks": {}, "general": ""}
    if ws is None:
        return out
    for r in range(1, ws.max_row + 1):
        a = clean(str(ws.cell(r, 1).value or ""))
        if not a:
            continue
        low = a.lower()
        raw_a = str(ws.cell(r, 1).value or "")
        lines = [re.sub(r"^[^\w*(]+", "", l.strip()) for l in raw_a.split("\n") if l.strip(" \uf0b7•·-")]
        block = "\n".join(("• " + l) if i else l for i, l in enumerate(lines) if l)
        if low.startswith("full stage makeup"):
            out["blocks"]["Full"] = block
        elif low.startswith("light stage makeup"):
            out["blocks"]["Light"] = block
        elif low.startswith("men's stage makeup") or low.startswith("mens stage makeup"):
            out["blocks"]["Men's"] = block
        elif low.startswith("costume care"):
            out["blocks"]["Care"] = block
        elif low.startswith("dancers of color"):
            out["general"] = a
        else:
            for key, roles in GUIDE_MAP:
                if key.lower() in low and "act " not in low[:4]:
                    vals = [clean(str(ws.cell(r, c).value or "")) for c in range(2, 8)]
                    parts = [f"{lab}: {v}" for lab, v in zip(GUIDE_COLS, vals) if v]
                    txt = f"{a} — " + "; ".join(parts)
                    for ro in roles:
                        out["roles"].setdefault(ro, txt)
                    break
    return out


def build_details(e, policy, backup):
    """Notes for one event: the day-specific part first and clearly marked, the standing rules after."""
    today, rules = [], []
    t = e["type"]
    if t.startswith("1st"):
        today.append(f"MANDATORY — this is the first rehearsal for {e['what']}.")
        rules.append(f"BA: {policy['mandatory']}.")
    elif t.startswith("Costume"):
        if e["what"].startswith("Backup"):
            today.append("BACKUP FITTING DATE — only for dancers who missed their scheduled costume fitting. Not needed if your dancer has already been fitted.")
        else:
            today.append(f"MANDATORY COSTUME FITTING for {e['what']}.")
            if backup:
                today.append(f"If this fitting is missed, the backup fitting date is {backup}.")
            if len(policy["fitting"]) > 25:
                rules.append(f"BA: {policy['fitting']}.")
    elif t.startswith("Mandatory – w/"):
        today.append("MANDATORY rehearsal with the Company.")
        rules.append(f"BA: {policy['mandatory']}.")
    elif t.startswith("Mandatory – Production"):
        if e["time"].startswith("Call"):
            call = e["time"].split("·")[0].replace("Call", "").strip()
            h, m = map(int, e["start"].split(":"))
            arrive = fmt12(f"{(h * 60 + m - 15) // 60:02d}:{(h * 60 + m - 15) % 60:02d}")
            today.append(f"CALL TIME {call} — be checked in and backstage by then. ARRIVE BY {arrive}.")
            today.append(f"On stage: {e['time'].split('·', 1)[1].strip()}.")
        else:
            today.append(f"Time: {e['time']}.")
        if e["note"]:
            today.append(e["note"] + ".")
        today.append("MANDATORY — production week.")
        rules.append(f"BA: {policy['calltime']}.")
        rules.append(f"BA: {policy['mandatory']}.")
    else:
        today.append("Regular rehearsal.")
        if len(policy["regular"]) > 25:
            rules.append(f"BA: {policy['regular']}.")
    if e["note"] and not t.startswith("Mandatory – Production") and not e["what"].startswith("Backup"):
        today.append(e["note"] + ".")
    rules.append(f"{policy['casts']}.")
    rules.append(f"SUBJECT TO CHANGE — the BA master schedule, the weekly BA emails and the BA portal are the official source. Master: https://docs.google.com/spreadsheets/d/{SHEET_ID}/")
    return "*** THIS DAY ***\n" + "\n".join(today) + "\n\n*** STANDING RULES ***\n" + "\n".join(rules)


def is_dress(e):
    return e["week"] == "Production Week" and re.search(r"Dress|Performance|School Show|@\s*\d", e["what"]) is not None \
        and "Spacing" not in e["what"]


def role_costume_text(role, guide):
    g = guide["roles"].get(role)
    if not g:
        return ""
    out = [f"*** COSTUME / HAIR / MAKEUP — {role} ***", g]
    mk = re.search(r"Makeup: ([^;]+)", g)
    if mk:
        want = mk.group(1)
        for k in ["Full", "Light", "Men's"]:
            if k.lower() in want.lower() and k in guide["blocks"]:
                out.append(guide["blocks"][k])
    if guide["general"]:
        out.append(guide["general"] + ".")
    if "Care" in guide["blocks"]:
        out.append(guide["blocks"]["Care"])
    return "\n".join(out)


# ----------------------------------------------------------------------------- outputs
def esc(s):
    return str(s).replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")


def fold(line):
    out, b = [], line.encode("utf-8")
    while len(b) > 73:
        cut = 73
        while (b[cut] & 0xC0) == 0x80:
            cut -= 1
        out.append(b[:cut].decode())
        b = b" " + b[cut:]
    out.append(b.decode())
    return "\r\n".join(out)


VTIMEZONE = ["BEGIN:VTIMEZONE", f"TZID:{TZ}",
             "BEGIN:DAYLIGHT", "TZOFFSETFROM:-0600", "TZOFFSETTO:-0500", "TZNAME:CDT", "DTSTART:19700308T020000",
             "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU", "END:DAYLIGHT",
             "BEGIN:STANDARD", "TZOFFSETFROM:-0500", "TZOFFSETTO:-0600", "TZNAME:CST", "DTSTART:19701101T020000",
             "RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU", "END:STANDARD", "END:VTIMEZONE"]


def slug(role, cast):
    return f"{role.replace(' ', '')}-Cast{cast}"


def ics(role, cast, events, stamp, seq, guide):
    rows = [e for e in events if role in e["tags"] and e["cast"] in ("All", cast)]
    L = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Nutcracker CC Schedule 2026//EN", "CALSCALE:GREGORIAN",
         "METHOD:PUBLISH", f"X-WR-CALNAME:Nutcracker 2026 – {role} Cast {cast}", f"X-WR-TIMEZONE:{TZ}",
         "REFRESH-INTERVAL;VALUE=DURATION:PT6H", "X-PUBLISHED-TTL:PT6H"] + VTIMEZONE
    for e in rows:
        ds = e["date"].replace("-", "")
        title = "Nutcracker: " + e["what"] + ("" if e["type"].startswith("Regular") else f" ({e['type']})")
        uid = hashlib.sha1(f"{slug(role, cast)}|{e['date']}|{e['what']}".encode()).hexdigest()[:16]
        L += ["BEGIN:VEVENT", f"UID:{uid}@nutcracker2026", f"DTSTAMP:{stamp}", f"SEQUENCE:{seq}",
              f"DTSTART;TZID={TZ}:{ds}T{e['start'].replace(':', '')}00",
              f"DTEND;TZID={TZ}:{ds}T{e['end'].replace(':', '')}00",
              f"SUMMARY:{esc(title)}", f"LOCATION:{esc(e['loc'])}",
              "DESCRIPTION:" + esc(e["details"] + ("\n\n" + role_costume_text(role, guide) if is_dress(e) and role_costume_text(role, guide) else "")),
              "BEGIN:VALARM", "TRIGGER:-PT1H", "ACTION:DISPLAY", "DESCRIPTION:Nutcracker in 1 hour", "END:VALARM",
              "END:VEVENT"]
    L.append("END:VCALENDAR")
    return "\r\n".join(fold(l) for l in L) + "\r\n"


def event_key(e):
    return f"{e['date']} {e['start']}–{e['end']} | {e['what']} | {e['cast']} | {e['loc']}"


def write_outputs(events, synced_label, guide):
    os.makedirs(os.path.join(DOCS, "ics"), exist_ok=True)
    now = datetime.datetime.now(datetime.timezone.utc)
    stamp, seq = now.strftime("%Y%m%dT%H%M%SZ"), int(now.strftime("%Y%m%d%H"))
    for role in ROLES:
        for cast in "AB":
            with open(os.path.join(DOCS, "ics", f"{slug(role, cast)}.ics"), "w", newline="") as f:
                f.write(ics(role, cast, events, stamp, seq, guide))
    costume = {ro: role_costume_text(ro, guide) for ro in ROLES}
    data = json.dumps({"roles": ROLES, "rows": events, "synced": synced_label, "costume": costume}, ensure_ascii=False)
    with open(os.path.join(DOCS, "data.json"), "w") as f:
        f.write(data)
    tpl = open(os.path.join(HERE, "template.html")).read()
    with open(os.path.join(DOCS, "index.html"), "w") as f:
        f.write(tpl.replace("__DATA__", data))
    open(os.path.join(DOCS, ".nojekyll"), "w").close()


def append_changes(old, new, synced_label):
    ok, nk = {event_key(e) for e in old}, {event_key(e) for e in new}
    added, removed = sorted(nk - ok), sorted(ok - nk)
    if not added and not removed:
        return False
    with open(os.path.join(HERE, "CHANGES.md"), "a") as f:
        f.write(f"\n## {synced_label}\n")
        for k in removed:
            f.write(f"- REMOVED: {k}\n")
        for k in added:
            f.write(f"- ADDED: {k}\n")
    return True


# ----------------------------------------------------------------------------- main
def load_workbook_bytes(path=None):
    if path:
        return load_workbook(path, data_only=True)
    import requests
    r = requests.get(EXPORT_URL, timeout=60)
    r.raise_for_status()
    if not r.content.startswith(b"PK"):
        raise ParseError("download did not return an .xlsx – is the sheet still shared as 'anyone with the link'?")
    return load_workbook(io.BytesIO(r.content), data_only=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", help="parse a local .xlsx instead of downloading")
    ap.add_argument("--check", action="store_true", help="parse and validate only; write nothing")
    args = ap.parse_args()
    try:
        wb = load_workbook_bytes(args.file)
        ws = wb.worksheets[0]
        grid, pw_row = parse_grid(ws)
        prod = parse_production(ws, pw_row)
        events = sorted(grid + prod, key=lambda e: (e["date"], e["start"], e["loc"]))
        policy = parse_policies(ws)
        guide = parse_guidelines(wb)
        bk = next((e for e in events if e["what"].startswith("Backup")), None)
        backup = f"{bk['day']}, {bk['time']}, {bk['loc']}" if bk else ""
        for e in events:
            e["details"] = build_details(e, policy, backup)
            e["dress"] = is_dress(e)
        problems = validate(events)
        if problems:
            raise ParseError("validation failed:\n  " + "\n  ".join(problems))
    except ParseError as e:
        print(f"SYNC ABORTED – {e}", file=sys.stderr)
        return 1
    print(f"parsed {len(events)} events across {len({e['date'] for e in events})} dates")
    if args.check:
        return 0
    old_path = os.path.join(DOCS, "data.json")
    old = json.load(open(old_path))["rows"] if os.path.exists(old_path) else []
    if old and [event_key(e) for e in old] == [event_key(e) for e in events] and all(
            (a["tags"], a["type"], a["note"], a["time"], a.get("details")) == (b["tags"], b["type"], b["note"], b["time"], b["details"])
            for a, b in zip(old, events)):
        print("no changes since last sync – nothing written")
        return 0
    synced_label = datetime.datetime.now(datetime.timezone.utc).astimezone(
        datetime.timezone(datetime.timedelta(hours=-6 if datetime.date.today() > datetime.date(YEAR, 11, 1) else -5))
    ).strftime("%b %-d, %Y %-I:%M %p")
    write_outputs(events, synced_label, guide)
    if append_changes(old, events, synced_label):
        print("changes recorded in CHANGES.md")
    print("site rebuilt")
    return 0


if __name__ == "__main__":
    sys.exit(main())
