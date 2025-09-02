import re
import uuid
import html
import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta, timezone

# ----------------------------
# Réglages
# ----------------------------
CODE_DEPARTEMENT = "L"  # code pour le MS PAPDD
DATE_DEBUT = "02/09/2025"  # inclus
DATE_FIN   = "24/12/2025"  # inclus
CAL_NAME   = f"ENPC - {CODE_DEPARTEMENT}"
ICS_OUTPUT = f"enpc_{CODE_DEPARTEMENT}_{DATE_DEBUT.replace('/','-')}_to_{DATE_FIN.replace('/','-')}.ics"

#Url de base de l'emploi du temps des ponts
BASE_URL = "https://emploidutemps.enpc.fr/?code_departement={code}&mydate={date}" 

# ----------------------------
# Outils date & fuseau
# ----------------------------

def iter_dates(d1_str: str, d2_str: str):
    d1 = datetime.strptime(d1_str, "%d/%m/%Y").date()
    d2 = datetime.strptime(d2_str, "%d/%m/%Y").date()
    day = d1
    while day <= d2:
        yield day
        day += timedelta(days=1)

# ----------------------------
# Récupération + parsing HTML
# ----------------------------
TIME_RANGE_RE = re.compile(r"^\s*(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})\s*$")
today = datetime.today().strftime('%d/%m/%Y')

def fetch_html_for_date(day):
    date_str = day.strftime("%d/%m/%Y")
    url = BASE_URL.format(code=CODE_DEPARTEMENT, date=requests.utils.quote(date_str, safe=""))
    resp = requests.get(url, timeout=20)
    resp.raise_for_status()
    return resp.text, url

def parse_events(html_text: str, day):
    """
    Retourne une liste de dicts:
    [{
      'start': datetime,
      'end': datetime,
      'location': str,
      'summary': str,
      'department': str,
      'raw_title': str,
      'source_date': str (dd/mm/yyyy),
    }, ...]
    """
    soup = BeautifulSoup(html_text, "html.parser")
    events = []

    for tr in soup.find_all("tr"):
        tds = tr.find_all("td")
        if len(tds) < 5:
            continue

        time_text = tds[0].get_text(strip=True)
        m = TIME_RANGE_RE.match(time_text)
        if not m:
            continue

        start_hm, end_hm = m.groups()
        dep = tds[1].get_text(strip=True) or ""
        location = tds[2].get_text(strip=True) or ""
        last_col_html = str(tds[4])
        last_col = tds[4].get_text(" ", strip=True)
        last_col = html.unescape(last_col)
        if last_col.startswith("\xa0"):
            last_col = last_col.lstrip("\xa0").strip()
        if last_col.startswith("&nbsp;"):
            last_col = last_col.replace("&nbsp;", "").strip()

        d = day
        start_dt = datetime.strptime(f"{d.strftime('%d/%m/%Y')} {start_hm}", "%d/%m/%Y %H:%M")
        end_dt   = datetime.strptime(f"{d.strftime('%d/%m/%Y')} {end_hm}", "%d/%m/%Y %H:%M")

        summary = last_col if last_col else (dep or "Cours")

        events.append({
            "start": start_dt,
            "end": end_dt,
            "location": location,
            "summary": summary,
            "department": dep,
            "raw_title": last_col_html,
            "source_date": d.strftime("%d/%m/%Y"),
        })

    return events

# ----------------------------
# Génération ICS
# ----------------------------

def ics_escape(text: str) -> str:
    """Échappement selon RFC 5545 pour TEXT."""
    return (text or "").replace("\\", "\\\\").replace("\n", "\\n").replace(",", "\\,").replace(";", "\\;")


def fold_ical_line(line: str) -> str:
    """
    Pliage des lignes >75 octets (CRLF + espace), en évitant de couper un
    caractère UTF-8 au milieu.
    """
    b = line.encode("utf-8")
    if len(b) <= 75:
        return line

    out_chunks = []
    i = 0
    first = True
    while i < len(b):
        j = min(i + 75, len(b))

        while j > i and (b[j-1] & 0xC0) == 0x80:
            j -= 1

        if j == i:
            j = min(i + 75, len(b))

        chunk = b[i:j].decode("utf-8")
        if first:
            out_chunks.append(chunk)
            first = False
        else:
            out_chunks.append(" " + chunk)  # espace initial = continuation
        i = j

    return "\r\n".join(out_chunks)


def format_dt_local_with_tzid(dt: datetime, tzid="Europe/Paris") -> str:
    return f";TZID={tzid}:{dt.strftime('%Y%m%dT%H%M%S')}"


def vtimezone_europe_paris_lines():
    """Bloc VTIMEZONE minimal pour Europe/Paris (CET/CEST) en liste de lignes."""
    return [
        "BEGIN:VTIMEZONE",
        "TZID:Europe/Paris",
        "X-LIC-LOCATION:Europe/Paris",
        "BEGIN:DAYLIGHT",
        "TZOFFSETFROM:+0100",
        "TZOFFSETTO:+0200",
        "TZNAME:CEST",
        "DTSTART:19700329T020000",
        "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU",
        "END:DAYLIGHT",
        "BEGIN:STANDARD",
        "TZOFFSETFROM:+0200",
        "TZOFFSETTO:+0100",
        "TZNAME:CET",
        "DTSTART:19701025T030000",
        "RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU",
        "END:STANDARD",
        "END:VTIMEZONE",
    ]


def make_vevent_lines(ev: dict):
    uid = str(uuid.uuid5(uuid.NAMESPACE_URL, f"{ev['source_date']}-{ev['summary']}-{ev['location']}-{ev['start']}-{ev['end']}"))
    dtstamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")

    lines = []
    lines.append("BEGIN:VEVENT")
    lines.append(f"UID:{uid}")
    lines.append(f"DTSTAMP:{dtstamp}")
    lines.append(f"DTSTART{format_dt_local_with_tzid(ev['start'])}")
    lines.append(f"DTEND{format_dt_local_with_tzid(ev['end'])}")

    if ev.get("location"):
        lines.append(fold_ical_line(f"LOCATION:{ics_escape(ev['location'])}"))

    summary = ev["summary"]
    if ev.get("department") and ev["department"] and ev["department"] not in summary:
        summary = f"[{ev['department']}] {summary}"
    lines.append(fold_ical_line(f"SUMMARY:{ics_escape(summary)}"))

    desc = "" #f"Source: emploi du temps ENPC (téléchargé le {today})."
    lines.append(fold_ical_line(f"DESCRIPTION:{ics_escape(desc)}"))
    lines.append("END:VEVENT")
    return lines


def build_ics(events: list) -> str:
    lines = [
        "BEGIN:VCALENDAR",
        "PRODID:-//ENPC Parser//FR//",
        "VERSION:2.0",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        fold_ical_line(f"X-WR-CALNAME:{ics_escape(CAL_NAME)}"),
        "X-WR-TIMEZONE:Europe/Paris",
    ]

    # VTIMEZONE
    lines.extend(vtimezone_europe_paris_lines())

    # VEVENTs
    for ev in events:
        lines.extend(make_vevent_lines(ev))

    lines.append("END:VCALENDAR")

    return "\r\n".join(lines)

# ----------------------------
# Exécution
# ----------------------------


def main():
    all_events = []
    for day in iter_dates(DATE_DEBUT, DATE_FIN):
        html_text, url = fetch_html_for_date(day)
        day_events = parse_events(html_text, day)
        all_events.extend(day_events)

    if not all_events:
        print("Aucun évènement trouvé sur la période.")
        return

    ics_text = build_ics(all_events)

    with open(ICS_OUTPUT, "w", encoding="utf-8", newline="") as f:
        f.write(ics_text)

    print(f"OK : {len(all_events)} évènement(s) exporté(s) → {ICS_OUTPUT}")


if __name__ == "__main__":
    main()
