import os, io, csv, uuid
from datetime import datetime, timedelta
from dateutil import tz
import pandas as pd
from flask import Flask, request, render_template_string, send_file, redirect, url_for, flash

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret")

TZ = tz.gettz("Europe/Berlin")

HTML_INDEX = """
<!doctype html>
<html lang="de">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width,initial-scale=1">
    <title>Kabelzieher – Planung</title>
    <style>
      body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;margin:24px;max-width:1100px}
      h1{margin:0 0 16px}
      form{display:grid;gap:12px;grid-template-columns:1fr 1fr}
      label{font-size:14px;color:#444}
      input,select{padding:10px;border:1px solid #ddd;border-radius:8px}
      .row{grid-column:1/-1}
      .card{border:1px solid #eee;border-radius:12px;padding:16px;margin:16px 0;background:#fafafa}
      .note{color:#666;font-size:13px}
      button{padding:12px 16px;border:0;border-radius:10px;background:#111;color:#fff;cursor:pointer}
      table{border-collapse:collapse;width:100%}
      th,td{border-bottom:1px solid #eee;padding:8px 10px;text-align:left}
      .ok{background:#0a7e3a;color:#fff;border-radius:6px;padding:2px 8px;font-size:12px}
    </style>
  </head>
  <body>
    <h1>📅 Kabelzieher – Terminplanung (MVP)</h1>
    {% with messages = get_flashed_messages() %}
      {% if messages %}
        <div class="card">
          {% for m in messages %}<div>{{m}}</div>{% endfor %}
        </div>
      {% endif %}
    {% endwith %}
    <div class="card">
      <form action="{{ url_for('plan') }}" method="post" enctype="multipart/form-data">
        <div class="row">
          <label><b>🗂️ NVT-Liste hochladen (CSV oder Excel)</b></label>
          <input type="file" name="file" required>
          <div class="note">Spalten erwartet: <b>Name, Adresse, Email, Handy, NVT, Team</b> (Team: "Team 1" / "Team 2" oder leer)</div>
        </div>

        <div>
          <label>Startzeit (z. B. 08:00)</label>
          <input type="time" name="start" value="08:00" required>
        </div>
        <div>
          <label>Endzeit (z. B. 17:00)</label>
          <input type="time" name="end" value="17:00" required>
        </div>

        <div>
          <label>Länge pro Termin (Min.)</label>
          <input type="number" name="slot" value="45" min="15" step="5" required>
        </div>
        <div>
          <label>Mittagspause (Start, z. B. 12:00)</label>
          <input type="time" name="lunch_start" value="12:00">
        </div>
        <div>
          <label>Mittagspause (Minuten)</label>
          <input type="number" name="lunch_len" value="45" min="0" step="5">
        </div>

        <div>
          <label>Datum</label>
          <input type="date" name="date" value="{{ today }}" required>
        </div>
        <div>
          <label>Team-Zuordnung</label>
          <select name="team_mode">
            <option value="auto" selected>Automatisch (nach Spalte „Team“ / NVT-Optimierung)</option>
            <option value="team1">Alles Team 1</option>
            <option value="team2">Alles Team 2</option>
            <option value="split">Beide Teams (System teilt fair auf, NVT-Wechsel minimieren)</option>
          </select>
        </div>

        <div class="row">
          <button type="submit">🔧 Planung erstellen</button>
        </div>
      </form>
    </div>

    <div class="note">Status: <span class="ok">läuft</span></div>
  </body>
</html>
"""

HTML_RESULT = """
<!doctype html>
<html lang="de"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Planung – Ergebnis</title>
<style>
 body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;margin:24px;max-width:1100px}
 h1{margin:0 0 16px}
 .grid{display:grid;grid-template-columns:1fr 1fr;gap:24px}
 table{border-collapse:collapse;width:100%}
 th,td{border-bottom:1px solid #eee;padding:8px 10px;text-align:left}
 .card{border:1px solid #eee;border-radius:12px;padding:16px;margin:16px 0;background:#fafafa}
 a.btn{display:inline-block;margin-top:8px;padding:10px 14px;border-radius:10px;background:#111;color:#fff;text-decoration:none}
 .note{color:#666;font-size:13px}
</style></head>
<body>
  <h1>✅ Planung erstellt für {{ date_str }}</h1>
  <div class="grid">
    <div class="card">
      <h3>Team 1</h3>
      <table>
        <thead><tr><th>Uhrzeit</th><th>Name</th><th>NVT</th><th>Adresse</th><th>Email</th><th>Handy</th></tr></thead>
        <tbody>
          {% for r in team1 %}<tr>
            <td>{{ r['time'] }}</td><td>{{ r['Name'] }}</td><td>{{ r['NVT'] }}</td><td>{{ r['Adresse'] }}</td><td>{{ r['Email'] }}</td><td>{{ r['Handy'] }}</td>
          </tr>{% endfor %}
        </tbody>
      </table>
      <a class="btn" href="{{ url_for('download_ics', plan_id=plan_id, team='team1') }}">📥 ICS für Team 1 herunterladen</a>
    </div>
    <div class="card">
      <h3>Team 2</h3>
      <table>
        <thead><tr><th>Uhrzeit</th><th>Name</th><th>NVT</th><th>Adresse</th><th>Email</th><th>Handy</th></tr></thead>
        <tbody>
          {% for r in team2 %}<tr>
            <td>{{ r['time'] }}</td><td>{{ r['Name'] }}</td><td>{{ r['NVT'] }}</td><td>{{ r['Adresse'] }}</td><td>{{ r['Email'] }}</td><td>{{ r['Handy'] }}</td>
          </tr>{% endfor %}
        </tbody>
      </table>
      <a class="btn" href="{{ url_for('download_ics', plan_id=plan_id, team='team2') }}">📥 ICS für Team 2 herunterladen</a>
    </div>
  </div>
  <p class="note">Tipp: ICS-Datei im Browser öffnen → Google Kalender fragt automatisch, in welchen Kalender (Team 1 / Team 2) importiert werden soll.</p>
  <p><a href="{{ url_for('index') }}">← Neue Planung</a></p>
</body></html>
"""

# In-Memory Speicher der letzten Planung
PLANS = {}

def parse_table(file_storage):
    name = file_storage.filename.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(file_storage)
    elif name.endswith(".xlsx") or name.endswith(".xls"):
        df = pd.read_excel(file_storage)
    else:
        # Versuch: CSV trotzdem lesen
        df = pd.read_csv(file_storage)
    # Spalten normalisieren
    rename = {c:c.strip().title() for c in df.columns}
    df = df.rename(columns=rename)
    needed = ["Name","Adresse","Email","Handy","Nvt","Team"]
    for n in needed:
        if n not in df.columns:
            if n=="Team" and "Team" not in df.columns:
                df["Team"] = ""
            elif n=="Nvt" and "NVT" in df.columns:
                df.rename(columns={"NVT":"Nvt"}, inplace=True)
            elif n not in df.columns:
                df[n] = ""
    # Duplikate nach Email/Handy entfernen
    key_cols = []
    if "Email" in df.columns: key_cols.append("Email")
    if "Handy" in df.columns: key_cols.append("Handy")
    if key_cols:
        df = df.drop_duplicates(subset=key_cols, keep="first")
    return df.fillna("")

def time_slots(day: datetime, start_str: str, end_str: str, slot_min: int, lunch_start: str, lunch_len: int):
    s = datetime.combine(day.date(), datetime.strptime(start_str,"%H:%M").time()).replace(tzinfo=TZ)
    e = datetime.combine(day.date(), datetime.strptime(end_str,"%H:%M").time()).replace(tzinfo=TZ)
    lunch_s = datetime.combine(day.date(), datetime.strptime(lunch_start,"%H:%M").time()).replace(tzinfo=TZ) if lunch_start else None
    slots = []
    cur = s
    while cur + timedelta(minutes=slot_min) <= e:
        if lunch_s and cur >= lunch_s and cur < lunch_s + timedelta(minutes=lunch_len):
            cur = lunch_s + timedelta(minutes=lunch_len)
            continue
        slots.append(cur)
        cur += timedelta(minutes=slot_min)
    return slots

def assign(df, slots, mode):
    # Einfache Heuristik:
    # 1) Nach NVT gruppieren → zuerst NVT-Blöcke füllen (minimiert Wechsel)
    # 2) Team-Modus berücksichtigen
    team1, team2 = [], []
    groups = [ (nvt, g.copy()) for nvt, g in df.groupby(df["Nvt"].astype(str)) ]
    # NVTs mit vielen Einträgen zuerst
    groups.sort(key=lambda x: len(x[1]), reverse=True)

    s1 = slots.copy()
    s2 = slots.copy()

    def take_slot(team_slots):
        return team_slots.pop(0) if team_slots else None

    for nvt, g in groups:
        rows = g.to_dict("records")
        # ziel-team
        for r in rows:
            t_pref = str(r.get("Team","")).strip().lower()
            if mode=="team1" or t_pref=="team 1":
                t = "team1"
            elif mode=="team2" or t_pref=="team 2":
                t = "team2"
            else:
                # auto/split: fülle das Team, das mehr freie Slots hat
                t = "team1" if len(s1) >= len(s2) else "team2"

            if t=="team1":
                tslot = take_slot(s1)
                if tslot is None:
                    # fallback: anderes team
                    tslot = take_slot(s2)
                    tgt = team2
                else:
                    tgt = team1
            else:
                tslot = take_slot(s2)
                if tslot is None:
                    tslot = take_slot(s1)
                    tgt = team1
                else:
                    tgt = team2

            if tslot is None:
                # kein Slot mehr – ignoriere (oder späterer Tag)
                continue

            tgt.append({
                "time": tslot.strftime("%H:%M"),
                "dt_start": tslot,
                "Name": r.get("Name",""),
                "Adresse": r.get("Adresse",""),
                "Email": r.get("Email",""),
                "Handy": r.get("Handy",""),
                "NVT": nvt or ""
            })
    # nach Zeit sortieren
    team1.sort(key=lambda x: x["dt_start"])
    team2.sort(key=lambda x: x["dt_start"])
    return team1, team2

def ics_for_team(date_str, events, slot_min):
    # Simple ICS
    lines = ["BEGIN:VCALENDAR","VERSION:2.0","PRODID:-//Kabelzieher//Plan//DE"]
    for ev in events:
        start = ev["dt_start"]
        end = start + timedelta(minutes=slot_min)
        dtstamp = datetime.now(tz=TZ).strftime("%Y%m%dT%H%M%S")
        uid = str(uuid.uuid4())
        def fmt(dt): return dt.strftime("%Y%m%dT%H%M%S")
        summary = f"Glasfaser-Termin: {ev['Name']} – {ev['NVT']}"
        desc = f"Adresse: {ev['Adresse']}\\nEmail: {ev['Email']}\\nHandy: {ev['Handy']}"
        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{dtstamp}",
            f"DTSTART;TZID=Europe/Berlin:{fmt(start)}",
            f"DTEND;TZID=Europe/Berlin:{fmt(end)}",
            f"SUMMARY:{summary}",
            f"DESCRIPTION:{desc}",
            "END:VEVENT"
        ]
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines).encode("utf-8")

@app.get("/")
def index():
    today = datetime.now(tz=TZ).strftime("%Y-%m-%d")
    return render_template_string(HTML_INDEX, today=today)

@app.post("/plan")
def plan():
    f = request.files.get("file")
    if not f: 
        flash("Bitte eine CSV- oder Excel-Datei wählen.")
        return redirect(url_for("index"))
    df = parse_table(f)

    date_str = request.form.get("date")
    day = datetime.strptime(date_str,"%Y-%m-%d").replace(tzinfo=TZ)
    start = request.form.get("start","08:00")
    end = request.form.get("end","17:00")
    slot = int(request.form.get("slot","45"))
    lunch_start = request.form.get("lunch_start","12:00")
    lunch_len = int(request.form.get("lunch_len","45"))
    mode = request.form.get("team_mode","auto")

    slots = time_slots(day, start, end, slot, lunch_start, lunch_len)
    t1, t2 = assign(df, slots, mode)

    plan_id = str(uuid.uuid4())
    PLANS[plan_id] = {"t1": t1, "t2": t2, "slot": slot, "date": day}

    return render_template_string(HTML_RESULT, team1=t1, team2=t2, date_str=day.strftime("%d.%m.%Y"), plan_id=plan_id)

@app.get("/download/<plan_id>/<team>.ics")
def download_ics(plan_id, team):
    plan = PLANS.get(plan_id)
    if not plan:
        flash("Plan nicht gefunden.")
        return redirect(url_for("index"))
    events = plan["t1"] if team=="team1" else plan["t2"]
    data = ics_for_team(plan["date"].strftime("%Y-%m-%d"), events, plan["slot"])
    filename = f"{team}_{plan['date'].strftime('%Y-%m-%d')}.ics"
    return send_file(io.BytesIO(data), as_attachment=True, download_name=filename, mimetype="text/calendar")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
