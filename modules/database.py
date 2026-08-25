
import sqlite3
from pathlib import Path
from datetime import datetime
import streamlit as st

DB_PATH = Path(__file__).resolve().parents[1] / "data.db"


def _use_supabase():
    try:
        return bool(st.secrets.get("supabase", {}).get("url")) and bool(st.secrets.get("supabase", {}).get("key"))
    except Exception:
        return False


def _supabase():
    from supabase import create_client
    return create_client(st.secrets["supabase"]["url"], st.secrets["supabase"]["key"])


def init_db():
    if _use_supabase():
        return
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.executescript("""
    CREATE TABLE IF NOT EXISTS teams (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL
    );
    CREATE TABLE IF NOT EXISTS matches (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        team TEXT NOT NULL,
        match_date TEXT NOT NULL,
        opponent TEXT NOT NULL,
        competition TEXT,
        home_away TEXT,
        created_by TEXT,
        created_at TEXT NOT NULL,
        click_count INTEGER DEFAULT 3,
        finish_touch TEXT,
        set_piece_type TEXT
    );
    CREATE TABLE IF NOT EXISTS app_meta (
        key TEXT PRIMARY KEY,
        value TEXT
    );

    CREATE TABLE IF NOT EXISTS goal_events (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        match_id INTEGER NOT NULL,
        team TEXT NOT NULL,
        event_type TEXT NOT NULL,
        minute INTEGER,
        scorer TEXT,
        assister TEXT,
        phase TEXT,
        creation_type TEXT,
        start_x REAL,
        start_y REAL,
        start_zone TEXT,
        assist_x REAL,
        assist_y REAL,
        assist_zone TEXT,
        finish_x REAL,
        finish_y REAL,
        finish_zone TEXT,
        video_url TEXT,
        comment TEXT,
        created_by TEXT,
        created_at TEXT NOT NULL
    );
    """)

    try:
        cur.execute("ALTER TABLE goal_events ADD COLUMN click_count INTEGER DEFAULT 3")
    except sqlite3.OperationalError:
        pass

    try:
        cur.execute("ALTER TABLE goal_events ADD COLUMN finish_touch TEXT")
    except sqlite3.OperationalError:
        pass

    try:
        cur.execute("ALTER TABLE goal_events ADD COLUMN set_piece_type TEXT")
    except sqlite3.OperationalError:
        pass

    migrated = cur.execute("SELECT value FROM app_meta WHERE key='coord_system_v2'").fetchone()
    if not migrated:
        rows = cur.execute("""
            SELECT id, start_x, start_y, assist_x, assist_y, finish_x, finish_y
            FROM goal_events
        """).fetchall()

        def _old_to_pitch_x(v):
            if v is None:
                return None
            return max(0.0, min(100.0, (float(v) - 3.5) / 93.0 * 100.0))

        def _old_to_pitch_y(v):
            if v is None:
                return None
            top = 35.0 / 650.0 * 100.0
            height = 580.0 / 650.0 * 100.0
            return max(0.0, min(100.0, (float(v) - top) / height * 100.0))

        for row in rows:
            event_id, sx, sy, ax, ay, fx, fy = row
            cur.execute("""
                UPDATE goal_events
                SET start_x=?, start_y=?, assist_x=?, assist_y=?, finish_x=?, finish_y=?
                WHERE id=?
            """, (
                _old_to_pitch_x(sx), _old_to_pitch_y(sy),
                _old_to_pitch_x(ax), _old_to_pitch_y(ay),
                _old_to_pitch_x(fx), _old_to_pitch_y(fy),
                event_id
            ))

        cur.execute(
            "INSERT OR REPLACE INTO app_meta(key, value) VALUES('coord_system_v2', 'pitch_normalized')"
        )

    for team in ("U15", "U16", "U18", "JWR"):
        cur.execute("INSERT OR IGNORE INTO teams(name) VALUES (?)", (team,))
    con.commit()
    con.close()


def create_match(data):
    row = {**data, "created_at": datetime.utcnow().isoformat(timespec="seconds")}
    if _use_supabase():
        return _supabase().table("matches").insert(row).execute().data[0]
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""INSERT INTO matches(team, match_date, opponent, competition, home_away, created_by, created_at)
                   VALUES (?,?,?,?,?,?,?)""",
                (row["team"], row["match_date"], row["opponent"], row.get("competition"),
                 row.get("home_away"), row.get("created_by"), row["created_at"]))
    row["id"] = cur.lastrowid
    con.commit(); con.close()
    return row


def get_matches(team=None):
    if _use_supabase():
        q = _supabase().table("matches").select("*").order("match_date", desc=True)
        if team:
            q = q.eq("team", team)
        return q.execute().data
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    if team:
        rows = con.execute("SELECT * FROM matches WHERE team=? ORDER BY match_date DESC, id DESC", (team,)).fetchall()
    else:
        rows = con.execute("SELECT * FROM matches ORDER BY match_date DESC, id DESC").fetchall()
    con.close()
    return [dict(r) for r in rows]


def insert_event(data):
    row = {**data, "created_at": datetime.utcnow().isoformat(timespec="seconds")}
    if _use_supabase():
        return _supabase().table("goal_events").insert(row).execute().data[0]
    cols = list(row.keys())
    vals = [row[c] for c in cols]
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute(
        f"INSERT INTO goal_events ({','.join(cols)}) VALUES ({','.join(['?']*len(cols))})",
        vals
    )
    row["id"] = cur.lastrowid
    con.commit(); con.close()
    return row


def get_events(team=None, match_id=None):
    if _use_supabase():
        q = _supabase().table("goal_events").select("*").order("created_at", desc=False)
        if team:
            q = q.eq("team", team)
        if match_id:
            q = q.eq("match_id", match_id)
        return q.execute().data
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    sql, params = "SELECT * FROM goal_events WHERE 1=1", []
    if team:
        sql += " AND team=?"; params.append(team)
    if match_id:
        sql += " AND match_id=?"; params.append(match_id)
    sql += " ORDER BY created_at ASC"
    rows = con.execute(sql, params).fetchall()
    con.close()
    return [dict(r) for r in rows]


def delete_event(event_id):
    if _use_supabase():
        _supabase().table("goal_events").delete().eq("id", event_id).execute()
        return True

    con = sqlite3.connect(DB_PATH)
    con.execute("DELETE FROM goal_events WHERE id=?", (event_id,))
    con.commit()
    con.close()
    return True


def update_event(event_id, data):
    """Update an existing goal event in SQLite or Supabase."""
    if _use_supabase():
        payload = dict(data)
        return _supabase().table("goal_events").update(payload).eq("id", event_id).execute().data

    if not data:
        return True

    cols = list(data.keys())
    values = [data[c] for c in cols]
    assignments = ", ".join(f"{c}=?" for c in cols)

    con = sqlite3.connect(DB_PATH)
    con.execute(
        f"UPDATE goal_events SET {assignments} WHERE id=?",
        values + [event_id]
    )
    con.commit()
    con.close()
    return True


def delete_match(match_id):
    """Delete a match and all related goal events."""
    if _use_supabase():
        client = _supabase()
        client.table("goal_events").delete().eq("match_id", match_id).execute()
        client.table("matches").delete().eq("id", match_id).execute()
        return True

    con = sqlite3.connect(DB_PATH)
    con.execute("DELETE FROM goal_events WHERE match_id=?", (match_id,))
    con.execute("DELETE FROM matches WHERE id=?", (match_id,))
    con.commit()
    con.close()
    return True
