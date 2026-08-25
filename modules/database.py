import sqlite3
from pathlib import Path
from datetime import datetime, timezone
import streamlit as st

DB_PATH = Path(__file__).resolve().parents[1] / "data.db"
TEAMS = ("U15", "U16", "U18", "JWR")


def _now_iso():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _supabase_config():
    """
    Return Supabase config if both URL and server-side key are configured.
    On Streamlit Community Cloud these values belong in App -> Settings -> Secrets.
    """
    try:
        cfg = st.secrets.get("supabase", {})
        url = str(cfg.get("url", "")).strip()
        key = str(cfg.get("key", "")).strip()
        if url and key:
            return {"url": url, "key": key}
    except Exception:
        pass
    return None


def using_supabase():
    return _supabase_config() is not None


@st.cache_resource
def _supabase_client(url, key):
    from supabase import create_client
    return create_client(url, key)


def _supabase():
    cfg = _supabase_config()
    if not cfg:
        raise RuntimeError("Supabase ist nicht konfiguriert.")
    return _supabase_client(cfg["url"], cfg["key"])


def backend_name():
    return "Supabase / PostgreSQL" if using_supabase() else "SQLite (lokal)"


def init_db():
    """
    Supabase schema is created once via schema.sql in the Supabase SQL Editor.
    SQLite is retained only as a local development fallback.
    """
    if using_supabase():
        # Lightweight connectivity check. Fails early with a readable error.
        try:
            _supabase().table("matches").select("id").limit(1).execute()
        except Exception as exc:
            raise RuntimeError(
                "Supabase ist konfiguriert, aber nicht erreichbar bzw. schema.sql wurde "
                "noch nicht ausgeführt. Bitte Supabase URL/Key und Tabellen prüfen."
            ) from exc
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
        created_at TEXT NOT NULL
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
        click_count INTEGER DEFAULT 3,
        finish_touch TEXT,
        set_piece_type TEXT,
        created_by TEXT,
        created_at TEXT NOT NULL,
        FOREIGN KEY(match_id) REFERENCES matches(id) ON DELETE CASCADE
    );
    """)

    # Upgrade older local databases safely.
    for sql in (
        "ALTER TABLE goal_events ADD COLUMN click_count INTEGER DEFAULT 3",
        "ALTER TABLE goal_events ADD COLUMN finish_touch TEXT",
        "ALTER TABLE goal_events ADD COLUMN set_piece_type TEXT",
    ):
        try:
            cur.execute(sql)
        except sqlite3.OperationalError:
            pass

    # v23 coordinate migration: old values were normalized to the whole image.
    migrated = cur.execute(
        "SELECT value FROM app_meta WHERE key='coord_system_v2'"
    ).fetchone()

    if not migrated:
        rows = cur.execute("""
            SELECT id, start_x, start_y, assist_x, assist_y, finish_x, finish_y
            FROM goal_events
        """).fetchall()

        def old_to_pitch_x(value):
            if value is None:
                return None
            return max(0.0, min(100.0, (float(value) - 3.5) / 93.0 * 100.0))

        def old_to_pitch_y(value):
            if value is None:
                return None
            top = 35.0 / 650.0 * 100.0
            height = 580.0 / 650.0 * 100.0
            return max(0.0, min(100.0, (float(value) - top) / height * 100.0))

        for event_id, sx, sy, ax, ay, fx, fy in rows:
            cur.execute("""
                UPDATE goal_events
                SET start_x=?, start_y=?, assist_x=?, assist_y=?, finish_x=?, finish_y=?
                WHERE id=?
            """, (
                old_to_pitch_x(sx), old_to_pitch_y(sy),
                old_to_pitch_x(ax), old_to_pitch_y(ay),
                old_to_pitch_x(fx), old_to_pitch_y(fy),
                event_id,
            ))

        cur.execute(
            "INSERT OR REPLACE INTO app_meta(key, value) VALUES('coord_system_v2','pitch_normalized')"
        )

    for team in TEAMS:
        cur.execute("INSERT OR IGNORE INTO teams(name) VALUES (?)", (team,))

    con.commit()
    con.close()


def create_match(data):
    row = {**data, "created_at": _now_iso()}

    if using_supabase():
        result = _supabase().table("matches").insert(row).execute()
        return result.data[0]

    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute(
        """INSERT INTO matches(
            team, match_date, opponent, competition, home_away, created_by, created_at
        ) VALUES (?,?,?,?,?,?,?)""",
        (
            row["team"], row["match_date"], row["opponent"],
            row.get("competition"), row.get("home_away"),
            row.get("created_by"), row["created_at"],
        ),
    )
    row["id"] = cur.lastrowid
    con.commit()
    con.close()
    return row


def get_matches(team=None):
    if using_supabase():
        q = _supabase().table("matches").select("*").order(
            "match_date", desc=True
        ).order("id", desc=True)
        if team:
            q = q.eq("team", team)
        return q.execute().data

    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    if team:
        rows = con.execute(
            "SELECT * FROM matches WHERE team=? ORDER BY match_date DESC, id DESC",
            (team,),
        ).fetchall()
    else:
        rows = con.execute(
            "SELECT * FROM matches ORDER BY match_date DESC, id DESC"
        ).fetchall()
    con.close()
    return [dict(r) for r in rows]


def insert_event(data):
    row = {**data, "created_at": _now_iso()}

    if using_supabase():
        result = _supabase().table("goal_events").insert(row).execute()
        return result.data[0]

    cols = list(row.keys())
    vals = [row[c] for c in cols]
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute(
        f"INSERT INTO goal_events ({','.join(cols)}) "
        f"VALUES ({','.join(['?'] * len(cols))})",
        vals,
    )
    row["id"] = cur.lastrowid
    con.commit()
    con.close()
    return row


def get_events(team=None, match_id=None):
    if using_supabase():
        q = _supabase().table("goal_events").select("*").order(
            "created_at", desc=False
        )
        if team:
            q = q.eq("team", team)
        if match_id is not None:
            q = q.eq("match_id", match_id)
        return q.execute().data

    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    sql = "SELECT * FROM goal_events WHERE 1=1"
    params = []

    if team:
        sql += " AND team=?"
        params.append(team)
    if match_id is not None:
        sql += " AND match_id=?"
        params.append(match_id)

    sql += " ORDER BY created_at ASC"
    rows = con.execute(sql, params).fetchall()
    con.close()
    return [dict(r) for r in rows]


def delete_event(event_id):
    if using_supabase():
        _supabase().table("goal_events").delete().eq("id", event_id).execute()
        return True

    con = sqlite3.connect(DB_PATH)
    con.execute("DELETE FROM goal_events WHERE id=?", (event_id,))
    con.commit()
    con.close()
    return True


def update_event(event_id, data):
    if not data:
        return True

    if using_supabase():
        return _supabase().table("goal_events").update(
            dict(data)
        ).eq("id", event_id).execute().data

    cols = list(data.keys())
    values = [data[c] for c in cols]
    assignments = ", ".join(f"{c}=?" for c in cols)

    con = sqlite3.connect(DB_PATH)
    con.execute(
        f"UPDATE goal_events SET {assignments} WHERE id=?",
        values + [event_id],
    )
    con.commit()
    con.close()
    return True


def delete_match(match_id):
    if using_supabase():
        # goal_events are removed automatically by FK ON DELETE CASCADE.
        _supabase().table("matches").delete().eq("id", match_id).execute()
        return True

    con = sqlite3.connect(DB_PATH)
    con.execute("PRAGMA foreign_keys = ON")
    con.execute("DELETE FROM goal_events WHERE match_id=?", (match_id,))
    con.execute("DELETE FROM matches WHERE id=?", (match_id,))
    con.commit()
    con.close()
    return True
