"""
One-time migration: local SQLite data.db -> Supabase/PostgreSQL.

Usage:
    python migrate_sqlite_to_supabase.py

Prerequisites:
1. Run schema.sql in Supabase SQL Editor.
2. Create .streamlit/secrets.toml with [supabase] url/key.
3. Keep the old data.db in the project folder while running this script.

The script does not copy SQLite IDs. It creates new Supabase match IDs and
maps all goal_events to the correct new match IDs.
"""

from pathlib import Path
import sqlite3
import tomllib
from supabase import create_client

ROOT = Path(__file__).resolve().parent
DB_PATH = ROOT / "data.db"
SECRETS_PATH = ROOT / ".streamlit" / "secrets.toml"


def load_supabase():
    if not SECRETS_PATH.exists():
        raise SystemExit(
            f"Fehlt: {SECRETS_PATH}\n"
            "Bitte zuerst .streamlit/secrets.toml anlegen."
        )

    with SECRETS_PATH.open("rb") as f:
        cfg = tomllib.load(f)

    supabase_cfg = cfg.get("supabase", {})
    url = str(supabase_cfg.get("url", "")).strip()
    key = str(supabase_cfg.get("key", "")).strip()

    if not url or not key:
        raise SystemExit("Supabase url/key fehlen in .streamlit/secrets.toml.")

    return create_client(url, key)


def sqlite_rows(table):
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    rows = [dict(r) for r in con.execute(f"SELECT * FROM {table} ORDER BY id").fetchall()]
    con.close()
    return rows


def table_columns(rows, allowed):
    return [{k: row.get(k) for k in allowed if k in row} for row in rows]


def main():
    if not DB_PATH.exists():
        raise SystemExit(f"Keine lokale Datenbank gefunden: {DB_PATH}")

    client = load_supabase()

    local_matches = sqlite_rows("matches")
    local_events = sqlite_rows("goal_events")

    remote_matches = client.table("matches").select("id").limit(1).execute().data
    remote_events = client.table("goal_events").select("id").limit(1).execute().data

    if remote_matches or remote_events:
        print("WARNUNG: Supabase enthält bereits Daten.")
        answer = input("Trotzdem fortfahren? Es können Duplikate entstehen. [j/N]: ").strip().lower()
        if answer not in {"j", "ja", "y", "yes"}:
            print("Migration abgebrochen.")
            return

    match_allowed = {
        "team", "match_date", "opponent", "competition",
        "home_away", "created_by", "created_at"
    }
    event_allowed = {
        "team", "event_type", "minute", "scorer", "assister",
        "phase", "creation_type",
        "start_x", "start_y", "start_zone",
        "assist_x", "assist_y", "assist_zone",
        "finish_x", "finish_y", "finish_zone",
        "video_url", "comment",
        "click_count", "finish_touch", "set_piece_type",
        "created_by", "created_at"
    }

    id_map = {}

    print(f"Übertrage {len(local_matches)} Spiele ...")
    for match in local_matches:
        old_id = match["id"]
        payload = {k: match.get(k) for k in match_allowed if k in match}
        result = client.table("matches").insert(payload).execute()
        if not result.data:
            raise RuntimeError(f"Spiel {old_id} konnte nicht übertragen werden.")
        id_map[old_id] = result.data[0]["id"]

    print(f"Übertrage {len(local_events)} Ereignisse ...")
    for event in local_events:
        old_match_id = event["match_id"]
        if old_match_id not in id_map:
            raise RuntimeError(
                f"Ereignis {event['id']}: match_id {old_match_id} wurde nicht gefunden."
            )

        payload = {k: event.get(k) for k in event_allowed if k in event}
        payload["match_id"] = id_map[old_match_id]
        client.table("goal_events").insert(payload).execute()

    print("")
    print("Migration abgeschlossen.")
    print(f"Spiele:     {len(local_matches)}")
    print(f"Ereignisse: {len(local_events)}")
    print("")
    print("Danach kannst du die lokale data.db als Backup aufbewahren.")
    print("Sie bleibt durch .gitignore von GitHub ausgeschlossen.")


if __name__ == "__main__":
    main()
