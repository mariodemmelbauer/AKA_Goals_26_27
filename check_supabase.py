"""Quick local Supabase connection/schema check."""

from pathlib import Path
import tomllib
from supabase import create_client

ROOT = Path(__file__).resolve().parent
SECRETS = ROOT / ".streamlit" / "secrets.toml"

if not SECRETS.exists():
    raise SystemExit("Fehlt: .streamlit/secrets.toml")

with SECRETS.open("rb") as f:
    cfg = tomllib.load(f)["supabase"]

client = create_client(cfg["url"], cfg["key"])

matches = client.table("matches").select("id", count="exact").limit(1).execute()
events = client.table("goal_events").select("id", count="exact").limit(1).execute()

print("Supabase-Verbindung OK")
print("matches:", matches.count)
print("goal_events:", events.count)
