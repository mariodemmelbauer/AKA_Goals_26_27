-- v42: Profis als gültiges Team in der bestehenden Supabase-Datenbank ergänzen
-- In Supabase -> SQL Editor ausführen.

begin;

alter table public.matches
  drop constraint if exists matches_team_check;

alter table public.matches
  add constraint matches_team_check
  check (team in ('U15','U16','U18','JWR','Profis'));

alter table public.goal_events
  drop constraint if exists goal_events_team_check;

alter table public.goal_events
  add constraint goal_events_team_check
  check (team in ('U15','U16','U18','JWR','Profis'));

commit;
