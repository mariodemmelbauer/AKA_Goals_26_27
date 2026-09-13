-- v50: Direkter Freistoß als gültige Standardsituation ergänzen
-- Einmal in Supabase -> SQL Editor ausführen.

begin;

alter table public.goal_events
  drop constraint if exists goal_events_set_piece_type_check;

alter table public.goal_events
  add constraint goal_events_set_piece_type_check
  check (
    set_piece_type is null
    or set_piece_type in (
      'Eckball links',
      'Eckball rechts',
      'Direkter Freistoß',
      'Elfmeter'
    )
  );

commit;
