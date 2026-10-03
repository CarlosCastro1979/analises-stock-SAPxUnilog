-- Aulas PT — schema for shared household calendars
-- Run once in the Supabase SQL editor (same project as Performance Logística).

-- Settings: one row per household
create table if not exists public.aulas_pt_settings (
  household_code text primary key,
  name_a text not null default 'Pessoa 1',
  name_b text not null default 'Pessoa 2',
  price_cents integer not null default 2500,
  updated_at timestamptz not null default now()
);

-- Sessions: one row per household + person + day (presence = horário feito)
create table if not exists public.aulas_pt_sessions (
  household_code text not null references public.aulas_pt_settings (household_code) on delete cascade,
  person text not null check (person in ('a', 'b')),
  day date not null,
  primary key (household_code, person, day)
);

create index if not exists aulas_pt_sessions_household_day_idx
  on public.aulas_pt_sessions (household_code, day);

-- RLS: anon full access (same pattern as logistica_excel_files)
alter table public.aulas_pt_settings enable row level security;
alter table public.aulas_pt_sessions enable row level security;

drop policy if exists "anon_all" on public.aulas_pt_settings;
create policy "anon_all" on public.aulas_pt_settings
  for all to anon
  using (true)
  with check (true);

drop policy if exists "anon_all" on public.aulas_pt_sessions;
create policy "anon_all" on public.aulas_pt_sessions
  for all to anon
  using (true)
  with check (true);

-- Realtime so both phones update live (idempotent)
do $$
begin
  if not exists (
    select 1 from pg_publication_tables
    where pubname = 'supabase_realtime' and schemaname = 'public' and tablename = 'aulas_pt_settings'
  ) then
    alter publication supabase_realtime add table public.aulas_pt_settings;
  end if;
  if not exists (
    select 1 from pg_publication_tables
    where pubname = 'supabase_realtime' and schemaname = 'public' and tablename = 'aulas_pt_sessions'
  ) then
    alter publication supabase_realtime add table public.aulas_pt_sessions;
  end if;
end $$;
