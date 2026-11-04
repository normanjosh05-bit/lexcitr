-- Run once in Supabase: SQL Editor > New query > paste > Run.
create table if not exists public.records (
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  id text not null,
  kind text not null check (kind in ('style','type','citation','library','pinpoint','document')),
  data jsonb not null,
  updated_at timestamptz not null default now(),
  primary key (user_id, id)
);
alter table public.records enable row level security;
drop policy if exists "own records only" on public.records;
create policy "own records only" on public.records
  for all to authenticated
  using (auth.uid() = user_id)
  with check (auth.uid() = user_id);
grant select, insert, update, delete on public.records to authenticated;
