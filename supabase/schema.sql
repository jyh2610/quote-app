-- Run this once in the Supabase SQL Editor (Project -> SQL Editor -> New query)
-- to create the table the quote app saves/loads quotes from.

create table if not exists public.quotes (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  header jsonb not null,
  groups jsonb not null,
  freight text not null default '1000',
  margin_rate text not null default '15',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

-- Gated by login (Supabase Auth), not per-user ownership: any signed-in
-- account can read/write/delete any row, so the whole team shares one quote
-- list. Accounts are created directly in the dashboard (Authentication ->
-- Users -> Add user) — there's no sign-up form in the app.
alter table public.quotes enable row level security;

drop policy if exists "public can read quotes" on public.quotes;
drop policy if exists "public can insert quotes" on public.quotes;
drop policy if exists "public can update quotes" on public.quotes;
drop policy if exists "public can delete quotes" on public.quotes;

drop policy if exists "authenticated can read quotes" on public.quotes;
create policy "authenticated can read quotes" on public.quotes
  for select using (auth.uid() is not null);

drop policy if exists "authenticated can insert quotes" on public.quotes;
create policy "authenticated can insert quotes" on public.quotes
  for insert with check (auth.uid() is not null);

drop policy if exists "authenticated can update quotes" on public.quotes;
create policy "authenticated can update quotes" on public.quotes
  for update using (auth.uid() is not null);

drop policy if exists "authenticated can delete quotes" on public.quotes;
create policy "authenticated can delete quotes" on public.quotes
  for delete using (auth.uid() is not null);

-- Keep updated_at current on every edit.
create or replace function public.set_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

drop trigger if exists quotes_set_updated_at on public.quotes;
create trigger quotes_set_updated_at
  before update on public.quotes
  for each row execute function public.set_updated_at();
