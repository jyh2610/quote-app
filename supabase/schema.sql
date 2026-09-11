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

-- No login system: every visitor uses the same public anon key, so RLS is
-- left open (anyone with the key can read/write/delete any row). That's the
-- accepted tradeoff for device-to-device sync without accounts. If this
-- table ever holds sensitive customer data, add auth + narrower policies.
alter table public.quotes enable row level security;

create policy "public can read quotes" on public.quotes
  for select using (true);

create policy "public can insert quotes" on public.quotes
  for insert with check (true);

create policy "public can update quotes" on public.quotes
  for update using (true);

create policy "public can delete quotes" on public.quotes
  for delete using (true);

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
