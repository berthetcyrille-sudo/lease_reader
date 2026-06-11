-- Migration : création de la table extractions
-- À exécuter dans Supabase > SQL Editor

create table if not exists public.extractions (
  id          uuid primary key default gen_random_uuid(),
  created_at  timestamptz not null default now(),
  file_name   text not null,
  data        jsonb not null
);

-- Index pour trier par date
create index if not exists extractions_created_at_idx
  on public.extractions (created_at desc);

-- RLS : activé mais politique ouverte pour l'instant (pas d'auth)
-- Ajuster quand l'auth sera activée
alter table public.extractions enable row level security;

create policy "allow_all_for_now"
  on public.extractions
  for all
  using (true)
  with check (true);
