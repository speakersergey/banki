-- Коллекция банок: таблица, хранилище фото и правила доступа.
-- Каждый пользователь видит и меняет только свои банки и свои фото.

create table if not exists public.cans (
  id          text primary key default gen_random_uuid()::text,
  user_id     uuid not null default auth.uid() references auth.users on delete cascade,
  num         integer,
  name        text,
  country     text,
  notes       text,
  flavor      text,
  event       text,
  volume      text,
  height      text,
  date        text,
  photo_path  text,
  thumb_path  text,
  created_at  timestamptz not null default now(),
  updated_at  timestamptz not null default now()
);
create index if not exists cans_user_idx on public.cans (user_id);

alter table public.cans enable row level security;

drop policy if exists "cans: read own"   on public.cans;
drop policy if exists "cans: add own"    on public.cans;
drop policy if exists "cans: change own" on public.cans;
drop policy if exists "cans: delete own" on public.cans;
create policy "cans: read own"   on public.cans for select to authenticated using (user_id = auth.uid());
create policy "cans: add own"    on public.cans for insert to authenticated with check (user_id = auth.uid());
create policy "cans: change own" on public.cans for update to authenticated using (user_id = auth.uid()) with check (user_id = auth.uid());
create policy "cans: delete own" on public.cans for delete to authenticated using (user_id = auth.uid());

-- Фото лежат в закрытом хранилище, в папке с id пользователя: <user_id>/<файл>.jpg
insert into storage.buckets (id, name, public)
values ('can-photos', 'can-photos', false)
on conflict (id) do nothing;

drop policy if exists "photos: read own"   on storage.objects;
drop policy if exists "photos: add own"    on storage.objects;
drop policy if exists "photos: change own" on storage.objects;
drop policy if exists "photos: delete own" on storage.objects;
create policy "photos: read own"   on storage.objects for select to authenticated
  using (bucket_id = 'can-photos' and (storage.foldername(name))[1] = auth.uid()::text);
create policy "photos: add own"    on storage.objects for insert to authenticated
  with check (bucket_id = 'can-photos' and (storage.foldername(name))[1] = auth.uid()::text);
create policy "photos: change own" on storage.objects for update to authenticated
  using (bucket_id = 'can-photos' and (storage.foldername(name))[1] = auth.uid()::text);
create policy "photos: delete own" on storage.objects for delete to authenticated
  using (bucket_id = 'can-photos' and (storage.foldername(name))[1] = auth.uid()::text);
