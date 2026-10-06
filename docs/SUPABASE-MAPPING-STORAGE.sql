-- Jalankan sekali di Supabase SQL Editor project Artha Megah.
insert into storage.buckets (id, name, public)
values ('mapping-files', 'mapping-files', false)
on conflict (id) do update set public = excluded.public;

do $$
begin
  if not exists (select 1 from pg_policies where schemaname = 'storage' and tablename = 'objects' and policyname = 'Mapping files select own') then
    create policy "Mapping files select own" on storage.objects for select to authenticated
      using (bucket_id = 'mapping-files' and (storage.foldername(name))[1] = (select auth.uid()::text));
  end if;
  if not exists (select 1 from pg_policies where schemaname = 'storage' and tablename = 'objects' and policyname = 'Mapping files insert own') then
    create policy "Mapping files insert own" on storage.objects for insert to authenticated
      with check (bucket_id = 'mapping-files' and (storage.foldername(name))[1] = (select auth.uid()::text));
  end if;
  if not exists (select 1 from pg_policies where schemaname = 'storage' and tablename = 'objects' and policyname = 'Mapping files update own') then
    create policy "Mapping files update own" on storage.objects for update to authenticated
      using (bucket_id = 'mapping-files' and (storage.foldername(name))[1] = (select auth.uid()::text))
      with check (bucket_id = 'mapping-files' and (storage.foldername(name))[1] = (select auth.uid()::text));
  end if;
  if not exists (select 1 from pg_policies where schemaname = 'storage' and tablename = 'objects' and policyname = 'Mapping files delete own') then
    create policy "Mapping files delete own" on storage.objects for delete to authenticated
      using (bucket_id = 'mapping-files' and (storage.foldername(name))[1] = (select auth.uid()::text));
  end if;
end $$;
