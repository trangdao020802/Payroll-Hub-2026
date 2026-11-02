create table if not exists public.promotion_workspaces (
  owner_id uuid not null references auth.users(id) on delete cascade,
  workspace_key text not null default 'default',
  records jsonb not null default '[]'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint promotion_workspaces_pkey primary key (owner_id, workspace_key),
  constraint promotion_workspaces_records_array check (jsonb_typeof(records) = 'array')
);

alter table public.promotion_workspaces enable row level security;
alter table public.promotion_workspaces force row level security;

revoke all on table public.promotion_workspaces from public, anon;
grant select, insert, update, delete on table public.promotion_workspaces to authenticated;

create policy promotion_workspaces_select_own
  on public.promotion_workspaces for select to authenticated
  using (owner_id = (select auth.uid())
    and coalesce((select auth.jwt() ->> 'is_anonymous'), 'false') = 'false');

create policy promotion_workspaces_insert_own
  on public.promotion_workspaces for insert to authenticated
  with check (owner_id = (select auth.uid())
    and coalesce((select auth.jwt() ->> 'is_anonymous'), 'false') = 'false');

create policy promotion_workspaces_update_own
  on public.promotion_workspaces for update to authenticated
  using (owner_id = (select auth.uid())
    and coalesce((select auth.jwt() ->> 'is_anonymous'), 'false') = 'false')
  with check (owner_id = (select auth.uid())
    and coalesce((select auth.jwt() ->> 'is_anonymous'), 'false') = 'false');

create policy promotion_workspaces_delete_own
  on public.promotion_workspaces for delete to authenticated
  using (owner_id = (select auth.uid())
    and coalesce((select auth.jwt() ->> 'is_anonymous'), 'false') = 'false');
