import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';
import test from 'node:test';
import { PGlite } from '@electric-sql/pglite';

test('the complete migration chain replays once, preserves existing roster data and isolates shared-project workspaces', async () => {
  const db = new PGlite();
  const owner = '00000000-0000-0000-0000-000000000001';
  const other = '00000000-0000-0000-0000-000000000002';
  try {
    // Supabase-managed schemas and the legacy roster predate tracked migrations.
    await db.exec(`create role anon; create role authenticated;
      create schema auth; create schema storage; create schema supabase_migrations;
      create table supabase_migrations.schema_migrations(version text primary key);
      create table auth.users(id uuid primary key);
      create function auth.uid() returns uuid language sql stable as
        $$ select nullif(current_setting('request.jwt.claim.sub',true),'')::uuid $$;
      create function auth.jwt() returns jsonb language sql stable as
        $$ select coalesce(nullif(current_setting('request.jwt.claims',true),''),'{}')::jsonb $$;
      create table storage.buckets(id text primary key,name text,public boolean,file_size_limit bigint,allowed_mime_types text[]);
      create table storage.objects(bucket_id text,name text,metadata jsonb,created_at timestamptz default now());
      alter table storage.objects enable row level security;
      create function storage.foldername(name text) returns text[] language sql immutable as $$ select string_to_array(name,'/') $$;
      grant usage on schema auth,storage,public to anon,authenticated;
      grant select,insert,update,delete on storage.objects to authenticated;
      create table public.roster_cham_cong(id integer primary key, employee_id text);
      insert into public.roster_cham_cong values (7,'SYNTHETIC');
      insert into auth.users values ('${owner}'),('${other}');`);
    const directory = new URL('../supabase/migrations/', import.meta.url);
    const files = readdirSync(directory).filter(name => /^\d{14}_.+\.sql$/.test(name)).sort();
    const versions = files.map(name => name.split('_')[0]);
    assert.equal(new Set(versions).size, files.length, 'migration versions must be unique');
    for (const file of files) {
      await db.exec(readFileSync(new URL(file, directory), 'utf8'));
      await db.query('insert into supabase_migrations.schema_migrations values ($1)', [file.split('_')[0]]);
    }
    assert.deepEqual((await db.query<{ version: string }>('select version from supabase_migrations.schema_migrations order by version')).rows.map(row => row.version), versions);
    const applied = new Set((await db.query<{ version: string }>('select version from supabase_migrations.schema_migrations')).rows.map(row => row.version));
    assert.deepEqual(files.filter(file => !applied.has(file.split('_')[0])), [], 'a second deployment must have no migrations left to execute');
    assert.deepEqual((await db.query('select id,employee_id from public.roster_cham_cong')).rows, [{ id: 7, employee_id: 'SYNTHETIC' }]);
    assert.deepEqual((await db.query('select raw_data from public.roster_cham_cong')).rows, [{ raw_data: null }]);

    await db.exec(`set role authenticated; select set_config('request.jwt.claim.sub','${owner}',false);`);
    await db.query('insert into public.promotion_workspaces(owner_id,records) values ($1,$2::jsonb)', [owner, '[{"test":true}]']);
    assert.equal((await db.query('select * from public.promotion_workspaces')).rows.length, 1);
    await db.exec(`select set_config('request.jwt.claim.sub','${other}',false)`);
    assert.equal((await db.query('select * from public.promotion_workspaces')).rows.length, 0);
    assert.equal((await db.query("update public.promotion_workspaces set records='[]' returning owner_id")).rows.length, 0);
    await assert.rejects(db.query('insert into public.promotion_workspaces(owner_id,workspace_key) values ($1,$2)', [owner, 'cross-owner']), { code: '42501' });
    await db.exec(`select set_config('request.jwt.claim.sub','${owner}',false); select set_config('request.jwt.claims','{"is_anonymous":true}',false);`);
    assert.equal((await db.query('select * from public.promotion_workspaces')).rows.length, 0);
    await assert.rejects(db.query('insert into public.promotion_workspaces(owner_id,workspace_key) values ($1,$2)', [owner, 'anonymous']), { code: '42501' });
    await db.exec('reset role; set role anon');
    await assert.rejects(db.query('select * from public.promotion_workspaces'), { code: '42501' });
    await assert.rejects(db.query("select public.commit_payroll_workspace(0,'{}'::jsonb)"), { code: '42501' });
    await db.exec('reset role');
    assert.equal((await db.query('select * from public.payroll_workspaces')).rows.length, 0);
  } finally { await db.close(); }
});
