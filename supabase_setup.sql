-- ============================================================
-- 螺旋桨测试数据记录器 · Supabase 建表脚本
-- 用法：Supabase 控制台 → 左侧 SQL Editor → New query
--       把本文件内容全部粘贴进去 → 点 Run
-- ============================================================

-- 记录表：每条记录一行，记录内容以 jsonb 存放
create table if not exists public.prop_records (
  id         bigint generated always as identity primary key,
  record_id  text,
  created_at timestamptz not null default now(),
  data       jsonb not null
);

-- 打开行级安全（RLS）
alter table public.prop_records enable row level security;

-- 允许匿名（网页里的 anon key）读取 / 新增 / 删除
-- 说明：等于"拿到网页链接的人都能增删记录"，团队内部共享够用。
-- 想限制为"只有知道口令的人能写"，见 README 文末说明。
create policy "anon read"   on public.prop_records for select using (true);
create policy "anon insert" on public.prop_records for insert with check (true);
create policy "anon delete" on public.prop_records for delete using (true);

-- 权限（一般 Supabase 已默认授予，这里显式写一遍更保险）
grant usage on schema public to anon;
grant select, insert, delete on public.prop_records to anon;

-- 可选：查看当前数据
-- select id, record_id, created_at, data from public.prop_records order by id desc limit 10;
