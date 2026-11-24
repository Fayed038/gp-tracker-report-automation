-- HQ admins can delete a saved daily report (to rebuild a wrong one). Everything else about saved reports stays read-only.
-- Not allowed once a lead has justified any row of it. Logged in hb_audit as "report_deleted".

-- Saved reports: still read-only, except a delete done inside hb_delete_report.
create or replace function public.hb_report_ro()
returns trigger language plpgsql set search_path to 'public' as $$
begin
  if tg_op = 'DELETE' and current_setting('hb.report_delete', true) = 'on' then return old; end if;
  raise exception 'Saved records cannot be changed';
end $$;
drop trigger if exists hb_rr_nodel on public.hb_report_rows;
create trigger hb_rr_nodel before delete on public.hb_report_rows for each row execute function public.hb_report_ro();
drop trigger if exists hb_reports_ro on public.hb_reports;
create trigger hb_reports_ro before delete or update on public.hb_reports for each row execute function public.hb_report_ro();

-- Returns the storage paths of that day's raw files, so the app can remove the files too.
create or replace function public.hb_delete_report(p_id bigint)
returns jsonb language plpgsql security definer set search_path to 'public' as $$
declare r hb_reports; n int; j int; s int; f jsonb;
begin
  if not hb_is_admin() then raise exception 'Only HQ admins can delete a report'; end if;
  select * into r from hb_reports where id = p_id for update;
  if not found then raise exception 'Report not found'; end if;
  select count(*), count(*) filter (where justified_at is not null or reason is not null), count(*) filter (where sent_at is not null)
    into n, j, s from hb_report_rows where report_id = p_id;
  if j > 0 then raise exception 'Leads have already justified % rows of this report, so it cannot be deleted', j; end if;
  select coalesce(jsonb_agg(path), '[]'::jsonb) into f from hb_files where kind = 'daily_raw' and report_date = r.report_date;
  insert into hb_audit(action, details) values ('report_deleted',
    jsonb_build_object('report_id', r.id, 'report_date', r.report_date, 'rows', n, 'sent_to_leads', s, 'files', r.summary->'files'));
  perform set_config('hb.report_delete', 'on', true);
  delete from hb_report_rows where report_id = p_id;
  delete from hb_reports where id = p_id;
  perform set_config('hb.report_delete', 'off', true);
  delete from hb_files where kind = 'daily_raw' and report_date = r.report_date;
  return f;
end $$;
revoke all on function public.hb_delete_report(bigint) from public, anon;
grant execute on function public.hb_delete_report(bigint) to authenticated;
revoke all on function public.hb_report_ro() from public, anon, authenticated;
