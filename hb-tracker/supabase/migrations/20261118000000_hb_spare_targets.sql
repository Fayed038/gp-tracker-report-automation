-- Spare SIM targets per region.
-- Rule: needed = active officers without a SIM + pending joiners who still need one + buffer (default 2).
-- HQ can change a region's buffer, or set a fixed target that replaces the rule. Leads see only their own region.

alter table public.hb_regions add column if not exists spare_buffer integer not null default 2;
alter table public.hb_regions add column if not exists spare_target integer; -- null = use the rule
do $$ begin
  alter table public.hb_regions add constraint hb_regions_spare_buffer_check check (spare_buffer between 0 and 50);
  alter table public.hb_regions add constraint hb_regions_spare_target_check check (spare_target is null or spare_target between 0 and 500);
exception when duplicate_object then null; end $$;
-- HQ has no officers of its own; its spares are the shared pool, so no buffer there.
update public.hb_regions set spare_buffer = 0 where kind = 'hq' and spare_buffer = 2;

-- Log target changes like every other change.
drop trigger if exists a4 on public.hb_regions;
create trigger a4 after update of spare_buffer, spare_target on public.hb_regions for each row execute function hb_audit_trg();

-- One row per region. security_invoker: the caller's row-level security applies, so a lead sees only their region.
create or replace view public.hb_spare_status with (security_invoker = true) as
select r.id as region_id, r.name, r.kind, r.spare_buffer, r.spare_target,
  coalesce(o.officers, 0) as officers, coalesce(o.with_sim, 0) as with_sim, coalesce(o.no_sim, 0) as no_sim,
  coalesce(q.pending, 0) as pending_joiners,
  coalesce(s.ready, 0) as spares_ready, coalesce(s.inactive, 0) as spares_inactive,
  coalesce(r.spare_target, coalesce(o.no_sim, 0) + coalesce(q.pending, 0) + r.spare_buffer) as needed,
  coalesce(s.ready, 0) - coalesce(r.spare_target, coalesce(o.no_sim, 0) + coalesce(q.pending, 0) + r.spare_buffer) as balance
from hb_regions r
left join (select region_id, count(*) as officers, count(sim_number) as with_sim, count(*) - count(sim_number) as no_sim
           from hb_officers where status = 'active' group by region_id) o on o.region_id = r.id
left join (select region_id, count(*) as pending from hb_requests
           where status = 'pending' and (kind = 'need_sim' or (kind = 'new_joiner' and coalesce(payload->>'sim_number', '') = ''))
           group by region_id) q on q.region_id = r.id
left join (select region_id, count(*) filter (where status = 'spare') as ready, count(*) filter (where status = 'spare_inactive') as inactive
           from hb_sims where not locked and officer_id is null group by region_id) s on s.region_id = r.id;
revoke all on public.hb_spare_status from anon;
grant select on public.hb_spare_status to authenticated;

-- A lead may only offer a spare SIM of their own region to a new joiner (checked when the request is sent).
create or replace function public.hb_req_sim_check()
returns trigger language plpgsql security definer set search_path to 'public' as $$
begin
  if new.kind = 'new_joiner' and coalesce(new.payload->>'sim_number', '') <> '' and not hb_is_admin()
     and not exists(select 1 from hb_sims where number = new.payload->>'sim_number' and region_id = new.region_id
                    and status = 'spare' and officer_id is null and not locked) then
    raise exception 'SIM % is not a spare SIM of your region', new.payload->>'sim_number';
  end if;
  return new;
end $$;
revoke all on function public.hb_req_sim_check() from public, anon, authenticated;
drop trigger if exists hb_req_sim_check on public.hb_requests;
create trigger hb_req_sim_check before insert on public.hb_requests for each row execute function public.hb_req_sim_check();

-- hb_decide: unchanged except the block between "spare checks begin" and "spare checks end".
CREATE OR REPLACE FUNCTION public.hb_decide(p_id bigint, p_approve boolean, p_note text DEFAULT NULL::text, p_sim text DEFAULT NULL::text)
 RETURNS void
 LANGUAGE plpgsql
 SECURITY DEFINER
 SET search_path TO 'public'
AS $function$
declare q hb_requests; p jsonb; oid bigint; o hb_officers; x jsonb;
begin
  if not hb_is_admin() then raise exception 'admin only'; end if;
  select * into q from hb_requests where id=p_id for update;
  if q.status<>'pending' then raise exception 'already decided'; end if;
  p:=q.payload;
  if p_approve then
    -- spare checks begin
    -- a SIM HQ picks must be free: not locked, not held by an officer, a ready spare or in the HQ pool
    if p_sim is not null and q.kind in ('new_joiner','need_sim') and not exists(select 1 from hb_sims where number=p_sim and not locked and officer_id is null and status in ('spare','hq_pool')) then
      raise exception 'SIM % is not free to give', p_sim; end if;
    -- a SIM the lead chose for a joiner must still be a ready spare of that lead's region
    if q.kind='new_joiner' and p_sim is null and coalesce(p->>'sim_number','')<>'' and not exists(select 1 from hb_sims where number=p->>'sim_number' and region_id=q.region_id and status='spare' and officer_id is null and not locked) then
      raise exception 'SIM % is no longer a spare of this region; pick a SIM for this joiner', p->>'sim_number'; end if;
    -- an officer list may not take an unheld SIM of another region, a shared HQ spare, an inactive spare or an HQ-pool SIM
    if q.kind='submit_list' and exists(
        select 1 from hb_sims s join (select u->'changes'->>'sim_number' as n from jsonb_array_elements(coalesce(p->'update','[]')) u
                                      union all select a->>'sim_number' from jsonb_array_elements(coalesce(p->'add','[]')) a) l on l.n=s.number
        where s.officer_id is null and not s.locked and (s.region_id is distinct from q.region_id or s.status in ('spare_inactive','hq_pool'))) then
      raise exception 'The list uses a SIM that is not a spare of this region. Ask HQ to give the SIM to the region first.'; end if;
    -- spare checks end
    if q.kind='new_joiner' then
      insert into hb_officers(name,designation,region_id,territory,personal_number,sim_number,verified,joined_on)
        values(p->>'name',p->>'designation',q.region_id,p->>'territory',p->>'personal_number',nullif(coalesce(p_sim,p->>'sim_number'),''),true,nullif(p->>'joined_on','')::date) returning id into oid;
      if coalesce(p_sim,p->>'sim_number','')<>'' then
        insert into hb_sims(number,region_id,status,officer_id) values(coalesce(p_sim,p->>'sim_number'),q.region_id,'active',oid)
          on conflict(number) do update set status='active', officer_id=oid, region_id=q.region_id;
      end if;
    elsif q.kind='resignation' then
      select * into o from hb_officers where id=(p->>'officer_id')::bigint and region_id=q.region_id;
      if not found then raise exception 'officer not in region'; end if;
      update hb_officers set status='left', left_on=nullif(p->>'last_day','')::date, sim_number=null where id=o.id;
      if o.sim_number is not null then
        update hb_sims set officer_id=null,
          status=case when (p->>'sim_returned')::boolean then 'spare' else 'hq_pool' end,
          note=case when (p->>'sim_returned')::boolean then 'returned by '||o.name else 'not returned by '||o.name end
          where number=o.sim_number;
      end if;
    elsif q.kind='transfer' then
      update hb_officers set region_id=coalesce((p->>'to_region_id')::int, p_sim::int) where id=(p->>'officer_id')::bigint and region_id=q.region_id returning * into o;
      if o.sim_number is not null then update hb_sims set region_id=o.region_id where number=o.sim_number; end if;
    elsif q.kind='promotion' then
      update hb_officers set designation=p->>'designation' where id=(p->>'officer_id')::bigint and region_id=q.region_id;
    elsif q.kind='edit_officer' then
      update hb_officers set name=coalesce(p->>'name',name), territory=coalesce(p->>'territory',territory), personal_number=coalesce(p->>'personal_number',personal_number)
        where id=(p->>'officer_id')::bigint and region_id=q.region_id;
    elsif q.kind='need_sim' then
      if p_sim is null then raise exception 'choose a SIM to give'; end if;
      update hb_sims set region_id=q.region_id, status='spare', officer_id=null, note='given by HQ' where number=p_sim and status in ('hq_pool','spare','spare_inactive');
      if not found then raise exception 'SIM not available'; end if;
    elsif q.kind='submit_list' and p ? 'officers' then
      for x in select * from jsonb_array_elements(p->'officers') loop
        insert into hb_officers(name,designation,region_id,territory,sim_number,personal_number,verified)
          values(x->>'name',x->>'designation',q.region_id,x->>'territory',nullif(x->>'sim_number',''),x->>'personal_number',false) returning id into oid;
        if coalesce(x->>'sim_number','')<>'' then
          insert into hb_sims(number,region_id,status,officer_id) values(x->>'sim_number',q.region_id,'active',oid) on conflict(number) do nothing;
        end if;
      end loop;
    elsif q.kind='submit_list' then
      for x in select * from jsonb_array_elements(coalesce(p->'left','[]')) loop
        select * into o from hb_officers where id=(x->>'officer_id')::bigint and region_id=q.region_id and status='active';
        if found then
          update hb_officers set status='left', left_on=current_date, sim_number=null where id=o.id;
          if o.sim_number is not null then update hb_sims set officer_id=null, status='hq_pool', note='Left (list update) '||o.name||coalesce(': '||nullif(x->>'remarks',''),'') where number=o.sim_number and not locked; end if;
        end if;
      end loop;
      for x in select * from jsonb_array_elements(coalesce(p->'update','[]')) loop
        select * into o from hb_officers where id=(x->>'officer_id')::bigint and region_id=q.region_id and status='active';
        if not found then continue; end if;
        if (x->'changes') ? 'sim_number' and nullif(x->'changes'->>'sim_number','') is not null then
          if exists(select 1 from hb_officers where sim_number=x->'changes'->>'sim_number' and status='active' and id<>o.id) then
            raise exception 'SIM % already belongs to another officer', x->'changes'->>'sim_number'; end if;
          if exists(select 1 from hb_sims where number=x->'changes'->>'sim_number' and (locked or status in ('suspended','closed'))) then
            raise exception 'SIM % is locked, suspended or closed', x->'changes'->>'sim_number'; end if;
        end if;
        update hb_officers set
          name=case when (x->'changes') ? 'name' then coalesce(nullif(x->'changes'->>'name',''),name) else name end,
          designation=case when (x->'changes') ? 'designation' then nullif(x->'changes'->>'designation','') else designation end,
          territory=case when (x->'changes') ? 'territory' then nullif(x->'changes'->>'territory','') else territory end,
          personal_number=case when (x->'changes') ? 'personal_number' then nullif(x->'changes'->>'personal_number','') else personal_number end,
          sim_number=case when (x->'changes') ? 'sim_number' then nullif(x->'changes'->>'sim_number','') else sim_number end
          where id=o.id;
        if (x->'changes') ? 'sim_number' then
          if o.sim_number is not null then update hb_sims set officer_id=null, status='spare', note='Freed from '||o.name||' (list update)' where number=o.sim_number and not locked; end if;
          if nullif(x->'changes'->>'sim_number','') is not null then
            insert into hb_sims(number,region_id,status,officer_id) values(x->'changes'->>'sim_number',q.region_id,'active',o.id)
              on conflict(number) do update set status='active', officer_id=o.id, region_id=q.region_id, note=null;
          end if;
        end if;
      end loop;
      for x in select * from jsonb_array_elements(coalesce(p->'add','[]')) loop
        if nullif(x->>'sim_number','') is not null and (exists(select 1 from hb_officers where sim_number=x->>'sim_number' and status='active')
           or exists(select 1 from hb_sims where number=x->>'sim_number' and (locked or status in ('suspended','closed')))) then
          raise exception 'New officer %: SIM % is already in use, locked or suspended', x->>'name', x->>'sim_number'; end if;
        insert into hb_officers(name,designation,region_id,territory,sim_number,personal_number,verified)
          values(x->>'name',nullif(x->>'designation',''),q.region_id,nullif(x->>'territory',''),nullif(x->>'sim_number',''),nullif(x->>'personal_number',''),false) returning id into oid;
        if nullif(x->>'sim_number','') is not null then
          insert into hb_sims(number,region_id,status,officer_id) values(x->>'sim_number',q.region_id,'active',oid)
            on conflict(number) do update set status='active', officer_id=oid, region_id=q.region_id, note=null;
        end if;
      end loop;
    end if;
  end if;
  update hb_requests set status=case when p_approve then 'approved' else 'rejected' end, decided_by=auth.uid(), decided_at=now(), decision_note=p_note where id=p_id;
end $function$;

-- Data changes approved by HQ on 18-11-2026, logged in hb_audit under the owner account.
select set_config('request.jwt.claim.sub', '49cf46ff-6a1c-40fe-9b95-fab4c671cd44', true);
-- Rangamati: fixed target of 2 (Shanti Chakma needs no SIM: no network in Juraichori).
update public.hb_regions set spare_target = 2 where name = 'Rangamati' and spare_target is distinct from 2;
-- South Bengal - Madaripur & Faridpur: the "N/A" placeholder row is not an officer.
update public.hb_officers set status = 'left', left_on = current_date
  where id = 283 and name = 'N/A' and status = 'active' and sim_number is null
    and region_id = (select id from public.hb_regions where name = 'South Bengal - Madaripur & Faridpur');
