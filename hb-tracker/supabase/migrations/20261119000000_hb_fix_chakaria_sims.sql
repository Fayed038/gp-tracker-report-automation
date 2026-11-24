-- Chakaria officer list (request #16, 4 Nov 2026): the Corporate SIM and Personal Number columns were swapped.
-- Fix: give each officer the corporate number that Team Tracker shows under their name, and keep their phone as the personal number.
-- Update only, nothing deleted. Every change is logged in hb_audit under the owner account.
select set_config('request.jwt.claim.sub', '49cf46ff-6a1c-40fe-9b95-fab4c671cd44', true);

drop table if exists pg_temp.hb_fix;
create temp table hb_fix as
select o.id, o.sim_number as old_sim, o.personal_number as new_sim, o.sim_number as new_personal
from hb_officers o
where o.region_id = 10 and o.status = 'active'
  and left(o.sim_number, 8) not in ('01709661','01313475','01329708','01332819','01755669','01765095')
  and left(o.personal_number, 8) in ('01709661','01313475','01329708','01332819','01755669','01765095')
  and exists(select 1 from hb_tt t where right(regexp_replace(t.mobile, '\D', '', 'g'), 10) = right(o.personal_number, 10));
-- Ariful islam (Santirhat): personal column had 01332919945, a typo of 01332819945 ("Ariful Islam_TSO_Shantirhat" in Team Tracker)
insert into hb_fix select id, sim_number, '01332819945', sim_number from hb_officers
  where id = 364 and sim_number = '01797839491' and personal_number = '01332919945';

do $$ begin
  if (select count(*) from hb_fix) <> 40 then raise exception 'expected 40 officers to fix, found %; nothing changed', (select count(*) from hb_fix); end if;
  if exists(select 1 from hb_fix f join hb_sims s on s.number = f.new_sim) then raise exception 'a corporate number is already in the SIM register; nothing changed'; end if;
  if exists(select 1 from hb_fix f join hb_officers o on o.sim_number = f.new_sim and o.id <> f.id and o.status = 'active') then raise exception 'a corporate number belongs to another officer; nothing changed'; end if;
end $$;

update hb_sims s set number = f.new_sim, note = 'Corrected 19-11-2026: was entered as ' || f.old_sim || ' (Chakaria list had SIM and personal number swapped)'
  from hb_fix f where s.number = f.old_sim and s.officer_id = f.id;
update hb_officers o set sim_number = f.new_sim, personal_number = f.new_personal from hb_fix f where o.id = f.id;

-- Md Liakot ali (Chakaria): no corporate SIM in the list or in Team Tracker; his phone was entered as the SIM.
update hb_officers set personal_number = sim_number, sim_number = null
  where id = 337 and region_id = 10 and sim_number = '01830160865' and personal_number is null;
update hb_sims set status = 'closed', officer_id = null, note = 'Not a GP SIM: personal number of Md Liakot ali entered as SIM in the Chakaria list (closed 19-11-2026)'
  where number = '01830160865' and region_id = 10;

drop table hb_fix;
