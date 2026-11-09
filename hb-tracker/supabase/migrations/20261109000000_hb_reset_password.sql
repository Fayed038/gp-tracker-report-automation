-- Admin "Reset password": forced change on next sign-in + audit helpers for the hb-reset-password Edge Function.
-- The temporary password is never stored here; only who reset whom, and when, goes to hb_audit.

alter table public.hb_profiles add column if not exists must_change_password boolean not null default false;

-- Audit row for a reset (or a failed attempt). p_admin = the admin who did it, p_user = whose password.
-- Called by the Edge Function BEFORE the password changes, so no reset can happen without a log row.
create or replace function public.hb_reset_log(p_admin uuid, p_user uuid, p_action text)
returns void language plpgsql security definer set search_path to 'public' as $$
declare t hb_profiles;
begin
  if p_action not in ('password_reset', 'password_reset_failed') then raise exception 'bad action'; end if;
  select * into t from hb_profiles where user_id = p_user;
  if not found then raise exception 'user not found'; end if;
  insert into hb_audit(user_id, action, details, region_id)
    values (p_admin, p_action, jsonb_build_object('email', t.email, 'name', t.full_name, 'user', p_user), t.region_id);
end $$;

-- After the new password is set: require a change on next sign-in and sign them out everywhere (same as Disable).
create or replace function public.hb_reset_mark(p_user uuid)
returns void language plpgsql security definer set search_path to 'public' as $$
begin
  update hb_profiles set must_change_password = true where user_id = p_user;
  delete from auth.sessions where user_id = p_user;
end $$;

-- Only the Edge Function (service role) may call these; browsers cannot.
revoke all on function public.hb_reset_log(uuid, uuid, text) from public, anon, authenticated;
revoke all on function public.hb_reset_mark(uuid) from public, anon, authenticated;
grant execute on function public.hb_reset_log(uuid, uuid, text) to service_role;
grant execute on function public.hb_reset_mark(uuid) to service_role;

-- The flag clears only when the password itself really changes, so it cannot be skipped from the browser.
-- (The Edge Function sets the flag AFTER its own password update, so that update does not clear it.)
create or replace function public.hb_pw_changed()
returns trigger language plpgsql security definer set search_path to 'public' as $$
begin
  update hb_profiles set must_change_password = false where user_id = new.id and must_change_password;
  return new;
end $$;
revoke all on function public.hb_pw_changed() from public, anon, authenticated;

drop trigger if exists hb_pw_changed on auth.users;
create trigger hb_pw_changed after update of encrypted_password on auth.users
  for each row when (old.encrypted_password is distinct from new.encrypted_password)
  execute function public.hb_pw_changed();
