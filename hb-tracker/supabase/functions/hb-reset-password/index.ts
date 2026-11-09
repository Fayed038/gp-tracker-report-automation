// hb-reset-password: an HQ admin sets a one-time temporary password for another HB-Tracker user.
// The service-role key is read from the Edge Function environment only; it never reaches the browser.
// The temporary password is returned once to the admin and is never logged or stored.
import { createClient } from "npm:@supabase/supabase-js@2";

// Keys are provided by Supabase automatically. Prefer the new key system (legacy anon/service_role keys retire at the end of 2026).
const envKey = (json: string, legacy: string) => { try { const k = JSON.parse(Deno.env.get(json) || "{}").default; if (k) return k as string; } catch { /* fall back */ } return Deno.env.get(legacy)!; };
const URL_ = Deno.env.get("SUPABASE_URL")!, ANON = envKey("SUPABASE_PUBLISHABLE_KEYS", "SUPABASE_ANON_KEY"), SERVICE = envKey("SUPABASE_SECRET_KEYS", "SUPABASE_SERVICE_ROLE_KEY");
const CORS = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type", "Access-Control-Allow-Methods": "POST, OPTIONS" };
const reply = (body: Record<string, unknown>, status = 200) => new Response(JSON.stringify(body), { status, headers: { ...CORS, "Content-Type": "application/json", "Cache-Control": "no-store" } });
const noSession = { persistSession: false, autoRefreshToken: false };

// 16 random characters from an alphabet without look-alikes (0/O, 1/l/I), shown as 4 groups: "Kx7m-Qp3w-Zt9r-Hb4n" (~93 bits).
function tempPassword() {
  const U = "ABCDEFGHJKLMNPQRSTUVWXYZ", L = "abcdefghijkmnopqrstuvwxyz", D = "23456789", A = U + L + D;
  const max = 256 - (256 % A.length), b = new Uint8Array(1);
  const pick = () => { do crypto.getRandomValues(b); while (b[0] >= max); return A[b[0] % A.length]; }; // rejection sampling: no bias
  for (;;) {
    const s = Array.from({ length: 16 }, pick).join("");
    if (/[A-Z]/.test(s) && /[a-z]/.test(s) && /[2-9]/.test(s)) return s.match(/.{4}/g)!.join("-");
  }
}

Deno.serve(async req => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  if (req.method !== "POST") return reply({ error: "Method not allowed" }, 405);

  // 1. Who is calling: verify the JWT with the auth server, then apply the database's own admin rule (hb_is_admin) as that user.
  const jwt = (req.headers.get("Authorization") || "").replace(/^Bearer\s+/i, "");
  if (!jwt) return reply({ error: "Not signed in" }, 401);
  const caller = createClient(URL_, ANON, { global: { headers: { Authorization: `Bearer ${jwt}` } }, auth: noSession });
  const { data: { user }, error: uErr } = await caller.auth.getUser(jwt);
  if (uErr || !user) return reply({ error: "Not signed in" }, 401);
  const { data: isAdmin, error: aErr } = await caller.rpc("hb_is_admin");
  if (aErr || isAdmin !== true) return reply({ error: "Admins only" }, 403);

  // 2. Whose password: must be an active, non-owner HB-Tracker user other than the caller.
  let userId: unknown;
  try { ({ user_id: userId } = await req.json()); } catch { return reply({ error: "Bad request" }, 400); }
  if (typeof userId !== "string" || !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(userId)) return reply({ error: "Bad request" }, 400);
  if (userId === user.id) return reply({ error: "Use Change password for your own account" }, 400);
  const admin = createClient(URL_, SERVICE, { auth: noSession });
  const { data: t, error: tErr } = await admin.from("hb_profiles").select("user_id,status,is_owner").eq("user_id", userId).maybeSingle();
  if (tErr) { console.error("hb-reset-password: profile lookup failed", tErr.message); return reply({ error: "Could not load this user" }, 500); }
  if (!t) return reply({ error: "User not found" }, 404);
  if (t.is_owner) return reply({ error: "The owner account cannot be reset" }, 403);
  if (t.status !== "active") return reply({ error: "Enable this user before resetting their password" }, 400);

  // 3. Audit first: if the log cannot be written, nothing is changed.
  const log = (action: string) => admin.rpc("hb_reset_log", { p_admin: user.id, p_user: userId, p_action: action });
  const { error: lErr } = await log("password_reset");
  if (lErr) { console.error("hb-reset-password: audit failed", lErr.message); return reply({ error: "Could not write the activity log; nothing was changed" }, 500); }

  // 4. Set the temporary password, then require a change on next sign-in and sign them out everywhere.
  const pw = tempPassword();
  const { error: pErr } = await admin.auth.admin.updateUserById(userId, { password: pw });
  if (pErr) { console.error("hb-reset-password: update failed for", userId, pErr.message); await log("password_reset_failed"); return reply({ error: "Could not set the new password" }, 500); }
  let { error: mErr } = await admin.rpc("hb_reset_mark", { p_user: userId });
  if (mErr) ({ error: mErr } = await admin.rpc("hb_reset_mark", { p_user: userId }));
  if (mErr) console.error("hb-reset-password: could not flag", userId, mErr.message);
  return reply({ temp_password: pw, ...(mErr ? { warning: "The password was reset, but the app could not force them to change it or sign out their other devices. Ask them to use Change password after signing in." } : {}) });
});
