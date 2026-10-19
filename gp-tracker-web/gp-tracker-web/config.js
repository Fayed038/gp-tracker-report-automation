// Public settings (the publishable key is safe to ship; data is protected by login + row-level security)
window.GP_CONFIG = {
  url: "https://iglscobmkgeupfmbgzkc.supabase.co",
  // Google relay (Apps Script web app URL ending in /exec). When set, all database calls go through Google,
  // so mobile operators that block supabase.co still work.
  relay: "PASTE_RELAY_URL_HERE",
  key: "sb_publishable_VHKjJ2uE0wkT2oQ94WcuuA_K5pqG3DF",
  checkpoints: [["8:00 AM", 8, 0], ["12:00 PM", 12, 0], ["3:00 PM", 15, 0], ["5:00 PM", 17, 0]],
  toleranceMinutes: 20,
  idleMinutes: 120
};
