import { createClient } from "@supabase/supabase-js";

// This is the public/publishable key, meant to be embedded in client-side
// code — it's safe to expose as long as Row Level Security policies on the
// tables it touches are set up correctly (see supabase/schema.sql).
const SUPABASE_URL = "https://dvjlzbmosaybzozvcodp.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_xqYwWxnWYbMCDcxjEHuovQ_HwnoqYKS";

export const supabase = createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY, {
  auth: {
    // Keeps the login in localStorage and silently refreshes it, so once
    // someone signs in on a device they stay signed in on later visits
    // without seeing the login screen again — until they log out.
    persistSession: true,
    autoRefreshToken: true,
    detectSessionInUrl: true,
  },
});
