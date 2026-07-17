import { createClient } from "@supabase/supabase-js";

const supabaseUrl = (import.meta.env.VITE_SUPABASE_URL || "https://ncmdpcxswcomnisxmlmo.supabase.co").trim();
const supabaseAnonKey = (import.meta.env.VITE_SUPABASE_ANON_KEY || "sb_publishable_8ns2Q-DT5tz5gnJt88ao7A_CBgZtOMA").trim();

export const isCloudEnabled = !!(supabaseUrl && supabaseAnonKey);

export const supabase = isCloudEnabled
  ? createClient(supabaseUrl, supabaseAnonKey)
  : null;
