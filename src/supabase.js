
import { createClient } from "@supabase/supabase-js";

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseKey = import.meta.env.VITE_SUPABASE_KEY;

// ✅ حماية ضد القيم الفاضية
if (!supabaseUrl || !supabaseKey) {
  console.error("Supabase ENV variables missing!");
}

export const supabase = createClient(
  supabaseUrl,
  supabaseKey
);
