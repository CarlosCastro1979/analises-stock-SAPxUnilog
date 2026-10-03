import { createClient } from '@supabase/supabase-js';
import AsyncStorage from '@react-native-async-storage/async-storage';

/** Same Supabase project as Performance Logística (anon key from index.html). */
export const SB_URL = 'https://qnscwppgljobelplgbkp.supabase.co';
export const SB_KEY =
  'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InFuc2N3cHBnbGpvYmVscGxnYmtwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYxMDc3NzgsImV4cCI6MjA5MTY4Mzc3OH0.9ysq0ibsn3qDPHe5WYF-yyq9-vEKjc_hIn9BNKZccYY';

export const supabase = createClient(SB_URL, SB_KEY, {
  auth: {
    storage: AsyncStorage,
    autoRefreshToken: false,
    persistSession: false,
    detectSessionInUrl: false,
  },
});
