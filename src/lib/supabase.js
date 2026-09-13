import { createClient } from '@supabase/supabase-js';

const SUPABASE_URL = 'https://qcubsjcttikjtlsicaii.supabase.co';

const SUPABASE_ANON_KEY =
  'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InFjdWJzamN0dGlranRsc2ljYWlpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODgwODc5MjIsImV4cCI6MjEwMzY2MzkyMn0.tZ0Yekxc2b9BM_mgPOmQ1Gv8cqTBRUQArmznP_WDCek';

export const supabase = createClient(
  SUPABASE_URL,
  SUPABASE_ANON_KEY
);