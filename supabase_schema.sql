-- Users table is handled by Supabase Auth (auth.users)
-- We can create a public.profiles table if needed, but for this app, 
-- we'll focus on the files and operations history.

-- 1. Create a bucket for PDF files
-- Note: This usually needs to be done in the Supabase Dashboard UI, 
-- but here is the concept for Storage RLS.
-- Bucket name: 'pdf-files'

-- 2. Operations History Table
CREATE TABLE IF NOT EXISTS public.operations_history (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    user_id UUID REFERENCES auth.users(id) ON DELETE CASCADE,
    operation_type TEXT NOT NULL,
    status TEXT DEFAULT 'processing',
    processed_file_url TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 3. Enable Row Level Security
ALTER TABLE public.operations_history ENABLE ROW LEVEL SECURITY;

-- 4. Create Policies
-- Users can only see their own history
CREATE POLICY "Users can view their own history" 
ON public.operations_history 
FOR SELECT 
USING (auth.uid() = user_id);

-- Users can only insert their own history
CREATE POLICY "Users can insert their own history" 
ON public.operations_history 
FOR INSERT 
WITH CHECK (auth.uid() = user_id);

-- Storage Policies (for 'pdf-files' bucket)
-- These are typically set in the storage section of Supabase
-- Policy: "Allow authenticated uploads"
-- Policy: "Allow users to read their own files"
