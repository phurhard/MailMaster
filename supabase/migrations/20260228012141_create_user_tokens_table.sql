create table public.user_tokens (
  id text primary key,
  email text not null,
  access_token text,
  refresh_token text,
  token_uri text,
  scopes jsonb,
  created_at timestamp with time zone default timezone('utc'::text, now()) not null
);
