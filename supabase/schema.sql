-- =============================================
-- IL SIPARIO MUSICALE — Database Schema
-- Run this in Supabase SQL Editor
-- =============================================

-- Enable UUID extension
create extension if not exists "uuid-ossp";

-- Profiles table (extends Supabase auth.users)
create table public.profiles (
  id uuid references auth.users on delete cascade primary key,
  email text not null,
  full_name text not null default '',
  role text not null default 'operator' check (role in ('admin', 'operator')),
  created_at timestamptz not null default now()
);

-- Enable RLS
alter table public.profiles enable row level security;

-- Everyone can read all profiles
create policy "Profiles are viewable by authenticated users" on public.profiles
  for select using (auth.role() = 'authenticated');

-- Users can update their own profile
create policy "Users can update own profile" on public.profiles
  for update using (auth.uid() = id);

-- Admin can insert profiles (handled via trigger)
create policy "Enable insert for auth" on public.profiles
  for insert with check (true);

-- Auto-create profile on signup
create or replace function public.handle_new_user()
returns trigger as $$
begin
  insert into public.profiles (id, email, full_name)
  values (new.id, new.email, coalesce(new.raw_user_meta_data->>'full_name', ''));
  return new;
end;
$$ language plpgsql security definer;

create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();

-- Trips table — stores entire trip as JSONB
create table public.trips (
  id uuid primary key default uuid_generate_v4(),
  data jsonb not null default '{}',
  status text not null default 'planning' check (status in ('planning', 'published', 'completed')),
  trip_name text generated always as (data->>'name') stored,
  trip_type text generated always as (data->>'type') stored,
  created_by uuid references public.profiles(id),
  created_by_name text not null default '',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

-- Enable RLS
alter table public.trips enable row level security;

-- All authenticated users can do everything with trips
create policy "Trips are viewable by authenticated users" on public.trips
  for select using (auth.role() = 'authenticated');

create policy "Trips are insertable by authenticated users" on public.trips
  for insert with check (auth.role() = 'authenticated');

create policy "Trips are updatable by authenticated users" on public.trips
  for update using (auth.role() = 'authenticated');

create policy "Trips are deletable by authenticated users" on public.trips
  for delete using (auth.role() = 'authenticated');

-- Auto-update updated_at
create or replace function public.update_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

create trigger trips_updated_at
  before update on public.trips
  for each row execute function public.update_updated_at();

-- Index for fast queries
create index trips_status_idx on public.trips(status);
create index trips_created_by_idx on public.trips(created_by);
