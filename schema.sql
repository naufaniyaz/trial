create table if not exists public.spi_master (
  spi_code text primary key,
  department text,
  department_code text,
  section text,
  risk_area text,
  spi_name text not null,
  measuring_criteria text,
  measure_type text,
  denominator_basis text,
  original_multiplier numeric,
  standard_multiplier numeric,
  indicator_type text,
  direction text,
  alert text,
  target text,
  data_source text,
  updated_at timestamptz default now()
);

create table if not exists public.spi_monthly_data (
  id bigint generated always as identity primary key,
  year integer not null,
  month_no integer not null check (month_no between 1 and 12),
  month text not null,
  spi_code text not null references public.spi_master(spi_code) on delete cascade,
  spi_name text,
  department text,
  denominator_basis text,
  original_multiplier numeric,
  standard_multiplier numeric,
  alert text,
  target text,
  denominator numeric,
  events numeric,
  original_value numeric,
  standardized_rate_per_100 numeric,
  percent_value numeric,
  status_value numeric,
  measure_type text,
  status text,
  direction text,
  updated_at timestamptz default now(),
  unique (year, month_no, spi_code)
);

alter table public.spi_master enable row level security;
alter table public.spi_monthly_data enable row level security;

create policy "Authenticated users can read SPI master"
on public.spi_master
for select
to authenticated
using (true);

create policy "Authenticated users can read SPI monthly data"
on public.spi_monthly_data
for select
to authenticated
using (true);

create index if not exists spi_monthly_data_spi_code_idx
on public.spi_monthly_data (spi_code);

create index if not exists spi_monthly_data_year_month_idx
on public.spi_monthly_data (year, month_no);
