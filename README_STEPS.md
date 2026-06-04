# SPI Dashboard Build Steps

This starter matches the screenshot layout: beOnd header, compact summary cards, missing-data table, department/status filters, and KPI cards with small line charts.

## 1. Keep The SPI Workbook Structure

Use your existing workbook:

- `SPI 2026 Final.xlsx`
- `SPI Master`
- `Monthly Data Entry`

Useful columns:

```text
SPI Code
SPI Name
Department
Risk Area
Measure Type
Direction
Alert
Target
Data Source
Month No
Month
Status Value
Status
```

## 2. Use VBA To Prepare Dashboard Data

Open `SPI 2026 Final.xlsx`, press `Alt + F11`, import `vba/ExportSPIDashboardData.bas`, then run:

```text
ExportSPIDashboardJson
```

Choose this dashboard project folder when prompted. The macro exports:

```text
data/spi-dashboard-data.json
```

That JSON contains only the dashboard-ready data from `SPI Master` and `Monthly Data Entry`.

## 3. Preview Locally

Run the local preview server from this folder:

```text
node server.mjs
```

Then open:

```text
http://localhost:4173
```

The page loads `data/spi-dashboard-data.json`.

You can also use `Load Excel` to test the workbook manually without running the VBA export.

## 4. Put Code on GitHub

Create a GitHub repository, then add these files:

```text
index.html
styles.css
dashboard.js
data/spi-dashboard-data.json
vba/ExportSPIDashboardData.bas
supabase/schema.sql
server.mjs
netlify.toml
README_STEPS.md
```

## 5. Deploy on Netlify

Connect the GitHub repository to Netlify.

Build command:

```text
leave blank
```

Publish directory:

```text
.
```

## 6. Add Supabase For Secure Storage

Use Supabase as the secure database and Netlify as the web host.

Run `supabase/schema.sql` in Supabase SQL Editor to create:

```text
public.spi_master
public.spi_monthly_data
```

The schema enables Row Level Security and gives authenticated users read-only access.

Avoid placing any Supabase service-role key in browser code.
