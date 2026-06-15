# Sales Heatmap

A standalone web app that turns a sales spreadsheet into an interactive
geographic heatmap. Import an Excel/CSV file with company, sales amount,
zipcode, and sales engineer, and the app maps where your sales are coming
from — by region, by industry, and by sales engineer.

## Features

- **Excel / CSV import** — drag & drop or browse for `.xlsx`, `.xls`, or `.csv`.
- **Fuzzy column detection** — headers are matched automatically, so
  `Company` / `Client` / `Account`, `Sales Amount` / `Revenue` / `Bookings`,
  `Zipcode` / `Zip` / `Postal`, `Sales Engineer` / `Rep` / `Owner`, and an
  optional `Industry` / `Sector` / `Vertical` column all work.
- **Heat layer** — intensity weighted by sales amount (Leaflet + leaflet.heat).
- **Company markers** — sized by sales, color-coded per sales engineer, with
  click-through popups (company, amount, engineer, industry, location).
- **Filters** — by sales engineer, region (US state), and (when present) by
  industry; toggle the heat layer and markers independently.
- **Account search** — type-ahead box that finds an account by company name or
  zip and flies the map to its marker, opening the popup.
- **Region rollups** — sales aggregated by US state, with a clickable
  "Sales by Region" breakdown that drills the map into a single state.
- **Drill-down breakdowns** — clicking any engineer, region, or industry row
  applies it as a filter (click again to clear).
- **Summary** — total sales, company count, # engineers, # states, geolocated
  count, plus sales breakdowns by engineer, region, and industry.
- **Downloadable template** — generates a correctly-shaped starter `.xlsx`.

## Expected columns

| Column         | Required | Example         |
| -------------- | -------- | --------------- |
| Company        | no       | Acme Robotics   |
| Sales Amount   | no       | 125000          |
| Zipcode        | **yes**  | 14623           |
| Sales Engineer | no       | Jordan Lee      |
| Industry       | no       | Manufacturing   |

Only the zipcode is strictly required to place a row on the map. Rows without
a valid US zipcode are skipped and reported.

## Getting started

```bash
npm install
npm run dev      # http://localhost:5173
```

Build for production:

```bash
npm run build    # outputs to dist/
npm run preview  # serve the production build locally
```

## How it works

- **Parsing**: [SheetJS (`xlsx`)](https://sheetjs.com) reads the workbook
  entirely in the browser — no file ever leaves the user's machine.
- **Mapping**: [Leaflet](https://leafletjs.com) with the
  [`leaflet.heat`](https://github.com/Leaflet/Leaflet.heat) plugin and
  OpenStreetMap tiles.
- **Geocoding**: zipcodes are resolved to coordinates via the free,
  CORS-enabled [Zippopotam.us](https://www.zippopotam.us) API, with results
  cached in memory and `localStorage` so repeat imports are instant. This is
  the only outbound network call and it runs client-side.

## Tech stack

React 18 · TypeScript · Vite 6 · Leaflet · leaflet.heat · SheetJS · lucide-react

## Notes & extension ideas

- Geocoding is at zipcode-centroid granularity, which suits regional analysis.
- Zippopotam.us covers US zipcodes. To support other countries or to avoid the
  API entirely, swap `src/lib/geocode.ts` for a bundled zip→lat/lng dataset.
- The heat gradient and marker palette live at the top of
  `src/components/SalesHeatmap.tsx`.
