import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import L from 'leaflet';
import 'leaflet/dist/leaflet.css';
import 'leaflet.heat';
import {
  Upload, MapPin, Users, Building2, DollarSign, AlertTriangle,
  Download, Loader2, X, Flame, Layers, EyeOff,
} from 'lucide-react';
import { geocodeZips, normalizeZip, type LatLng } from '../lib/geocode';

// ---------------------------------------------------------------------------
// Data model + Excel parsing
// ---------------------------------------------------------------------------

interface SalesRow {
  id: string; // stable id so a marker can edit/delete its exact record
  company: string;
  sales: number;
  zip: string;
  engineer: string;
  industry?: string;
}

// Map a spreadsheet header to one of our known fields using fuzzy matching.
function detectColumns(headers: string[]) {
  const norm = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '');
  const find = (candidates: string[]) =>
    headers.find((h) => {
      const n = norm(h);
      return candidates.some((c) => n.includes(c));
    });

  return {
    company: find(['company', 'client', 'account', 'customer', 'name']),
    sales: find(['sales', 'revenue', 'amount', 'total', 'value', 'bookings']),
    zip: find(['zip', 'postal', 'zipcode']),
    engineer: find(['salesengineer', 'engineer', 'rep', 'salesperson', 'seller', 'owner', 'se']),
    industry: find(['industry', 'sector', 'vertical', 'segment']),
  };
}

function parseSales(raw: unknown): number {
  if (typeof raw === 'number') return raw;
  if (raw === null || raw === undefined) return 0;
  // Strip currency symbols, commas, spaces.
  const cleaned = String(raw).replace(/[^0-9.\-]/g, '');
  const n = parseFloat(cleaned);
  return Number.isNaN(n) ? 0 : n;
}

interface ParseResult {
  rows: SalesRow[];
  skipped: number; // rows dropped for missing zip
  hasIndustry: boolean;
}

function parseWorkbook(data: ArrayBuffer): ParseResult {
  const wb = XLSX.read(data, { type: 'array' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' });
  if (json.length === 0) return { rows: [], skipped: 0, hasIndustry: false };

  const headers = Object.keys(json[0]);
  const cols = detectColumns(headers);

  const rows: SalesRow[] = [];
  let skipped = 0;

  for (const r of json) {
    const zip = cols.zip ? normalizeZip(r[cols.zip]) : null;
    if (!zip) {
      skipped++;
      continue;
    }
    rows.push({
      id: crypto.randomUUID(),
      company: cols.company ? String(r[cols.company] || 'Unknown').trim() : 'Unknown',
      sales: cols.sales ? parseSales(r[cols.sales]) : 0,
      zip,
      engineer: cols.engineer ? String(r[cols.engineer] || 'Unassigned').trim() || 'Unassigned' : 'Unassigned',
      industry: cols.industry ? String(r[cols.industry] || '').trim() || undefined : undefined,
    });
  }

  return { rows, skipped, hasIndustry: !!cols.industry };
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const ENGINEER_PALETTE = [
  '#F76902', '#2563EB', '#16A34A', '#9333EA', '#DC2626',
  '#0891B2', '#CA8A04', '#DB2777', '#4F46E5', '#65A30D',
];

function fmtMoney(n: number): string {
  if (n >= 1_000_000) return `$${(n / 1_000_000).toFixed(2)}M`;
  if (n >= 1_000) return `$${(n / 1_000).toFixed(1)}K`;
  return `$${n.toLocaleString(undefined, { maximumFractionDigits: 0 })}`;
}

const ALL = '__all__';

// Multiple companies can share a zipcode, and zipcodes geocode to a single
// centroid — so those markers would stack on the exact same point, leaving all
// but the top one hidden and unclickable. Fan any such collisions out on a
// small spiral around the shared centroid so each marker is individually
// visible and clickable. (The heat layer keeps the true coordinates; only the
// marker positions are nudged.)
interface PlacedPoint {
  row: SalesRow;
  ll: LatLng;
  at: [number, number]; // display coordinate (may be jittered)
}

function spreadOverlaps(points: { row: SalesRow; ll: LatLng }[]): PlacedPoint[] {
  const groups = new Map<string, { row: SalesRow; ll: LatLng }[]>();
  for (const p of points) {
    const key = `${p.ll.lat.toFixed(5)},${p.ll.lng.toFixed(5)}`;
    const g = groups.get(key);
    if (g) g.push(p);
    else groups.set(key, [p]);
  }

  const GOLDEN_ANGLE = 2.399963229728653; // even, non-clumping angular spacing
  const STEP = 0.025; // degrees; controls how far apart fanned markers sit
  const out: PlacedPoint[] = [];

  for (const g of groups.values()) {
    if (g.length === 1) {
      out.push({ ...g[0], at: [g[0].ll.lat, g[0].ll.lng] });
      continue;
    }
    // Sunflower spiral: radius grows as sqrt(i), angle steps by the golden
    // angle, giving an even fan that scales to any group size.
    g.forEach((p, i) => {
      const radius = STEP * Math.sqrt(i + 0.5);
      const angle = i * GOLDEN_ANGLE;
      const lat = p.ll.lat + radius * Math.cos(angle);
      // Compensate longitude for latitude so the fan stays roughly circular.
      const lng = p.ll.lng + (radius * Math.sin(angle)) / Math.cos((p.ll.lat * Math.PI) / 180);
      out.push({ ...p, at: [lat, lng] });
    });
  }

  return out;
}

// Build an interactive marker popup: the account details plus an inline editor
// for fixing the zipcode (re-geocodes on save) or deleting the record. Returns
// a DOM node so we can wire real event handlers, and stops clicks/scrolls from
// leaking through to the map underneath.
function buildMarkerPopup(
  p: PlacedPoint,
  color: string,
  handlers: { onSaveZip: (zip: string) => void; onDelete: () => void }
): HTMLElement {
  const el = document.createElement('div');
  el.style.cssText = 'font-family:system-ui;min-width:200px';
  el.innerHTML = `
    <div style="font-weight:700;color:#402E32;margin-bottom:4px">${escapeHtml(p.row.company)}</div>
    <div style="color:#F76902;font-weight:600;font-size:15px">${fmtMoney(p.row.sales)}</div>
    <div style="color:#6b5b56;font-size:12px;margin-top:4px">
      <span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${color};margin-right:5px"></span>
      ${escapeHtml(p.row.engineer)}
    </div>
    ${p.row.industry ? `<div style="color:#6b5b56;font-size:12px">${escapeHtml(p.row.industry)}</div>` : ''}
    <div style="color:#9c8a84;font-size:11px;margin-top:4px">${escapeHtml(p.ll.place ?? p.row.zip)}</div>`;

  const editor = document.createElement('div');
  editor.style.cssText = 'margin-top:8px;padding-top:8px;border-top:1px solid #E8D5C4;display:flex;flex-direction:column;gap:6px';

  const zipRow = document.createElement('div');
  zipRow.style.cssText = 'display:flex;align-items:center;gap:6px';
  const label = document.createElement('span');
  label.textContent = 'Zip';
  label.style.cssText = 'font-size:12px;color:#6b5b56';
  const input = document.createElement('input');
  input.type = 'text';
  input.value = p.row.zip;
  input.maxLength = 10;
  input.style.cssText = 'flex:1;min-width:0;border:1px solid #E8D5C4;border-radius:6px;padding:4px 6px;font-size:12px;color:#402E32';
  const saveBtn = document.createElement('button');
  saveBtn.type = 'button';
  saveBtn.textContent = 'Save';
  saveBtn.style.cssText = 'background:#F76902;border:none;color:#fff;border-radius:6px;padding:4px 10px;font-size:12px;font-weight:600;cursor:pointer';
  zipRow.append(label, input, saveBtn);

  const msg = document.createElement('div');
  msg.style.cssText = 'font-size:11px;color:#B91C1C;display:none';

  const delBtn = document.createElement('button');
  delBtn.type = 'button';
  delBtn.textContent = 'Delete account';
  delBtn.style.cssText = 'background:transparent;border:1px solid #FECACA;color:#B91C1C;border-radius:6px;padding:4px 10px;font-size:12px;font-weight:600;cursor:pointer';

  const clearMsg = () => {
    msg.style.display = 'none';
    input.style.borderColor = '#E8D5C4';
  };
  const showMsg = (text: string) => {
    msg.textContent = text;
    msg.style.display = 'block';
  };

  saveBtn.addEventListener('click', () => {
    const norm = normalizeZip(input.value);
    if (!norm) {
      input.style.borderColor = '#FECACA';
      showMsg('Enter a valid 5-digit US zip.');
      return;
    }
    if (norm === p.row.zip) {
      showMsg('That is already the current zip.');
      return;
    }
    saveBtn.textContent = 'Saving…';
    saveBtn.disabled = true;
    handlers.onSaveZip(norm);
  });
  input.addEventListener('input', clearMsg);
  input.addEventListener('keydown', (e) => { if (e.key === 'Enter') saveBtn.click(); });

  // Two-click confirm so a stray click can't wipe a record.
  let armed = false;
  delBtn.addEventListener('click', () => {
    if (!armed) {
      armed = true;
      delBtn.textContent = 'Click again to confirm';
      return;
    }
    handlers.onDelete();
  });

  editor.append(zipRow, msg, delBtn);
  el.appendChild(editor);

  L.DomEvent.disableClickPropagation(el);
  L.DomEvent.disableScrollPropagation(el);
  return el;
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export function SalesHeatmap() {
  const [rows, setRows] = useState<SalesRow[]>([]);
  const [hasIndustry, setHasIndustry] = useState(false);
  const [skipped, setSkipped] = useState(0);
  const [fileName, setFileName] = useState<string>('');
  const [geo, setGeo] = useState<Record<string, LatLng | null>>({});
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState<{ done: number; total: number } | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [dragOver, setDragOver] = useState(false);

  const [engineerFilter, setEngineerFilter] = useState<string>(ALL);
  const [industryFilter, setIndustryFilter] = useState<string>(ALL);
  const [showHeat, setShowHeat] = useState(true);
  const [showMarkers, setShowMarkers] = useState(true);
  const [hideNoSales, setHideNoSales] = useState(false);

  const mapEl = useRef<HTMLDivElement>(null);
  const mapRef = useRef<L.Map | null>(null);
  const heatRef = useRef<L.HeatLayer | null>(null);
  const markersRef = useRef<L.LayerGroup | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // -- Derived lists ---------------------------------------------------------

  const engineers = useMemo(
    () => Array.from(new Set(rows.map((r) => r.engineer))).sort(),
    [rows]
  );
  const industries = useMemo(
    () => Array.from(new Set(rows.map((r) => r.industry).filter(Boolean) as string[])).sort(),
    [rows]
  );

  const engineerColor = useMemo(() => {
    const map: Record<string, string> = {};
    engineers.forEach((e, i) => { map[e] = ENGINEER_PALETTE[i % ENGINEER_PALETTE.length]; });
    return map;
  }, [engineers]);

  const filtered = useMemo(
    () =>
      rows.filter(
        (r) =>
          (engineerFilter === ALL || r.engineer === engineerFilter) &&
          (industryFilter === ALL || r.industry === industryFilter)
      ),
    [rows, engineerFilter, industryFilter]
  );

  // -- Summary stats ---------------------------------------------------------

  const stats = useMemo(() => {
    const totalSales = filtered.reduce((s, r) => s + r.sales, 0);
    const located = filtered.filter((r) => geo[r.zip]).length;

    const byEngineer = new Map<string, number>();
    const byIndustry = new Map<string, number>();
    for (const r of filtered) {
      byEngineer.set(r.engineer, (byEngineer.get(r.engineer) ?? 0) + r.sales);
      if (r.industry) byIndustry.set(r.industry, (byIndustry.get(r.industry) ?? 0) + r.sales);
    }
    const sortDesc = (m: Map<string, number>) =>
      Array.from(m.entries()).sort((a, b) => b[1] - a[1]);

    return {
      totalSales,
      companies: filtered.length,
      located,
      byEngineer: sortDesc(byEngineer),
      byIndustry: sortDesc(byIndustry),
    };
  }, [filtered, geo]);

  // -- File handling ---------------------------------------------------------

  async function handleFile(file: File) {
    setError(null);
    setLoading(true);
    setProgress(null);
    try {
      const buf = await file.arrayBuffer();
      const { rows: parsed, skipped: sk, hasIndustry: hi } = parseWorkbook(buf);
      if (parsed.length === 0) {
        setError('No usable rows found. Make sure the sheet has a zipcode column.');
        setLoading(false);
        return;
      }
      setFileName(file.name);
      setRows(parsed);
      setSkipped(sk);
      setHasIndustry(hi);
      setEngineerFilter(ALL);
      setIndustryFilter(ALL);

      const zips = parsed.map((r) => r.zip);
      setProgress({ done: 0, total: new Set(zips).size });
      const resolved = await geocodeZips(zips, (done, total) => setProgress({ done, total }));
      setGeo(resolved);
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Failed to read file.');
    } finally {
      setLoading(false);
      setProgress(null);
    }
  }

  function onInputChange(e: React.ChangeEvent<HTMLInputElement>) {
    const f = e.target.files?.[0];
    if (f) handleFile(f);
    e.target.value = '';
  }

  function downloadTemplate() {
    const sample = [
      { Company: 'Acme Robotics', 'Sales Amount': 125000, Zipcode: '14623', 'Sales Engineer': 'Jordan Lee', Industry: 'Manufacturing' },
      { Company: 'Northside Health', 'Sales Amount': 89000, Zipcode: '14620', 'Sales Engineer': 'Priya Shah', Industry: 'Healthcare' },
      { Company: 'Lakeshore Logistics', 'Sales Amount': 64500, Zipcode: '14534', 'Sales Engineer': 'Jordan Lee', Industry: 'Logistics' },
      { Company: 'Summit Financial', 'Sales Amount': 210000, Zipcode: '10001', 'Sales Engineer': 'Marcus Cole', Industry: 'Finance' },
    ];
    const ws = XLSX.utils.json_to_sheet(sample);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sales');
    XLSX.writeFile(wb, 'sales-heatmap-template.xlsx');
  }

  function reset() {
    setRows([]);
    setGeo({});
    setFileName('');
    setSkipped(0);
    setHasIndustry(false);
    setError(null);
  }

  // -- Editing data from a marker --------------------------------------------

  const deleteRow = useCallback((id: string) => {
    setRows((rs) => rs.filter((r) => r.id !== id));
  }, []);

  // Update a record's zipcode and re-geocode it (cached lookups are instant).
  const updateRowZip = useCallback(async (id: string, zip: string) => {
    setRows((rs) => rs.map((r) => (r.id === id ? { ...r, zip } : r)));
    const resolved = await geocodeZips([zip]);
    setGeo((g) => ({ ...g, ...resolved }));
  }, []);

  // -- Map init --------------------------------------------------------------

  useEffect(() => {
    if (rows.length === 0 || !mapEl.current || mapRef.current) return;
    const map = L.map(mapEl.current, { scrollWheelZoom: true }).setView([39.5, -98.35], 4);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '&copy; OpenStreetMap contributors',
      maxZoom: 18,
    }).addTo(map);
    markersRef.current = L.layerGroup().addTo(map);
    mapRef.current = map;
  }, [rows.length]);

  // Tear down map when data is cleared.
  useEffect(() => {
    if (rows.length === 0 && mapRef.current) {
      mapRef.current.remove();
      mapRef.current = null;
      heatRef.current = null;
      markersRef.current = null;
    }
  }, [rows.length]);

  // -- Render layers when data / filters / geo change ------------------------

  useEffect(() => {
    const map = mapRef.current;
    if (!map) return;

    const points = filtered
      .map((r) => ({ row: r, ll: geo[r.zip] }))
      .filter((p): p is { row: SalesRow; ll: LatLng } => !!p.ll);

    const maxSales = points.reduce((m, p) => Math.max(m, p.row.sales), 0) || 1;

    // Heat layer — driven purely by sales, not by location/density. We sum
    // each location's sales (so a spot's intensity reflects its *total* sales,
    // not how many accounts happen to sit there) and weight by that total with
    // no presence floor, so $0 locations produce no heat at all.
    if (heatRef.current) {
      map.removeLayer(heatRef.current);
      heatRef.current = null;
    }
    const salesByZip = new Map<string, { lat: number; lng: number; sales: number }>();
    for (const p of points) {
      const e = salesByZip.get(p.row.zip);
      if (e) e.sales += p.row.sales;
      else salesByZip.set(p.row.zip, { lat: p.ll.lat, lng: p.ll.lng, sales: p.row.sales });
    }
    let maxZipSales = 0;
    for (const e of salesByZip.values()) maxZipSales = Math.max(maxZipSales, e.sales);
    maxZipSales = maxZipSales || 1;

    const heatData: [number, number, number][] = Array.from(salesByZip.values())
      .filter((e) => e.sales > 0)
      .map((e) => [e.lat, e.lng, e.sales / maxZipSales]);

    if (showHeat && heatData.length) {
      heatRef.current = L.heatLayer(heatData, {
        radius: 32,
        blur: 22,
        maxZoom: 12,
        minOpacity: 0.35,
        gradient: { 0.2: '#2563EB', 0.4: '#16A34A', 0.6: '#FACC15', 0.8: '#F76902', 1.0: '#DC2626' },
      }).addTo(map);

      // leaflet.heat appends its <canvas> to the overlayPane on top of the
      // circle markers, and re-appends it (back to the top) every time the
      // layer is re-added. That canvas would otherwise swallow clicks meant
      // for the markers underneath, making them unclickable whenever the heat
      // layer is showing. The heat layer is purely decorative, so let pointer
      // events fall straight through to the markers below.
      const heatCanvas = (heatRef.current as unknown as { _canvas?: HTMLCanvasElement })._canvas;
      if (heatCanvas) heatCanvas.style.pointerEvents = 'none';
    }

    // Markers
    const layer = markersRef.current;
    if (layer) {
      layer.clearLayers();
      if (showMarkers) {
        // Optionally hide accounts with no sales (the heat layer is unaffected).
        const markerPoints = hideNoSales ? points.filter((p) => p.row.sales > 0) : points;
        for (const p of spreadOverlaps(markerPoints)) {
          const radius = 6 + (p.row.sales / maxSales) * 22;
          const color = engineerColor[p.row.engineer] ?? '#F76902';
          const marker = L.circleMarker(p.at, {
            radius,
            color: '#ffffff',
            weight: 1.5,
            fillColor: color,
            fillOpacity: 0.8,
          });
          marker.bindPopup(
            buildMarkerPopup(p, color, {
              onSaveZip: (zip) => updateRowZip(p.row.id, zip),
              onDelete: () => deleteRow(p.row.id),
            }),
            { minWidth: 200 }
          );
          layer.addLayer(marker);
        }
      }
    }

    // Fit bounds to visible points (once we actually have some).
    if (points.length) {
      const bounds = L.latLngBounds(points.map((p) => [p.ll.lat, p.ll.lng] as [number, number]));
      map.fitBounds(bounds.pad(0.2), { maxZoom: 11 });
    }
  }, [filtered, geo, showHeat, showMarkers, hideNoSales, engineerColor, updateRowZip, deleteRow]);

  // Make sure Leaflet recalculates size after the container mounts.
  useEffect(() => {
    if (mapRef.current) setTimeout(() => mapRef.current?.invalidateSize(), 100);
  }, [rows.length]);

  // -------------------------------------------------------------------------
  // Render
  // -------------------------------------------------------------------------

  const hasData = rows.length > 0;

  return (
    <div style={{ minHeight: '100vh', backgroundColor: '#FFF6EE' }}>
      {/* App bar */}
      <header style={{ backgroundColor: '#FFFFFF', borderBottom: '1px solid #E8D5C4', padding: '0 24px' }}>
        <div style={{ maxWidth: '1280px', margin: '0 auto', height: '64px', display: 'flex', alignItems: 'center', gap: '10px' }}>
          <Flame size={24} style={{ color: '#F76902' }} />
          <span style={{ fontSize: '20px', fontWeight: 700, color: '#402E32' }}>Sales Heatmap</span>
        </div>
      </header>

      <div style={{ maxWidth: '1280px', margin: '0 auto', padding: '24px 16px 80px' }}>
        <p style={{ color: '#B5866E', marginTop: 0, marginBottom: '24px', fontSize: '15px' }}>
          Import an Excel file to visualize sales by region, industry, and sales engineer.
        </p>

        {/* Upload zone */}
        {!hasData && (
          <div
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => {
              e.preventDefault();
              setDragOver(false);
              const f = e.dataTransfer.files?.[0];
              if (f) handleFile(f);
            }}
            onClick={() => fileInputRef.current?.click()}
            style={{
              border: `2px dashed ${dragOver ? '#F76902' : '#E8D5C4'}`,
              borderRadius: '16px',
              backgroundColor: dragOver ? '#FFF1E6' : '#FFFFFF',
              padding: '56px 24px',
              textAlign: 'center',
              cursor: 'pointer',
              transition: 'all 150ms ease',
            }}
          >
            {loading ? (
              <div>
                <Loader2 size={40} className="spin" style={{ color: '#F76902', margin: '0 auto 16px' }} />
                <p style={{ color: '#402E32', fontWeight: 600 }}>
                  {progress ? `Locating zipcodes… ${progress.done}/${progress.total}` : 'Reading file…'}
                </p>
              </div>
            ) : (
              <>
                <Upload size={40} style={{ color: '#F76902', margin: '0 auto 16px' }} />
                <p style={{ color: '#402E32', fontWeight: 600, fontSize: '17px' }}>
                  Drop an Excel file here, or click to browse
                </p>
                <p style={{ color: '#B5866E', fontSize: '14px', marginTop: '6px' }}>
                  .xlsx / .xls / .csv with columns: Company, Sales Amount, Zipcode, Sales Engineer
                </p>
                <button
                  onClick={(e) => { e.stopPropagation(); downloadTemplate(); }}
                  style={{
                    marginTop: '20px', display: 'inline-flex', alignItems: 'center', gap: '8px',
                    backgroundColor: 'transparent', border: '1px solid #F76902', color: '#F76902',
                    padding: '8px 16px', borderRadius: '8px', cursor: 'pointer', fontSize: '14px', fontWeight: 600,
                  }}
                >
                  <Download size={16} /> Download template
                </button>
              </>
            )}
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={onInputChange}
              style={{ display: 'none' }}
            />
          </div>
        )}

        {error && (
          <div style={{ marginTop: '16px', display: 'flex', alignItems: 'center', gap: '10px', backgroundColor: '#FEF2F2', border: '1px solid #FECACA', borderRadius: '10px', padding: '12px 16px', color: '#B91C1C' }}>
            <AlertTriangle size={18} /> {error}
          </div>
        )}

        {/* Loaded view */}
        {hasData && (
          <>
            {/* Toolbar */}
            <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: '12px', marginBottom: '16px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '14px', color: '#402E32', backgroundColor: '#FFFFFF', border: '1px solid #E8D5C4', borderRadius: '8px', padding: '8px 12px' }}>
                <Building2 size={16} style={{ color: '#B5866E' }} />
                <strong>{fileName}</strong>
                <button onClick={reset} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#B5866E', display: 'flex', padding: 0 }} title="Clear">
                  <X size={16} />
                </button>
              </div>

              <Select label="Sales Engineer" value={engineerFilter} onChange={setEngineerFilter}
                options={[{ value: ALL, label: 'All Engineers' }, ...engineers.map((e) => ({ value: e, label: e }))]} />

              {hasIndustry && (
                <Select label="Industry" value={industryFilter} onChange={setIndustryFilter}
                  options={[{ value: ALL, label: 'All Industries' }, ...industries.map((i) => ({ value: i, label: i }))]} />
              )}

              <Toggle active={showHeat} onClick={() => setShowHeat((v) => !v)} icon={<Flame size={15} />} label="Heat" />
              <Toggle active={showMarkers} onClick={() => setShowMarkers((v) => !v)} icon={<Layers size={15} />} label="Markers" />
              <Toggle active={hideNoSales} onClick={() => setHideNoSales((v) => !v)} icon={<EyeOff size={15} />} label="Hide $0" />
            </div>

            {/* Stat cards */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: '12px', marginBottom: '16px' }}>
              <StatCard icon={<DollarSign size={18} />} label="Total Sales" value={fmtMoney(stats.totalSales)} />
              <StatCard icon={<Building2 size={18} />} label="Companies" value={String(stats.companies)} />
              <StatCard icon={<Users size={18} />} label="Sales Engineers" value={String(engineers.length)} />
              <StatCard icon={<MapPin size={18} />} label="Located" value={`${stats.located}/${stats.companies}`} />
            </div>

            {(skipped > 0 || stats.located < stats.companies) && (
              <div style={{ marginBottom: '16px', fontSize: '13px', color: '#9A6A4E', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <AlertTriangle size={15} />
                {skipped > 0 && <span>{skipped} row(s) skipped (missing/invalid zipcode). </span>}
                {stats.located < stats.companies && <span>{stats.companies - stats.located} company(ies) couldn't be placed on the map.</span>}
              </div>
            )}

            {/* Map + breakdown */}
            <div style={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1fr) 320px', gap: '16px' }} className="heatmap-layout">
              <div style={{ position: 'relative', borderRadius: '14px', overflow: 'hidden', border: '1px solid #E8D5C4', boxShadow: '0 2px 12px rgba(64,46,50,0.06)' }}>
                <div ref={mapEl} style={{ width: '100%', height: '560px', backgroundColor: '#EAE3DC' }} />
                {/* Legend */}
                <div style={{ position: 'absolute', bottom: '16px', left: '16px', zIndex: 500, backgroundColor: 'rgba(255,255,255,0.95)', borderRadius: '8px', padding: '8px 12px', fontSize: '11px', color: '#402E32', boxShadow: '0 1px 6px rgba(0,0,0,0.15)' }}>
                  <div style={{ fontWeight: 600, marginBottom: '4px' }}>Sales intensity</div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <span>Low</span>
                    <div style={{ width: '90px', height: '8px', borderRadius: '4px', background: 'linear-gradient(90deg,#2563EB,#16A34A,#FACC15,#F76902,#DC2626)' }} />
                    <span>High</span>
                  </div>
                </div>
              </div>

              {/* Breakdown panel */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                <Breakdown
                  title="Sales by Engineer"
                  entries={stats.byEngineer}
                  total={stats.totalSales}
                  colorFor={(name) => engineerColor[name] ?? '#F76902'}
                />
                {hasIndustry && stats.byIndustry.length > 0 && (
                  <Breakdown
                    title="Sales by Industry"
                    entries={stats.byIndustry}
                    total={stats.totalSales}
                    colorFor={() => '#B5866E'}
                  />
                )}
              </div>
            </div>
          </>
        )}
      </div>

      <style>{`
        @media (max-width: 900px) {
          .heatmap-layout { grid-template-columns: 1fr !important; }
        }
      `}</style>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Small presentational helpers
// ---------------------------------------------------------------------------

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]!
  ));
}

function StatCard({ icon, label, value }: { icon: React.ReactNode; label: string; value: string }) {
  return (
    <div style={{ backgroundColor: '#FFFFFF', border: '1px solid #E8D5C4', borderRadius: '12px', padding: '14px 16px' }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: '8px', color: '#B5866E', fontSize: '13px', marginBottom: '6px' }}>
        <span style={{ color: '#F76902' }}>{icon}</span>{label}
      </div>
      <div style={{ fontSize: '22px', fontWeight: 700, color: '#402E32' }}>{value}</div>
    </div>
  );
}

function Select({ label, value, onChange, options }: {
  label: string; value: string; onChange: (v: string) => void; options: { value: string; label: string }[];
}) {
  return (
    <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px', color: '#B5866E' }}>
      {label}
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        style={{ backgroundColor: '#FFFFFF', border: '1px solid #E8D5C4', borderRadius: '8px', padding: '8px 10px', fontSize: '14px', color: '#402E32', cursor: 'pointer' }}
      >
        {options.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
      </select>
    </label>
  );
}

function Toggle({ active, onClick, icon, label }: { active: boolean; onClick: () => void; icon: React.ReactNode; label: string }) {
  return (
    <button
      onClick={onClick}
      style={{
        display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', fontWeight: 600,
        padding: '8px 12px', borderRadius: '8px', cursor: 'pointer',
        border: `1px solid ${active ? '#F76902' : '#E8D5C4'}`,
        backgroundColor: active ? '#FFF1E6' : '#FFFFFF',
        color: active ? '#F76902' : '#B5866E',
      }}
    >
      {icon}{label}
    </button>
  );
}

function Breakdown({ title, entries, total, colorFor }: {
  title: string; entries: [string, number][]; total: number; colorFor: (name: string) => string;
}) {
  return (
    <div style={{ backgroundColor: '#FFFFFF', border: '1px solid #E8D5C4', borderRadius: '12px', padding: '16px' }}>
      <div style={{ fontSize: '15px', fontWeight: 700, color: '#402E32', marginBottom: '12px' }}>{title}</div>
      <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', maxHeight: '260px', overflowY: 'auto' }}>
        {entries.map(([name, amount]) => {
          const pct = total > 0 ? (amount / total) * 100 : 0;
          return (
            <div key={name}>
              <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px', marginBottom: '3px' }}>
                <span style={{ color: '#402E32', fontWeight: 500, display: 'flex', alignItems: 'center', gap: '6px', overflow: 'hidden' }}>
                  <span style={{ width: '8px', height: '8px', borderRadius: '50%', backgroundColor: colorFor(name), flexShrink: 0 }} />
                  <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{name}</span>
                </span>
                <span style={{ color: '#402E32', fontWeight: 600, flexShrink: 0, marginLeft: '8px' }}>{fmtMoney(amount)}</span>
              </div>
              <div style={{ height: '6px', backgroundColor: '#F3EBE4', borderRadius: '3px', overflow: 'hidden' }}>
                <div style={{ width: `${pct}%`, height: '100%', backgroundColor: colorFor(name), borderRadius: '3px' }} />
              </div>
            </div>
          );
        })}
        {entries.length === 0 && <div style={{ color: '#B5866E', fontSize: '13px' }}>No data</div>}
      </div>
    </div>
  );
}
