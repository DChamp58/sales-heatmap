// Zipcode → latitude/longitude geocoding for the sales heatmap.
//
// Uses the free, CORS-enabled Zippopotam.us API and caches every lookup in
// memory and in localStorage so repeated imports of the same zipcodes don't
// hit the network twice. All lookups run client-side in the user's browser.

export interface LatLng {
  lat: number;
  lng: number;
  place?: string; // e.g. "Rochester, NY"
  state?: string; // two-letter state abbreviation, e.g. "NY"
}

const STORAGE_KEY = 'sales_heatmap_zip_cache_v1';

// In-memory cache, hydrated from localStorage on first use.
let memCache: Record<string, LatLng | null> | null = null;

function loadCache(): Record<string, LatLng | null> {
  if (memCache) return memCache;
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    memCache = raw ? JSON.parse(raw) : {};
  } catch {
    memCache = {};
  }
  return memCache!;
}

function persistCache() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(memCache ?? {}));
  } catch {
    /* localStorage may be full or unavailable — ignore */
  }
}

/** Normalize a raw zipcode value (number or string) to a 5-digit US zip. */
export function normalizeZip(raw: unknown): string | null {
  if (raw === null || raw === undefined) return null;
  // Strip anything that isn't a digit, then take the first 5 (handles ZIP+4).
  const digits = String(raw).replace(/\D/g, '');
  if (digits.length < 5) {
    // Some leading zeros may have been dropped by Excel (e.g. "1001" → "01001").
    return digits.length === 4 ? digits.padStart(5, '0') : null;
  }
  return digits.slice(0, 5);
}

async function fetchZip(zip: string): Promise<LatLng | null> {
  try {
    const res = await fetch(`https://api.zippopotam.us/us/${zip}`);
    if (!res.ok) return null;
    const data = await res.json();
    const place = data?.places?.[0];
    if (!place) return null;
    const lat = parseFloat(place.latitude);
    const lng = parseFloat(place.longitude);
    if (Number.isNaN(lat) || Number.isNaN(lng)) return null;
    return {
      lat,
      lng,
      place: `${place['place name']}, ${place['state abbreviation']}`,
      state: place['state abbreviation'] || undefined,
    };
  } catch {
    return null;
  }
}

/**
 * Geocode a list of zipcodes, returning a map of zip → LatLng (or null when a
 * zip couldn't be resolved). Cached zips resolve instantly; the rest are
 * fetched with limited concurrency to be polite to the API.
 *
 * @param onProgress optional callback fired as lookups complete (done, total)
 */
export async function geocodeZips(
  zips: string[],
  onProgress?: (done: number, total: number) => void
): Promise<Record<string, LatLng | null>> {
  const cache = loadCache();
  const unique = Array.from(new Set(zips));
  const result: Record<string, LatLng | null> = {};

  const toFetch: string[] = [];
  for (const zip of unique) {
    if (zip in cache) {
      result[zip] = cache[zip];
    } else {
      toFetch.push(zip);
    }
  }

  let done = unique.length - toFetch.length;
  onProgress?.(done, unique.length);

  const CONCURRENCY = 6;
  let index = 0;

  async function worker() {
    while (index < toFetch.length) {
      const zip = toFetch[index++];
      const latlng = await fetchZip(zip);
      cache[zip] = latlng;
      result[zip] = latlng;
      done++;
      onProgress?.(done, unique.length);
    }
  }

  await Promise.all(
    Array.from({ length: Math.min(CONCURRENCY, toFetch.length) }, worker)
  );

  persistCache();
  return result;
}
