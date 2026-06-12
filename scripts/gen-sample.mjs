// Generates sample-data/sample-sales.xlsx — fake sales data for trying out the
// heatmap. Run with: node scripts/gen-sample.mjs
import * as XLSX from 'xlsx';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dir = dirname(fileURLToPath(import.meta.url));

// Real US zipcodes spread across regions so the heatmap shows clusters + spread.
const ZIPS = [
  '14623', '14620', '14534', '14450', '14564', // Rochester NY area (cluster)
  '10001', '10016', '11201', '10701',          // NYC metro
  '02108', '02139', '01803',                   // Boston area
  '19103', '08540',                            // Philly / Princeton
  '20001', '22102',                            // DC / N. Virginia
  '30303', '30339',                            // Atlanta
  '33101', '33602',                            // Miami / Tampa
  '60601', '60606', '53703',                   // Chicago / Madison
  '48226', '44114',                            // Detroit / Cleveland
  '75201', '77002', '78701',                   // Dallas / Houston / Austin
  '80202', '84101',                            // Denver / Salt Lake
  '85004', '89101',                            // Phoenix / Vegas
  '94105', '94301', '95113', '90017', '92101', // SF / Palo Alto / San Jose / LA / San Diego
  '98101', '97201',                            // Seattle / Portland
];

const ENGINEERS = ['Jordan Lee', 'Priya Shah', 'Marcus Cole', 'Dana Wright', 'Sofia Reyes'];
const INDUSTRIES = ['Manufacturing', 'Healthcare', 'Logistics', 'Finance', 'Technology', 'Retail', 'Energy', 'Education'];

const PREFIX = ['Acme', 'Northside', 'Lakeshore', 'Summit', 'Vertex', 'Pioneer', 'Cobalt', 'Granite', 'Riverbend', 'Apex',
  'Brightway', 'Cedar', 'Delta', 'Evergreen', 'Harbor', 'Ironclad', 'Keystone', 'Meridian', 'Quantum', 'Sterling',
  'Tradewind', 'Union', 'Vantage', 'Westgate', 'Zenith'];
const SUFFIX = ['Robotics', 'Health', 'Logistics', 'Financial', 'Systems', 'Retail Group', 'Energy', 'Labs',
  'Industries', 'Solutions', 'Partners', 'Technologies', 'Dynamics', 'Manufacturing', 'Networks'];

const pick = (arr) => arr[Math.floor(Math.random() * arr.length)];
const rand = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;

const used = new Set();
function uniqueCompany() {
  let name;
  do { name = `${pick(PREFIX)} ${pick(SUFFIX)}`; } while (used.has(name));
  used.add(name);
  return name;
}

const rows = [];
for (let i = 0; i < 60; i++) {
  rows.push({
    Company: uniqueCompany(),
    'Sales Amount': rand(8, 480) * 1000, // $8k – $480k
    Zipcode: pick(ZIPS),
    'Sales Engineer': pick(ENGINEERS),
    Industry: pick(INDUSTRIES),
  });
}

const ws = XLSX.utils.json_to_sheet(rows);
ws['!cols'] = [{ wch: 24 }, { wch: 14 }, { wch: 10 }, { wch: 16 }, { wch: 16 }];
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Sales');
const out = join(__dir, '..', 'sample-data', 'sample-sales.xlsx');
XLSX.writeFile(wb, out);
console.log(`Wrote ${rows.length} rows -> ${out}`);
