import express from 'express';
import { createRemoteJWKSet, jwtVerify } from 'jose';
import JSZip from 'jszip';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { readFile, stat } from 'node:fs/promises';
import { readFileSync } from 'node:fs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
try {
  const localEnv = readFileSync(path.join(__dirname, '.env'), 'utf8');
  for (const line of localEnv.split(/\r?\n/)) {
    const match = line.match(/^\s*([\w.-]+)\s*=\s*(.*?)\s*$/);
    if (!match || Object.hasOwn(process.env, match[1])) continue;
    const value = match[2];
    process.env[match[1]] = value.length >= 2 && ['"', "'"].includes(value[0]) && value.at(-1) === value[0]
      ? value.slice(1, -1)
      : value;
  }
} catch (error) {
  if (error.code !== 'ENOENT') throw error;
}

const app = express();
app.disable('x-powered-by');
app.use(express.json({ limit: '32kb' }));

const required = ['OIDC_ISSUER', 'OIDC_AUDIENCE', 'OIDC_JWKS_URL', 'OIDC_AUTHORITY', 'OIDC_CLIENT_ID', 'OIDC_API_SCOPE'];
const missing = required.filter((key) => !process.env[key]);
if (missing.length) console.warn(`Portal configuration is incomplete: ${missing.join(', ')}`);

let partnerMap = {};
try {
  partnerMap = JSON.parse(process.env.PARTNER_ACCOUNT_MAP_JSON || '{}');
  if (!partnerMap || Array.isArray(partnerMap) || typeof partnerMap !== 'object') throw new Error('Expected a JSON object');
} catch (error) {
  console.error(`Invalid PARTNER_ACCOUNT_MAP_JSON: ${error.message}`);
  process.exit(1);
}

const jwks = process.env.OIDC_JWKS_URL ? createRemoteJWKSet(new URL(process.env.OIDC_JWKS_URL)) : null;
const exportDirectory = path.resolve(process.env.EXPORT_DATA_DIR || path.join(__dirname, 'data'));
const adminFilename = path.basename(process.env.ADMIN_EXPORT_FILENAME || 'Sampled X Orignal.xlsx');
const partnerFilename = path.basename(process.env.PARTNER_EXPORT_FILENAME || 'Report Level.xlsx');
const xmlUnescape = (value) => value.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const normalizePartner = (value) => String(value ?? '').trim().replace(/\s+/g, ' ').toLocaleLowerCase('en');

function cellValue(xml) {
  if (!xml) return null;
  const textParts = [...xml.matchAll(/<(?:[\w.-]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?t>/g)].map((match) => xmlUnescape(match[1]));
  if (textParts.length) return textParts.join('');
  const value = xml.match(/<(?:[\w.-]+:)?v\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?v>/)?.[1];
  if (value === undefined || value === '') return null;
  const decoded = xmlUnescape(value);
  const number = Number(decoded);
  return Number.isFinite(number) ? number : decoded;
}

async function readExport(filePath) {
  const info = await stat(filePath);
  if (!info.isFile()) throw new Error('Export path is not a file.');
  const zip = await JSZip.loadAsync(await readFile(filePath));
  const sheetEntry = zip.file('xl/worksheets/sheet1.xml');
  if (!sheetEntry) throw new Error('Workbook has no first worksheet.');
  const xml = await sheetEntry.async('string');
  const rows = [...xml.matchAll(/<(?:[\w.-]+:)?row\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?row>/g)]
    .map(([, rowXml]) => [...rowXml.matchAll(/<(?:[\w.-]+:)?c\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?c>)/g)]
      .map(([, body = '']) => cellValue(body)));

  const headerIndex = rows.findIndex((row) => normalizePartner(row[0]) === 'partner name');
  if (headerIndex < 0) throw new Error('Export is missing the Partner Name header.');
  const columns = rows[headerIndex].map((value, index) => {
    const label = String(value ?? '').replace(/^C2B Affiliate\s*\|\s*/i, '').trim();
    return label || `Column ${index + 1}`;
  });
  const partnerColumn = columns.findIndex((column) => normalizePartner(column) === 'partner name');
  if (partnerColumn < 0) throw new Error('Export is missing a partner name column.');

  const dataRows = rows.slice(headerIndex + 1)
    .filter((row) => normalizePartner(row[partnerColumn]) && normalizePartner(row[partnerColumn]) !== 'partner name')
    .map((row) => Object.fromEntries(columns.map((column, index) => [column, row[index] ?? null])));
  const summaryRow = rows.slice(headerIndex + 1)
    .find((row) => !normalizePartner(row[partnerColumn]) && row.some((value) => value !== null && value !== ''));
  const summary = summaryRow
    ? Object.fromEntries(columns.map((column, index) => [column, summaryRow[index] ?? null]))
    : null;

  const periodRows = rows.slice(0, headerIndex);
  const period = periodRows.flat().find((value) => typeof value === 'string' && /^[A-Za-z]{3}-\d{2,4}$/.test(value)) || null;
  return {
    columns,
    rows: dataRows,
    summary,
    period,
    updatedAt: info.mtime.toISOString(),
    filename: path.basename(filePath)
  };
}

async function getDashboardData(admin) {
  const filename = admin ? adminFilename : partnerFilename;
  return readExport(path.join(exportDirectory, filename));
}

async function authenticate(req, res, next) {
  if (!jwks || !process.env.OIDC_ISSUER || !process.env.OIDC_AUDIENCE) {
    return res.status(503).json({ error: 'Identity configuration is incomplete.' });
  }
  const token = req.headers.authorization?.match(/^Bearer\s+(.+)$/i)?.[1];
  if (!token) return res.status(401).json({ error: 'Sign-in required.' });
  try {
    const { payload } = await jwtVerify(token, jwks, { issuer: process.env.OIDC_ISSUER, audience: process.env.OIDC_AUDIENCE });
    if (typeof payload.sub !== 'string' || !payload.sub) return res.status(401).json({ error: 'Invalid user identity.' });
    req.identity = payload;
    next();
  } catch {
    return res.status(401).json({ error: 'Your sign-in session is invalid or expired.' });
  }
}

function isAdmin(identity) {
  const claim = identity[process.env.ADMIN_ROLE_CLAIM || 'roles'];
  const roles = Array.isArray(claim) ? claim : typeof claim === 'string' ? [claim] : [];
  return roles.includes(process.env.ADMIN_ROLE_NAME || 'affiliate-admin');
}

function accessFor(identity) {
  const admin = isAdmin(identity);
  const partnerName = partnerMap[identity.sub];
  return admin ? { role: 'admin', partnerName: null } : partnerName
    ? { role: 'partner', partnerName }
    : null;
}

app.get('/api/public-config', (_req, res) => {
  if (!process.env.OIDC_AUTHORITY || !process.env.OIDC_CLIENT_ID || !process.env.OIDC_API_SCOPE) {
    return res.status(503).json({ error: 'Sign-in configuration is incomplete.' });
  }
  res.set('Cache-Control', 'no-store');
  res.json({ authority: process.env.OIDC_AUTHORITY, clientId: process.env.OIDC_CLIENT_ID, apiScope: process.env.OIDC_API_SCOPE });
});

app.get('/api/session', authenticate, (req, res) => {
  const access = accessFor(req.identity);
  if (!access) return res.status(403).json({ error: 'No partner account is linked to this sign-in.' });
  res.set('Cache-Control', 'no-store');
  res.json({ ...access, displayName: req.identity.name || req.identity.preferred_username || '' });
});

app.get('/api/data', authenticate, async (req, res) => {
  const access = accessFor(req.identity);
  if (!access) return res.status(403).json({ error: 'No partner account is linked to this sign-in.' });
  try {
    const dashboard = await getDashboardData(access.role === 'admin');
    const visibleRows = access.role === 'admin'
      ? dashboard.rows
      : dashboard.rows.filter((row) => normalizePartner(row['Partner Name']) === normalizePartner(access.partnerName));
    if (access.role === 'partner' && visibleRows.length === 0) {
      return res.status(404).json({ error: 'Your partner is not present in the latest export. Contact the portal administrator.' });
    }
    res.set('Cache-Control', 'no-store');
    res.json({
      ...access,
      period: dashboard.period,
      updatedAt: dashboard.updatedAt,
      rowCount: visibleRows.length,
      columns: dashboard.columns,
      summary: access.role === 'admin' ? dashboard.summary : visibleRows[0],
      rows: visibleRows
    });
  } catch (error) {
    console.error('Export load failed:', error.message);
    res.status(503).json({ error: 'The latest Apex export is unavailable. Contact the portal administrator.' });
  }
});

app.use(express.static(path.join(__dirname, 'public'), { etag: true, maxAge: '5m' }));
app.get(/.*/, (_req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

const port = Number(process.env.PORT || 3000);
app.listen(port, () => console.log(`Affiliate portal listening on http://localhost:${port}`));

