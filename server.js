import express from 'express';
import dotenv from 'dotenv';
import { google } from 'googleapis';
import path from 'path';

dotenv.config();

const app = express();
const PORT = Number(process.env.PORT || 8080);

app.use(express.json({ limit: '1mb' }));
app.use(express.static('public'));

app.get('/', (_req, res) => {
  res.sendFile(path.join(process.cwd(), 'index.html'));
});

function normalizeText(value) {
  return String(value ?? '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function parseGoogleSheetUrl(input) {
  const raw = String(input || '').trim();
  const idMatch = raw.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/) || raw.match(/^([a-zA-Z0-9-_]{20,})$/);
  if (!idMatch) throw new Error('Le lien doit être un lien Google Sheets valide.');
  let gid = null;
  try {
    const url = new URL(raw);
    gid = url.searchParams.get('gid');
  } catch (_) {
    const gidMatch = raw.match(/[?#&]gid=(\d+)/);
    gid = gidMatch ? gidMatch[1] : null;
  }
  return { spreadsheetId: idMatch[1], gid };
}

function createSheetsClient() {
  const scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly'];

  const auth = new google.auth.GoogleAuth({
    keyFile: path.join(process.cwd(), 'gvask8-9f7d1bbddfc2.json'),
    scopes
  });

  return google.sheets({
    version: 'v4',
    auth
  });
}

const HEADER_ALIASES = {
  firstName: [
  'prenom de l eleve',
  'prenom de l enfant',
  'prenom enfant',
  'first name'
],

lastName: [
  'nom de l eleve',
  'nom de l enfant',
  'nom enfant',
  'last name'
],

fullName: [
  'prenom et nom de l eleve',
  'nom et prenom de l eleve',
  'prenom et nom de l enfant',
  'nom et prenom de l enfant',
  'nom complet'
],
  email: ['adresse email', 'e mail', 'email', 'mail'],
  dob: ['date de naissance de l enfant', 'date de naissance enfant', 'date de naissance', 'birth date', 'date naissance'],
  age: ['age de l enfant', 'age', 'column 14'],
  day: ['jour', 'jour du cours', 'cours choisi', 'choix du cours', 'creneau', 'jour horaire'],
  guardian: ['prenom et nom du responsable legal', 'responsable legal', 'nom du responsable legal', 'parent responsable', 'responsable'],
  phone: ['telephone en cas urgence ou infos', 'telephone en cas d urgence ou infos', 'telephone', 'numero de telephone', 'tel'],
  medical: ['maladies allergies', 'maladies allergies ?', 'allergies', 'maladie', 'informations medicales', 'sante'],
  remark: ['remarque commentaire', 'remarques', 'remarque', 'commentaire', 'commentaires'],
  image: ['j autorise l utilisation de l image de mon enfant', 'autorisation utilisation de l image', 'droit a l image', 'droit image', 'image']
};

function findColumn(headers, aliases) {
  const normalized = headers.map(normalizeText);

  for (const alias of aliases) {
    const a = normalizeText(alias);
    const exact = normalized.findIndex(h => h === a);

    if (exact >= 0) return exact;
  }

  for (const alias of aliases) {
    const a = normalizeText(alias);
    const partial = normalized.findIndex(
      h => h.length > 0 && h.includes(a)
    );

    if (partial >= 0) return partial;
  }

  return -1;
}
function getColumnMap(headers) {
  return Object.fromEntries(Object.entries(HEADER_ALIASES).map(([key, aliases]) => [key, findColumn(headers, aliases)]));
}

function cell(row, index) {
  return index >= 0 ? String(row[index] ?? '').trim() : '';
}

function imageAllowedToImageFlag(value) {
  const v = normalizeText(value);
  // In the app, image=true means image use is refused / blocked.
  if (!v) return false;
  return /\b(non|no|pas d accord|refus|refuse|not authorized|not authorised)\b/.test(v);
}

function isLikelyMedical(value) {
  const v = normalizeText(value);
  return v && !['non', 'aucun', 'aucune', 'none', 'neant', 'n a pas'].includes(v);
}

function rowToStudent(row, map) {
  const firstName = cell(row, map.firstName);
  const lastName = cell(row, map.lastName);
  const fullName = cell(row, map.fullName);
  const name = fullName || [firstName, lastName].filter(Boolean).join(' ').trim();
  const medical = cell(row, map.medical);
  const imageSource = cell(row, map.image);
  return {
    name,
    email: cell(row, map.email),
    dob: cell(row, map.dob),
    age: cell(row, map.age),
    day: cell(row, map.day),
    guardian: cell(row, map.guardian),
    phone: cell(row, map.phone),
    medical,
    remark: cell(row, map.remark),
    image: imageAllowedToImageFlag(imageSource),
    allergy: Boolean(isLikelyMedical(medical))
  };
}

function escapeSheetName(title) {
  return `'${String(title).replace(/'/g, "''")}'`;
}

async function readSheet(url) {
  const { spreadsheetId, gid } = parseGoogleSheetUrl(url);
  const sheets = createSheetsClient();
  const meta = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets.properties' });
  const tabs = meta.data.sheets || [];
  if (!tabs.length) throw new Error('Ce Google Sheet ne contient aucun onglet.');
  let tab = gid ? tabs.find(s => String(s.properties?.sheetId) === String(gid)) : null;
  tab ||= tabs.find(s => !s.properties?.hidden) || tabs[0];
  const title = tab.properties?.title;
  if (!title) throw new Error('Impossible de déterminer l’onglet à lire.');

  const valuesResponse = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${escapeSheetName(title)}!A:ZZ`,
    valueRenderOption: 'FORMATTED_VALUE',
    dateTimeRenderOption: 'FORMATTED_STRING'
  });
  const rows = valuesResponse.data.values || [];
  if (!rows.length) throw new Error('Cet onglet est vide.');

  const headers = rows[0];
  const map = getColumnMap(headers);
  const warnings = [];
  if (map.fullName < 0 && (map.firstName < 0 || map.lastName < 0)) warnings.push('Colonnes prénom/nom non reconnues automatiquement.');
  if (map.email < 0) warnings.push('Colonne e-mail non reconnue.');
  if (map.dob < 0) warnings.push('Colonne date de naissance non reconnue.');

  const students = rows.slice(1)
    .map(row => rowToStudent(row, map))
    .filter(student => student.name || student.email);

  return { students, warnings, sheetTitle: title, headers, columnMap: map };
}

app.post('/api/sheets/import', async (req, res) => {
  try {
    const { url } = req.body || {};
    const result = await readSheet(url);
    res.json({ ok: true, count: result.students.length, students: result.students, warnings: result.warnings, sheetTitle: result.sheetTitle });
  } catch (error) {
    const message = error?.message || 'Impossible de lire ce Google Sheet.';
    console.error('[Sheets import]', error);
    res.status(400).json({ ok: false, error: message });
  }
});

const USERS = [
  {
    username: "admin",
    password: "1111",
    name: "Daniel Navarro",
    role: "admin"
  }
];

app.post('/api/login', (req, res) => {
  const { username, password } = req.body || {};

  const user = USERS.find(
    u =>
      u.username.toLowerCase() === String(username || '').trim().toLowerCase() &&
      u.password === String(password || '')
  );

  if (!user) {
    return res.status(401).json({
      ok: false,
      error: 'Identifiant ou code d’accès incorrect.'
    });
  }

  res.json({
    ok: true,
    user: {
      username: user.username,
      name: user.name,
      role: user.role
    }
  });
});

app.get('/api/health', (_req, res) => res.json({ ok: true }));

app.listen(PORT, () => {
  console.log(`GVASK8 running on http://localhost:${PORT}`);
});
