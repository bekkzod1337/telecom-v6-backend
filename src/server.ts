import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import multer from 'multer';
import * as XLSX from 'xlsx';
import { Pool } from 'pg';

type Cell = string | number | boolean | null;
type CalculationRow = Record<string, Cell>;

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });
const pool = new Pool({ connectionString: process.env.DATABASE_URL });
app.use(cors({ origin: process.env.FRONTEND_URL ?? 'http://localhost:3000' }));
app.use(express.json({ limit: '2mb' }));

const sourceHeaders = [
  '№', 'Стандарт', 'Регион', 'Дилер', 'Торговая точка', 'Наименование клиента',
  'Тип документа', 'Серия и номер документа', 'Дата рождения', 'Тарифный план',
  'Номер телефона', 'Номер Sim-карты (ICC)', 'Номер контракта', 'Дата подключения',
  'ФИО оператора', 'Логин', 'Сумма оплат в день подключения', 'Статус заявки'
];
const tariffDefaults: Record<string, number> = {
  'Bonus Super Salom': 0.6, 'Super Salom': 0.6, 'Bonus Super Lux': 0.5,
  'Super Lux': 0.55, 'Super Mini': 0.55, 'Super Bonus': 0.6,
  'Mobile Mini M': 0.55, 'Mobile Elite': 0.7, 'Mobile Lux': 0.7,
  'Optimal': 0.55, 'Farzand': 0.5, 'Mini Voice': 0.55,
  'Data 1TB': 0.15, 'Uzmobile M2M yil': 0.15
};
const bonusBasePrices: Record<string, number> = {
  'Bonus Super Salom': 70000,
  'Bonus Super Lux': 77001,
  'Bonus Ideal Plus': 85001
};

function text(value: Cell): string { return value == null ? '' : String(value); }
function number(value: Cell): number {
  const source = String(value ?? '').replace(/\s/g, '').replace(',', '.');
  const parsed = typeof value === 'number' ? value : Number(source.replace('%', ''));
  if (source.endsWith('%') && Number.isFinite(parsed)) return parsed / 100;
  return Number.isFinite(parsed) ? parsed : 0;
}
function calculate(row: CalculationRow): CalculationRow {
  const tariff = text(row['Тарифный план']).trim();
  const payment = bonusBasePrices[tariff] ?? number(row['Сумма оплат в день подключения']);
  const rate = number(row['Комиссия %']);
  const net = payment / 1.12;
  const commission = tariff === 'Bonus Super Salom' ? 8000
    : tariff === 'Bonus Super Lux' || tariff === 'Bonus Ideal Plus' ? 12000
    : net * rate;
  return { ...row, 'Сумма оплат в день подключения': payment, 'Сумма без НДС': net, 'Комиссия %': rate, 'Сумма комиссии': commission,
    'Комиссия без НДС': commission / 1.12, 'К выплате с коэффициентом': commission * 1.06 };
}

app.get('/api/health', (_req, res) => res.json({ ok: true }));

app.post('/api/reestr/import', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Excel fayl yuborilmadi' });
  try {
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
    const sheetName = workbook.SheetNames.find((name) => /реестр|reestr/i.test(name)) ?? workbook.SheetNames[0];
    if (!sheetName) return res.status(400).json({ error: 'Excel ichida sheet topilmadi' });
    const sheet = workbook.Sheets[sheetName];
    const raw = XLSX.utils.sheet_to_json<Cell[]>(sheet, { header: 1, defval: null, raw: false }) as Cell[][];
    const headerIndex = raw.findIndex((row) => row.some((cell) => text(cell).includes('Тарифный')));
    if (headerIndex < 0) return res.status(422).json({ error: 'Reestr sarlavhasi topilmadi' });
    const headers = raw[headerIndex].map((value, index) => (text(value).replace(/\s+/g, ' ').trim() || sourceHeaders[index] || `Ustun ${index + 1}`));
    const rows = raw.slice(headerIndex + 1).filter((row) => row.some((value) => text(value).trim() !== '')).map((row, index) => {
      const result: CalculationRow = {};
      headers.forEach((header, column) => { result[header] = row[column] ?? null; });
      result['№'] = index + 1;
      result['Komissiya %'] = tariffDefaults[text(result['Тарифный план'])] ?? 0.6;
      return calculate(result);
    });
    const loginValues = [...new Set(rows.map((row) => text(row['Логин']).trim()).filter(Boolean))];
    const dealerByLogin = new Map<string, string>();
    if (loginValues.length > 0) {
      const dealerResult = await pool.query(
        `SELECT dl.login, d.name
         FROM dealer_logins dl
         INNER JOIN dealers d ON d.id = dl.dealer_id
         WHERE dl.login = ANY($1::text[])`,
        [loginValues]
      );
      for (const item of dealerResult.rows) dealerByLogin.set(item.login, item.name);
    }
    for (const row of rows) row['Diler nomi'] = dealerByLogin.get(text(row['Логин']).trim()) ?? '';
    res.json({ sheetName, columns: [...headers, 'Сумма без НДС', 'Комиссия %', 'Сумма комиссии', 'Комиссия без НДС', 'К выплате с коэффициентом'], rows });
  } catch (error) {
    res.status(422).json({ error: error instanceof Error ? error.message : 'Excel o‘qishda xatolik' });
  }
});

app.post('/api/dealers', async (req, res) => {
  const { entries } = req.body as { entries?: Array<{ login?: string; dealer?: string }> };
  const validEntries = Array.isArray(entries)
    ? entries.map((entry) => ({ login: entry.login?.trim() ?? '', dealer: entry.dealer?.trim() ?? '' }))
      .filter((entry) => entry.login && entry.dealer)
    : [];
  if (validEntries.length === 0) {
    return res.status(400).json({ error: 'Har bir qatorda login va diler nomini TAB bilan ajrating' });
  }
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    const savedEntries: Array<{ login: string; dealer: string }> = [];
    for (const entry of validEntries) {
      const dealerResult = await client.query(
        'INSERT INTO dealers(name, company) VALUES($1,$2) ON CONFLICT(name) DO UPDATE SET company=EXCLUDED.company RETURNING id,name,company',
        [entry.dealer, entry.dealer]
      );
      await client.query('INSERT INTO dealer_logins(login,dealer_id) VALUES($1,$2) ON CONFLICT(login) DO UPDATE SET dealer_id=EXCLUDED.dealer_id',
        [entry.login, dealerResult.rows[0].id]);
      savedEntries.push(entry);
    }
    await client.query('COMMIT');
    res.status(201).json({ entries: savedEntries });
  } catch (error) {
    await client.query('ROLLBACK');
    res.status(500).json({ error: error instanceof Error ? error.message : 'Diler saqlanmadi' });
  } finally { client.release(); }
});

app.get('/api/dealers', async (_req, res) => {
  try {
    const result = await pool.query(`SELECT d.id,d.name,d.company,COALESCE(array_agg(l.login) FILTER (WHERE l.login IS NOT NULL),'{}') AS logins
      FROM dealers d LEFT JOIN dealer_logins l ON l.dealer_id=d.id GROUP BY d.id ORDER BY d.name`);
    res.json(result.rows);
  } catch (error) { res.status(500).json({ error: error instanceof Error ? error.message : 'Dilerlar olinmadi' }); }
});

async function start() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS dealers (
      id SERIAL PRIMARY KEY,
      name TEXT NOT NULL UNIQUE,
      company TEXT NOT NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
    CREATE TABLE IF NOT EXISTS dealer_logins (
      id SERIAL PRIMARY KEY,
      login TEXT NOT NULL UNIQUE,
      dealer_id INTEGER NOT NULL REFERENCES dealers(id) ON DELETE CASCADE,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
  `);
  app.listen(Number(process.env.PORT ?? 4000), () => console.log('Backend http://localhost:4000'));
}

start().catch((error) => {
  console.error('Database initialization failed:', error);
  process.exit(1);
});
