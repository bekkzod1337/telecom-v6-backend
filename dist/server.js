"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
require("dotenv/config");
const express_1 = __importDefault(require("express"));
const cors_1 = __importDefault(require("cors"));
const multer_1 = __importDefault(require("multer"));
const XLSX = __importStar(require("xlsx"));
const pg_1 = require("pg");
const app = (0, express_1.default)();
const upload = (0, multer_1.default)({ storage: multer_1.default.memoryStorage(), limits: { fileSize: 250 * 1024 * 1024 } });
const pool = new pg_1.Pool({ connectionString: process.env.DATABASE_URL });
app.use((0, cors_1.default)({ origin: process.env.FRONTEND_URL ?? 'http://localhost:3000' }));
app.use(express_1.default.json({ limit: '2mb' }));
const sourceHeaders = [
    '№', 'Стандарт', 'Регион', 'Дилер', 'Торговая точка', 'Наименование клиента',
    'Тип документа', 'Серия и номер документа', 'Дата рождения', 'Тарифный план',
    'Номер телефона', 'Номер Sim-карты (ICC)', 'Номер контракта', 'Дата подключения',
    'ФИО оператора', 'Логин', 'Сумма оплат в день подключения', 'Статус заявки'
];
const tariffDefaults = {
    'Bonus Super Salom': 0.6, 'Super Salom': 0.6, 'Bonus Super Lux': 0.5,
    'Super Lux': 0.55, 'Super Mini': 0.55, 'Super Bonus': 0.6,
    'Mobile Mini M': 0.55, 'Mobile Elite': 0.7, 'Mobile Lux': 0.7,
    'Optimal': 0.55, 'Farzand': 0.5, 'Mini Voice': 0.55,
    'Data 1TB': 0.15, 'Data 300GB': 0.15, 'Data 600GB': 0.15, 'Uzmobile M2M yil': 0.15,
};
const tariffPrices = {
    'Super Lux': 70000, 'Bonus Super Lux': 70000, 'Super Salom': 70000, 'Bonus Super Salom': 70000,
    'Mobile Lux': 101000, 'Mobile Elite': 150000, 'Ideal Plus': 85000, 'Bonus Ideal Plus': 85000,
    'Mini Voice': 45000, 'Bayramona 35': 70000, 'Mobile Sport': 70000, Balance: 65000,
    Optimal: 55000, Farzand: 55000, 'Mobile Mini M': 45000, 'Uzmobile M2M yil': 100000,
    'Data 300GB': 500000, 'Data 600GB': 900000, 'Data 1TB': 1300000, 'Ideal yil': 770000,
    'Super Bonus': 70000, 'Super Mini': 45000
};
function text(value) { return value == null ? '' : String(value); }
function normalizeLogin(value) { return text(value).trim().toLowerCase(); }
function number(value) {
    const source = String(value ?? '').replace(/\s/g, '').replace(',', '.');
    const parsed = typeof value === 'number' ? value : Number(source.replace('%', ''));
    if (source.endsWith('%') && Number.isFinite(parsed))
        return parsed / 100;
    return Number.isFinite(parsed) ? parsed : 0;
}
function calculate(row) {
    const tariff = text(row['Тарифный план']).trim();
    const currentPayment = number(row['Сумма оплат в день подключения']);
    const listedPrice = tariffPrices[tariff];
    const payment = listedPrice ?? currentPayment;
    const rate = number(row['Комиссия %']);
    const net = tariff === 'Bayrammona 35' ? 62499 : (listedPrice ?? payment) / 1.12;
    const commission = tariff === 'Bonus Super Salom' ? 8000
        : tariff === 'Bonus Super Lux' || tariff === 'Bonus Ideal Plus' ? 12000
            : net * rate;
    return { ...row, 'Сумма оплат в день подключения': payment, 'Сумма без НДС': Math.trunc(net), 'Комиссия %': rate, 'Сумма комиссии': Math.trunc(commission),
        'Комиссия без НДС': Math.trunc(commission / 1.12), 'К выплате с коэффициентом': Math.trunc(commission * 1.06) };
}
function monthOf(value) {
    const match = text(value).match(/(?:^|\D)(\d{1,2})[./-]\d{1,2}[./-]\d{4}/);
    return match ? Number(match[1]) : new Date(text(value)).getMonth() + 1 || 1;
}
function percent(value) {
    const parsed = number(value);
    return parsed > 1 ? parsed / 100 : parsed;
}
function isAnnualTariff(tariff) {
    return /(?:^|\s)yil(?:\s|$)/i.test(tariff) || /(?:^|\s)yillik(?:\s|$)/i.test(tariff);
}
function isFixedFifteenPercentTariff(tariff) {
    return new Set(['Data 1TB', 'Data 300GB', 'Data 600GB', 'Uzmobile M2M yil']).has(tariff);
}
app.get('/api/health', (_req, res) => res.json({ ok: true }));
app.post('/api/reestr/import', upload.single('file'), async (req, res) => {
    if (!req.file)
        return res.status(400).json({ error: 'Excel fayl yuborilmadi' });
    try {
        const workbook = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
        const sheetName = workbook.SheetNames.find((name) => /реестр|reestr/i.test(name)) ?? workbook.SheetNames[0];
        if (!sheetName)
            return res.status(400).json({ error: 'Excel ichida sheet topilmadi' });
        const sheet = workbook.Sheets[sheetName];
        const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: false });
        const headerIndex = raw.findIndex((row) => row.some((cell) => text(cell).includes('Тарифный')));
        if (headerIndex < 0)
            return res.status(422).json({ error: 'Reestr sarlavhasi topilmadi' });
        const headers = raw[headerIndex].map((value, index) => (text(value).replace(/\s+/g, ' ').trim() || sourceHeaders[index] || `Ustun ${index + 1}`));
        const rows = raw.slice(headerIndex + 1).filter((row) => row.some((value) => text(value).trim() !== '')).map((row, index) => {
            const result = {};
            headers.forEach((header, column) => { result[header] = row[column] ?? null; });
            result['№'] = index + 1;
            result['Komissiya %'] = tariffDefaults[text(result['Тарифный план'])] ?? 0.6;
            return calculate(result);
        }).filter((row) => text(row['Статус заявки']).trim().toLocaleLowerCase() !== 'ожидание оплаты');
        const loginValues = [...new Set(rows.map((row) => normalizeLogin(row['Логин'])).filter(Boolean))];
        const dealerByLogin = new Map();
        const settingsByLogin = new Map();
        if (loginValues.length > 0) {
            const dealerResult = await pool.query(`SELECT dl.login, d.name, dl.base_rate, dl.quarter_rate_1, dl.quarter_rate_2, dl.quarter_rate_3, dl.actual_count, d.plan_count
         FROM dealer_logins dl
         INNER JOIN dealers d ON d.id = dl.dealer_id
         WHERE lower(dl.login) = ANY($1::text[])`, [loginValues]);
            for (const item of dealerResult.rows) {
                const login = normalizeLogin(item.login);
                dealerByLogin.set(login, item.name);
                settingsByLogin.set(login, {
                    baseRate: Number(item.base_rate), quarterRates: [Number(item.quarter_rate_1), Number(item.quarter_rate_2), Number(item.quarter_rate_3)],
                    plan: Number(item.plan_count), actual: Number(item.actual_count)
                });
            }
        }
        const facts = new Map();
        for (const row of rows) {
            const login = normalizeLogin(row['Логин']);
            const dealer = dealerByLogin.get(login);
            if (dealer)
                facts.set(dealer, (facts.get(dealer) ?? 0) + number(row['Сони'] || 1));
        }
        for (const row of rows) {
            const login = normalizeLogin(row['Логин']);
            const dealer = dealerByLogin.get(login) ?? text(row['Логин']).trim();
            const settings = settingsByLogin.get(login);
            const fact = facts.get(dealer) ?? settings?.actual ?? 0;
            const tariff = text(row['Тарифный план']).trim();
            const monthRateTariff = tariff === 'Super Lux' || tariff === 'Bonus Super Lux';
            const planCompleted = settings && settings.plan > 0 && fact >= settings.plan;
            const rate = isFixedFifteenPercentTariff(tariff) ? 0.15
                : settings && ((monthRateTariff && planCompleted) || (isAnnualTariff(tariff) && planCompleted))
                    ? settings.quarterRates[(monthOf(row['Дата подключения']) - 1) % 3] : settings?.baseRate;
            row['Diler nomi'] = dealer;
            row['Asl foiz'] = settings?.baseRate ?? null;
            row['1-oy foiz'] = settings?.quarterRates[0] ?? null;
            row['2-oy foiz'] = settings?.quarterRates[1] ?? null;
            row['3-oy foiz'] = settings?.quarterRates[2] ?? null;
            row['Fakt'] = settings ? fact : null;
            row['Plan'] = settings?.plan ?? null;
            if (rate !== undefined) {
                row['Комиссия %'] = rate;
                Object.assign(row, calculate(row));
            }
        }
        res.json({ sheetName, columns: [...headers, 'Сумма без НДС', 'Комиссия %', 'Сумма комиссии', 'Комиссия без НДС', 'К выплате с коэффициентом'], rows });
    }
    catch (error) {
        res.status(422).json({ error: error instanceof Error ? error.message : 'Excel o‘qishda xatolik' });
    }
});
app.post('/api/dealers', async (req, res) => {
    const { entries } = req.body;
    const validEntries = Array.isArray(entries)
        ? entries.map((entry) => ({
            login: entry.login?.trim() ?? '',
            dealer: entry.dealer?.trim() ?? '',
            baseRate: entry.baseRate,
            rates: entry.rates,
            plan: entry.plan
        }))
            .filter((entry) => entry.login && entry.dealer)
        : [];
    if (validEntries.length === 0) {
        return res.status(400).json({ error: 'Har bir qatorda login va diler nomini TAB bilan ajrating' });
    }
    const client = await pool.connect();
    try {
        await client.query('BEGIN');
        const savedEntries = [];
        for (const entry of validEntries) {
            const dealerResult = await client.query(`INSERT INTO dealers(name, company, plan_count) VALUES($1,$2,$3)
         ON CONFLICT(name) DO UPDATE SET company=EXCLUDED.company, plan_count=EXCLUDED.plan_count
         RETURNING id,name,company`, [entry.dealer, entry.dealer, number(entry.plan ?? 0)]);
            const rates = entry.rates ?? [];
            const quarterRates = [rates[0] ?? 0.6, rates[1] ?? 0.65, rates[2] ?? 0.7];
            await client.query(`INSERT INTO dealer_logins(login,dealer_id,base_rate,quarter_rate_1,quarter_rate_2,quarter_rate_3,actual_count,plan_count)
         VALUES($1,$2,$3,$4,$5,$6,$7,$8)
         ON CONFLICT(login) DO UPDATE SET dealer_id=EXCLUDED.dealer_id, base_rate=EXCLUDED.base_rate,
         quarter_rate_1=EXCLUDED.quarter_rate_1, quarter_rate_2=EXCLUDED.quarter_rate_2,
         quarter_rate_3=EXCLUDED.quarter_rate_3, actual_count=EXCLUDED.actual_count, plan_count=EXCLUDED.plan_count`, [
                entry.login,
                dealerResult.rows[0].id,
                percent(entry.baseRate ?? 0.45),
                ...quarterRates.map(percent),
                0,
                number(entry.plan ?? 0)
            ]);
            savedEntries.push(entry);
        }
        await client.query('COMMIT');
        res.status(201).json({ entries: savedEntries });
    }
    catch (error) {
        await client.query('ROLLBACK');
        res.status(500).json({ error: error instanceof Error ? error.message : 'Diler saqlanmadi' });
    }
    finally {
        client.release();
    }
});
app.get('/api/dealers', async (_req, res) => {
    try {
        const result = await pool.query(`SELECT d.id,d.name,d.company,COALESCE(array_agg(l.login) FILTER (WHERE l.login IS NOT NULL),'{}') AS logins
      FROM dealers d LEFT JOIN dealer_logins l ON l.dealer_id=d.id GROUP BY d.id ORDER BY d.name`);
        res.json(result.rows);
    }
    catch (error) {
        res.status(500).json({ error: error instanceof Error ? error.message : 'Dilerlar olinmadi' });
    }
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
    ALTER TABLE dealer_logins ADD COLUMN IF NOT EXISTS base_rate NUMERIC(8,4) NOT NULL DEFAULT 0.45;
    ALTER TABLE dealer_logins ADD COLUMN IF NOT EXISTS quarter_rate_1 NUMERIC(8,4) NOT NULL DEFAULT 0.60;
    ALTER TABLE dealer_logins ADD COLUMN IF NOT EXISTS quarter_rate_2 NUMERIC(8,4) NOT NULL DEFAULT 0.65;
    ALTER TABLE dealer_logins ADD COLUMN IF NOT EXISTS quarter_rate_3 NUMERIC(8,4) NOT NULL DEFAULT 0.70;
    ALTER TABLE dealer_logins ADD COLUMN IF NOT EXISTS actual_count INTEGER NOT NULL DEFAULT 0;
    ALTER TABLE dealer_logins ADD COLUMN IF NOT EXISTS plan_count INTEGER NOT NULL DEFAULT 0;
    ALTER TABLE dealers ADD COLUMN IF NOT EXISTS base_rate NUMERIC(8,4) NOT NULL DEFAULT 0.45;
    ALTER TABLE dealers ADD COLUMN IF NOT EXISTS quarter_rate_1 NUMERIC(8,4) NOT NULL DEFAULT 0.60;
    ALTER TABLE dealers ADD COLUMN IF NOT EXISTS quarter_rate_2 NUMERIC(8,4) NOT NULL DEFAULT 0.65;
    ALTER TABLE dealers ADD COLUMN IF NOT EXISTS quarter_rate_3 NUMERIC(8,4) NOT NULL DEFAULT 0.70;
    ALTER TABLE dealers ADD COLUMN IF NOT EXISTS plan_count INTEGER NOT NULL DEFAULT 0;
  `);
    app.listen(Number(process.env.PORT ?? 4000), () => console.log('Backend http://localhost:4000'));
}
start().catch((error) => {
    console.error('Database initialization failed:', error);
    process.exit(1);
});
