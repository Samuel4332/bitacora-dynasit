const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { Pool } = require('pg');

try {
    const envFile = fs.readFileSync(path.join(__dirname, '.env'), 'utf8');
    envFile.split('\n').forEach(function (line) {
        line = line.trim();
        if (!line || line.startsWith('#')) return;
        const i = line.indexOf('=');
        if (i < 0) return;
        const k = line.slice(0, i).trim();
        let v = line.slice(i + 1).trim();
        if (v.startsWith('"') && v.endsWith('"')) v = v.slice(1, -1);
        if (!process.env[k]) process.env[k] = v;
    });
} catch (e) { }

const PORT = process.env.PORT || 5000;
const HOST = process.env.HOST || '0.0.0.0';
const DATABASE_URL = process.env.DATABASE_URL || '';
const APP_PASSWORD = process.env.APP_PASSWORD || '';
const AUTH_SECRET = process.env.AUTH_SECRET || 'bitacora-dynasit-cambiar-en-render';

const pool = new Pool({
    connectionString: DATABASE_URL,
    ssl: { rejectUnauthorized: false },
    max: 5,
    idleTimeoutMillis: 30000
});

const MIME_TYPES = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.svg': 'image/svg+xml',
    '.ico': 'image/x-icon'
};

const DEFAULT_DATA = {
    rooms: ['A101', 'A203', 'B104', 'C302', 'D101', 'E205', 'F108'],
    categories: [
        'Conexion WiFi / Internet',
        'Cerradura Electronica',
        'Television / IPTV',
        'Telefono de Habitacion',
        'Caja Fuerte Digital',
        'Enchufes / Cableado de Red'
    ],
    novelties: []
};

async function initDb() {
    await pool.query('CREATE TABLE IF NOT EXISTS bitacora_state (id smallint PRIMARY KEY, data jsonb NOT NULL)');
    const r = await pool.query('SELECT data FROM bitacora_state WHERE id = 1');
    if (r.rows.length === 0) {
        let seed = DEFAULT_DATA;
        try {
            const raw = fs.readFileSync(path.join(__dirname, 'data.json'), 'utf8');
            const parsed = JSON.parse(raw);
            if (parsed && parsed.rooms && parsed.categories && parsed.novelties) seed = parsed;
        } catch (e) { }
        await pool.query('INSERT INTO bitacora_state (id, data) VALUES (1, $1)', [JSON.stringify(seed)]);
        console.log('Base de datos inicializada con los datos actuales.');
    }
}

// --- AUTENTICACION (token HMAC) ---
function makeToken(days) {
    const exp = Math.floor(Date.now() / 1000) + 60 * 60 * 24 * (days || 7);
    const payload = Buffer.from(JSON.stringify({ e: exp })).toString('base64');
    const sig = crypto.createHmac('sha256', AUTH_SECRET).update(payload).digest('base64');
    return payload + '.' + sig;
}

function validToken(token) {
    if (!token || typeof token !== 'string') return false;
    const parts = token.split('.');
    if (parts.length !== 2) return false;
    const [payload, sig] = parts;
    const expected = crypto.createHmac('sha256', AUTH_SECRET).update(payload).digest('base64');
    let ok = false;
    try { ok = crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(expected)); } catch (e) { ok = false; }
    if (!ok) return false;
    try {
        const parsed = JSON.parse(Buffer.from(payload, 'base64').toString('utf8'));
        return parsed.e && parsed.e > Math.floor(Date.now() / 1000);
    } catch (e) {
        return false;
    }
}

function checkPassword(pass) {
    if (typeof pass !== 'string' || typeof APP_PASSWORD !== 'string') return false;
    const a = Buffer.from(pass);
    const b = Buffer.from(APP_PASSWORD);
    if (a.length !== b.length) return false;
    return crypto.timingSafeEqual(a, b);
}

// --- HELPERS ---
function sendJSON(res, status, payload) {
    const body = JSON.stringify(payload);
    res.writeHead(status, {
        'Content-Type': 'application/json; charset=utf-8',
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, X-Token',
        'Cache-Control': 'no-store'
    });
    res.end(body);
}

function getRequestBody(req) {
    return new Promise((resolve, reject) => {
        let body = '';
        req.on('data', (chunk) => {
            body += chunk;
            if (body.length > 5e6) {
                reject(new Error('Payload demasiado grande'));
            }
        });
        req.on('end', () => {
            try {
                resolve(body ? JSON.parse(body) : {});
            } catch (e) {
                reject(new Error('JSON invalido'));
            }
        });
        req.on('error', reject);
    });
}

async function loadState() {
    const r = await pool.query('SELECT data FROM bitacora_state WHERE id = 1');
    if (r.rows.length === 0) return DEFAULT_DATA;
    return r.rows[0].data;
}

async function saveState(data) {
    await pool.query(
        'INSERT INTO bitacora_state (id, data) VALUES (1, $1) ON CONFLICT (id) DO UPDATE SET data = EXCLUDED.data',
        [JSON.stringify(data)]
    );
}

function needsAuth() {
    return typeof APP_PASSWORD === 'string' && APP_PASSWORD.length > 0;
}

const server = http.createServer(async (req, res) => {
    const url = new URL(req.url, `http://${req.headers.host}`);
    const pathname = url.pathname;

    if (req.method === 'OPTIONS') {
        res.writeHead(204, {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, X-Token'
        });
        res.end();
        return;
    }

    // === API ===
    if (pathname === '/api/ping' && req.method === 'GET') {
        sendJSON(res, 200, { ok: true });
        return;
    }

    if (pathname === '/api/login' && req.method === 'POST') {
        try {
            const body = await getRequestBody(req);
            const tokenDays = body.remember ? 90 : 7;
            if (needsAuth() && checkPassword(body.password)) {
                sendJSON(res, 200, { token: makeToken(tokenDays), days: tokenDays });
            } else if (needsAuth()) {
                sendJSON(res, 401, { error: 'Contrasena incorrecta' });
            } else {
                sendJSON(res, 200, { token: makeToken(tokenDays), days: tokenDays });
            }
        } catch (e) {
            sendJSON(res, 400, { error: e.message });
        }
        return;
    }

    if (pathname === '/api/data') {
        if (req.method === 'GET') {
            try {
                if (needsAuth() && !validToken(req.headers['x-token'])) {
                    sendJSON(res, 401, { error: 'AUTENTICACION_REQUERIDA' });
                    return;
                }
                sendJSON(res, 200, await loadState());
            } catch (e) {
                sendJSON(res, 500, { error: e.message });
            }
            return;
        }
        if (req.method === 'POST') {
            try {
                if (needsAuth() && !validToken(req.headers['x-token'])) {
                    sendJSON(res, 401, { error: 'AUTENTICACION_REQUERIDA' });
                    return;
                }
                const incoming = await getRequestBody(req);
                const current = await loadState();
                const next = {
                    rooms: Array.isArray(incoming.rooms) ? incoming.rooms : current.rooms,
                    categories: Array.isArray(incoming.categories) ? incoming.categories : current.categories,
                    novelties: Array.isArray(incoming.novelties) ? incoming.novelties : current.novelties
                };
                await saveState(next);
                sendJSON(res, 200, next);
            } catch (e) {
                sendJSON(res, 400, { error: e.message });
            }
            return;
        }
    }

    // === STATIC FILES ===
    let filePath;
    if (pathname === '/' || pathname === '/index.html') {
        filePath = path.join(__dirname, 'index.html');
    } else {
        filePath = path.join(__dirname, pathname);
    }

    const ext = path.extname(filePath).toLowerCase();
    const mimeType = MIME_TYPES[ext] || 'application/octet-stream';

    fs.readFile(filePath, (err, content) => {
        if (err) {
            if (err.code === 'ENOENT') {
                res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
                res.end('404 Not Found');
            } else {
                res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
                res.end('500 Internal Server Error');
            }
            return;
        }
        res.writeHead(200, { 'Content-Type': mimeType, 'Cache-Control': 'no-store' });
        res.end(content);
    });
});

pool.on('error', (err) => {
    console.error('Error en pool PostgreSQL:', err.message);
});

initDb()
    .then(() => {
        server.listen(PORT, HOST, () => {
            console.log('==========================================');
            console.log('  Bitacora Dynas-IT - NUBE (Postgres)');
            console.log('==========================================');
            console.log('  Puerto:           ' + PORT);
            console.log('  Auth requerida:   ' + (needsAuth() ? 'SI' : 'NO'));
            console.log('==========================================');
        });
    })
    .catch((err) => {
        console.error('No se pudo iniciar la base de datos: ' + err.message);
        process.exit(1);
    });