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

const EXPORT_TICKETS = new Map();
const EXPORT_TICKET_TTL = 4 * 60 * 1000; // 4 minutos

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

function two(x) {
    x = String(x);
    return x.length < 2 ? '0' + x : x;
}

// ============================================================
// ===== GENERADOR XLSX (4 hojas, sin dependencias) =====
// ============================================================
var CRC_TABLE = (function () {
    var t = [];
    for (var n = 0; n < 256; n++) {
        var c = n;
        for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
        t[n] = c >>> 0;
    }
    return t;
})();

function crc32(buf) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
}

function xlsxEsc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function xlsxColName(i) {
    var s = '';
    i++;
    while (i > 0) {
        var m = (i - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        i = Math.floor((i - 1) / 26);
    }
    return s;
}

function xlsxSheetXml(rows) {
    var out = '';
    for (var r = 0; r < rows.length; r++) {
        out += '<row r="' + (r + 1) + '">';
        var row = rows[r];
        for (var c = 0; c < row.length; c++) {
            var ref = xlsxColName(c) + (r + 1);
            var cell = row[c];
            if (cell && cell.t === 'n') {
                out += '<c r="' + ref + '"><v>' + cell.v + '</v></c>';
            } else {
                out += '<c r="' + ref + '" t="inlineStr"><is><t>' + xlsxEsc(cell == null ? '' : cell.v) + '</t></is></c>';
            }
        }
        out += '</row>';
    }
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + out + '</sheetData></worksheet>';
}

function zipStore(entries) {
    var now = new Date();
    var dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    var dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    var parts = [], central = [], offset = 0;
    entries.forEach(function (e) {
        var name = Buffer.from(e.name, 'utf8');
        var data = e.data;
        var crc = crc32(data);
        var lh = Buffer.alloc(30);
        lh.writeUInt32LE(0x04034b50, 0);
        lh.writeUInt16LE(20, 4);
        lh.writeUInt16LE(dosTime, 10);
        lh.writeUInt16LE(dosDate, 12);
        lh.writeUInt32LE(crc, 14);
        lh.writeUInt32LE(data.length, 18);
        lh.writeUInt32LE(data.length, 22);
        lh.writeUInt16LE(name.length, 26);
        parts.push(lh, name, data);
        var ch = Buffer.alloc(46);
        ch.writeUInt32LE(0x02014b50, 0);
        ch.writeUInt16LE(20, 4);
        ch.writeUInt16LE(20, 6);
        ch.writeUInt16LE(dosTime, 14);
        ch.writeUInt16LE(dosDate, 16);
        ch.writeUInt32LE(crc, 16);
        ch.writeUInt32LE(data.length, 20);
        ch.writeUInt32LE(data.length, 24);
        ch.writeUInt16LE(name.length, 28);
        ch.writeUInt32LE(offset, 42);
        central.push(ch, name);
        offset += 30 + name.length + data.length;
    });
    var cdBuf = Buffer.concat(central);
    var cdStart = 0;
    parts.forEach(function (b) { cdStart += b.length; });
    var eocd = Buffer.alloc(22);
    eocd.writeUInt32LE(0x06054b50, 0);
    eocd.writeUInt16LE(entries.length, 8);
    eocd.writeUInt16LE(entries.length, 10);
    eocd.writeUInt32LE(cdBuf.length, 12);
    eocd.writeUInt32LE(cdStart, 16);
    return Buffer.concat(parts.concat([cdBuf, eocd]));
}

function buildXlsx(sheets) {
    var entries = [], ctOverride = '', wbSheets = '', wbRels = '';
    sheets.forEach(function (s, i) {
        var n = i + 1;
        ctOverride += '<Override PartName="/xl/worksheets/sheet' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
        wbSheets += '<sheet name="' + xlsxEsc(s.name) + '" sheetId="' + n + '" r:id="rId' + n + '"/>';
        wbRels += '<Relationship Id="rId' + n + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + n + '.xml"/>';
        entries.push({ name: 'xl/worksheets/sheet' + n + '.xml', data: Buffer.from(xlsxSheetXml(s.rows), 'utf8') });
    });
    var CT = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' + ctOverride + '</Types>';
    var RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    var WB = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + wbSheets + '</sheets></workbook>';
    var WBR = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + wbRels + '</Relationships>';
    entries.unshift(
        { name: '[Content_Types].xml', data: Buffer.from(CT, 'utf8') },
        { name: '_rels/.rels', data: Buffer.from(RELS, 'utf8') },
        { name: 'xl/workbook.xml', data: Buffer.from(WB, 'utf8') },
        { name: 'xl/_rels/workbook.xml.rels', data: Buffer.from(WBR, 'utf8') }
    );
    return zipStore(entries);
}

function prepareExport(noveltyList) {
    var list = (Array.isArray(noveltyList) ? noveltyList : []).slice();
    list.sort(function (a, b) {
        return String(b.dateLocal || '').localeCompare(String(a.dateLocal || '')) || String(a.room || '').localeCompare(String(b.room || ''));
    });
    var total = list.length;
    var cM = {}, rM = {}, mM = {};
    list.forEach(function (n) {
        var c = String(n.category || 'Sin categoria');
        var r = String(n.room || 'Sin habitacion');
        var key = String(n.dateLocal || '').slice(0, 7);
        cM[c] = (cM[c] || 0) + 1;
        rM[r] = (rM[r] || 0) + 1;
        if (key) mM[key] = (mM[key] || 0) + 1;
    });
    function sorted(map) {
        return Object.keys(map).map(function (k) { return { k: k, n: map[k] }; }).sort(function (a, b) { return b.n - a.n || String(a.k).localeCompare(String(b.k)); });
    }
    return { list: list, total: total, byCat: sorted(cM), byRoom: sorted(rM), byMonth: sorted(mM) };
}

function pct(n, total) {
    return total ? Math.round(n * 100 / total) + '%' : '0%';
}

function exportToXlsx(exported) {
    var detalle = [[{ v: 'Fecha' }, { v: 'Hora' }, { v: 'Habitacion' }, { v: 'Categoria' }, { v: 'Descripcion' }]];
    exported.list.forEach(function (n) {
        var dl = String(n.dateLocal || '');
        detalle.push([{ v: dl.slice(0, 10) }, { v: dl.slice(11, 16) }, { v: n.room }, { v: n.category }, { v: n.description }]);
    });
    function table(title, rows, showPct) {
        var hdr = showPct ? [{ v: title }, { v: 'Cantidad' }, { v: 'Porcentaje' }] : [{ v: title }, { v: 'Cantidad' }];
        var body = [hdr];
        rows.forEach(function (i) {
            var r = [{ v: i.k }, { v: i.n, t: 'n' }];
            if (showPct) r.push({ v: pct(i.n, exported.total) });
            body.push(r);
        });
        var tot = [{ v: 'TOTAL' }, { v: exported.total, t: 'n' }];
        if (showPct) tot.push({ v: '100%' });
        body.push(tot);
        return body;
    }
    return buildXlsx([
        { name: 'Detalle', rows: detalle },
        { name: 'Por Categoria', rows: table('Categoria', exported.byCat, true) },
        { name: 'Por Habitacion', rows: table('Habitacion', exported.byRoom, true) },
        { name: 'Por Mes', rows: table('Mes', exported.byMonth, false) }
    ]);
}

// ============================================================
// ===== GENERADOR PDF (sin dependencias) =====
// ============================================================
var WINANSI = (function () {
    var m = {};
    [0xC1, 0xC9, 0xCD, 0xD3, 0xDA, 0xD1, 0xDC, 0xC7, 0xE1, 0xE9, 0xED, 0xF3, 0xFA, 0xF1, 0xFC, 0xE7, 0xBB, 0xBF, 0xA1, 0xB0, 0xAA, 0xBA, 0xB7, 0xBC, 0xA7].forEach(function (c) { m[c] = c; });
    m[0x2018] = 0x91; m[0x2019] = 0x92; m[0x201C] = 0x93; m[0x201D] = 0x94;
    m[0x2022] = 0x95; m[0x2013] = 0x96; m[0x2014] = 0x97; m[0x20AC] = 0x80;
    return m;
})();

function winansiBytes(s) {
    var arr = [];
    s = String(s);
    for (var i = 0; i < s.length; i++) {
        var ch = s.charCodeAt(i);
        if (ch < 128) { arr.push(ch); continue; }
        var mapped = WINANSI[ch];
        arr.push(mapped == null ? 0x3F : mapped);
    }
    return arr;
}

function pdfStr(s) {
    var arr = winansiBytes(String(s).replace(/\\/g, '\\\\').replace(/\(/g, '\\(').replace(/\)/g, '\\)'));
    var out = '';
    for (var i = 0; i < arr.length; i++) out += String.fromCharCode(arr[i]);
    return out;
}

var CHARW = (function () {
    var w = {};
    var letters = 'acemnorsuvwxz';
    for (var i = 0; i < letters.length; i++) w[letters[i]] = 0.5;
    var wide = 'bdghkpqyfijlt';
    for (i = 0; i < wide.length; i++) w[wide[i]] = 0.5;
    w['m'] = 0.78; w['w'] = 0.78;
    'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('').forEach(function (c) { w[c] = 0.667; });
    '0123456789'.split('').forEach(function (c) { w[c] = 0.556; });
    w[' '] = 0.278; w['.'] = 0.222; w[','] = 0.222; w[':'] = 0.333; w[';'] = 0.333;
    w['('] = 0.333; w[')'] = 0.333; w['-'] = 0.333; w['/'] = 0.278; w['\u00b7'] = 0.278; w['|'] = 0.278;
    return w;
})();

function textWidth(t, size) {
    var tot = 0;
    t = String(t);
    for (var i = 0; i < t.length; i++) {
        var c = t[i];
        tot += (CHARW[c] != null ? CHARW[c] : 0.5) * size;
    }
    return tot;
}

function wrapText(t, size, maxW) {
    var words = String(t || '').split(/\s+/);
    var lines = [], cur = '';
    for (var i = 0; i < words.length; i++) {
        var w = words[i];
        var test = cur ? cur + ' ' + w : w;
        if (cur && textWidth(test, size) > maxW) { lines.push(cur); cur = w; }
        else cur = test;
    }
    if (cur) lines.push(cur);
    return lines.length ? lines : [''];
}

function PdfDoc() {
    this.pages = [[]];
    this.y = 756;
    this.x0 = 50;
    this.x1 = 562;
    this.font = 'F1';
    this.size = 10;
}
PdfDoc.prototype.newStream = function () {
    this.pages.push([]);
    this.y = 756;
};
PdfDoc.prototype.ensure = function (needed) {
    if (this.y - needed < 44) this.newStream();
};
PdfDoc.prototype.setFont = function (f) { this.font = f; };
PdfDoc.prototype.setText = function (s) { this.size = s; };
PdfDoc.prototype.writeRaw = function (text, x, align) {
    var size = this.size;
    var px = x;
    if (align === 'right') px = x - textWidth(text, size);
    this.pages[this.pages.length - 1].push('BT /' + this.font + ' ' + size + ' Tf ' + Math.round(px) + ' ' + Math.round(this.y) + ' Td (' + pdfStr(text) + ') Tj ET');
};
PdfDoc.prototype.nextLine = function (n) {
    this.y -= this.size * 1.35 * (n || 1);
};
PdfDoc.prototype.row = function (left, right) {
    this.ensure(this.size * 1.35);
    this.writeRaw(left, this.x0, 'left');
    if (right != null) this.writeRaw(right, this.x1, 'right');
    this.nextLine(1);
};
PdfDoc.prototype.drawLine = function (strokeW) {
    var yy = Math.round(this.y - 1);
    this.pages[this.pages.length - 1].push((strokeW || 0.6) + ' w ' + this.x0 + ' ' + yy + ' m ' + this.x1 + ' ' + yy + ' l S');
    this.pages[this.pages.length - 1].push('0 w');
};

function buildPdf(exported, title) {
    var doc = new PdfDoc();
    var now = new Date();
    var MAXW = doc.x1 - doc.x0;
    doc.setFont('F2'); doc.setText(15);
    doc.row('Bitacora Dynas-IT - Reporte de Novedades');
    doc.setFont('F1'); doc.setText(9);
    wrapText(title || 'Reporte general', 9, MAXW).forEach(function (l) { doc.row(l); });
    doc.row('Generado: ' + two(now.getDate()) + '/' + two(now.getMonth() + 1) + '/' + now.getFullYear() + '  -  ' + two(now.getHours()) + ':' + two(now.getMinutes()));
    doc.drawLine(0.8);
    doc.nextLine(0.6);

    doc.ensure(90);
    doc.setFont('F2'); doc.setText(11);
    doc.row('Resumen (' + exported.total + ' novedades)');
    function countTable(label, rows, showPct) {
        doc.ensure(80);
        doc.setFont('F2'); doc.setText(9.5);
        doc.row(label);
        doc.setFont('F1'); doc.setText(9);
        doc.row('Concepto', 'Cantidad   ' + (showPct ? ' %' : ''));
        doc.drawLine(0.4);
        rows.forEach(function (i) {
            doc.row(String(i.k), String(i.n) + (showPct ? '   ' + pct(i.n, exported.total) : ''));
        });
        doc.row('TOTAL', String(exported.total));
        doc.nextLine(0.9);
    }
    countTable('Por Categoria', exported.byCat, true);
    countTable('Por Habitacion', exported.byRoom, true);
    countTable('Por Mes', exported.byMonth, false);

    doc.ensure(120);
    doc.setFont('F2'); doc.setText(11);
    doc.row('Detalle de novedades (' + exported.total + ')');
    doc.setFont('F1');
    exported.list.forEach(function (n) {
        doc.ensure(70);
        doc.setFont('F2'); doc.setText(9.5);
        doc.row('Hab. ' + n.room + '  -  ' + n.category, String(n.dateLocal || ''));
        doc.setFont('F1'); doc.setText(9);
        wrapText(String(n.description || ''), 9, MAXW).forEach(function (l) { doc.row(l); });
        doc.nextLine(0.4);
    });

    var pageCount = doc.pages.length;
    doc.pages.forEach(function (pg, idx) {
        pg.push('BT /F1 8 Tf 306 ' + Math.round(28) + ' Td (' + pdfStr('Bitacora Dynas-IT · Pagina ' + (idx + 1) + ' de ' + pageCount) + ') Tj ET');
    });
    return assemblePdf(doc);
}

function assemblePdf(doc) {
    var objs = [];
    objs.push({ n: 1, body: '<< /Type /Catalog /Pages 2 0 R >>' });
    var kids = [], i;
    for (i = 0; i < doc.pages.length; i++) kids.push((3 + i) + ' 0 R');
    objs.push({ n: 2, body: '<< /Type /Pages /Kids [' + kids.join(' ') + '] /Count ' + doc.pages.length + ' >>' });
    var fontF1 = 3 + doc.pages.length;
    var fontF2 = fontF1 + 1;
    objs.push({ n: fontF1, body: '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>' });
    objs.push({ n: fontF2, body: '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>' });
    var next = fontF2 + 1;
    for (i = 0; i < doc.pages.length; i++) {
        var stream = doc.pages[i].join('\n');
        var contentId = next + i;
        objs.push({ n: contentId, body: '<< /Length ' + stream.length + ' >>\nstream\n' + stream + '\nendstream' });
        objs.push({ n: 3 + i, body: '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 ' + fontF1 + ' 0 R /F2 ' + fontF2 + ' 0 R >> >> /Contents ' + contentId + ' 0 R >>' });
    }
    objs.sort(function (a, b) { return a.n - b.n; });
    var out = '%PDF-1.4\n';
    var offsets = {};
    objs.forEach(function (o) {
        offsets[o.n] = Buffer.byteLength(out, 'latin1');
        out += o.n + ' 0 obj\n' + o.body + '\nendobj\n';
    });
    var xrefAt = Buffer.byteLength(out, 'latin1');
    out += 'xref\n0 ' + (objs.length + 1) + '\n0000000000 65535 f \n';
    for (i = 1; i <= objs.length; i++) out += ('0000000000' + offsets[i]).slice(-10) + ' 00000 n \n';
    out += 'trailer\n<< /Size ' + (objs.length + 1) + ' /Root 1 0 R >>\nstartxref\n' + xrefAt + '\n%%EOF';
    return Buffer.from(out, 'latin1');
}

// --- TICKETS DE EXPORTACION (archivos de descarga) ---
function createExportTicket(fmt, title, novelties) {
    const key = crypto.randomBytes(12).toString('hex');
    const sig = crypto.createHmac('sha256', AUTH_SECRET).update(key + '\u0001' + fmt).digest('hex').slice(0, 24);
    EXPORT_TICKETS.set(key, { fmt: fmt, title: title, novelties: novelties, exp: Date.now() + EXPORT_TICKET_TTL });
    return { url: '/api/export/file?k=' + key + '&s=' + sig };
}

function ticketSigOk(key, s, fmt) {
    try {
        const expected = crypto.createHmac('sha256', AUTH_SECRET).update(key + '\u0001' + fmt).digest('hex').slice(0, 24);
        return crypto.timingSafeEqual(Buffer.from(s), Buffer.from(expected));
    } catch (e) {
        return false;
    }
}

setInterval(function () {
    const now = Date.now();
    EXPORT_TICKETS.forEach(function (t, k) {
        if (now > t.exp) EXPORT_TICKETS.delete(k);
    });
}, 60000).unref();

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

    if (pathname === '/api/export/xlsx' && req.method === 'POST') {
        try {
            if (needsAuth() && !validToken(req.headers['x-token'])) {
                sendJSON(res, 401, { error: 'AUTENTICACION_REQUERIDA' });
                return;
            }
            const body = await getRequestBody(req);
            const exported = prepareExport(body.novelties);
            const buf = exportToXlsx(exported);
            const d = new Date();
            const fname = 'bitacora_' + d.getFullYear() + two(d.getMonth() + 1) + two(d.getDate()) + '_' + two(d.getHours()) + two(d.getMinutes()) + '.xlsx';
            res.writeHead(200, {
                'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'Content-Disposition': 'attachment; filename="' + fname + '"',
                'Content-Length': buf.length,
                'Cache-Control': 'no-store',
                'Access-Control-Allow-Origin': '*',
                'Access-Control-Allow-Headers': 'Content-Type, X-Token'
            });
            res.end(buf);
        } catch (e) {
            sendJSON(res, 400, { error: e.message });
        }
        return;
    }

    if (pathname === '/api/export/pdf' && req.method === 'POST') {
        try {
            if (needsAuth() && !validToken(req.headers['x-token'])) {
                sendJSON(res, 401, { error: 'AUTENTICACION_REQUERIDA' });
                return;
            }
            const body = await getRequestBody(req);
            const exported = prepareExport(body.novelties);
            const buf = buildPdf(exported, body.title);
            const d = new Date();
            const fname = 'bitacora_' + d.getFullYear() + two(d.getMonth() + 1) + two(d.getDate()) + '_' + two(d.getHours()) + two(d.getMinutes()) + '.pdf';
            res.writeHead(200, {
                'Content-Type': 'application/pdf',
                'Content-Disposition': 'attachment; filename="' + fname + '"',
                'Content-Length': buf.length,
                'Cache-Control': 'no-store',
                'Access-Control-Allow-Origin': '*',
                'Access-Control-Allow-Headers': 'Content-Type, X-Token'
            });
            res.end(buf);
        } catch (e) {
            sendJSON(res, 400, { error: e.message });
        }
        return;
    }

    // === EXPORTACION POR TICKET (descarga en celular / WebView) ===
    if (pathname === '/api/export/url/xlsx' || pathname === '/api/export/url/pdf') {
        if (req.method === 'POST') {
            try {
                if (needsAuth() && !validToken(req.headers['x-token'])) {
                    sendJSON(res, 401, { error: 'AUTENTICACION_REQUERIDA' });
                    return;
                }
                const fmt = pathname.split('/')[4];
                const body = await getRequestBody(req);
                sendJSON(res, 200, createExportTicket(fmt, body.title, body.novelties));
            } catch (e) {
                sendJSON(res, 400, { error: e.message });
            }
        }
        return;
    }

    if (pathname === '/api/export/file' && req.method === 'GET') {
        const k = url.searchParams.get('k') || '';
        const s = url.searchParams.get('s') || '';
        const t = EXPORT_TICKETS.get(k);
        if (!t) {
            res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
            res.end('404 Enlace expirado o invalido');
            return;
        }
        if (Date.now() > t.exp) {
            EXPORT_TICKETS.delete(k);
            res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
            res.end('404 Enlace expirado o invalido');
            return;
        }
        if (!ticketSigOk(k, s, t.fmt)) {
            res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
            res.end('404 Enlace expirado o invalido');
            return;
        }
        EXPORT_TICKETS.delete(k);
        try {
            const exported = prepareExport(t.novelties);
            const isPdf = t.fmt === 'pdf';
            const buf = isPdf ? buildPdf(exported, t.title) : exportToXlsx(exported);
            const d = new Date();
            const fname = 'bitacora_' + d.getFullYear() + two(d.getMonth() + 1) + two(d.getDate()) + '_' + two(d.getHours()) + two(d.getMinutes()) + (isPdf ? '.pdf' : '.xlsx');
            res.writeHead(200, {
                'Content-Type': isPdf ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'Content-Disposition': 'attachment; filename="' + fname + '"',
                'Content-Length': buf.length,
                'Cache-Control': 'no-store, no-cache',
                'Access-Control-Allow-Origin': '*',
                'Access-Control-Allow-Headers': 'Content-Type, X-Token'
            });
            res.end(buf);
        } catch (e) {
            sendJSON(res, 500, { error: e.message });
        }
        return;
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