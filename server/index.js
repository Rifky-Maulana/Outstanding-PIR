require('dotenv').config();

const path = require('path');
const express = require('express');
const db = require('./db');
const { STOCK_TAKE, INFO } = require('./queries');
const murni = require('./queries-murni');

// Dua cara menjalankan query yang sama:
//   default  tabel temporer  -> lebih cepat, butuh izin CREATE TABLE
//   murni    CTE             -> lebih lambat, cukup izin SELECT
// Pakai DB_MODE=murni kalau login ke database hanya diberi hak baca.
const MODE_MURNI = String(process.env.DB_MODE).toLowerCase() === 'murni';
const QUERY = MODE_MURNI ? murni.STOCK_TAKE : STOCK_TAKE;

const app = express();
app.use('/vendor', express.static(path.join(__dirname, '..', 'node_modules', 'exceljs', 'dist')));
app.use(express.static(path.join(__dirname, '..', 'public')));

// Hasil di-cache di memori supaya tabel besar tidak dijalankan ulang tiap klik.
let cache = { at: null, rows: null };

function jelaskanError(err) {
  const kode = err.code || '';
  if (kode === 'ELOGIN') {
    return 'Login ditolak. Cek DB_USER dan DB_PASSWORD di .env, dan pastikan SQL Server mengizinkan SQL Authentication.';
  }
  if (kode === 'ESOCKET' || kode === 'ETIMEOUT') {
    return 'Tidak bisa menjangkau server. Cek DB_SERVER, pastikan service SQL Server Browser jalan (untuk named instance) dan TCP/IP aktif di SQL Server Configuration Manager.';
  }
  if (kode === 'EREQUEST') {
    return `SQL menolak query: ${err.message}`;
  }
  return err.message;
}

app.get('/api/info', async (req, res) => {
  try {
    const hasil = await db.query(INFO);
    res.json({ ok: true, info: hasil.recordset[0], cached_at: cache.at });
  } catch (err) {
    res.status(500).json({ ok: false, error: jelaskanError(err) });
  }
});

app.get('/api/coils', async (req, res) => {
  try {
    if (req.query.refresh === '1' || !cache.rows) {
      const hasil = await db.query(QUERY);
      // Batch menghasilkan beberapa recordset; ambil yang paling besar.
      const sets = hasil.recordsets || [hasil.recordset];
      const rows = sets.reduce((a, b) => (b && b.length > (a ? a.length : 0) ? b : a), []);
      cache = { at: new Date().toISOString(), rows };
    }
    res.json({ ok: true, cached_at: cache.at, rows: cache.rows });
  } catch (err) {
    res.status(500).json({ ok: false, error: jelaskanError(err) });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Outstanding PIR jalan di http://localhost:${port}`);
  console.log(`Server DB: ${process.env.DB_SERVER}  DB: ${process.env.DB_DATABASE}`);
  console.log(`Mode query: ${MODE_MURNI ? 'SELECT murni (tanpa tabel temporer)' : 'tabel temporer'}`);
});
