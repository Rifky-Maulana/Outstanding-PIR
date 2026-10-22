const sql = require('mssql');

// "LTP-2017-023\SQL2017" harus dipecah jadi host + instanceName.
// Kalau tidak ada backslash, berarti instance default.
function parseServer(raw) {
  const value = (raw || '').trim();
  const idx = value.indexOf('\\');
  if (idx === -1) return { server: value, instanceName: null };
  return {
    server: value.slice(0, idx),
    instanceName: value.slice(idx + 1),
  };
}

const { server, instanceName } = parseServer(process.env.DB_SERVER);

const config = {
  server,
  database: process.env.DB_DATABASE,
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  options: {
    encrypt: String(process.env.DB_ENCRYPT).toLowerCase() === 'true',
    trustServerCertificate: true,
    enableArithAbort: true,
  },
  pool: { max: 4, min: 0, idleTimeoutMillis: 30000 },
  requestTimeout: 180000,
  connectionTimeout: 30000,
};

if (instanceName) {
  config.options.instanceName = instanceName;
}

let poolPromise = null;

function getPool() {
  if (!poolPromise) {
    poolPromise = new sql.ConnectionPool(config)
      .connect()
      .catch((err) => {
        poolPromise = null;
        throw err;
      });
  }
  return poolPromise;
}

async function query(text) {
  const pool = await getPool();
  const result = await pool.request().query(text);
  return result;
}

module.exports = { getPool, query, config, sql };
