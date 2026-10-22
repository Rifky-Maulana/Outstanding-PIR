require('dotenv').config();

const fs = require('fs');
const path = require('path');
const db = require('../server/db');
const { sqlCoilBersih } = require('../server/config');

const backupArg = process.argv[2];
if (!backupArg) {
  console.error('Pemakaian: node scripts/rollback-scrap-nextprocess.js <backup.json>');
  process.exit(1);
}

const backup = JSON.parse(fs.readFileSync(path.resolve(backupArg), 'utf8'));
if (!Array.isArray(backup.targets) || !backup.targets.length) throw new Error('Backup tidak valid.');

const values = backup.targets.map((row) => {
  const coil = String(row.coil_db).replace(/'/g, "''");
  const nr = String(row.prdp_nr).replace(/'/g, "''");
  const oldValue = row.prdp_nextprocess == null
    ? 'NULL'
    : `'${String(row.prdp_nextprocess).replace(/'/g, "''")}'`;
  return `('${coil}', '${nr}', ${oldValue})`;
}).join(',\n');

async function main() {
  const pool = await db.getPool();
  const transaction = new db.sql.Transaction(pool);
  await transaction.begin(db.sql.ISOLATION_LEVEL.READ_COMMITTED);
  try {
    const result = await new db.sql.Request(transaction).query(`
      CREATE TABLE #rollback (coil_db varchar(100), prdp_nr varchar(100), old_nextprocess varchar(255) NULL);
      INSERT INTO #rollback VALUES ${values};
      UPDATE p
      SET p.prdp_nextprocess = r.old_nextprocess
      OUTPUT deleted.prdp_nr, deleted.prdp_coilno,
             deleted.prdp_nextprocess AS from_value,
             inserted.prdp_nextprocess AS restored_value
      FROM dbo.bnm_prdp p
      JOIN #rollback r
        ON p.prdp_nr = r.prdp_nr
       AND ${sqlCoilBersih('p.prdp_coilno')} = r.coil_db
      WHERE p.prdp_nextprocess = 'PACK';
    `);
    if (result.rowsAffected[0] !== backup.targets.length) {
      throw new Error(`Backup ${backup.targets.length}, tetapi yang dipulihkan ${result.rowsAffected[0]}.`);
    }
    await transaction.commit();
    console.log(`Rollback berhasil: ${result.rowsAffected[0]} baris dipulihkan.`);
  } catch (error) {
    if (!transaction._aborted) await transaction.rollback();
    throw error;
  }
}

main().then(() => process.exit(0)).catch((error) => {
  console.error(error.message);
  process.exit(1);
});
