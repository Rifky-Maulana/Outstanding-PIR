require('dotenv').config();

const fs = require('fs');
const path = require('path');
const db = require('../server/db');
const { sqlCoilBersih, sqlIdFisik } = require('../server/config');

const args = process.argv.slice(2);
const commit = args.includes('--commit');
const inputArg = args.find((arg) => arg !== '--commit');

if (!inputArg) {
  console.error('Pemakaian: node scripts/update-scrap-nextprocess.js <daftar.txt> [--commit]');
  process.exit(1);
}

const inputPath = path.resolve(inputArg);
const coils = fs.readFileSync(inputPath, 'utf8')
  .split(/\r?\n/)
  .map((line) => line.trim())
  .filter(Boolean);

if (!coils.length) throw new Error('Daftar coil kosong.');
if (new Set(coils.map((coil) => coil.toUpperCase())).size !== coils.length) {
  throw new Error('Daftar mengandung nomor coil duplikat.');
}
if (coils.some((coil) => !/^[A-Za-z]{2}-[A-Za-z0-9-]+$/.test(coil))) {
  throw new Error('Daftar mengandung format nomor coil yang tidak dikenal.');
}

const values = coils.map((coil) => {
  const escaped = coil.replace(/'/g, "''");
  const id = coil.slice(3).toUpperCase().replace(/'/g, "''");
  return `('${escaped}', '${id}')`;
}).join(',\n');

const prepareSql = `
CREATE TABLE #daftar (
  coil varchar(100) NOT NULL,
  id_fisik varchar(100) NOT NULL PRIMARY KEY
);
INSERT INTO #daftar (coil, id_fisik) VALUES ${values};

WITH calon AS (
  SELECT d.coil AS coil_daftar, d.id_fisik,
         ${sqlCoilBersih('p.prdp_coilno')} AS coil_db,
         p.prdp_nr, p.prdp_nextprocess, p.prdp_datestart,
         ROW_NUMBER() OVER (
           PARTITION BY d.id_fisik
           ORDER BY p.prdp_datestart DESC,
                    TRY_CONVERT(bigint, p.prdp_nr) DESC,
                    p.prdp_nr DESC
         ) AS rn
  FROM #daftar d
  JOIN dbo.bnm_prdp p
    ON ${sqlIdFisik('p.prdp_coilno')} = d.id_fisik
   AND p.prdp_type = 'Prime'
)
SELECT coil_daftar, id_fisik, coil_db, prdp_nr, prdp_nextprocess, prdp_datestart
INTO #langsung
FROM calon WHERE rn = 1;

WITH sumber AS (
  SELECT d.coil AS coil_daftar, d.id_fisik,
         ${sqlCoilBersih('p.prdp_coilno')} AS coil_db,
         p.prdp_nr, p.prdp_nextprocess, p.prdp_datestart,
         ROW_NUMBER() OVER (
           PARTITION BY d.id_fisik
           ORDER BY p.prdp_datestart DESC,
                    TRY_CONVERT(bigint, p.prdp_nr) DESC,
                    p.prdp_nr DESC
         ) AS rn
  FROM #daftar d
  JOIN dbo.bnm_prdp p
    ON p.prdp_type = 'Prime'
  CROSS APPLY STRING_SPLIT(REPLACE(ISNULL(p.prdp_coilnonext, ''), ',', ';'), ';') s
  WHERE NOT EXISTS (SELECT 1 FROM #langsung l WHERE l.id_fisik = d.id_fisik)
    AND ${sqlCoilBersih('s.value')} <> ''
    AND ${sqlIdFisik('s.value')} = d.id_fisik
)
SELECT coil_daftar, id_fisik, coil_db, prdp_nr, prdp_nextprocess, prdp_datestart
INTO #sumber
FROM sumber WHERE rn = 1;

SELECT coil_daftar, id_fisik, coil_db, prdp_nr, prdp_nextprocess,
       prdp_datestart, CAST('latest_input' AS varchar(20)) AS mapping_type
INTO #mapping
FROM #langsung
UNION ALL
SELECT coil_daftar, id_fisik, coil_db, prdp_nr, prdp_nextprocess,
       prdp_datestart, CAST('source_parent' AS varchar(20)) AS mapping_type
FROM #sumber;

SELECT DISTINCT coil_db, prdp_nr, prdp_nextprocess
INTO #target
FROM #mapping;

SELECT * FROM #mapping ORDER BY mapping_type, coil_daftar;
SELECT * FROM #target ORDER BY TRY_CONVERT(bigint, prdp_nr), prdp_nr;
SELECT d.* FROM #daftar d
WHERE NOT EXISTS (SELECT 1 FROM #mapping m WHERE m.id_fisik = d.id_fisik)
ORDER BY d.coil;
`;

function inspectPrepared(prepared) {
    const mappings = prepared.recordsets[0];
    const targets = prepared.recordsets[1];
    const unresolved = prepared.recordsets[2];
    const direct = mappings.filter((row) => row.mapping_type === 'latest_input').length;
    const sourceOnly = mappings.filter((row) => row.mapping_type === 'source_parent').length;

    console.log({ input: coils.length, mappings: mappings.length, direct, sourceOnly,
      targetRows: targets.length, unresolved: unresolved.length, mode: commit ? 'COMMIT' : 'PREVIEW' });
    if (unresolved.length) {
      console.table(unresolved);
      throw new Error('Ada coil yang tidak dapat dipetakan. Update dibatalkan.');
    }
    if (mappings.length !== coils.length) {
      throw new Error(`Jumlah mapping ${mappings.length} tidak sama dengan input ${coils.length}.`);
    }
    console.table(mappings.filter((row) => row.mapping_type === 'source_parent'));
    return { mappings, targets };
}

async function main() {
  const pool = await db.getPool();

  if (!commit) {
    const prepared = await pool.request().query(prepareSql);
    inspectPrepared(prepared);
    console.log('PREVIEW selesai. Tidak ada data yang diubah.');
    return;
  }

  const transaction = new db.sql.Transaction(pool);
  await transaction.begin(db.sql.ISOLATION_LEVEL.READ_COMMITTED);

  try {
    const prepared = await new db.sql.Request(transaction).query(prepareSql);
    const { mappings, targets } = inspectPrepared(prepared);

    const backupDir = path.join(__dirname, '..', 'backups');
    fs.mkdirSync(backupDir, { recursive: true });
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    const backupPath = path.join(backupDir, `scrap-update-${stamp}.json`);
    fs.writeFileSync(backupPath, JSON.stringify({ createdAt: new Date().toISOString(), inputPath,
      mappings, targets }, null, 2));

    const updated = await new db.sql.Request(transaction).query(`
      UPDATE p
      SET p.prdp_nextprocess = 'PACK'
      OUTPUT deleted.prdp_nr, deleted.prdp_coilno,
             deleted.prdp_nextprocess AS old_nextprocess,
             inserted.prdp_nextprocess AS new_nextprocess
      FROM dbo.bnm_prdp p
      JOIN #target t
        ON p.prdp_nr = t.prdp_nr
       AND ${sqlCoilBersih('p.prdp_coilno')} = t.coil_db
      WHERE p.prdp_type = 'Prime';
    `);
    if (updated.rowsAffected[0] !== targets.length) {
      throw new Error(`Target ${targets.length}, tetapi yang berubah ${updated.rowsAffected[0]}.`);
    }

    await transaction.commit();
    console.log(`COMMIT berhasil: ${updated.rowsAffected[0]} baris diubah menjadi PACK.`);
    console.log(`Backup rollback: ${backupPath}`);
  } catch (error) {
    if (!transaction._aborted) await transaction.rollback();
    throw error;
  }
}

main().then(() => process.exit(0)).catch((error) => {
  console.error(error.message);
  process.exit(1);
});
