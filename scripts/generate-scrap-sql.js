const fs = require('fs');
const path = require('path');

const inputArg = process.argv[2];
const outputArg = process.argv[3] || 'sql/update-scrap-702.sql';
if (!inputArg) {
  console.error('Pemakaian: node scripts/generate-scrap-sql.js <daftar.txt> [output.sql]');
  process.exit(1);
}

const coils = fs.readFileSync(path.resolve(inputArg), 'utf8')
  .split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
if (coils.length !== 702) throw new Error(`Diharapkan 702 coil, ditemukan ${coils.length}.`);
if (new Set(coils.map((coil) => coil.toUpperCase())).size !== coils.length) throw new Error('Ada coil duplikat.');
if (coils.some((coil) => !/^[A-Za-z]{2}-[A-Za-z0-9-]+$/.test(coil))) throw new Error('Format coil tidak valid.');

const list = coils.join('\n').replace(/'/g, "''");
const sql = `USE [bnm-data];
GO

SET NOCOUNT ON;
SET XACT_ABORT ON;

/*
  DEFAULT AMAN: @commit = 0 hanya preview, tanpa membuka transaksi UPDATE.
  IT/DBA mengubah menjadi 1 hanya setelah seluruh hasil preview diperiksa.

  Hasil audit 22 Oktober 2026:
    702 coil daftar
    696 memiliki baris input sendiri
      6 output-only, dipetakan ke 5 baris induk
    701 baris bnm_prdp menjadi target UPDATE
*/
DECLARE @commit bit = 0;
DECLARE @daftar_text varchar(max) = '${list}';

CREATE TABLE #daftar (
    coil varchar(100) NOT NULL,
    id_fisik varchar(100) NOT NULL PRIMARY KEY
);

INSERT INTO #daftar (coil, id_fisik)
SELECT coil,
       CASE WHEN CHARINDEX('-', coil) = 3 THEN SUBSTRING(coil, 4, 100) ELSE coil END
FROM (
    SELECT REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(value)), CHAR(13), ''), CHAR(10), ''), CHAR(9), '') AS coil
    FROM STRING_SPLIT(REPLACE(@daftar_text, CHAR(13), ''), CHAR(10))
) x
WHERE coil <> '';

IF (SELECT COUNT(*) FROM #daftar) <> 702
    THROW 50001, 'Validasi gagal: daftar harus berisi 702 id fisik unik.', 1;

WITH calon AS (
    SELECT d.coil AS coil_daftar, d.id_fisik,
           REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(CAST(p.prdp_coilno AS varchar(100)))), CHAR(13), ''), CHAR(10), ''), CHAR(9), '') AS coil_db,
           p.prdp_nr, p.prdp_nextprocess, p.prdp_datestart,
           ROW_NUMBER() OVER (
               PARTITION BY d.id_fisik
               ORDER BY p.prdp_datestart DESC,
                        TRY_CONVERT(bigint, p.prdp_nr) DESC,
                        p.prdp_nr DESC
           ) AS rn
    FROM #daftar d
    JOIN dbo.bnm_prdp p
      ON CASE WHEN CHARINDEX('-', REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(CAST(p.prdp_coilno AS varchar(100)))), CHAR(13), ''), CHAR(10), ''), CHAR(9), '')) = 3
              THEN SUBSTRING(REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(CAST(p.prdp_coilno AS varchar(100)))), CHAR(13), ''), CHAR(10), ''), CHAR(9), ''), 4, 100)
              ELSE REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(CAST(p.prdp_coilno AS varchar(100)))), CHAR(13), ''), CHAR(10), ''), CHAR(9), '') END = d.id_fisik
     AND p.prdp_type = 'Prime'
)
SELECT coil_daftar, id_fisik, coil_db, prdp_nr, prdp_nextprocess, prdp_datestart
INTO #langsung
FROM calon WHERE rn = 1;

WITH sumber AS (
    SELECT d.coil AS coil_daftar, d.id_fisik,
           REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(CAST(p.prdp_coilno AS varchar(100)))), CHAR(13), ''), CHAR(10), ''), CHAR(9), '') AS coil_db,
           p.prdp_nr, p.prdp_nextprocess, p.prdp_datestart,
           ROW_NUMBER() OVER (
               PARTITION BY d.id_fisik
               ORDER BY p.prdp_datestart DESC,
                        TRY_CONVERT(bigint, p.prdp_nr) DESC,
                        p.prdp_nr DESC
           ) AS rn
    FROM #daftar d
    JOIN dbo.bnm_prdp p ON p.prdp_type = 'Prime'
    CROSS APPLY STRING_SPLIT(REPLACE(ISNULL(p.prdp_coilnonext, ''), ',', ';'), ';') s
    WHERE NOT EXISTS (SELECT 1 FROM #langsung l WHERE l.id_fisik = d.id_fisik)
      AND REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(CAST(s.value AS varchar(100)))), CHAR(13), ''), CHAR(10), ''), CHAR(9), '') <> ''
      AND CASE WHEN CHARINDEX('-', REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(CAST(s.value AS varchar(100)))), CHAR(13), ''), CHAR(10), ''), CHAR(9), '')) = 3
               THEN SUBSTRING(REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(CAST(s.value AS varchar(100)))), CHAR(13), ''), CHAR(10), ''), CHAR(9), ''), 4, 100)
               ELSE REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(CAST(s.value AS varchar(100)))), CHAR(13), ''), CHAR(10), ''), CHAR(9), '') END = d.id_fisik
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

SELECT DISTINCT coil_db, prdp_nr, prdp_nextprocess AS old_nextprocess
INTO #target
FROM #mapping;

IF (SELECT COUNT(*) FROM #mapping) <> 702
    THROW 50002, 'Validasi gagal: tidak seluruh 702 coil dapat dipetakan.', 1;
IF (SELECT COUNT(*) FROM #langsung) <> 696
    THROW 50003, 'Validasi gagal: jumlah direct mapping berubah dari 696.', 1;
IF (SELECT COUNT(*) FROM #sumber) <> 6
    THROW 50004, 'Validasi gagal: jumlah output-only berubah dari 6.', 1;
IF (SELECT COUNT(*) FROM #target) <> 701
    THROW 50005, 'Validasi gagal: target UPDATE harus tepat 701 baris.', 1;

/* HASIL 1: ringkasan wajib sama dengan 702 / 696 / 6 / 701. */
SELECT
    (SELECT COUNT(*) FROM #daftar) AS daftar,
    (SELECT COUNT(*) FROM #langsung) AS direct_mapping,
    (SELECT COUNT(*) FROM #sumber) AS output_only,
    (SELECT COUNT(*) FROM #target) AS target_update,
    CASE WHEN @commit = 1 THEN 'COMMIT' ELSE 'PREVIEW / NO UPDATE' END AS mode;

/* HASIL 2: enam output-only yang memakai lima baris induk. */
SELECT * FROM #mapping WHERE mapping_type = 'source_parent' ORDER BY coil_daftar;

/* HASIL 3: semua baris yang akan diubah. */
SELECT * FROM #target ORDER BY TRY_CONVERT(bigint, prdp_nr), prdp_nr;

/* HASIL 4: simpan hasil ini sebagai rollback.sql sebelum COMMIT. */
SELECT 'UPDATE dbo.bnm_prdp SET prdp_nextprocess = '
       + CASE WHEN old_nextprocess IS NULL THEN 'NULL'
              ELSE '''' + REPLACE(old_nextprocess, '''', '''''') + '''' END
       + ' WHERE prdp_nr = ''' + REPLACE(CAST(prdp_nr AS varchar(100)), '''', '''''')
       + ''' AND REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(CAST(prdp_coilno AS varchar(100)))), CHAR(13), ''''), CHAR(10), ''''), CHAR(9), '''') = '''
       + REPLACE(coil_db, '''', '''''') + ''' AND prdp_nextprocess = ''PACK'';' AS rollback_sql
FROM #target
ORDER BY TRY_CONVERT(bigint, prdp_nr), prdp_nr;

IF @commit = 1
BEGIN
    BEGIN TRANSACTION;

    UPDATE p
       SET p.prdp_nextprocess = 'PACK'
    OUTPUT deleted.prdp_nr, deleted.prdp_coilno,
           deleted.prdp_nextprocess AS old_nextprocess,
           inserted.prdp_nextprocess AS new_nextprocess
    FROM dbo.bnm_prdp p
    JOIN #target t
      ON p.prdp_nr = t.prdp_nr
     AND REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(CAST(p.prdp_coilno AS varchar(100)))), CHAR(13), ''), CHAR(10), ''), CHAR(9), '') = t.coil_db
    WHERE p.prdp_type = 'Prime'
      AND ISNULL(p.prdp_nextprocess, '') = ISNULL(t.old_nextprocess, '');

    IF @@ROWCOUNT <> 701
        THROW 50006, 'UPDATE tidak mengenai tepat 701 baris. Transaksi dibatalkan.', 1;

    COMMIT TRANSACTION;
    PRINT 'COMMIT BERHASIL: 701 baris diubah menjadi PACK.';
END
ELSE
BEGIN
    PRINT 'PREVIEW SELESAI: tidak ada data yang diubah.';
END;
GO
`;

const outputPath = path.resolve(outputArg);
fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, sql, 'utf8');
console.log(`SQL dibuat: ${outputPath}`);
