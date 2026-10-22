// Versi SELECT murni dari query stock take. Aturan bisnis harus selalu sama
// dengan queries.js; file ini hanya tidak memakai tabel temporer.

const { klausaBukanSelesai, sqlCoilBersih, sqlIdFisik } = require('./config');

const STOCK_TAKE = `
SET NOCOUNT ON;
WITH pecah AS (
    SELECT
        ${sqlCoilBersih('s.value')} AS anak,
        ${sqlIdFisik('s.value')} AS anak_id,
        ${sqlCoilBersih('p.prdp_coilno')} AS induk,
        ${sqlIdFisik('p.prdp_coilno')} AS induk_id,
        NULLIF(p.prdp_weoutputkg, 0) AS w_out,
        NULLIF(p.prdp_weinput, 0) AS w_in,
        NULLIF(p.prdp_toutputmm, 0) AS t_out,
        NULLIF(p.prdp_tinput, 0) AS t_in,
        NULLIF(p.prdp_wioutputmm, 0) AS l_out,
        NULLIF(p.prdp_wiinput, 0) AS l_in,
        NULLIF(LTRIM(RTRIM(p.prdp_location)), '') AS lok,
        ISNULL(NULLIF(LTRIM(RTRIM(p.prdp_nextprocess)), ''), '(kosong)') AS np,
        p.prdp_grade AS grade,
        p.prdp_temper AS temper,
        NULLIF(LTRIM(RTRIM(p.prdp_custnm)), '') AS customer,
        p.prdp_linenex AS line_cur,
        p.prdp_coilnomother AS mother,
        p.prdp_datestart AS waktu,
        p.prdp_nr AS nr
    FROM dbo.bnm_prdp p
    CROSS APPLY STRING_SPLIT(REPLACE(ISNULL(p.prdp_coilnonext, ''), ',', ';'), ';') s
    WHERE p.prdp_type = 'Prime'
      AND LTRIM(RTRIM(s.value)) <> ''
),
own AS (
    SELECT * FROM (
        SELECT
            ${sqlCoilBersih('prdp_coilno')} AS coil,
            ${sqlIdFisik('prdp_coilno')} AS id_fisik,
            NULLIF(prdp_weoutputkg, 0) AS w_out,
            NULLIF(prdp_weinput, 0) AS w_in,
            NULLIF(prdp_toutputmm, 0) AS t_out,
            NULLIF(prdp_tinput, 0) AS t_in,
            NULLIF(prdp_wioutputmm, 0) AS l_out,
            NULLIF(prdp_wiinput, 0) AS l_in,
            NULLIF(LTRIM(RTRIM(prdp_location)), '') AS lok,
            ISNULL(NULLIF(LTRIM(RTRIM(prdp_nextprocess)), ''), '(kosong)') AS np,
            prdp_grade AS grade,
            prdp_temper AS temper,
            NULLIF(LTRIM(RTRIM(prdp_custnm)), '') AS customer,
            prdp_linecur AS line_cur,
            prdp_coilnomother AS mother,
            prdp_status AS status,
            prdp_datestart AS waktu,
            prdp_nr AS nr,
            ROW_NUMBER() OVER (
                PARTITION BY ${sqlIdFisik('prdp_coilno')}
                ORDER BY prdp_datestart DESC,
                         TRY_CONVERT(bigint, prdp_nr) DESC,
                         prdp_nr DESC
            ) AS rn
        FROM dbo.bnm_prdp
        WHERE prdp_type = 'Prime'
    ) z WHERE rn = 1
),
src AS (
    SELECT * FROM (
        SELECT *, ROW_NUMBER() OVER (
            PARTITION BY anak_id
            ORDER BY waktu DESC, TRY_CONVERT(bigint, nr) DESC, nr DESC
        ) AS rn_src
        FROM pecah
    ) z WHERE rn_src = 1
),
node_dasar AS (
    SELECT id_fisik, MAX(waktu) AS waktu
    FROM (
        SELECT id_fisik, waktu FROM own
        UNION ALL
        SELECT anak_id AS id_fisik, waktu FROM src
    ) n
    GROUP BY id_fisik
),
node AS (
    SELECT id_fisik, waktu,
           CASE WHEN RIGHT(id_fisik, 1) LIKE '[A-Z]'
                THEN LEFT(id_fisik, LEN(id_fisik) - 1) END AS parent_id
    FROM node_dasar
),
anak AS (
    SELECT DISTINCT induk_id AS id_fisik FROM pecah WHERE anak_id <> induk_id
    UNION
    SELECT DISTINCT induk.id_fisik
    FROM node induk
    JOIN node anak_node
      ON anak_node.parent_id = induk.id_fisik
     AND anak_node.waktu > induk.waktu
),
ujung AS (
    SELECT n.id_fisik FROM node n
    WHERE NOT EXISTS (SELECT 1 FROM anak a WHERE a.id_fisik = n.id_fisik)
),
stck AS (
    SELECT DISTINCT
        ${sqlIdFisik('stck_artsernr')} AS id_fisik
    FROM dbo.bnm_stck
    WHERE ISNULL(stck_artsernr, '') <> ''
)
SELECT
    COALESCE(o.coil, sr.anak) AS coil,
    COALESCE(o.grade, sr.grade) AS grade,
    COALESCE(o.temper, sr.temper) AS temper,
    COALESCE(o.customer, sr.customer) AS customer,
    COALESCE(o.t_out, o.t_in, sr.t_out, sr.t_in) AS tebal,
    COALESCE(o.l_out, o.l_in, sr.l_out, sr.l_in) AS lebar,
    COALESCE(o.w_out, o.w_in, sr.w_out, sr.w_in) AS berat,
    COALESCE(o.line_cur, sr.line_cur) AS line_skrg,
    COALESCE(o.np, sr.np) AS next_process,
    COALESCE(o.lok, sr.lok) AS lokasi,
    COALESCE(o.mother, sr.mother) AS mother,
    o.status AS status,
    COALESCE(o.waktu, sr.waktu) AS waktu,
    CASE WHEN o.id_fisik IS NULL THEN 1 ELSE 0 END AS hasil_potong
FROM ujung u
LEFT JOIN own o ON o.id_fisik = u.id_fisik
LEFT JOIN src sr ON sr.anak_id = u.id_fisik
WHERE NOT EXISTS (SELECT 1 FROM stck st WHERE st.id_fisik = u.id_fisik)
  AND COALESCE(o.lok, sr.lok, '') NOT LIKE '%hapus%'
  AND ${klausaBukanSelesai('COALESCE(o.np, sr.np)')}
ORDER BY COALESCE(o.waktu, sr.waktu) DESC;
`;

module.exports = { STOCK_TAKE };
