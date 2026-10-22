// Aturan bisnis Outstanding PIR.
//
// Ditaruh di sini, bukan di dalam string SQL. Sistem lama menanam aturan
// langsung di SQL dan itu membuatnya tidak bisa diubah tanpa membongkar query.
//
// Tidak boleh menambah tabel di `bnm-data` — akses ke sana harus tetap
// read-only. Kalau nanti aturan ini perlu diubah lewat UI, pindahkan ke
// database aplikasi sendiri, bukan ke bnm-data.

// Next process yang berarti coil TIDAK IKUT stock take.
//
// Dua sebab coil dikeluarkan:
//   (a) sudah selesai / sudah keluar pabrik
//   (b) barisnya bukan coil sama sekali
//
// `prdp_nextprocess` adalah teks bebas tanpa validasi, jadi memakai daftar
// nilai persis akan bocor. Contoh variasi yang benar-benar ada di data
// (per backup Maret 2026):
//
//   PACKING (55.923)          FG (49.178)         CUST (66.082)
//   Packing - FGR Cancel      FG - FGR Cancel     CUST - TRANSFER
//   Packing - REPROSES        FG - REPROSES       CUST-TRANSFER
//   PACKINGReprocess          FG - TRANSFER       CUST - RETURN
//   Packing-TRANSFER          FG-TRANSFER         CUST - Reproses
//   pack                      Packing -  Reprocess (spasi dobel)
//   Customer - TRANSFER
//
// Semuanya bermuara ke barang keluar — sudah dicek: alur "REPROSES" berakhir
// di `FG - REPROSES`, dan `CUST - RETURN` (barang balik dari customer) pun
// berakhir di `Packing` atau `CUST` lagi setelah diproses ulang. Karena itu
// dipakai pola, bukan daftar nilai.
//
// Perbandingan case-insensitive mengikuti collation SQL Server.
const POLA_SELESAI = [
  '%pack%',    // PACKING, pack, Packing - REPROSES, PACKINGReprocess, ...
  'FG%',       // FG, FG - FGR Cancel, FG-TRANSFER, ...
  '%deliv%',   // Delivery
  '%cust%',    // CUST = sudah sampai customer. Juga Customer - TRANSFER.
  '%nego%',    // NEGO = sedang dinegosiasi, bukan stock
  '%cancel%',  // CANCEL = dibatalkan
  '%maint%',   // MAINTENANCE = catatan kegiatan, barisnya bukan coil
  '%trial%',   // TRIAL = uji coba, bukan coil stock take
];

// Dipakai membangun potongan WHERE. Diletakkan di sini supaya penambahan
// pola tidak perlu menyentuh queries.js.
function klausaBukanSelesai(kolom) {
  return POLA_SELESAI.map((p) => `${kolom} NOT LIKE '${p}'`).join('\n    AND ');
}

function sqlCoilBersih(kolom) {
  return `REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(CAST(${kolom} AS varchar(100)))), CHAR(13), ''), CHAR(10), ''), CHAR(9), '')`;
}

function sqlIdFisik(kolom) {
  const coil = sqlCoilBersih(kolom);
  return `CASE WHEN CHARINDEX('-', ${coil}) = 3 THEN SUBSTRING(${coil}, 4, 100) ELSE ${coil} END`;
}

module.exports = { POLA_SELESAI, klausaBukanSelesai, sqlCoilBersih, sqlIdFisik };
