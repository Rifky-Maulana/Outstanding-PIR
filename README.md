# Outstanding PIR — stock take coil WIP

Aplikasi web internal untuk menampilkan coil **Prime** yang masih termasuk WIP
dan belum tercatat sebagai Raw Material atau Finished Goods.

Definisi WIP yang digunakan saat ini:

1. coil merupakan generasi terakhir atau tidak mempunyai keturunan;
2. `next process` belum menunjukkan proses selesai atau barang keluar; dan
3. nomor coil belum tercatat di `bnm_stck`.

Identitas fisik ditentukan dari bagian nomor setelah prefix proses. Karena itu
pergantian prefix seperti `MD-7167-211` → `DD-7167-211` → `LD-7167-211`
tetap diperlakukan sebagai satu fisik. Relasi input/output juga dibandingkan
dengan identitas tersebut, sehingga `SB-7741-324A` dan `FB-7741-324A` tidak
lagi dianggap sebagai dua coil berbeda.

Nomor coil dibersihkan dari spasi, tab, dan karakter baris baru sebelum
dibandingkan. Jika waktu operasi terakhir sama, `prdp_nr` numerik digunakan
sebagai urutan transaksi. Coil yang lokasinya ditandai telah dihapus
berdasarkan stock take tidak ikut hasil.

Aplikasi hanya membaca database `bnm-data`. Aturan bisnis dan riwayat
investigasi yang lebih lengkap tersedia di `KONTEKS.md`.

## Fitur

- Menampilkan nomor coil, mother coil, grade, temper, dimensi, berat, line,
  next process, lokasi, asal data, dan waktu operasi.
- Pencarian berdasarkan coil, mother coil, customer, grade, temper, lokasi,
  dan next process.
- Filter khusus coil hasil potong.
- Ringkasan jumlah coil, total tonase, dan jumlah hasil potong.
- Ekspor seluruh hasil filter ke Excel (`.xlsx`).
- Cache hasil query di memori agar database tidak dibaca ulang setiap kali
  halaman dibuka.
- Tombol **Ambil ulang dari database** untuk menghapus cache dan menjalankan
  query kembali.

Tabel di browser menampilkan maksimal 2.000 baris. Ekspor Excel tetap mencakup
seluruh baris yang sesuai dengan filter.

## Persyaratan

- Node.js 18 atau lebih baru.
- Akses ke Microsoft SQL Server.
- Login SQL Server yang memiliki akses baca ke `bnm-data`.

## Menjalankan aplikasi

```bash
npm install
copy .env.example .env
npm start
```

Buka <http://localhost:3000>.

Untuk menjalankan server dalam mode watch:

```bash
npm run dev
```

## Konfigurasi

Isi `.env` sesuai lingkungan:

```env
DB_SERVER=LTP-2017-023\SQL2017
DB_DATABASE=bnm-data
DB_USER=pir_reader
DB_PASSWORD=kata_sandi
DB_ENCRYPT=false
PORT=3000

# Opsional: gunakan query CTE tanpa tabel temporer
# DB_MODE=murni
```

`DB_SERVER` dapat berupa nama host untuk default instance atau
`HOST\INSTANCE` untuk named instance.

### Mode query

| `DB_MODE` | Cara kerja | Hak database |
|---|---|---|
| Kosong/default | Tabel temporer, umumnya lebih cepat | `SELECT` dan `CREATE TABLE` |
| `murni` | CTE tanpa tabel temporer, umumnya lebih lambat | `SELECT` saja |

Kedua mode menjalankan aturan bisnis dan menghasilkan bentuk data yang sama.
Gunakan `DB_MODE=murni` jika akun produksi hanya boleh melakukan `SELECT`.

## Membuat login database

Contoh akun untuk mode default:

```sql
USE master;
CREATE LOGIN pir_reader WITH PASSWORD = 'GantiPasswordIni_123',
    CHECK_POLICY = OFF;

USE [bnm-data];
CREATE USER pir_reader FOR LOGIN pir_reader;
ALTER ROLE db_datareader ADD MEMBER pir_reader;
GRANT CREATE TABLE TO pir_reader;
```

Untuk mode `murni`, izin `GRANT CREATE TABLE` tidak diperlukan.

Driver database yang digunakan tidak memakai Windows Authentication, sehingga
SQL Server harus mengizinkan SQL Server Authentication.

## Endpoint

| Endpoint | Keterangan |
|---|---|
| `GET /api/info` | Informasi server, database, rentang data, dan jumlah operasi |
| `GET /api/coils` | Mengambil hasil dari cache atau database |
| `GET /api/coils?refresh=1` | Menjalankan ulang query dan memperbarui cache |

Cache hanya disimpan di memori dan akan kosong ketika server dimulai ulang.

## Struktur project

```text
public/index.html          Frontend satu halaman
server/index.js            Server Express dan endpoint API
server/db.js               Koneksi dan connection pool SQL Server
server/config.js           Pola next process yang dikeluarkan
server/queries.js          Query cepat menggunakan tabel temporer
server/queries-murni.js    Query CTE untuk akses SELECT-only
KONTEKS.md                 Catatan investigasi dan keputusan bisnis
```

## Pemecahan masalah

| Gejala | Penyebab umum |
|---|---|
| `ELOGIN` | User/password salah atau SQL Authentication belum aktif |
| `ESOCKET` / timeout | TCP/IP belum aktif atau server tidak dapat dijangkau |
| Named instance tidak ditemukan | Service SQL Server Browser belum berjalan |
| Query ditolak pada mode default | Akun belum memiliki izin membuat tabel temporer; gunakan `DB_MODE=murni` |

## Maintenance daftar scrap

Perubahan `prdp_nextprocess` tidak dilakukan oleh aplikasi web. Untuk daftar
stock take 22 Oktober 2026, file siap pakai tersedia di
`sql/update-scrap-702.sql`. Buka melalui SSMS menggunakan akun IT/DBA.

File SQL secara default memakai `@commit = 0`. Mode ini hanya menampilkan
preview dan tidak menjalankan `UPDATE`. Hasil wajib menunjukkan:

- daftar: 702
- direct mapping: 696
- output-only: 6
- target update: 701

Simpan result set `rollback_sql` ke file sebelum eksekusi. Setelah seluruh
hasil diperiksa, IT mengubah `@commit = 1` dan menjalankan ulang script. Hanya
bagian `UPDATE` yang dibungkus transaction agar tidak menahan lock selama
proses preview.

Untuk daftar lain, SQL dapat dibuat ulang dengan:

```bash
node scripts/generate-scrap-sql.js daftar.txt sql/update-scrap.sql
```

Alternatif melalui Node.js tersedia di `scripts/update-scrap-nextprocess.js`.

Simpan daftar nomor coil satu per baris sebagai `scrap-coils.txt`, lalu jalankan
preview menggunakan akun database read-only:

```bash
node scripts/update-scrap-nextprocess.js scrap-coils.txt
```

Preview memetakan setiap coil ke transaksi terakhir. Coil output yang belum
memiliki baris sendiri akan dipetakan ke baris operasi induknya. Tidak ada data
yang berubah pada mode preview.

Setelah hasil diperiksa, IT/DBA dapat memakai akun yang memiliki izin `UPDATE`
dan menjalankan:

```bash
node scripts/update-scrap-nextprocess.js scrap-coils.txt --commit
```

Mode commit menggunakan transaksi, memvalidasi jumlah target, dan menyimpan
backup JSON lokal dalam folder `backups`. Rollback dijalankan dengan:

```bash
node scripts/rollback-scrap-nextprocess.js backups/nama-backup.json
```

Jangan memberikan izin write kepada akun aplikasi web. Gunakan kredensial
maintenance sementara milik IT/DBA, lalu kembalikan `.env` ke akun read-only.

## Batasan data

- Hanya data dengan `prdp_type = 'Prime'` yang disertakan.
- Status `WIP2` tidak digunakan sebagai penentu WIP.
- Nilai berat dapat kosong jika tidak tersedia pada operasi coil maupun operasi
  induknya. UI akan menampilkan jumlah coil tanpa berat pada ringkasan tonase.
- Kualitas hasil bergantung pada konsistensi nomor coil dan nilai bebas pada
  kolom `prdp_nextprocess`.
