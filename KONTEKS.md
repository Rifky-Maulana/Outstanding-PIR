# Outstanding PIR — Konteks Proyek

Dokumen serah-terima. Berisi latar belakang, hasil investigasi, dan pekerjaan
yang belum selesai. Ditulis supaya siapa pun (termasuk AI assistant) bisa masuk
ke proyek ini tanpa mengulang penyelidikan dari nol.

---

## 0. Tujuan aplikasi — baca ini dulu

Aplikasi web untuk **stock take bulanan**. Mendaftar coil yang masih ada
fisiknya di pabrik, beserta tonasenya.

### Definisi WIP menurut pemilik proses

> WIP adalah **coil paling akhir dari keturunannya** dan masih WIP.
> Kalau dia belum delivery dan next process-nya bukan Pack, biasanya dia WIP.

Kunci yang sebelumnya salah dipahami: **next process TIDAK bisa dipakai sebagai
penentu utama.** Di database, coil induk ikut tersimpan datanya dan tetap punya
next process, padahal fisiknya sudah tidak ada.

Contoh dari pemilik proses:

```
1 Jan   9520-100   next process BAL
2 Jan   9520-110, 9520-120     <- anak dari 9520-100
```

Ketiganya ada di database. Ketiganya terlihat seperti WIP kalau dilihat dari
next process. Tapi yang benar-benar ada fisiknya hanya `-110` dan `-120`.
`9520-100` sudah habis jadi kedua anak itu.

**Yang diambil: yang paling akhir.**

### Urutan penyaringan yang benar

1. **Ambil coil generasi terakhir** — coil yang tidak punya keturunan.
   Ini penentu utama, bukan next process.
2. **Buang yang sudah selesai** — next process `Pack` / delivery.
   Ini penyaring kedua, dipakai setelah langkah 1.

### Kenapa ini krusial

Kalau induk ikut terhitung, **tonase dobel**: berat induk + berat semua anaknya,
padahal barangnya sama. Stock take jadi salah.

### Konsekuensi untuk bug Berat

Coil generasi terakhir sebagian besar adalah **coil hasil potong** — dan
beratnya persis yang sekarang kosong di aplikasi (lihat bagian 9 dan 14).
Jadi bug Berat bukan masalah pinggiran; itu memblokir angka tonase yang
menjadi inti stock take.

---

## 1. Latar belakang

Ada aplikasi Excel VBA warisan di sebuah pabrik baja (coil stainless).
Developer sebelumnya sudah resign tanpa meninggalkan dokumentasi. Saya masuk
setelahnya dan ditugaskan melanjutkan modul **Outstanding PIR**.

File utamanya: `FG Status Report Rev.17 RV.xlsm` (ada juga varian `Rev.64 DV.xlsm`).

Rencana: pindahkan modul Outstanding PIR dari Excel ke aplikasi web.

### Apa itu Outstanding PIR

Laporan untuk **stock take**: mendaftar semua coil yang masih WIP (Work In
Progress) — masih dalam proses produksi, belum jadi Finished Good. Coil yang
sudah FG tidak ikut dihitung.

---

## 2. Arsitektur sistem lama

**Database: SQL Server.** Dua database yang dipakai:
- `bnm-data` — data produksi utama
- `ww-bnm` — master artikel, dipakai untuk kolom Grade & Temper lewat subquery

Excel murni lapisan tampilan. Semua data ada di SQL Server. Sekitar 80% kode VBA
hanya mengatur lebar kolom, merge sel, dan warna.

**Tabel penting di `bnm-data`:**

| Tabel | Isi | Jumlah kolom |
|---|---|---|
| `bnm_prdp` | Catatan operasi produksi. 253.043 baris. | 157 |
| `bnm_stck` | Stok (Raw Material dan Finished Good) | 58 |
| `bnm_type` | Master kategori (GRADE, MACHINE, TEMPER, dll.) | 10 |

**Modul VBA yang relevan:**
- `ModOutstanding.bas` — logika laporan Outstanding PIR (291 baris)
- `DBConn.bas` — koneksi database
- `VarDeclared.bas` — deklarasi variabel global
- `GeneralModule.bas`, `AdditionalFunction.bas` — helper
- `AutoRefreshData.bas` — penjadwal auto-refresh

**Alur `ProcessOutStandingPiR`:**

```
ProcessOutStandingPiR       entry point
  └─ CallDatabase           buka koneksi SQL
  └─ FormatOutStandingPiR   hapus sheet, tulis header
       └─ StoCkDetailBSOUt  siapkan blok resume
            └─ addDataOutstanding   query utama, isi data mulai baris 9
  └─ closeAllConnection
```

---

## 3. Struktur data `bnm_prdp` — ini kuncinya

**Setiap baris = satu operasi produksi**, bukan satu coil.

- `prdp_coilno` — coil **input** operasi tersebut
- `prdp_coilnonext` — coil **output**, dipisah titik koma kalau dipotong

Contoh nyata:

```
LB-9511-400  ->  FB-9511-410; FB-9511-420; FB-9511-430; FB-9511-440
Pq-9462-410  ->  FQ-9462-411A; FQ-9462-411B; ... (12 coil)
MF-9526-220  ->  DF-9526-220      (tidak dipotong, hanya pindah line)
```

Huruf pertama nomor coil menandakan line saat ini. Coil bisa **bolak-balik**
antar line dan kembali ke nomor yang sama:

```
PQ-9464-100  ->  BQ-9464-100   (REANNELING)
BQ-9464-100  ->  Pq-9464-100   (SPM)
```

Perhatikan perbedaan huruf besar/kecil (`Pq` vs `PQ`). Aman karena collation
SQL Server case-insensitive, tapi rapuh.

**Konsekuensi:** untuk tahu coil masih ada atau tidak, harus dibandingkan
waktunya — kapan terakhir diproduksi vs kapan terakhir dipakai.

**Kolom lain yang dipakai:**

| Kolom | Catatan |
|---|---|
| `prdp_status` | Isinya `''` (245.842), `WIP` (7.189), `WIP2` (12) |
| `prdp_type` | `Prime` atau `Jasa` |
| `prdp_linecur` / `prdp_linenex` | Kode line, 1 huruf |
| `prdp_nextprocess` | varchar(100), **teks bebas** — Packing, Slitting, Mill, DEG, BAL, Rewind, RE SPM, dll. |
| `prdp_datestart` | Waktu operasi |
| `prdp_dateend` | Selesainya operasi itu saja, **bukan** selesainya coil |
| `prdp_coilnomother` | Nomor mother coil (4 digit) |
| `prdp_remark` | varchar(MAX) — bisa sangat panjang |

---

## 4. Aturan bisnis (sudah dikonfirmasi)

1. **Coil `Jasa` TIDAK ikut stock take.** Filter `prdp_type = 'Prime'` benar.
   (1.750 dari 7.189 coil WIP bertipe Jasa — sekitar 24%.)
2. **Coil dengan `prdp_linecur = 'F'` IKUT dihitung**, meskipun tampil sebagai
   "FG" di laporan. Ada 1.740 coil seperti ini.
3. **Coil dengan next process kosong IKUT dihitung.** Di lapangan memang ada
   coil yang belum ditentukan proses berikutnya.
4. **Next process `Packing` / delivery = BUKAN WIP.** Coil sudah selesai.
5. **Coil yang sudah dipotong BUKAN WIP lagi.** Induk sudah tidak ada fisiknya;
   yang ada anak-anaknya. Kalau induk ikut dihitung, tonasenya dobel.
6. Status `WIP2` = penanda data bermasalah yang perlu diperbaiki lewat menu
   EPIR ADMIN. **Belum diputuskan** ikut atau tidak.

---

## 5. Temuan utama: definisi WIP di kode lama SALAH

Query lama memakai `prdp_status IN ('wip')`. Itu keliru.

Bukti: coil dengan next process **BAL** (jelas masih WIP) — 5.792 di antaranya
berstatus **kosong**, hanya 7 yang berstatus `WIP`. Tanggal terbaru keduanya
sama-sama Maret 2026, jadi status kosong itu aktif dipakai sekarang, bukan data
lama yang nyangkut.

Hipotesis yang sudah diuji dan **gugur**:
- `prdp_status = 'wip'` — melewatkan ribuan coil
- `prdp_dateend` kosong — gugur, 5.816 dari 5.820 coil BAL justru terisi
- "punya next process" — gugur, coil yang sudah dipotong juga punya next process

**Definisi yang dipakai sekarang (belum final):**

> Coil masih WIP kalau:
> (a) pernah diproduksi sebagai output DAN belum dipakai lagi sebagai input
>     (bandingkan waktunya), ATAU
> (b) jadi input operasi yang outputnya belum tercatat (sedang digarap),
> DAN tidak ada di `bnm_stck`.

---

## 6. Hasil perbandingan (per backup Maret 2026)

Aplikasi web menjalankan dua definisi berdampingan:

| Kelompok | Jumlah | Arti |
|---|---|---|
| `keduanya` | 3.120 | Lama dan baru sepakat |
| `lama` | 2.311 | Hanya definisi lama. **Next process-nya banyak "Packing"** — indikasi laporan lama menghitung coil yang sudah selesai |
| `baru` | 4.132 | Hanya definisi baru. Kandidat coil yang selama ini hilang dari laporan |

Total lama: 5.431. Total baru: 7.252.

**Belum diverifikasi ke fisik gudang.** Ini langkah yang masih tertunda karena
data yang ada cuma backup (per 28 Maret 2026), bukan produksi.

---

## 7. Bug di kode VBA lama (untuk referensi, jangan diulang)

- `On Error Resume Next` menutupi ~30 baris. Error ditelan diam-diam, laporan
  tetap bilang "DATA TELAH SELESAI DISIAPKAN!!!"
- Nama file di-hardcode: `ThisWorkbook.Name Like "FG Status Report Rev.17 RV.xlsm"`.
  Rename file = seluruh proses dilewati tanpa peringatan.
- Auto-refresh hanya jalan di PC `BNM-ERP` dengan user `csuser`. Di komputer lain
  kodenya justru **menulis** ke sheet lalu `ThisWorkbook.Save`.
- SQL injection: input dari sel G4 dirangkai langsung ke string SQL.
  `ADODB.Command` sudah dideklarasikan tapi tidak dipakai.
- Kredensial plaintext di komentar `DBConn.bas`: `User ID=csuser;Password=CScs1234`.
  **Perlu diangkat ke atasan.**
- `@EndDate` di-declare dan di-set tapi tidak pernah dipakai di WHERE.
  Judul laporan "UNTIL <tanggal>" itu cuma label, bukan filter.
- `VarDeclared.bas`: pola `Public a, b, c As String` — hanya `c` yang String,
  sisanya Variant. Ada di hampir semua baris. **Jangan diperbaiki massal**,
  akan memicu type mismatch di banyak tempat.
- Global state: `ValueData01` diisi di `ProcessOutStandingPiR`, ditimpa
  `StoCkDetailBSOUt` jadi `"GRADE"`.
- `Cells.Replace "1/2/1900"` berlaku ke seluruh sheet, bisa mengubah data lain.
- `Table_ResumePiR` dibuat dari `Range(DataSource.Range("M18").Value)` —
  alamat range disimpan di dalam sel, bukan di kode.
- Kolom "Next Process Date" di laporan isinya `2` semua. Rusak.
- Width Output = 0 karena tidak punya fallback, padahal Thickness punya.

---

## 8. Aplikasi web (kondisi sekarang)

**Stack:** Node.js + Express + `mssql` (tedious), frontend HTML/CSS/JS polos
tanpa build step.

```
outstanding-pir/
├── .env                 (tidak di-commit)
├── .env.example
├── package.json
├── README.md
├── server/
│   ├── db.js            koneksi pool, parsing named instance
│   ├── queries.js       SQL perbandingan dua definisi
│   └── index.js         Express, endpoint /api/info dan /api/coils
└── public/
    └── index.html       satu halaman: 3 kotak angka, filter, tabel, export CSV
```

**Koneksi:** instance lokal ternyata **default**, bukan named. `.env` isinya
`DB_SERVER=LTP-2017-023` tanpa backslash. Pakai login SQL `pir_reader`
(read-only + `GRANT CREATE TABLE` untuk tabel temporer).

**Performa:** versi CTE terlalu lambat (>4 menit). Sekarang pakai tabel temporer
`#prod`, `#used`, `#pending`, `#stck` dengan clustered index. `STRING_SPLIT`
harus di-`CAST` ke `varchar(100)` karena tipe MAX tidak bisa jadi index key.

---

## 9. Yang belum selesai

**Bug di aplikasi web:**
- Kolom **Tebal dan Lebar isinya 0 semua**, Lokasi kosong. Sumbernya diambil
  dari `prdp_toutputmm` / `prdp_wioutputmm` yang ternyata sering kosong.
  Perlu fallback ke kolom input, seperti yang dilakukan query VBA lama untuk
  thickness.
- Kolom **Berat kosong untuk coil hasil potong**. Berat anak coil tidak ada di
  baris operasi induknya. **Ini penting** — stock take butuh tonase. Perlu dicari
  di mana berat anak coil tersimpan.

**Fitur yang belum ada:**
- Raw Material dari `bnm_stck` belum ikut (laporan lama menempelkannya lewat
  UNION terpisah)
- Blok resume/pivot per grade
- Status `WIP2` belum ikut definisi mana pun
- Aturan next process (`Packing` dsb.) belum dijadikan konfigurasi. Rencananya
  taruh di file config atau tabel di database aplikasi sendiri —
  **jangan menambah tabel di `bnm-data`**, akses ke sana harus tetap read-only.

**Verifikasi yang tertunda:**
- Cek fisik ke gudang untuk sampel coil dari kelompok `lama` dan `baru`
- Butuh backup database yang lebih baru (yang ada tertinggal ~5 bulan)

---


## 10. Isi `bnm_type` dan pemetaan kode line

`bnm_type` adalah master kategori, dibaca dengan filter `type_unit`.

- `type_unit = 'GRADE'` → 19 entri, `type_spell` **terisi semua**. Dipakai blok
  resume Outstanding PIR. Salah satu entrinya bernama `Jasa`, sejajar dengan
  grade material (304, 316L, dst.) — jadi Jasa memang kategori resmi di sistem.
- `type_unit = 'MACHINE'` → 31 entri: M, D, RS, A, B, L, P, S, T, U, V, W, PACK,
  lalu alat bantu (Crane10T - M, Transfer Car - P1, CompGA45, MTC, GA, IT, QC,
  PROD, WH).
- Kategori lain (`GA INV`, TEMPER, dll.) punya `type_spell` **kosong**. Tidak
  dipakai laporan ini, jadi tidak masalah — tapi jangan diandalkan.

**Penting: `BAL` tidak ada di daftar MACHINE.** Jadi `prdp_nextprocess`
menyimpan teks bebas, bukan kode mesin dari master. Nilai yang muncul di data:
Packing, Slitting, Mill, DEG, BAL, Rewind, RE SPM, RE TL, Reclean, TLL, CUST,
HOLD, dan lainnya. Tidak ada validasi, jadi typo dan variasi penulisan mungkin
saja ada.

**Pemetaan kode line ke nama mesin** ditanam hardcoded di dalam string SQL VBA
lewat `CASE`, mencakup: A, B, R, M, D, L, p, S, T, U, V, W, H, F.

Catatan:
- Semua kode line yang benar-benar ada di data WIP (T, F, S, V, U, L, P, W, M,
  B, D, R) sudah tercakup. Pemetaannya lengkap.
- `A` dan `H` ada di `CASE` tapi tidak muncul di data WIP.
- `CASE` menulis `'p'` huruf kecil sementara data memakai `'P'`. Jalan hanya
  karena collation SQL Server case-insensitive. Rapuh.
- **Tidak ada cabang `ELSE`.** Ada 4 coil dengan `prdp_linecur` kosong — nama
  mesinnya jadi NULL dan Item Type-nya jadi `WIP-` tanpa huruf. Kalau nanti ada
  kode line baru ditambahkan, coil-nya muncul tanpa nama mesin tanpa peringatan.

---

## 11. Bug VBA tambahan

Selain yang di bagian 7:

- **Alias duplikat di SELECT.** `NextProcess` muncul dua kali (kolom 2 sebenarnya
  `prdp_datenext`, kolom 15 `prdp_nextprocess`). `CurrentLine` juga dua kali
  (kolom 5 dan 6). Tidak error karena `CopyFromRecordset` tidak peduli nama,
  tapi menyesatkan saat dibaca.
- **`Like` tanpa wildcard.** `ThisWorkbook.Name Like "FG Status Report Rev.17 RV.xlsm"`
  sama saja dengan `=`. Tidak ada toleransi sama sekali.
- **Timestamp disimpan sebagai teks.** `NumberFormat = "@"` diterapkan sebelum
  menulis `Now` ke H2/H4, jadi hasilnya string bukan datetime. Tidak bisa
  dibandingkan atau diurutkan.
- **`Application.OnTime` menjadwalkan dirinya sendiri**, dan saya belum menemukan
  rutin pembatalnya. Kalau tidak ada `Schedule:=False` di `Workbook_BeforeClose`,
  Excel akan membuka ulang file ini sendiri setelah ditutup.
  **Perlu dicek di `WorkbookFG`.**
- **`LogFile_WriteError` logikanya kacau.** Rangkaian `If Dir(A)="" Then MkDir A
  ElseIf Dir(B)="" ...` maksudnya fallback antar folder, tapi hasilnya bisa
  membuat folder di lokasi yang salah. Kebetulan sering jalan.
  Log error tersimpan di `C:\ErrorVBA\<nama-komputer>\`.
- **`Data From` di sheet di-hardcode `-`.** Kombinasi dengan `@EndDate` yang tak
  terpakai berarti judul "OUTSTANDIR PIR UNTIL <tanggal>" itu murni label. Kalau
  user mengira laporan menampilkan posisi per tanggal tertentu, mereka salah
  paham. **Perlu diluruskan sebelum mengubah apa pun.**
- **Typo di judul:** `OUTSTANDIR PIR` (baris 39 `ModOutstanding.bas`).
- **"Excel Total Line" menunjukkan 9** padahal data terlihat jauh lebih banyak.
  Kemungkinan formula belum dihitung ulang — cek urutan `Application.Calculation`.
- **`ProcessOutStandingPiR` tidak menyetel `Application.ScreenUpdating = False`,**
  dan memakai `.Select` di banyak tempat. Lambat, dan bergantung pada sheet aktif.

**Belum terverifikasi:** `FormatOutStandingPiR` melakukan `ClearContents` /
`ClearFormats` lalu `ListObjects.Add` dengan nama `Table_OutStandingPiR`.
Setahu saya ClearContents tidak menghapus objek tabel, jadi Add dengan nama sama
seharusnya error — dan error itu ditelan `On Error Resume Next`. Tapi kode ini
konon berjalan normal. Perlu dicek langsung di file salinan, lihat apakah file
log di `C:\ErrorVBA\` bertambah.

---

## 12. Fakta backup dan lingkungan

- **Rentang data backup:** 1899-12-30 sampai **2026-03-28**. Total 253.043 baris
  `bnm_prdp`. Backup ini tertinggal sekitar 5 bulan dari produksi — cukup untuk
  memahami struktur, **tidak cukup** untuk memverifikasi hasil akhir.
- Instance lokal: `LTP-2017-023`, SQL Server 2017, **instance default**
  (bukan named). Login Windows `BNMSTAINLESS\bahrul.ulum`.
- Kedua database sudah di-restore: `bnm-data` dan `ww-bnm`.
- `ww_art` punya 123 kolom, `ww_prd` 62 kolom. Dipakai laporan lama lewat
  subquery untuk kolom Grade & Temper. Aplikasi web belum memakainya.
- Kunci join ke stok: `bnm_stck.stck_artsernr` = nomor coil.
- Di `DBConn.bas` ada baris ter-comment `Server_Name = "LTP-2017-025\SQL2017E"`.
  Developer sebelumnya juga bekerja dengan salinan lokal — polanya sudah benar,
  hanya tidak didokumentasikan.
- Contoh mother coil dengan banyak anak (untuk uji logika potong): 9360, 9431,
  9545, 9307 (masing-masing 12 anak); 9183, 9426 (11); 9308, 9546 (10).

---

## 13. Prinsip kerja

- Akses ke `bnm-data` dan `ww-bnm` **read-only**, hanya SELECT.
- Jangan hardcode aturan bisnis di dalam string SQL. Itu kesalahan sistem lama.
- Setiap asumsi diuji ke data sebelum dipakai. Sejauh ini tiga asumsi sudah
  gugur saat diuji (`prdp_status`, `prdp_dateend`, "punya next process").
- Semua eksplorasi Excel dilakukan di **salinan file**, jangan yang asli.

---

## 14. Pembacaan kode `server/queries.js` (belum diuji ke DB)

Dibaca dari kode, belum dikonfirmasi ke data. Berkaitan dengan bug di bagian 9.

**Penyebab kolom Berat kosong untuk coil hasil potong — ketemu di kode.**

Di SELECT akhir, semua kolom detail memakai fallback `COALESCE(o.…, sr.…)`
kecuali satu:

```sql
o.berat AS berat,        -- tidak ada COALESCE
```

`#own` = detail dari operasi terakhir di mana coil jadi **input**.
`#src` = detail dari operasi yang **menghasilkan** coil itu.

Coil hasil potong belum pernah jadi input operasi apa pun — barisnya hanya ada
di `#src`. Jadi `o.berat` NULL, dan tanpa `COALESCE` hasilnya kosong. Cocok
persis dengan gejala yang dilaporkan.

**Tapi `COALESCE` saja tidak cukup:** `#src` memang tidak menyeleksi kolom berat
sama sekali. Harus ditambah dulu di sana.

**Masalah kedua:** `#own` memakai `prdp_weinput` — berat **input** operasi.
Untuk coil yang sudah diproses, berat sesudah operasi belum tentu sama dengan
sebelumnya. Perlu dicek apakah ada kolom berat output di antara 157 kolom
`bnm_prdp`.

**Tebal/Lebar:** `COALESCE` sudah ada, tapi kedua sisi (`#own` dan `#src`)
sama-sama mengambil `prdp_toutputmm` / `prdp_wioutputmm`. Karena keduanya
sering kosong, fallback ke kolom **input** yang disebut di bagian 9 memang
belum ada di kode mana pun.

**Langkah berikutnya (semua read-only SELECT):**
1. Daftar kolom `bnm_prdp` yang namanya mengandung `we` / `weight` / `berat`,
   lewat `INFORMATION_SCHEMA.COLUMNS`.
2. Cek keterisian tiap kandidat — berapa persen NULL atau 0.
3. Ambil sampel satu induk yang dipotong beserta anak-anaknya (pakai mother coil
   dari bagian 12: 9360, 9431, 9545, 9307), lihat kolom mana yang menyimpan
   berat anak.

Jangan menebak nama kolom sebelum langkah 1 dijalankan.

---

## 15. Hasil investigasi berat & dimensi (SUDAH diuji ke DB, backup Maret 2026)

Menjawab bug di bagian 9 dan dugaan di bagian 14. Semua angka di bawah hasil
SELECT read-only ke `bnm-data` di `LTP-2017-023`.

### Berat anak coil: `prdp_weoutputkg`, di baris anak sendiri

Berat anak **tidak** tersimpan di baris operasi induknya. Setiap anak punya
**baris sendiri** di `bnm_prdp`. Itu sebabnya aplikasi gagal menemukannya —
dicari di tempat yang salah, bukan karena kolomnya tidak ada.

Bukti, induk `LB-9511-400` (dipotong jadi 4, Maret 2026):

| Coil | `prdp_weoutputkg` |
|---|---|
| FB-9511-410 | 478 |
| FB-9511-420 | 161 |
| FB-9511-430 | 374 |
| FB-9511-440 | 463 |
| **jumlah anak** | **1.476** |
| induk LB-9511-400 | 1.530 |

Selisih 54 kg = susut potong. Angka masuk akal.

Di baris induk, `prdp_weinput` = `prdp_weoutputkg` = 1530 — baris induk hanya
mencatat berat induk, tidak memecah per anak.

### Jebakan: baris TERBARU sering ber-`weoutputkg` = 0

Contoh `FB-9511-410`:

```
2026-03-16 12:16   weinput 478   weoutputkg 478   next Packing
2026-03-16 14:47   weinput 478   weoutputkg 0     next FG   lokasi 223033
```

Kalau ambil baris terbaru mentah-mentah, berat jadi nol. Terjadi pada
**6.298 coil** (`weoutputkg` = 0 padahal `weinput` terisi).

**Fallback yang diperlukan:** `weoutputkg` kosong/nol → pakai `weinput`.

Setelah fallback, dari 111.912 coil ujung hanya **35 coil** yang beratnya
benar-benar tidak ada. Praktis selesai.

### Nama kolom input yang benar (penyebab Tebal/Lebar 0)

**Tanpa akhiran `mm`** — ini yang bikin bug bagian 9:

| Dimensi | Kolom output (dipakai kode sekarang) | Kolom input (belum dipakai) |
|---|---|---|
| Tebal | `prdp_toutputmm` | `prdp_tinput` |
| Lebar | `prdp_wioutputmm` | `prdp_wiinput` |

Percobaan memakai `prdp_wiinputmm` / `prdp_tinputmm` akan error
*Invalid column name*.

Butuh fallback ke kolom input: tebal **28.359 coil**, lebar **9.728 coil**.

### Lokasi kosong bukan bug aplikasi

**64.902 coil** memang `prdp_location` kosong di database. Datanya tidak diisi
di lapangan. Tidak bisa diperbaiki lewat kode — perlu dikonfirmasi ke bagian
produksi/gudang.

### Daftar lengkap kolom berat di `bnm_prdp`

| Kolom | Tipe | Terisi (dari 246.042 baris Prime) |
|---|---|---|
| `prdp_weoutputkg` | float | 239.518 (97%) |
| `prdp_weinput` | float | 228.441 (93%) |
| `prdp_wealloc` | float | 202.429 |
| `prdp_wediff` | float | 41.281 |
| `prdp_weoutputlbs` | float | — (satuan lbs, tidak dipakai) |
| `prdp_weinterm`, `prdp_wereproses`, `prdp_wenote` | varchar | bukan angka |

---

## 16. Aturan filter: `linecur` vs `nextprocess` (dikonfirmasi pemilik proses)

Dua kolom ini sering tertukar. **Hanya `nextprocess` yang menentukan selesai
atau belum.**

| Kolom | Arti | Peran di filter |
|---|---|---|
| `prdp_linecur` | posisi coil sekarang ada di line mana | **jangan dipakai memfilter** |
| `prdp_nextprocess` | mau diapakan berikutnya | penentu selesai / belum |

Penjelasan pemilik proses:

> Kalau `linecur` F tapi belum Packing/FG/Delivery/PACK, tidak apa-apa.
> Misal line F next process TL atau Slit — dia masih akan diproses,
> jadi masih WIP.

Contoh yang harus **IKUT** dihitung:

```
linecur = F, nextprocess = TL     -> masih akan dipotong, WIP
linecur = F, nextprocess = Slit   -> masih akan di-slitting, WIP
```

Contoh yang **KELUAR**:

```
nextprocess = Packing / FG / Delivery / PACK   -> sudah selesai
```

Ini mendamaikan bagian 4 no. 2 ("coil `linecur = 'F'` IKUT dihitung meskipun
tampil sebagai FG di laporan"): yang membingungkan hanya huruf F-nya, sedangkan
penentu sebenarnya ada di kolom yang berbeda.

**Catatan:** nilai `FG` muncul di `prdp_nextprocess` di data nyata, tapi belum
tercatat di daftar bagian 10. Coil ber-`nextprocess = FG` pada sampel sudah
punya nomor lokasi gudang (223033, 223034) — konsisten dengan barang jadi.

Daftar nilai yang dianggap selesai (`Packing`, `FG`, `Delivery`, `PACK`) masih
perlu diverifikasi menyeluruh ke seluruh nilai `prdp_nextprocess` yang ada di
data — teks bebas, jadi mungkin ada variasi penulisan dan typo.

---

## 17. Query stock take sudah dipasang (4 September 2026) — angka DIKOREKSI di bagian 18

`queries.js` ditulis ulang: dari perbandingan dua definisi → query stock take
tunggal. Aturan bisnis dipindah ke `server/config.js`.

### Hasil per backup Maret 2026

| | |
|---|---|
| Jumlah coil | **3.454** |
| Total tonase | **2.178,39 ton** |
| Berat kosong | 11 (0,3%) |
| Tebal/lebar kosong | 4 |
| Coil hasil potong | 802 |
| Lokasi kosong | 3.238 — data lapangan, bukan bug |
| Waktu query | ±15 detik |

### Cara menentukan coil ujung — Cara B, bukan perbandingan waktu

Diuji berdampingan lebih dulu:

| Cara | Hasil mentah |
|---|---|
| A — perbandingan waktu (kode lama) | 9.874 |
| B — tidak punya keturunan (definisi pemilik proses) | 74.728 |
| sepakat | 8.941 |

Angka B yang besar itu karena belum disaring. Setelah saringan:

```
74.728  coil ujung
65.014  setelah buang Packing/FG/Delivery
 2.652  setelah buang yang ada di bnm_stck   <- pemangkas terbesar
```

Query final menghasilkan 3.454 karena juga menangkap **802 coil hasil potong**
yang belum punya baris sendiri di `bnm_prdp` — uji bertahap di atas
melewatkannya.

**Penting:** output yang namanya SAMA dengan input tidak dihitung keturunan
(`FB-9511-410 -> FB-9511-410`). Itu rewind atau pindah line, coilnya barang
yang sama. Kalau dihitung sebagai keturunan, coil-nya hilang dari laporan
padahal fisiknya ada.

### Filter next process pakai POLA, bukan daftar nilai

Daftar `('Packing','FG','Delivery','PACK')` **bocor**. Di data ada 17 variasi
penulisan:

```
PACKING (55.923)      FG (49.178)
Packing - FGR Cancel  FG - FGR Cancel
Packing - REPROSES    FG - REPROSES
PACKINGReprocess      FG - TRANSFER
Packing-TRANSFER      FG-TRANSFER
pack                  Packing -  Reprocess   (spasi dobel)
```

Sudah dicek: alur "REPROSES" pun berakhir di `FG - REPROSES` — jadi semua
varian ini memang berarti selesai.

Pola yang dipakai (`server/config.js`): `%pack%`, `FG%`, `%deliv%`.
Menutup 3 coil yang lolos dari filter daftar.

### Rantai fallback

```
berat  : weoutputkg -> weinput -> (dari operasi induk) weoutputkg -> weinput
tebal  : toutputmm  -> tinput  -> (induk) toutputmm  -> tinput
lebar  : wioutputmm -> wiinput -> (induk) wioutputmm -> wiinput
```

`NULLIF(kolom, 0)` dipakai supaya nilai 0 diperlakukan sebagai kosong — baris
terakhir sebuah coil sering ber-`weoutputkg` = 0.

### Keputusan pemilik proses

- **Tahun tidak difilter.** Asal masih WIP dan belum FG, ikut — termasuk coil
  2022 yang mengendap (990 coil).
- `Packing`, `FG`, `Delivery`, `PACK` = selesai, tidak masuk stock take.
- Filter di `nextprocess`, tidak pernah di `linecur`.

### Perubahan file

| File | Perubahan |
|---|---|
| `server/config.js` | **baru** — pola next process selesai |
| `server/queries.js` | ditulis ulang, `COMPARE` → `STOCK_TAKE` |
| `server/index.js` | menyesuaikan nama query |

Backup `queries.js` versi lama ada di scratchpad sesi ini.

### Belum dikerjakan

- **`public/index.html` belum disesuaikan.** Masih menampilkan 3 kotak
  perbandingan lama/baru/keduanya dan memakai kolom `sumber` yang sudah tidak
  ada. Perlu diganti: jumlah coil + total tonase.
- Kolom `hasil_potong` (0/1) sudah tersedia di hasil query, belum dipakai UI.
- Raw Material dari `bnm_stck` belum ikut.
- Blok resume per grade belum ada.
- `WIP2` (12 baris) masih belum diputuskan.
- **Belum diverifikasi ke fisik gudang.** Angka 3.454 coil / 2.178 ton ini
  dari backup Maret 2026, bukan produksi.

---

## 18. Halaman web + perbaikan pemisah koma (4 September 2026)

### Angka final

| | |
|---|---|
| Jumlah coil | **2.821** |
| Total tonase | **1.302,8 ton** |
| Berat kosong | 6 |
| Coil hasil potong | 169 |
| Waktu query | ±16 detik |

Angka ini menggantikan 3.454 coil / 2.178 ton di bagian 17 — yang itu salah,
sebabnya di bawah.

### Temuan: pemisah coil kadang KOMA, bukan titik koma

`STRING_SPLIT` hanya memecah di `;`. Padahal **655 baris** memakai koma:

```
DD-8736-400  ->  'LD-8736-410,LD-8736-420'      <- koma
LB-9511-400  ->  'FB-9511-410; FB-9511-420'     <- titik koma (umum)
```

Akibatnya dua coil terbaca sebagai **satu nama gabungan**. Nama gabungan itu
tidak pernah cocok saat dicek apakah punya keturunan, jadi lolos ke laporan
sebagai WIP padahal anak-anaknya sudah diproses lebih lanjut.

Dampaknya: **642 baris hasil salah, 875 ton kelebihan hitung.**

**Perbaikan:** koma diseragamkan jadi titik koma sebelum dipecah —
`STRING_SPLIT(REPLACE(ISNULL(prdp_coilnonext,''), ',', ';'), ';')`.
Diterapkan di ketiga tempat `STRING_SPLIT` dipakai.

**Verifikasi.** Ditelusuri satu per satu:

```
DD-8736-400 -> LD-8736-410, LD-8736-420      (dipisah koma)
LD-8736-410 -> SD-8736-411; SD-8736-412      (sudah dipotong lagi)
SD-8736-411 -> FD-8736-411A; FD-8736-411B    (dipotong lagi)
```

`LD-8736-410` memang **bukan** WIP — sudah habis jadi cucu. Tidak ada di
`bnm_stck` juga. Jadi hilangnya 2.497 coil dari hasil itu benar, bukan
kehilangan data.

Sebaran pemakaian koma per tahun: 2022 → 644, 2023 → 7, 2024 → 5.
Praktis sudah ditinggalkan, tapi datanya masih ada dan tetap harus ditangani.

### Sisa keanehan data (bukan bug aplikasi)

Tiga baris yang nomor coilnya jelas bukan nomor coil:

```
TRIAL SAMPEL              next=TRIAL
ACC ENTRY TIDAK BISA ON   next=MAINTENANCE
PURGING OTOMATIS          next=MAINTENANCE
MAINTENANCE               next=MAINTENANCE
```

Operator memakai kolom nomor coil untuk mencatat kegiatan non-produksi.
Semuanya tanpa berat, jadi tidak merusak tonase. **Belum difilter** — perlu
diputuskan: buang berdasarkan `next_process = MAINTENANCE/TRIAL`, atau
biarkan tampil supaya kelihatan datanya perlu dibersihkan.

### Halaman web sudah disesuaikan

`public/index.html` — tiga kotak perbandingan lama/baru/keduanya diganti:

| Kotak | Isi |
|---|---|
| Coil WIP | jumlah coil |
| Total tonase | ton, + peringatan kalau ada coil tanpa berat |
| Hasil potong | jumlah, bisa diklik untuk memfilter |

Perubahan lain:
- Kolom `sumber` (tag lama/baru/keduanya) dihapus, diganti kolom **Asal**
  yang menandai coil hasil potong.
- Nilai kosong tampil sebagai **—** abu-abu, bukan sel kosong atau `0`.
  Supaya jelas datanya memang tidak ada, bukan gagal diambil.
- Label tonase otomatis menyebut jumlah coil tanpa berat — supaya tidak
  ada yang mengira totalnya lengkap padahal tidak.
- Angka pakai format Indonesia (`1.302,8`).
- CSV ikut membawa kolom `hasil_potong`.

Pencarian, unduh CSV, dan ambil-ulang tetap seperti semula.

### Belum dikerjakan

- Raw Material dari `bnm_stck` belum ikut
- Blok resume per grade belum ada
- `WIP2` (12 baris) belum diputuskan
- Baris non-produksi (MAINTENANCE/TRIAL) belum difilter
- **Belum diverifikasi ke fisik gudang** — angka 2.821 coil / 1.302,8 ton ini
  dari backup Maret 2026, bukan produksi

---

## 19. Dua mode query: tabel temporer vs SELECT murni (4 September 2026)

Disiapkan untuk menyambung ke database produksi, di mana login kemungkinan
hanya diberi hak baca.

### Pilihan lewat `.env`

```
# (kosong)      -> tabel temporer, butuh izin CREATE TABLE
DB_MODE=murni   -> CTE, cukup izin SELECT
```

| | Tabel temporer | SELECT murni |
|---|---|---|
| File | `server/queries.js` | `server/queries-murni.js` |
| Izin dibutuhkan | SELECT + **CREATE TABLE** | **SELECT saja** |
| Menulis ke server | ya, numpang `tempdb` | tidak sama sekali |
| Kecepatan | lebih cepat | ~60% lebih lambat |

**Hasil dua-duanya identik:** 2.821 coil, 1.302,8 ton, **nol selisih coil**.
Sudah diadu langsung, bukan diasumsikan.

### Soal angka waktu — jangan dijadikan patokan

Pengukuran di backup Maret 2026, laptop `LTP-2017-023`, tanpa pemakai lain:

```
jalan pertama (cache dingin) : temp 45,4 dtk   murni 17,7 dtk
putaran 1                    : temp 12,8 dtk   murni 17,8 dtk
putaran 2                    : temp 11,0 dtk   murni 17,7 dtk
putaran 3                    : temp 10,5 dtk   murni 16,6 dtk
```

Jalan pertama 45 detik, lalu turun ke ~11 detik untuk query yang sama persis.
Bedanya cuma cache database sudah panas.

**Kesimpulannya: angka detik itu relatif.** Tergantung cache, beban server,
hardware, jumlah data. Di produksi akan berbeda lagi. Yang tetap adalah
jumlah kerjanya — baca ~253 ribu baris, 2× `ROW_NUMBER()`, 1× `STRING_SPLIT`.

### Optimasi yang sudah dilakukan sebelumnya

`STRING_SPLIT` tadinya dipanggil **3 kali** atas seluruh tabel (untuk `#semua`,
`#anak`, `#src`). Profiling per langkah:

```
5,13 dtk  #semua        <- STRING_SPLIT
2,65 dtk  #anak         <- STRING_SPLIT
0,26 dtk  #ujung
0,56 dtk  #stck
0,85 dtk  #own
5,24 dtk  #src          <- STRING_SPLIT
0,56 dtk  SELECT akhir
```

13 dari 15 detik habis di pemecahan teks. Sekarang dipecah **sekali** di
`#pecah` / CTE `pecah`, sisanya diturunkan dari situ. `#semua` dihapus —
isinya bisa didapat dari gabungan `#own` + `#src`.

### Catatan: ejaan coil tidak konsisten

Saat mengadu dua versi, muncul `Pa-9273-310` vs `PA-9273-310` — **coil yang
sama**, ditulis beda huruf besar/kecil di dua baris berbeda:

```
PA-9273-300 -> 'Pa-9273-310; Pa-9273-320'
Pa-9273-310 -> 'PA-9273-310'
PA-9273-310 -> ''
```

Jumlah dan tonase tidak terpengaruh (tetap dihitung 1 coil, collation SQL
Server case-insensitive). Yang berbeda hanya ejaan mana yang kebetulan
ditampilkan. Ini kerapuhan yang sudah dicatat di bagian 3.

### Untuk menyambung ke produksi

1. Ganti `DB_SERVER` di `.env` ke server produksi
2. Set `DB_MODE=murni` — supaya izin yang diminta cukup SELECT
3. Minta DBA: login read-only ke `bnm-data` + `ww-bnm`, dan akses jaringan
   ke port SQL Server

**Jangan pakai** kredensial `csuser` / `CScs1234` yang ada di `DBConn.bas`.
Itu kredensial bocor yang justru sedang dilaporkan (lihat bagian 7).

### Yang masih perlu diingat

Aplikasi hanya query saat tombol **"Ambil ulang dari database"** ditekan —
buka halaman biasa membaca cache di memori, tidak menyentuh database. Untuk
stock take bulanan, itu berarti 1–3 query sebulan.

Cache masih di memori, hilang kalau server di-restart. Belum disimpan ke file.

---

## 20. Aturan penyaringan diperluas + hasil audit (4 September 2026)

### Angka final

| | |
|---|---|
| Jumlah coil | **2.789** |
| Total tonase | **1.295,2 ton** |

Turun dari 2.821 / 1.302,8 t setelah `CUST`, `NEGO`, `CANCEL`, `MAINTENANCE`,
`TRIAL` dibuang.

### Keputusan pemilik proses

> `CUST` itu customer, jadi sudah pasti selesai prosesnya — dia sudah sampai
> ke customer.
>
> maintenance buang, trial buang, nego buang, cancel buang.

Pola di `server/config.js` bertambah jadi 8:

```
%pack%    PACKING, pack, Packing - REPROSES, PACKINGReprocess
FG%       FG, FG - FGR Cancel, FG-TRANSFER
%deliv%   Delivery
%cust%    CUST (66.082 baris), CUST - TRANSFER, Customer - TRANSFER   <- BARU
%nego%    NEGO                                                        <- BARU
%cancel%  CANCEL                                                      <- BARU
%maint%   MAINTENANCE — barisnya bukan coil                           <- BARU
%trial%   TRIAL                                                       <- BARU
```

**`CUST` ternyata besar: 66.082 baris.** Sebelumnya ikut terhitung sebagai WIP.

`CUST - RETURN` (2 baris) dan `CUST - Reproses` (1 baris) sempat dicek — barang
balik dari customer. Ditelusuri riwayat lengkapnya, ketiganya berakhir di
`Packing` atau `CUST` lagi setelah diproses ulang. Aman dibuang, lagipula yang
menentukan hanya baris terakhir.

Verifikasi sesudah dipasang: **tidak ada satu pun** nilai mengandung
cust/nego/cancel/maint/trial/pack/fg/deliv yang lolos ke hasil.

### Audit penyaringan — hasilnya

Diuji dari beberapa arah, bukan hanya memastikan yang sudah jalan.

**Saringan next process: bersih.** 17 variasi Packing/FG/Delivery tertangkap
semua. Dicari nilai lolos yang bernuansa selesai (kirim, done, finish, keluar,
jual, close) — tidak ada.

**Penentuan keturunan: 16 dari 16 benar.** Dua arah:
- 8 coil acak dari hasil → terbukti tidak punya keturunan, tidak ada di stok
- 8 coil yang jelas punya anak → terbukti tidak masuk hasil

**Tidak ada pemisah lain.** Dicek `/`, `|`, `+`, `&`, kata "dan", spasi antar
coil — nol semua. Hanya `;` dan `,`, keduanya sudah ditangani.

**Tidak ada spasi/tab/enter tersembunyi** di nomor coil.

**HOLD (520 coil): tidak ada yang punya keturunan.** Dicek 60 sampel, nol.
Wajar — saringan keturunan jalan lebih dulu.

**Yang diambil memang baris terakhir**, bukan pencarian kata. Contoh
`FD-7476-312`: HOLD (23 Mei) → PACKING (24 & 31 Mei) → HOLD (7 Juni). Yang
dipakai HOLD 7 Juni. Kalau salah ambil, coil ini terbuang karena dikira PACKING.

### Kerapuhan yang ditemukan (belum jadi masalah)

**1. Tanggal kembar di baris terakhir — 1.283 coil.**

Untuk coil ini, urutan ditentukan `prdp_nr`. Masalahnya `prdp_nr` **tidak
selalu searah waktu**: ada 4.373 baris di mana nomor naik tapi tanggal mundur.

Dari 1.283 itu, **125 coil** kritis: satu baris `PACKING` (selesai), satu
`CUST` (dulu dianggap WIP), tanggal identik. Mana yang menang bergantung
nomor urut.

Dampak nyata saat diukur: **1 coil, 100 kg dari 1.302,8 ton** — praktis nol.
Sekarang `CUST` sudah ikut dibuang, jadi kasus ini hilang sendiri. Tapi polanya
tetap ada untuk pasangan nilai lain.

**2. Coil bertipe campuran Prime + Jasa — 6 coil.**

Contoh `RN-9461-000`:
```
Prime, 16 Okt : output RN-9461-000  (dirinya sendiri)
Jasa,  30 Okt : output VN-9461-100; VN-9461-200; VN-9461-300   <- dipotong
```
Query hanya melihat baris Prime, jadi pemotongan lewat baris Jasa tidak
terlihat. Ada **912 coil** yang punya keturunan lewat baris Jasa.

Dampak nyata: **nol** — tidak satu pun masuk hasil, semuanya sudah tersaring
aturan lain. Tapi ini titik rapuh kalau pola data berubah.

**3. Ejaan next process tidak konsisten.**

Dari 30 "jenis" yang tampil, sebenarnya hanya **23** — sisanya beda huruf
besar-kecil:

```
SLITTING   [Slitting / SLITTING]        539 coil, 487,0 t
REWIND     [Rewind / REWIND / rewind]   303 coil,  90,5 t
STRETCHING [Stretching / STRETCHING]    294 coil,  29,7 t
MILL       [Mill / MILL]                 67 coil, 141,6 t
DIVERT     [Divert / DIVERT]             67 coil,  17,6 t
SAMPLE     [Sample / SAMPLE]             10 coil,   0,6 t
```

Tidak salah hitung (collation SQL Server case-insensitive), hanya tampilannya
terpecah. **Sengaja TIDAK diseragamkan di halaman** — kalau ditutupi, hilang
petunjuk bahwa data mentahnya memang tidak konsisten.

### Sisa next process di hasil (23 jenis, sudah disatukan)

```
 663  BC          520  HOLD        539  SLITTING    303  REWIND
 294  STRETCHING  142  WIP          67  MILL         67  DIVERT
  40  TLL          34  SPM          33  PRIME        17  DEG
  15  BAL          15  REWORK       10  SAMPLE        8  RECLEAN
   6  RE SPM        4  MICS          3  RE CLEAN      3  RE TLL
   3  RE INSPECT    2  SHEARING      1  REINSPECT
```

**`WIP` (142 coil) dan `PRIME` (33 coil) — sudah dikonfirmasi IKUT dihitung.**

Pemilik proses: *"WIP belum ditentukan proses berikutnya."* Sesuai aturan
bagian 4 no. 3 (coil dengan next process kosong ikut dihitung), `WIP` adalah
bentuk eksplisit dari "belum ditentukan".

`PRIME` ditelusuri ke data dan polanya sama persis — baris terakhirnya selalu
**output kosong, berat 0, status `WIP`**:

```
Pa-9361-110   9 Des  Slitting -> Pa-9361-110   830 kg
              9 Des  Prime    -> (kosong)        0 kg   status WIP
MF-9269-112   2 Des  Prime    -> MF-9269-112   508 kg
              2 Des  Prime    -> (kosong)        0 kg   status WIP
```

Jadi `PRIME` juga berarti coil menganggur, bukan proses. Ikut dihitung —
33 coil, 11,4 ton. Tidak ada perubahan kode.

### Belum dikerjakan

- Raw Material dari `bnm_stck` belum ikut
- Blok resume per grade belum ada
- `WIP2` (12 baris) belum diputuskan
- **Belum diverifikasi ke fisik gudang** — angka 2.789 coil / 1.295,2 ton ini
  dari backup Maret 2026, bukan produksi

## 21. Audit identitas fisik di produksi (22 Oktober 2026)

Dua false positive ditemukan lewat pengecekan Excel produksi:

- `MD-7167-211` masih muncul walaupun fisiknya berlanjut melalui `DD`, `LD`,
  lalu pecah menjadi `FD-7167-211A` sampai `E`.
- `SB-7741-324A` masih muncul walaupun fisik yang sama sudah berganti prefix
  menjadi `FB-7741-324A`, pecah lagi menjadi `A1/A2`, lalu Packing.

Akar masalahnya adalah query lama memakai nama coil lengkap sebagai identitas.
Di produksi, prefix dua huruf menunjukkan tahap proses dan dapat berubah tanpa
selalu mengisi `prdp_coilnonext`.

Query sekarang memakai bagian setelah `XX-` sebagai `id_fisik`. Audit format
menemukan 287.221 dari 287.238 baris operasi memakai format tersebut; 17 baris
lainnya merupakan catatan non-coil seperti maintenance dan trial. Input,
output, dan `bnm_stck` dinormalisasi dengan aturan yang sama. Spasi, tab, CR,
dan LF tersembunyi juga dibuang sebelum perbandingan.

Hasil audit produksi setelah perbaikan:

- 2.108 coil / 878,36945 ton
- nol duplikat `id_fisik`
- nol hasil dengan next process selesai
- nol hasil berlokasi “dihapuskan”
- nol hasil yang masih cocok dengan `bnm_stck`
- nol hasil yang menjadi induk pemotongan lewat baris non-Prime
- keluarga `7167-211` dan `7741-324` tidak lagi muncul

Jika beberapa operasi memiliki `prdp_datestart` sama, urutan terakhir dipilih
dengan `prdp_nr` yang dikonversi ke angka. Ini penting karena banyak riwayat
reproses, return, dan transfer memakai timestamp identik.

## 22. Audit sampel produksi (22 Oktober 2026)

Audit independen dilakukan setelah desain ulang identitas fisik. Sampel dipilih
secara berlapis dan deterministik, bukan hanya mengambil baris teratas:

- 97 coil dari total 2.108 hasil
- mencakup seluruh 23 kelompok `next process`
- mencakup data paling lama dan terbaru
- mencakup coil hasil potong dan coil tanpa berat

Setiap sampel dicocokkan kembali ke baris mentah `bnm_prdp`, output hasil
`STRING_SPLIT`, anak inferred, dan `bnm_stck`. Pemeriksaan meliputi:

- nama coil dan next process sama dengan transaksi terakhir
- tidak punya anak eksplisit atau inferred yang lebih baru
- tidak tercatat di stok setelah normalisasi identitas
- next process bukan proses selesai
- lokasi tidak bertanda dihapus

Hasil: **0 kegagalan dari 97 sampel**.

Kontrol negatif juga diuji sebanyak 90 coil:

- 30 coil dengan proses terakhir selesai
- 30 coil dari `bnm_stck`
- 30 coil yang terbukti menjadi induk pemotongan

Hasil: **0 dari 90 kontrol negatif bocor ke laporan**.

Audit ini membuktikan konsistensi query terhadap data dan aturan yang sudah
disepakati. Audit belum menggantikan pengecekan fisik gudang; khusus coil lama
yang secara data masih WIP, keberadaan fisiknya tetap perlu spot check lapangan.

## 23. Arti `hasil_potong` setelah normalisasi (22 Oktober 2026)

Secara teknis, `hasil_potong = 1` berarti identitas coil ditemukan sebagai
output di `prdp_coilnonext`, tetapi belum pernah muncul sebagai input operasi
sendiri di `prdp_coilno`. Jadi nama “hasil potong” tidak selalu tepat.

Profil produksi saat audit:

- 9 dari 2.108 WIP (0,43%) memiliki flag tersebut
- 5 berasal dari operasi multi-output dan memang sesuai makna hasil split
- 4 berasal dari operasi single-output; lebih tepat disebut output baru atau
  output yang belum diproses lagi
- semua 9 memiliki berat, tetapi berat multi-output masih berasal dari satu
  baris operasi induk

Dua kelompok multi-output yang aktif saat audit:

- `SD-8097-311B` → 3 output, masing-masing tampil 548 kg
- `PA-7158-100` → 2 output, masing-masing tampil 1.730 kg

Database hanya menyediakan satu `prdp_weoutputkg` pada baris induk. Karena
nilai itu ditempelkan ke setiap anak, total tonase dapat terduplikasi. Potensi
duplikasi maksimum pada dua kelompok tersebut adalah 2.826 kg. Berat individual
anak tidak dapat ditentukan dengan aman dari baris itu saja.

Empat single-output yang perlu dipahami sebagai output belum diproses, bukan
otomatis hasil potong: `PA-8958-100A`, `UB-7804-321A110`, `FF-7699-22G`, dan
`FB-7735-123`.
