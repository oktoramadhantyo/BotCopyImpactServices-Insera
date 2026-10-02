# Bot Ambil Impacted Services (Insera)

Userscript Tampermonkey untuk mengambil **seluruh** data pada tab
**Impacted Services** di Insera, lalu mengunduhnya sebagai Excel `.xls`.

## Kebutuhan

- Browser Chrome/Edge/Firefox dengan ekstensi **Tampermonkey**.
- Akses ke Insera. Login tetap dilakukan manual.

Tidak ada Python, server, Google Apps Script, atau library tambahan.

## Pemasangan

1. Buka dashboard Tampermonkey.
2. Klik **Create a new script**.
3. Hapus seluruh isi editor, lalu tempel seluruh isi
   `insera-impact.user.js`.
4. Tekan **Save** (Ctrl+S).

## Cara pakai

Panel muncul sendiri di pojok kanan bawah begitu tab Impacted Services
ditemukan, lalu tombolnya berubah aktif setelah tabelnya terdeteksi.

1. Login ke Insera, cari tiket, buka tiketnya.
2. Klik **Scan & Download XLS**.

Tidak perlu membuka tab atau membuka section lebih dulu. Script melakukan
persiapannya sendiri:

- Kalau halaman membuka tab lain (misalnya `Related Records`), script
  mengklik tab **Impacted Service** sendiri.
- Kalau tabel berada di balik section yang tertutup, script membuka
  **Impacted Service Summary** lebih dulu.

Nomor tiket terdeteksi otomatis dan ditampilkan di panel, jadi tidak ada
yang perlu diketik. Nomor itu hanya dipakai pada nama file.

Selama proses berjalan tombol berubah menjadi **Batalkan** (merah). Klik
itu untuk menghentikan; data parsial tidak akan diunduh.

## Keluarannya

Berkas Excel berisi sheet `Impacted Services` dengan kolom:

| DATECREATED | SERVICE ID | REGION | IBOOSTER OPER STATUS | WA BLAST | PHONE NUMBER | REALM | SERVICE TYPE |
| --- | --- | --- | --- | --- | --- | --- | --- |

Semua sel disimpan sebagai teks. Ini disengaja: nomor telepon seperti
`6281291870453` akan diubah Excel menjadi scientific notation dan kehilangan
digit terakhirnya jika disimpan sebagai angka.

`PHONE NUMBER` tetap satu kolom. Bila sel sumber memuat dua nomor, keduanya
tetap dalam satu sel dengan spasi pemisah, misal
`6281291870453 6281291870453`.

Baris pertama header ditebalkan, diberi warna, dibekukan saat di-scroll, dan
punya filter otomatis.

### Kalau nama file berakhiran `_SEBAGIAN`

Artinya ada halaman yang gagal diambil, sehingga isinya **belum lengkap**.
Berkas seperti ini juga punya sheet kedua bernama `CATATAN` yang menuliskan
halaman mana yang gagal, berapa baris yang berhasil terkumpul, dan apakah
ada baris yang dibuang karena jumlah kolomnya tidak sesuai.

Tidak ada cara file ini terlihat lengkap padahal isinya kurang. Kalau
`_SEBAGIAN` muncul, ambil ulang setelah beberapa saat.

Kalau semua halaman berhasil tapi jumlah baris unik lebih sedikit dari
`N items found`, itu **tidak** berarti ada data hilang: beberapa baris punya
`SERVICE ID` kembar dan digabung menjadi satu. Jumlahnya disebut di panel.

## Cara kerja

### Memilih tabel yang benar

Tabel **tidak** diambil berdasarkan posisi atau urutan di halaman. Semua
`<table>` di dokumen (termasuk di dalam `iframe`) diperiksa, dan yang dipakai
harus memenuhi dua syarat:

1. Header-nya memuat `SERVICE ID` **dan** `IBOOSTER OPER STATUS`.
2. Header-nya memuat minimal satu kolom khas Impacted Services, yaitu
   `DATECREATED`, `WA BLAST`, `PHONE NUMBER`, `REALM`, atau `SERVICE TYPE`.

Halaman yang salah masih bisa punya `SERVICE ID`, jadi satu syarat saja tidak
cukup. Contoh salah yang sudah ditangani: tab `Related Records` punya tabel
lebih besar dan sering muncul lebih dahulu di DOM.

Kalau tabel yang benar ada tapi sedang disembunyikan tab lain, tabel itu tetap
dipakai untuk pembacaan. Tabel milik tab lain tidak pernah dikembalikan pada
jalur pembacaan data.

### Membaca jumlah baris

Total baris dibaca dari keterangan paginator yang **terlihat**, yaitu
`175 items found, displaying 1 to 10.` Keterangan milik tabel yang
disembunyikan diabaikan, karena kalau tidak `displaying` tidak akan pernah naik
saat berpindah halaman.

Pengambilan halaman ada dua jalur:

1. **Jalur cepat (request paralel).** Script mencoba menaikkan *page size*
   ke 100, lalu 50, lalu 25, lalu mengambil 8 halaman sekaligus.
   Tiap halaman diverifikasi terhadap keterangan `displaying X to Y` miliknya;
   yang tidak cocok atau gagal HTTP diulang sampai 3 kali.
2. **Jalur klik.** Kalau jalur cepat tidak bisa diverifikasi (paginasi
   ICEfaces/ADF biasanya butuh submit form, bukan GET bebas), script otomatis
   mengklik tombol "halaman berikutnya" di tabel dan membaca hasilnya dari
   DOM. Satu halaman per waktu, jadi 500 halaman bisa memakan waktu 5-15
   menit. Setelah selesai tabel dikembalikan ke halaman 1.

Tombol halaman berikutnya dipilih berdasarkan penilaian, bukan sekadar
posisi di DOM:

| Kontrol | Dipakai? | Alasan |
| --- | --- | --- |
| `Next`, `next page`, panah kanan | ya | selalu maju satu halaman |
| `Last` | hanya cadangan | melompat ke akhir, jadi tidak pernah jadi pilihan pertama |
| Nomor halaman (`1`, `2`, `3`) | tidak | bukan tombol navigasi |
| `First`, `Previous` | tidak | hanya memundahkan halaman |

Kalau tombol yang diklik tidak memindahkan halaman ke depan, script mencoba
kandidat lain sebelum menyerah, lalu melaporkannya sebagai macet dengan
alasan yang jelas.

Tabel juga dicari ulang setelah tiap klik, karena beberapa widget AJAX
mengganti seluruh elemen `<table>` alih-alih mengisi ulang isinya. Kalau tidak,
script akan menunggu perubahan yang tidak akan pernah datang.

Baris dideduplikasi berdasarkan `SERVICE ID`.

## Diagnosis

Tidak ada tombol diagnosis. Buka Console browser (F12) lalu ketik:

```
botimpactCek()
```

Hasilnya tampil di Console **dan otomatis disalin ke clipboard**, jadi bisa
dipaste ke chat. Laporan memuat:

- URL yang dipakai dan jumlah `iframe`.
- Nilai mentah setiap atribut `src` iframe beserta hasil normalisasinya. Ini
  penting karena widget Icefaces sering menuliskan `src` yang relatif, `#`, atau
  `javascript:void(0)`, yang tidak bisa langsung dipakai sebagai base URL.
- Tab Impacted Service ditemukan atau tidak, dan apakah sudah diklik.
- Tabel apa saja yang ada di halaman, mana yang terlihat, mana yang cocok,
  serta header yang terbaca dari masing-masing tabel.
- Bentuk link paginasi dan kandidat tombol berikutnya yang dipakai.
- Nomor tiket beserta sumbernya, dan jumlah baris di DOM.

Kalau tabel tidak terdeteksi, paste laporan itu. Daftar seluruh tabel di
laporan biasanya langsung menunjukkan tab mana yang sebenarnya aktif.

Semua pesan error juga ditulis ke Console dengan awalan `[BotImpact]`.

## Perbedaan link relatif di dalam iframe

Link relatif di dalam iframe diselesaikan terhadap dokumen iframe itu
sendiri, bukan terhadap dokumen paling atas. `next.xhtml?p=2` di dalam
`/jw/faces/detail.xhtml` harus menjadi `/jw/faces/next.xhtml?p=2`; kalau
dipakai dokumen atas, hasilnya `/faces/next.xhtml?p=2` dan link tersebut 404.

Nilai `href` yang tidak mungkin jadi URL - `#`, kosong, `javascript:`,
`about:blank` - dilewati, dan paginasi tidak berhenti hanya karena satu link
rusak. `href="#"` tidak pernah diubah menjadi URL dokumen, supaya script tidak
melakukan permintaan GET yang tidak pernah dimaksud.

## Catatan efisiensi

- Deteksi tabel hanya menghitung posisi tampilan untuk tabel yang benar-benar
  cocok, bukan untuk semua tabel di halaman.
- Konteks tabel dicari ulang hanya kalau tabel yang dipakai sudah hilang atau
  tidak terlihat, bukan setiap 250 ms.
- Polling panel memakai cek ringan saat panel sudah tampil, dan cek penuh hanya
  saat panel tersembunyi.
- Baris header dihitung satu kali per elemen tabel lalu dipakai ulang.
- Pencarian tab Impacted Service menyaring teks mentah dulu, sebelum teks itu
  dinormalisasi.

Di dalam script tidak ada kode mati yang aman dihapus: seluruh fungsi,
konstanta, dan variabelnya terpakai.

## Catatan

- Sheet bernama `Impacted Services`.
- Batas `MAKS_HALAMAN` = 800 halaman. Dengan 100 baris per halaman ini setara
  80.000 baris, jauh di atas kebutuhan normal.
- Hanya berjalan di domain Insera yang cocok dengan pola di metadata
  userscript.
