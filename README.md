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

Panel muncul sendiri di pojok kanan bawah layar, lalu tombolnya berubah aktif
setelah tabelnya terdeteksi. Hanya ada satu panel: script memakai `@noframes`
supaya tidak ikut disuntikkan ke dalam iframe, tempat tabel Impacted Services
berada. Tanpa itu panel ikut terpasang di dalam iframe dan terlihat seperti
nempel di halaman. Selain itu ada dua salinan script dengan state terpisah,
sehingga file bisa terunduh dua kali dari satu klik.

1. Login ke Insera, cari tiket, buka tiketnya.
2. Klik **Scan & Download XLS**.

Tidak perlu membuka tab atau membuka section lebih dulu. Script melakukan
persiapannya sendiri:

- Kalau halaman membuka tab lain (misalnya `Related Records`), script
  mengklik tab **Impacted Service** sendiri.
- Kalau tabel berada di balik section yang tertutup, script membuka
  **Impacted Service Summary** lebih dulu.

Nomor tiket terdeteksi otomatis dari baris ringkasan yang ada di atas tulisan
**Impacted Service**, lalu ditampilkan di panel. Tidak ada yang perlu diketik.
Nomor itu hanya dipakai pada nama file.

Baris ringkasan diperiksa **lebih dulu**, sebelum URL dan judul tab. Kalau
halaman menampilkan nomor INC-nya sendiri di situ, itu nomor yang benar untuk
tabel yang sedang dibaca, sementara URL dan judul tab kadang masih menyimpan
nomor lama.

Kalau ternyata berbeda, yang dipakai tetap yang dari baris ringkasan, dan
selisihnya ditulis satu per satu di Console:

```
[BotImpact] INC dari baris ringkasan (INC53545268) berbeda dari penyebutan lain di halaman (INC53540001, dari URL). Yang dipakai: INC53545268.
```

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

### Nama file

```
ImpactService_INC53545268_1910.xls
ImpactService_INC53545268_901_SEBAGIAN.xls
```

Formatnya `ImpactService_<INC>_<jumlah baris>.xls`.

Angka di tengah adalah jumlah baris yang **benar-benar masuk file**, bukan
total item yang tertulis di Insera. Untuk dataset yang bermasalah, kedua angka
itu berbeda jauh, dan nama filenya langsung menunjukkannya: `..._901_SEBAGIAN.xls`
untuk 1.910 item berarti file itu tidak lengkap tanpa harus membukanya dulu.

Nomor diambil dari baris ringkasan di atas tulisan **Impacted Service**, contoh:

```
INC53545268 | [SQM GAMAS] | AKSES | DISTRIBUSI | TIF-2| REG-2 | JAKUT | ...
```

Yang dipakai `INC53545268` saja. Syaratnya token INC diikuti tanda `|` dalam
jarak dekat. Halaman detail tiket sering menyebut nomor incident lain di tabel,
breadcrumb, atau panel Expanded; penyebutan seperti itu tidak ditandai pipeline
jadi tidak akan ikut terambil.

Kalau baris ringkasan tidak ditemukan, nomor dicari dari URL, judul tab, field
bertanda tiket, lalu link tabel, dan terakhir seluruh teks halaman. Kalau
semua gagal, namanya jadi `ImpactService_TANPA-TIKET_...`.

Nama file tidak memakai stempel waktu, jadi scan berulang pada tiket yang sama
akan ditangani browser dengan menambah `(1)`, `(2)`, dan seterusnya. Waktu scan
tetiap ada di sheet `CATATAN`.

### Kalau nama file berakhiran `_SEBAGIAN`

Artinya ada halaman yang gagal diambil, sehingga isinya **belum lengkap**.
Berkas seperti ini juga punya sheet kedua bernama `CATATAN` yang menuliskan
halaman mana yang gagal, berapa baris yang berhasil terkumpul, halaman mana
yang URL-nya diturunkan dari link lain, dan apakah ada baris yang dibuang
karena jumlah kolomnya tidak sesuai.

Tidak ada cara file ini terlihat lengkap padahal isinya kurang. Kalau
`_SEBAGIAN` muncul, ambil ulang setelah beberapa saat.

### Kalau jumlah baris lebih sedikit dari `N items found`

Sheet `CATATAN` dan panel membedakan dua hal yang sebelumnya tercampur jadi satu:

| Keadaan | Artinya | File |
| --- | --- | --- |
| Baris identik digabung | Data kembar sungguhan di sumber. Dua baris dengan isi seluruh kolom sama persis. | Utuh |
| `Item tidak masuk file` | Ada data yang benar-benar tidak terbaca: halaman gagal, halaman di luar batas, atau baris hilang tanpa sebab yang diketahui. | `_SEBAGIAN` |
| `Baris dibuang (kolom berlebih)` | Baris terbaca tapi jumlah kolomnya tidak sesuai, jadi tidak bisa dipakai di sheet ini. Dihitung terpisah dari data hilang. | `_SEBAGIAN` |

Ada satu jenis kehilangan yang tidak punya daftar halaman: jumlah baris di file
lebih sedikit dari `N items found`, padahal tidak ada halaman gagal dan tidak ada
halaman di luar batas. Keadaan itu dilaporkan sebagai `N item tidak terbaca
tanpa sebab yang diketahui` dan file tetap diberi `_SEBAGIAN`. Ini yang menangkap
kasus caption `1,910` terbaca `910`: 100 halaman tidak pernah diambil, jadi
tidak masuk daftar halaman gagal karena memang tidak dicoba.

Dua baris dengan `SERVICE ID` sama tapi kolom lain berbeda **keduanya disimpan**.
Kunci anti-duplikat adalah seluruh isi baris, bukan `SERVICE ID` - kalau kuncinya
`SERVICE ID`, entri berbeda yang memakai ID sama akan saling menimpa dan hilang
tanpa jejak.

Untuk dataset besar, concurrency diturunkan agar scan tidak timeout di tengah
jalan: di bawah 1.000 item dipakai 8 request sekaligus, 1.000 item ke atas
dipakai 4.

### Batas jumlah data

Batas keras script adalah 2.000 halaman, jadi 20.000 baris pada 10 baris per
halaman. Scan 10.000 baris (1.000 halaman) masih jauh di bawah batas itu.

Kalau data melebihi batas, baris yang tidak diambil dihitung dan ditulis ke sheet
`CATATAN`, dan file diberi akhiran `_SEBAGIAN`. Pemotongan tidak pernah terjadi
diam-diam.

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
`1,910 items found, displaying 1 to 10.` Keterangan milik tabel yang
disembunyikan diabaikan, karena kalau tidak `displaying` tidak akan pernah naik
saat berpindah halaman.

Angka pada keterangan bisa memakai pemisah ribuan, jadi `1,910` dibaca sebagai
1910, bukan 910. Bug lama di sini costing 1.009 baris yang hilang tanpa
peringatan: sebenarnya ada 191 halaman, script hanya mengambil 91 dan
menganggap file-nya lengkap. Jumlah item dan jumlah halaman selalu dihitung
ulang dari angka yang sudah dibersihkan.

Pengambilan halaman ada dua jalur:

1. **Jalur cepat (request paralel).** Script mencoba menaikkan *page size*
   ke 100, lalu 50, lalu 25, lalu mengambil beberapa halaman sekaligus.
   Tiap halaman diverifikasi terhadap keterangan `displaying X to Y` miliknya;
   yang tidak cocok atau gagal HTTP diulang sampai 3 kali.
2. **Jalur klik.** Kalau jalur cepat tidak bisa dipakai, script otomatis
   mengklik tombol "halaman berikutnya" di tabel dan membaca hasilnya dari
   DOM. Satu halaman per waktu, jadi 500 halaman bisa memakan waktu 5-15
   menit. Setelah selesai tabel dikembalikan ke halaman 1.

Jalur klik dipakai dalam dua keadaan:

- verifikasi halaman 2 gagal, jadi URL sama sekali tidak bisa dipakai; atau
- **sebagian halaman gagal**. Paginator Insera hanya menyediakan link untuk
  halaman 2-8 dan halaman terakhir, jadi halaman 9 dan seterusnya harus
  dibangun sendiri. Kalau URL buatan itu ditolak server, jalur klik melengkapinya
  tanpa membuang hasil jalur cepat yang sudah terkumpul - peta baris diurutkan
  berdasarkan seluruh isi baris, jadi baris kembar dari kedua jalur saling
  menimpa.

Halaman yang masih gagal dicoba ulang lewat URL-nya sampai dua putaran terpisah
sebelum jatuh ke jalur klik, karena untuk tabel 1.000 halaman fallback klik
berarti 1.000 klik berurutan. Kalau halaman yang gagal masih lebih dari 20,
jalur klik dilewati dan halaman-halaman tersebut dicatat di sheet `CATATAN`
sebagai data yang tidak terbaca.

Nomor halaman selalu diteruskan eksplisit ke pemeriksa. Menghitungnya dari
indeks array pernah membuat halaman 3 diperiksa terhadap rentang baris halaman
2, sehingga isinya yang benar ikut dibuang.

### Membangun URL halaman yang tidak punya link

`href` yang dipanen dari DOM dipakai langsung. Halaman yang tidak punya link
tidak dibangun dari base dokumen, melainkan dari link yang **terdekat**:
hanya nomor halaman yang diganti, semua param lain ikut terbawa.

Hal itu penting karena paginator ADF/ICEfaces hanya memahami param `p` kalau
disertai param state milik framework - prefix `d-1234-`, loop state, view state.
Base dokumen tidak punya satu pun itu, sehingga permintaan seperti `?p=9`
diabaikan diam-diam dan server mengirim ulang halaman 1. Gejalanya khas:
`displaying 1 to 10` terus-menerus di halaman mana pun yang dikejar.

Param `ps` juga tidak pernah ditimpa kalau sudah ada di `href`, karena bentuk
berprefiks (`:ps`) hanya itu yang dibaca server.

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

### Deduplikasi dan pemakaian memori

Baris dideduplikasi berdasarkan **seluruh isi baris**, bukan `SERVICE ID`. Dua
baris dengan `SERVICE ID` sama tapi kolom lain berbeda akan keduanya disimpan;
hanya baris yang identik seluruh kolomnya yang digabung.

Baris diekstrak di dalam worker, bukan dikembalikan sebagai HTML. HTML satu
halaman Insera adalah seluruh halaman - tab, script, header - jadi ratusan KB
sampai megabyte, bukan cuma tabel 10 barisnya. Kalau HTML disimpan sampai semua
halaman selesai, 1.000 halaman menahan ratusan MB dan tab bisa crash sebelum scan
selesai.

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
- Nomor tiket beserta sumbernya, dan jumlah baris di DOM. Sumbernya terbaca
  `baris ringkasan di atas Impacted Service` kalau baris INC-nya terdeteksi,
  dan `URL` atau `judul tab` kalau tidak.

Kalau tabel tidak terdeteksi, paste laporan itu. Daftar seluruh tabel di
laparan biasanya langsung menunjukkan tab mana yang sebenarnya aktif.

Laporan ini juga memastikan panel cuma satu: `jumlah_iframe` lebih dari nol
bersama `diIframe: true` pada tabel memastikan tabelnya memang ada di dalam
iframe, dan itu yang diharapkan - panel tetap satu karena script tidak
dijalankan di dalam frame.

Semua pesan error juga ditulis ke Console dengan awalan `[BotImpact]`.

## Uji ulang

Kalau sebelumnya keluar berkas `_SEBAGIAN`, jalankan lagi langkah ini untuk
memastikan perbaikannya bekerja.

1. Di Tampermonkey, buka script lalu tekan **Save** (Ctrl+S) supaya versi
   terbaru tersimpan, lalu **refresh** tab Insera.
2. Buka tiket yang sama seperti sebelumnya.
3. Klik **Scan & Download XLS**.

Yang diharapkan di Console:

```
[BotImpact] Pagination: 8 link dari DOM, sisanya dibangun (prefix ...)
[BotImpact] Halaman tanpa link DOM, URL-nya diturunkan dari link yang terdekat (mis. halaman 8): 9, 10, 11, ...
[BotImpact] Total halaman: 18 (total item 175)
[BotImpact] Konkurensi: 8 request sekaligus (175 item, ambang 1000)
[BotImpact] Verifikasi OK: halaman 2 berisi baris 11-20.
[BotImpact] Ringkasan: 175 baris dari 175 item (100%), 8 kolom, 0 halaman gagal, 0 baris dibuang, 0 baris identik digabung, tidak ada data hilang
```

- Tidak boleh ada baris `Halaman N gagal setelah 3x`.
- `Ringkasan` harus menyebut `0 halaman gagal` dan `tidak ada data hilang`.
- Jumlah halaman harus cocok dengan matematika: `total item / baris per halaman`.
  Untuk 1.910 item pada 10 baris per halaman hasilnya 191 halaman, bukan 91.
  Kalau angkanya 91, pemisah ribuan masih salah baca.
- Nama berkas **tanpa** `_SEBAGIAN`.

Kalau `Halaman tanpa link DOM` tidak muncul, berarti DOM menyediakan link untuk
semua halaman dan tidak ada URL yang perlu diturunkan sendiri.

Kalau masih keluar `_SEBAGIAN`, bandingkan dua baris ini:

- Console: `Halaman gagal lewat URL: ... Melengkapi lewat jalur klik DOM.`
  lalu `Jalur klik DOM selesai: N halaman, M baris unik terkumpul.`
- Sheet `CATATAN`: baris `Halaman gagal` dan `Cara diambil`

Kalau muncul `Halaman gagal lewat URL` tapi `Jalur klik DOM selesai` juga
muncul, berarti klik berhasil menutup semua halaman - dan `_SEBAGIAN` kalau
masih ada berasal dari `Baris dibuang (kolom berlebih)`, bukan dari halaman
yang gagal. Kalau `Jalur klik juga tidak berhasil` yang muncul, kirimkan isi
Console dan hasil `botimpactCek()`.

Kalau jumlah baris di file lebih sedikit dari `N items found`, jangan diasumsikan
itu `SERVICE ID` kembar. Baca baris `Item tidak masuk file` dan `Baris identik
digabung` di sheet `CATATAN`: yang pertama berarti data hilang, yang kedua
cuma kembar sungguhan. Untuk dataset 1.910 item, angka `901 baris` berarti ada
1.009 item yang tidak terbaca - bukan 9 baris kembar.

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
- HTML tiap halaman dibuang begitu barisnya diekstrak, jadi memori yang dipakai
  tidak tumbuh seiring jumlah halaman.

Di dalam script tidak ada kode mati yang aman dihapus: seluruh fungsi,
konstanta, dan variabelnya terpakai.

## Catatan

- Sheet bernama `Impacted Services`.
- Batas `MAKS_HALAMAN` = 2.000 halaman. Pada 10 baris per halaman ini setara
  20.000 baris, jadi scan 10.000 baris (1.000 halaman) masih di bawah batas.
- Jumlah request sekaligus menyesuaikan ukuran dataset: 8 di bawah 1.000 item,
  4 di atasnya.
- Hanya berjalan di domain Insera yang cocok dengan pola di metadata
  userscript.
