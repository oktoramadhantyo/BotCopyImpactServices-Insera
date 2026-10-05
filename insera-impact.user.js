  // ==UserScript==
  // @name         Bot Ambil Impacted Services - Insera
  // @namespace    http://tampermonkey.net/
  // @version      1.0.0
  // @description  Baca seluruh data tab "Impacted Services" di web Insera (semua halaman, bukan cuma yang tampil), lalu unduh jadi file Excel (.xls SpreadsheetML) berisi kolom DATECREATED, SERVICE ID, REGION, IBOOSTER OPER STATUS, WA BLAST, PHONE NUMBER, Realm, Service Type. Nama file memuat nomor tiket yang dicari.
  // @author       diana
  // @match        *://*oss-incident.telkom.co.id/*
  // @grant        GM_xmlhttpRequest
  // @connect      oss-incident.telkom.co.id
  // @run-at       document-end
  // @noframes
  // ==/UserScript==
  //
  // @noframes itu wajib, bukan kosmetik. Tanpa baris ini Tampermonkey
  // menyuntikkan script ke dokumen utama DAN ke setiap iframe same-origin,
  // sehingga panel muncul dua kali: satu di pojok layar, satu lagi di dalam
  // area iframe yang terlihat seperti nempel di halaman Impacted Service.
  // Dua salinan juga berarti dua state terpisah, jadi file bisa terunduh dua
  // kali dari satu klik.
  //
  // Membaca iframe tetap jalan: tabel dibaca lewat contentDocument dari
  // dokumen atas (lihat semuaDokumen()), bukan lewat script yang jalan di
  // dalam iframe.

  /*
  * ============================================================
  *  BOT AMBIL IMPACTED SERVICES - INSERA
  *
  *  Prinsip:
  *    - TIDAK melakukan login. User login manual ke Insera, lalu navigasi
  *      manual: search nomor tiket -> klik tab Impacted Services.
  *    - TIDAK memakai tombol export bawaan Insera. Data dibaca langsung
  *      dari elemen <tr>/<td> di tabel.
  *    - TIDAK butuh server / Apps Script / Python. Murni Tampermonkey.
  *
  *  Alur tombol "Download Excel":
  *    1. cek tabel Impacted Services ada di halaman (atau di dalam iframe)
  *    2. baca keterangan "N items found, displaying 1 to 10" -> total & page size
  *    3. coba naikkan page size (10 -> 100) supaya jumlah request berkurang
  *    4. panen href halaman 2..N dari link pagination di DOM, atau bangun URL
  *    5. verifikasi halaman 2 benar-benar berisi baris 11..20
  *    6. fetch semua halaman paralel, langsung gabung & buang duplikat per SERVICE ID
  *    7. susun XML .xls (SpreadsheetML 2003) lalu unduh
  *
  *  Struktur Project:
  *    insera-impact.user.js  -> script ini
  *    README.md              -> panduan pasang & pakai
  * ============================================================
  */

  (function () {
    "use strict";

    // ==================== KONFIGURASI ====================

    var TAG = "[BotImpact]";

    // Identitas tabel Impacted Services: semua header ini harus ADA.
    var HEADER_KUNCI = ["SERVICE ID", "IBOOSTER OPER STATUS"];

    // Selain dua header di atas, tabel Impacted Services minimal punya salah satu
    // kolom khas ini. Ini mencegah tabel tab lain yang kebetulan punya
    // "SERVICE ID" ikut terbaca.
    var HEADER_KHAMAT = [
      "DATECREATED",
      "REGION",
      "WA BLAST",
      "PHONE NUMBER",
      "REALM",
      "SERVICE TYPE",
    ];

    // Teks tab dan sub-section yang dicari.
    var NAMA_TAB = "IMPACTED SERVICE";
    var NAMA_RINGKASAN = "IMPACTED SERVICE SUMMARY";

    /*
    * Syarat murah untuk menyaring teks sebelum dinormalisasi.
    *
    * normalisasiJudul() hanya merapatkan spasi dan mengubah huruf besar, tidak
    * pernah menambah atau menghapus karakter. Jadi kalau teks mentah tidak
    * memuat NAMA_TAB, pasti tidak akan cocok setelah normalisasi.
    *
    * Dipakai sebagai saringan awal supaya alokasi string (toUpperCase, trim,
    * regex) tidak terjadi untuk tiap elemen di halaman.
    */
    var RE_NAMA_TAB_MURAH = new RegExp(
      NAMA_TAB.split(/\s+/)
        .map(function (bagian) {
          return bagian.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        })
        .join("\\s+"),
      "i"
    );

    // Class yang dipakai framework untuk menyembunyikan elemen di luar CSS
    // standar. Dipakai saat mengecek apakah sebuah tab benar-benar aktif.
    var RE_CLASS_TERSEMBUNYI =
      /\b(X-HIDE|UI-HIDE|HIDDEN|SUPPRESS|HIDDEN-ACCESSIBLE)\b/;

    // Class penanda tab yang sedang aktif.
    var RE_CLASS_AKTIF =
      /\b(ACTIVE|SELECTED|CURRENT|X-TAB-ACTIVE|UI-TABS-ACTIVE|TAB-ACTIVE|IS-OPEN|OPEN)\b/;

    // Class yang menandai batang header sebuah sub-section yang bisa dilipat.
    var RE_BATANG_HEADER =
      /(PANEL-HEADER|PANELHEADER|PANELHEAD|CARD-HEADER|SECTION-HEADER|ACCORDION|COLLAPSE|X-TOOL|(^|-)HEADER$|(^|-)BAR$|(^|-)TITLE$|^HEADING)/;

    // Class yang menandai isi panel milik sebuah batang header.
    var RE_ISI_PANEL =
      /(BODY|CONTENT|WRAPPER|INNER|AREA|PANEL-BODY|X-PANEL-BODY)/;

    // Dipakai hanya bila header tabel tidak terbaca.
    var HEADER_CADANGAN = [
      "DATECREATED",
      "SERVICE ID",
      "REGION",
      "IBOOSTER OPER STATUS",
      "WA BLAST",
      "PHONE NUMBER",
      "REALM",
      "SERVICE TYPE",
    ];

    // Lebar kolom (poin) per nama header, biar sheet enak dibaca.
    var LEBAR_KOLOM = {
      DATECREATED: 118,
      "SERVICE ID": 235,
      REGION: 62,
      "IBOOSTER OPER STATUS": 130,
      "WA BLAST": 72,
      "PHONE NUMBER": 165,
      REALM: 100,
      "SERVICE TYPE": 88,
    };
    var LEBAR_DEFAULT = 110;

    /*
    * Jumlah request yang jalan bersamaan.
    *
    * Dataset besar dikurangi. 1.000 halaman x 8 permintaan serentak bisa memicu
    * timeout session di tengah scan, dan scan yang sudah berjalan halfway lalu
    * gagal jauh lebih mahal daripada yang pelan sampai selesai. Di bawah ambang
    * jumlah halaman kecil, jadi tidak ada alasan menahan scan.
    */
    var KONKURENSI_KECIL = 8;
    var KONKURENSI_BESAR = 4;
    var AMBANG_KONKURENSI = 1000;

    /*
    * Batas jumlah halaman yang boleh diambil.
    *
    * 10.000 baris dengan 10 baris per halaman = 1.000 halaman, jadi batas lama
    * 800 akan memotong 2.000 baris. Batas ini dibuat 2.000 supaya ada ruang
    * sampai 20.000 baris.
    *
    * Pemotongan tidak boleh diam-diam: kalau total halaman melewati batas ini,
    * file WAJIB diberi akhiran _SEBAGIAN dan daftar halaman yang tidak diambil
    * ditulis di sheet CATATAN. Tanpa itu, scan 25.000 baris akan menghasilkan
    * file 20.000 baris yang terlihat lengkap padahal 5.000 baris hilang.
    */
    var MAKS_HALAMAN = 2000;
    var TIMEOUT_MS = 30000;

    /*
    * Berapa kali satu halaman dicoba ulang bila HTTP gagal ATAU isinya bukan
    * halaman yang diharapkan. 2x percobaan ulang = total 3 kali fetch.
    */
    var MAKS_RETRY = 2;

    /*
    * Setelah semua halaman selesai, halaman yang masih gagal dicoba lagi
    * satu per satu dalam putaran tersendiri. Untuk tabel 1.000 halaman,
    * fallback klik berarti 1.000 klik berurutan - tidak realistis - sedangkan
    * mengulang URL halaman yang gagal itu murah dan tidak bergantung pada tombol
    * navigasi sama sekali.
    */
    var PUTARAN_ULANG = 2;

    /*
    * Batas jumlah halaman gagal sebelum fallback klik dicoba. Di atas angka ini
    * klik berurutan akan makan waktu terlalu lama, jadi lebih baik file ditandai
    * sebagian dan daftar halaman gagalnya dibaca di sheet CATATAN.
    */
    var BATAS_LARANG_KLIK = 20;

    // Percobaan menaikkan items-per-page. Script mengetes satu per satu, lalu
    // memakai yang pertama yang benar-benar mengubah "displaying 1 to N".
    var PS_NAIK = [100, 50, 25];

    // Regex nomor tiket Insera, misalnya INC12345678.
    var RE_TIKET = /\bINC[-_ ]?[0-9A-Z]{4,}\b/i;

    // Karakter control yang tidak sah di XML 1.0, dan non-breaking space.
    var RE_CONTROL_XML = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;
    var RE_NBSP = /\u00a0/g;

    // ==================== UTILITAS UMUM ====================

    function log(pesan) {
      console.log(TAG, pesan);
    }

    function waktuSekarang() {
      return new Date();
    }

    // 20261002_143005 -> dipakai di nama file
    function stempelWaktu() {
      var d = waktuSekarang();
      function p(n, l) {
        var s = String(n);
        while (s.length < l) s = "0" + s;
        return s;
      }
      return (
        d.getFullYear() +
        p(d.getMonth() + 1, 2) +
        p(d.getDate(), 2) +
        "_" +
        p(d.getHours(), 2) +
        p(d.getMinutes(), 2) +
        p(d.getSeconds(), 2)
      );
    }

    // Buang karakter illegal di XML lalu escape 5 entity dasar.
    function xmlEsc(nilai) {
      return String(nilai == null ? "" : nilai)
        .replace(RE_CONTROL_XML, "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
    }

    // Samakan isi sel: <br> jadi spasi, rapikan spasi berlebih, buang newline.
    function rapiNilai(nilai) {
      return String(nilai == null ? "" : nilai)
        .replace(RE_NBSP, " ")
        .replace(/\s*\n+\s*/g, " ")
        .replace(/[ \t]{2,}/g, " ")
        .trim();
    }

    function normalisasiJudul(teks) {
      return String(teks || "")
        .replace(/\s+/g, " ")
        .trim()
        .toUpperCase();
    }

    // Buang satu huruf dari awal nama file yang tidak aman.
    function amanUntukNamaBerkas(teks) {
      return String(teks || "").replace(/[^A-Z0-9._-]/g, "-");
    }

    function jeda(ms) {
      return new Promise(function (r) {
        setTimeout(r, ms);
      });
    }

    // ==================== PENYAMAAN URL ====================

    /*
    * Jadikan sebuah nilai menjadi URL absolut, atau null kalau tidak mungkin.
    *
    * new URL() mensyaratkan base yang benar-benar valid. Nilai base dari
    * atribut src milik <iframe> belum tentu demikian: src bisa berupa path
    * relatif ("/jw/..."), "#", atau spasi. Semuanya membuat new URL melempar
    * "Invalid base URL" dan menghentikan seluruh proses pengambilan.
    *
    * location.href selalu berupa URL absolut, jadi dipakai sebagai base
    * terakhir. Dengan begitu fungsi ini tidak mungkin gagal total.
    */
    function absolutURL(nilai, baseCadangan) {
      /*
      * Base WAJIB dipakai untuk menyelesaikan nilai relatif, dan itu bukan soal
      * selera: isi iframe berada di path yang berbeda dari dokumen atas.
      * "next.xhtml?p=2" di dalam /jw/faces/detail.xhtml harus menjadi
      * /jw/faces/next.xhtml?p=2. Kalau dokumen atas dipakai sebagai base,
      * hasilnya /faces/next.xhtml?p=2 dan link itu akan 404.
      *
      * Nilai dari widget sering tidak bisa jadi URL sama sekali: "#", spasi,
      * "javascript:void(0)", atau "about:blank". Semua itu dilewati lalu
      * fallback ke base yang valid. new URL() karena itu tidak pernah menerima
      * base tidak valid - itulah sumber "Invalid base URL".
      */
      var mentah = nilai ? String(nilai).trim() : "";
      var dasar = [];
      var i;
      var u;

      if (mentah && mentah.charAt(0) !== "#") {
        // Sudah punya skema sendiri, tidak butuh base.
        try {
          u = new URL(mentah);
          if (u.protocol === "http:" || u.protocol === "https:") return u.href;
        } catch (e) {
          /* nilai relatif, lanjut ke base */
        }
        // Base dari iframe lebih tepat daripada dokumen atas.
        if (baseCadangan) dasar.push(baseCadangan);
        dasar.push(location.href);
        for (i = 0; i < dasar.length; i++) {
          try {
            u = new URL(mentah, dasar[i]);
            if (u.protocol === "http:" || u.protocol === "https:") return u.href;
          } catch (e) {
            /* base ini tidak valid, coba yang berikutnya */
          }
        }
      }

      // Tidak ada nilai yang bisa dipakai: kembalikan base yang valid saja.
      if (baseCadangan) dasar.push(baseCadangan);
      dasar.push(location.href);
      for (i = 0; i < dasar.length; i++) {
        try {
          u = new URL("", dasar[i]);
          if (u.protocol === "http:" || u.protocol === "https:") return u.href;
        } catch (e) {
          /* bukan base yang valid */
        }
      }
      return null;
    }

    /*
    * URL absolut dari sebuah <a>.
    *
    * Base dari caller dipakai lebih dulu karena itu base dokumen tempat anchor
    * itu ditemukan, jadi hasilnya konsisten dengan ctx.base. Properti .href
    * milik browser dipakai sebagai cadangan: untuk elemen yang belum punya base
    * dokumen - hasil DOMParser, atau iframe yang sr-nya "#" - .href bisa kosong
    * atau bahkan menunjuk ke dokumen lain.
    */
    function hrefTAnchor(a, base) {
      if (!a) return null;
      var mentah = String((a.getAttribute && a.getAttribute("href")) || "").trim();

      // "#" dan href kosong bukan navigasi. Kalau diteruskan ke absolutURL(),
      // nilai kosong akan dijawab dengan base dokumen, dan itu berubah jadi
      // permintaan GET yang tidak pernah dimaksud.
      if (!mentah || mentah.charAt(0) === "#") return null;

      var dariBase = absolutURL(mentah, base);
      if (dariBase) return dariBase;

      try {
        var absolut = a.href;
        if (absolut && /^(https?):/i.test(absolut)) return absolut;
      } catch (e) {
        /* elemen tidak punya base dokumen */
      }
      return null;
    }

    // ==================== VISIBILITAS: MENYINGKIRKAN TAB YANG NON-AKTIF ====================

    /*
    * Beberapa lingkungan tidak menghitung layout sama sekali, sehingga
    * getClientRects() selalu mengembalikan daftar kosong. Kalau sinyal itu ikut
    * dipakai, setiap elemen akan terbaca tersembunyi. Karena itu sinyal kotak
    * tampilan hanya dipercaya kalau document.body sendiri punya kotak, artinya
    * dokumen ini memang sedang dirender.
    */
    function dokumenPunyaLayout(doc) {
      try {
        return !!(
          doc &&
          doc.body &&
          typeof doc.body.getClientRects === "function" &&
          doc.body.getClientRects().length > 0
        );
      } catch (e) {
        return false;
      }
    }

    /*
    * Halaman detail tiket Insera punya beberapa tab (Incident, Related Records,
    * Customer Information, Impacted Service, Config ONT). Tabel milik tab yang
    * tidak aktif tetap ada di DOM, hanya disembunyikan lewat CSS. Karena itu
    * "cari tabel pertama" tidak akan pernah benar - harus disaring lebih dulu.
    *
    * Fungsi ini menelusuri elemen beserta seluruh induknya. Kalau salah satu
    * disembunyikan, elemen dianggap tidak terlihat.
    */
    function terlihatDiLayar(el) {
      if (!el || el.nodeType !== 1) return false;
      var doc = el.ownerDocument;
      var win = (doc && doc.defaultView) || window;
      if (!win || typeof win.getComputedStyle !== "function") return true;

      var node = el;
      var kedalaman = 0;
      while (node && node.nodeType === 1 && kedalaman < 100) {
        if (node.hasAttribute && node.hasAttribute("hidden")) return false;

        var gaya = null;
        try {
          gaya = win.getComputedStyle(node);
        } catch (e) {
          gaya = null;
        }
        if (gaya) {
          if (gaya.display === "none") return false;
          if (gaya.visibility === "hidden" || gaya.visibility === "collapse") return false;
          if (gaya.opacity === "0") return false;
        }

        // Sebagian widget menyembunyikan elemen lewat class, bukan CSS standar.
        var kelas = normalisasiJudul(node.getAttribute && node.getAttribute("class"));
        if (kelas && RE_CLASS_TERSEMBUNYI.test(kelas)) return false;

        if (node.tagName === "BODY" || node.tagName === "HTML") break;
        node = node.parentElement;
        kedalaman++;
      }

      // Terakhir: elemen harus benar-benar punya tempat di layar. Pemeriksaan ini
      // hanya sah kalau dokumen ini memang menghitung layout.
      if (dokumenPunyaLayout(el.ownerDocument)) {
        try {
          if (typeof el.getClientRects === "function" && el.getClientRects().length === 0) {
            return false;
          }
        } catch (e) {
          /* abaikan, dianggap terlihat */
        }
      }
      return true;
    }

    // ==================== DETEKSI TABEL ====================

    /*
    * Baris yang memuat header tabel.
    *
    * Icefaces biasanya memakai <thead>, tapi tidak selalu. Beberapa grid juga
    * menaruh baris filter atau checkbox DI ATAS baris header asli, jadi baris
    * pertama belum tentu yang benar. Karena itu baris paling banyak teksnya yang
    * dipilih, bukan sekadar baris pertama.
    */
    /*
    * Cache satu entri untuk baris header.
    *
    * Header sebuah tabel tidak berubah selama elemen <table>-nya masih
    * terpasang, sedangkan barisHeader() dipanggil berulang kali: saat deteksi
    * tabel, tiap polling render (250 ms), lalu lagi saat membaca data. Tanpa
    * cache, tiap panggilan mengulang querySelectorAll("tr") untuk seluruh
    * tabel beserta normalisasi teks tiap <th>-nya.
    *
    * Cache hanya dipakai ulang kalau elemen baris header masih benar-benar
    * berada di dalam kontainer, sehingga kasus AJAX yang mengganti <thead>
    * ikut tertangani.
    */
    var cacheBarisHeader = null;

    function barisHeader(container) {
      if (!container || !container.querySelectorAll) return null;

      if (
        cacheBarisHeader &&
        cacheBarisHeader.kontainer === container &&
        cacheBarisHeader.baris &&
        container.contains(cacheBarisHeader.baris)
      ) {
        return cacheBarisHeader.baris;
      }

      var hasil = cariBarisHeader(container);
      cacheBarisHeader = { kontainer: container, baris: hasil };
      return hasil;
    }

    function cariBarisHeader(container) {
      if (!container || !container.querySelectorAll) return null;
      var baris = [];
      var i, j;
      var thead;
      try {
        thead = container.querySelectorAll("thead tr");
      } catch (e) {
        thead = [];
      }
      for (i = 0; i < thead.length; i++) baris.push(thead[i]);

      if (!baris.length) {
        var trs = container.querySelectorAll("tr");
        for (i = 0; i < trs.length; i++) {
          if (trs[i].querySelectorAll("th").length > 0) baris.push(trs[i]);
        }
      }
      if (!baris.length) return null;
      if (baris.length === 1) return baris[0];

      var terbaik = baris[0];
      var skorTerbaik = -1;
      for (i = 0; i < baris.length; i++) {
        var ths = baris[i].querySelectorAll("th");
        var jumlahTeks = 0;
        for (j = 0; j < ths.length; j++) {
          if (normalisasiJudul(ths[j].innerText || ths[j].textContent)) jumlahTeks++;
        }
        // Teks paling banyak menang; jumlah <th> hanya sebagai pemutus seri.
        var skor = jumlahTeks * 1000 + ths.length;
        if (skor > skorTerbaik) {
          skorTerbaik = skor;
          terbaik = baris[i];
        }
      }
      return terbaik;
    }

    // Gabungkan semua teks header jadi satu string, dipisah penanda yang tidak
    // mungkin membentuk nama kolom baru ("SERVICE" + "ID" tidak jadi "SERVICE ID").
    function teksHeaderGabung(container) {
      var headTr = barisHeader(container);
      if (!headTr) return "";
      var ths = headTr.querySelectorAll("th");
      var bagian = [];
      for (var i = 0; i < ths.length; i++) {
        bagian.push(normalisasiJudul(ths[i].innerText || ths[i].textContent));
      }
      return bagian.join(" | ");
    }

    /*
    * Apakah tabel ini tabel Impacted Services?
    * Semua HEADER_KUNCI harus ada, dan minimal satu HEADER_KHAMAT juga ada.
    * Syarat kedua itu yang membuang tabel tab lain.
    */
    function adakahTabelImpact(container) {
      if (!container) return false;
      var gabung = teksHeaderGabung(container);
      if (!gabung) return false;

      var i;
      for (i = 0; i < HEADER_KUNCI.length; i++) {
        if (gabung.indexOf(normalisasiJudul(HEADER_KUNCI[i])) < 0) return false;
      }
      for (i = 0; i < HEADER_KHAMAT.length; i++) {
        if (gabung.indexOf(normalisasiJudul(HEADER_KHAMAT[i])) >= 0) return true;
      }
      return false;
    }

    // Berapa banyak kolom data yang dimiliki sebuah tabel.
    function jumlahKolomTabel(t) {
      var headTr = barisHeader(t);
      if (!headTr) return 0;
      return indeksTerlihat(headTr).length;
    }

    // Tabel navigasi/layout, bukan tabel data.
    function tabelNavigasi(t) {
      if (!t) return true;
      var cls = String((t.getAttribute && t.getAttribute("class")) || "");
      return /datepicker|calendar|pagelinks|pagination|navigation|layout/i.test(cls);
    }

    /*
    * Cari tabel Impacted Services di dalam satu dokumen.
    *
    * Urutan penilaian:
    * 1. Tabel yang header-nya cocok DAN sedang terlihat -> langsung dipakai.
    * 2. Tabel yang header-nya cocok tapi tersembunyi -> dipakai sebagai cadangan
    *    (isi HTML-nya tetap lengkap, jadi tetap aman dibaca).
    * 3. bestEffort: tabel non-cocok dengan kolom terbanyak, hanya untuk pesan
    *    error supaya user tahu kolom apa yang sebenarnya terbaca.
    */
    function cariTabelDiDokumen(doc, bestEffort) {
      if (!doc || !doc.querySelectorAll) return null;
      var tabel = doc.querySelectorAll("table");
      var cadangan = null;
      var tebakan = null;
      var kolomTebakan = 0;

      for (var i = 0; i < tabel.length; i++) {
        var t = tabel[i];
        if (tabelNavigasi(t)) continue;

        var cocok = adakahTabelImpact(t);

        if (cocok) {
          /*
          * Visibilitas hanya dibutuhkan untuk tabel yang header-nya cocok.
          *
          * terlihatDiLayar menelusuri sampai 100 elemen induk dan memanggil
          * getComputedStyle di tiap tingkat, yang memaksa browser menghitung
          * ulang tata letak. Memanggilnya untuk setiap tabel di halaman —
          * padahal sebagian besar bukan tabel Impacted Service — hanya
          * memboroskan waktu tanpa hasil.
          */
          if (terlihatDiLayar(t)) return t;
          if (!cadangan) cadangan = t;
          continue;
        }

        // bestEffort hanya untuk pesan error, jadi tabel yang tidak cocok tetap
        // perlu dinilai kelihatan dan lebar kolomnya.
        if (bestEffort && terlihatDiLayar(t)) {
          var n = jumlahKolomTabel(t);
          if (n > kolomTebakan) {
            kolomTebakan = n;
            tebakan = t;
          }
        }
      }
      return cadangan || tebakan || null;
    }

    function cariKonteksDiDokumen(doc, base, bestEffort) {
      var t = cariTabelDiDokumen(doc, bestEffort);
      if (!t) return null;
      return {
        doc: doc,
        tabel: t,
        // Dinormalkan satu kali di sini, jadi semua pemakai ctx.base (URL halaman
        // dan link paginasi) pasti menerima URL absolut yang valid. Kalau tidak,
        // new URL(href, ctx.base) bisa melempar "Invalid base URL".
        base: absolutURL(base, location.href) || location.href,
        identity: adakahTabelImpact(t),
        terlihat: terlihatDiLayar(t),
      };
    }

    /*
    * Kumpulkan dokumen yang boleh dicari: dokumen utama plus setiap iframe yang
    * bisa diakses (iframe cross-origin dilewati).
    */
    function semuaDokumen() {
      var hasil = [{ doc: document, base: location.href, bingkai: null }];
      var frames;
      try {
        frames = document.querySelectorAll("iframe");
      } catch (e) {
        return hasil;
      }
      for (var i = 0; i < frames.length; i++) {
        var fdoc = null;
        try {
          fdoc = frames[i].contentDocument;
          if (!fdoc && frames[i].contentWindow) fdoc = frames[i].contentWindow.document;
        } catch (e) {
          fdoc = null;
        }
        if (!fdoc || !fdoc.querySelectorAll) continue;
        var src = frames[i].getAttribute("src");
        hasil.push({
          doc: fdoc,
          // src wajib dinormalkan: atributnya bisa relatif ("/jw/...") atau "#",
          // yang keduanya tidak sah dipakai sebagai base new URL().
          base: absolutURL(src, location.href) || location.href,
          bingkai: frames[i],
        });
      }
      return hasil;
    }

    /*
    * Cari konteks tabel Impacted Services.
    *
    * Ini TIDAK lagi menerima tabel tab lain. Sebuah konteks hanya dikembalikan
    * bila header-nya benar-benar cocok, sehingga tabel "Related Records" yang
    * punya 9-10 kolom tidak pernah terpilih.
    *
    * Parameter bestEffort hanya untuk pesan diagnosis, tidak untuk pembacaan data.
    */
    function cariKonteksTabel(bestEffort) {
      var docs = semuaDokumen();
      var ctx = null;
      var tersembunyi = null;
      var tebakan = null;
      var i;

      /*
      * Satu kali loop untuk dua kebutuhan sekaligus: yang terlihat langsung
      * dipakai, yang tersembunyi disimpan sebagai cadangan.
      *
      * Versi lama memakai dua pass terpisah dengan pemanggilan
      * cariKonteksDiDokumen() yang persis sama. Pass kedua hanya berbeda di
      * syarat keluar, sehingga seluruh pemindaian tabel terulang sekali lagi
      * persis ketika tabel target sedang tersembunyi - yaitu keadaan paling
      * umum, karena tab Impacted Service belum diklik.
      */
      for (i = 0; i < docs.length; i++) {
        ctx = cariKonteksDiDokumen(docs[i].doc, docs[i].base, false);
        if (ctx && ctx.identity) {
          if (ctx.terlihat) return ctx;
          if (!tersembunyi) tersembunyi = ctx;
        }
      }
      if (tersembunyi) return tersembunyi;

      // Best effort: hanya untuk menjelaskan ke user kolom apa yang terbaca.
      if (bestEffort) {
        for (i = 0; i < docs.length; i++) {
          ctx = cariKonteksDiDokumen(docs[i].doc, docs[i].base, true);
          if (ctx) return ctx;
        }
      }
      return null;
    }

    // ==================== TAB & SUB-SECTION ====================

    /*
    * Halaman detail tiket punya tab: Incident, Related Records,
    * Customer Information, Impacted Service, Config ONT.
    *
    * Script tidak lagi berasumsi user sudah mengklik tab yang benar. Blok di
    * bawah mencari tab "Impacted Service" lalu mengkliknya, dan membuka
    * sub-section "Impacted Service Summary" bila masih tertutup.
    */

    // Selector yang umum dipakai framework tab, dari yang paling spesifik.
    var SEL_TAB = [
      '[role="tab"]',
      '[role="tablist"] a',
      '[role="tablist"] li',
      ".x-tab-strip .x-tab",
      ".x-tab-strip .x-tab-item",
      ".ui-tabs-anchor",
      ".nav-tabs a",
      ".nav-tabs li",
      ".tab-nav a",
      ".tab-nav li",
      ".tabs li a",
      ".tabs li",
      '[data-toggle="tab"]',
      '[data-tab]',
      '[data-tab-name]',
      "a[href*='mpacted']",
      "button[title*='mpacted']",
      "input[value*='mpacted']",
      "span[title*='mpacted']",
      "a[onclick]",
      "li a",
      "a",
    ];

    var SEL_SECTION = [
      '[aria-expanded]',
      "summary",
      ".x-panel-header",
      ".panel-heading",
      ".card-header",
      ".section-header",
      ".toggle",
      ".toggler",
      "[data-toggle='collapse']",
      "[class*='collaps']",
      "[class*='expand']",
      "h1, h2, h3, h4, h5, h6",
      "legend",
    ];

    // Teks elemen setelah dibersihkan dari spasi/icon/angka navigasi.
    function teksElement(el) {
      if (!el) return "";
      var t = el.innerText || el.textContent || "";
      return normalisasiJudul(t);
    }

    /*
    * Cari elemen tab "Impacted Service" di dalam satu dokumen.
    * Mengembalikan { el, aktif } atau null.
    */
    function cariTabImpacted(doc) {
      if (!doc || !doc.querySelectorAll) return null;
      var i, j, list, el;

      // Pass 1: pakai selector framework yang sudah umum dipakai.
      for (i = 0; i < SEL_TAB.length; i++) {
        try {
          list = doc.querySelectorAll(SEL_TAB[i]);
        } catch (e) {
          continue;
        }
        for (j = 0; j < list.length; j++) {
          el = list[j];
          var t = teksElement(el);
          if (t.indexOf(NAMA_TAB) >= 0 && el.querySelectorAll("table").length === 0) {
            return { el: el, aktif: tabSudahAktif(el) };
          }
        }
      }

      /*
      * Pass 2: pindai elemen yang teksnya cocok, lalu ambil yang paling kecil.
      * Ini menangkap tab yang dibuat tanpa class framework biasa.
      *
      * Tag dibatasi pada yang memang bisa jadi pemicu tab. Tanpa batas ini,
      * heading "Impacted Service Summary" ikut cocok dan dianggap sebagai tab.
      */
      try {
        var semua = doc.querySelectorAll(
          "a, button, li, span, input[type='submit'], input[type='button'], [role='tab']"
        );
        var terbaik = null;
        var panjangTerpendek = 9999;
        for (i = 0; i < semua.length; i++) {
          el = semua[i];
          // Saringan murah dulu: baca teks mentah, dan hanya normalisasi kalau
          // memang memuat nama tab. Halaman Insera punya ratusan elemen
          // a/span/li, dan menormalisasi semuanya tiap polling itu mahal.
          var mentah = el.tagName === "INPUT"
            ? (el.value || "")
            : (el.innerText || el.textContent || "");
          if (!RE_NAMA_TAB_MURAH.test(mentah)) continue;

          var teks = normalisasiJudul(mentah);
          if (teks.indexOf(NAMA_TAB) < 0) continue;
          if (teks.length > 40) continue; // terlalu besar, itu bukan tombol tab
          // Jangan pilih elemen yang hanya membungkus teks, bukan pemicunya.
          if (el.tagName === "SPAN" && !el.hasAttribute("title")) continue;
          if (teks.length < panjangTerpendek) {
            panjangTerpendek = teks.length;
            terbaik = el;
          }
        }
        if (terbaik) return { el: terbaik, aktif: tabSudahAktif(terbaik) };
      } catch (e) {
        /* abaikan */
      }
      return null;
    }

    // Apakah tab ini sudah aktif?
    function tabSudahAktif(el) {
      if (!el) return false;
      if (el.getAttribute("aria-selected") === "true") return true;
      if (el.getAttribute("aria-expanded") === "true") return true;
      if (el.getAttribute("selected") === "true") return true;
      var kelas = normalisasiJudul(el.getAttribute && el.getAttribute("class"));
      if (kelas && RE_CLASS_AKTIF.test(kelas)) return true;
      return false;
    }

    // Klik elemen dengan urutan event yang lengkap, jadi framework JSF/ADF pun
    // bereaksi seperti klik mouse asli.
    function klikElemen(el) {
      if (!el) return false;
      try {
        var win = (el.ownerDocument && el.ownerDocument.defaultView) || window;
        var urutan = ["mousedown", "mouseup", "click"];
        for (var i = 0; i < urutan.length; i++) {
          var ev = new win.MouseEvent(urutan[i], {
            bubbles: true,
            cancelable: true,
            view: win,
          });
          el.dispatchEvent(ev);
        }
        return true;
      } catch (e) {
        try {
          el.click();
          return true;
        } catch (e2) {
          return false;
        }
      }
    }

    /*
    * Cari elemen yang mewakili sub-section "Impacted Service Summary".
    * Mengembalikan null kalau teksnya tidak ditemukan, supaya script tidak
    * menebak dan menutup bagian yang salah.
    */
    function cariElemenRingkasan(doc) {
      if (!doc || !doc.querySelectorAll) return null;
      var semua;
      try {
        semua = doc.querySelectorAll(SEL_SECTION.join(","));
      } catch (e) {
        return null;
      }
      // Cari yang paling spesifik: teks persis lebih diutamakan daripada teks
      // yang memuat frasa tersebut di tengah.
      var memuat = null;
      for (var i = 0; i < semua.length; i++) {
        var t = teksElement(semua[i]);
        if (!t) continue;
        if (t === NAMA_RINGKASAN) return semua[i];
        if (t.indexOf(NAMA_RINGKASAN) >= 0 && !memuat) memuat = semua[i];
      }
      if (memuat) return memuat;

      // Selector terakhir: elemen yang teksnya persis sama, apa pun strukturnya.
      try {
        var semua2 = doc.querySelectorAll(
          "div, span, h1, h2, h3, h4, h5, h6, legend, summary, td, th, label"
        );
        for (var j = 0; j < semua2.length; j++) {
          if (teksElement(semua2[j]) === NAMA_RINGKASAN) return semua2[j];
        }
      } catch (e) {
        /* abaikan */
      }
      return null;
    }

    /*
    * Induk terdekat yang berperan sebagai batang header sebuah sub-section.
    *
    * Penting: pada pola markup yang umum, judul dan tombol panah adalah SAUDARA,
    * bukan anak. Contohnya:
    *
    *   <div class="x-panel-header" aria-expanded="false">
    *     <h3>Impacted Service Summary</h3>
    *     <a class="x-tool-expand">...</a>
    *   </div>
    *
    * Kalau tombol dicari di dalam <h3> saja, tombolnya tidak akan ketemu dan
    * klik justru jatuh ke teks judul, yang tidak melakukan apa-apa.
    */
    function batangHeader(el) {
      if (!el) return null;
      if (el.tagName === "A" || el.tagName === "BUTTON" || el.tagName === "SUMMARY") {
        return el; // elemen itu sendiri sudah jadi pemicunya
      }
      var node = el;
      for (var i = 0; i < 6 && node; i++) {
        var kelas = normalisasiJudul(node.getAttribute && node.getAttribute("class"));
        if (kelas && RE_BATANG_HEADER.test(kelas)) return node;
        if (node.hasAttribute && node.hasAttribute("aria-expanded")) return node;
        if (/^(HEADER|H3|H4|H5|H6|LEGEND|DT)$/.test(node.tagName) === false) {
          // Lanjut naik sampai ketemu batang.
        }
        node = node.parentElement;
      }
      return el;
    }

    /*
    * Cari isi panel milik sebuah batang header: saudara tepat setelah batang itu.
    */
    function isiPanelDariBatang(batang) {
      if (!batang) return null;
      var n = batang.nextElementSibling;
      while (n) {
        var kelas = normalisasiJudul(n.getAttribute("class"));
        if (kelas && RE_ISI_PANEL.test(kelas)) return n;
        if (n.querySelector && n.querySelector("table")) return n;
        n = n.nextElementSibling;
      }
      // Cadangan: cari di dalam induk batang.
      try {
        var induk = batang.parentElement;
        if (!induk) return null;
        var kandidat = induk.querySelectorAll("div, td, section");
        for (var i = 0; i < kandidat.length; i++) {
          var k = normalisasiJudul(kandidat[i].getAttribute("class"));
          if (k && RE_ISI_PANEL.test(k)) return kandidat[i];
        }
      } catch (e) {
        /* abaikan */
      }
      return null;
    }

    /*
    * Cari sub-section "Impacted Service Summary" yang masih tertutup, lalu buka.
    * Mengembalikan true bila ada yang diklik.
    */
    function bukaRingkasanImpacted(doc) {
      var target = cariElemenRingkasan(doc);
      if (!target) {
        // Tanpa teks yang cocok, jangan menebak. Biarkan user yang membuka manual.
        return false;
      }

      /*
      * <details> menyimpan statusnya di atribut "open", bukan di computed style.
      * Isi <details> yang tertutup tetap punya display:block, jadi statusnya
      * harus dibaca dari atributnya.
      */
      var details = null;
      try {
        details = target.closest ? target.closest("details") : null;
      } catch (e) {
        details = null;
      }
      if (details) {
        if (details.hasAttribute("open")) return false; // sudah terbuka
        var ringkasan = details.querySelector("summary");
        klikElemen(ringkasan || details);
        return true;
      }

      var batang = batangHeader(target) || target;

      // Sudah terbuka? Jangan diklik, atau section-nya malah tertutup lagi.
      if (batang.getAttribute("aria-expanded") === "true") return false;
      var isi = isiPanelDariBatang(batang);
      if (isi && terlihatDiLayar(isi)) return false; // sudah terbuka

      // Cari tombol toggle di seluruh batang header, bukan cuma di dalam judul.
      var tombol = cariTombolToggle(batang);
      if (tombol) {
        klikElemen(tombol);
        return true;
      }
      // Tidak ada tombol terpisah: batang header itu sendiri yang jadi pemicu.
      klikElemen(batang);
      return true;
    }

    // Tombol toggle (ikon panah atau chevron) milik sebuah batang header.
    function cariTombolToggle(header) {
      if (!header) return null;
      var candidates;
      try {
        candidates = header.querySelectorAll(
          "a, button, .x-tool-expand, .x-tool-toggle, .toggle-icon, .icon, i, span, img"
        );
      } catch (e) {
        return null;
      }
      for (var i = 0; i < candidates.length; i++) {
        var el = candidates[i];
        if (el === header) continue;
        var kelas = normalisasiJudul(el.getAttribute("class"));
        var judul = normalisasiJudul(el.getAttribute("title") || el.getAttribute("aria-label") || "");
        if (kelas && /(EXPAND|COLLAPS|TOGGLE|CHEVRON|CARET|ARROW)/.test(kelas)) return el;
        if (judul && /(EXPAND|COLLAPS|TOGGLE|SHOW|HIDE)/.test(judul)) return el;
        // Ikon kecil tanpa teks, misalnya chevron atau ikon panah.
        if (!teksElement(el) && i < 6 && (el.tagName === "I" || el.tagName === "IMG" || el.tagName === "SPAN")) {
          return el;
        }
      }
      return null;
    }

    /*
    * Pastikan tab "Impacted Service" aktif dan bagian ringkasannya terbuka.
    * Dipanggil dari jalankan() sebelum tabel dicari.
    *
    * Balikan: { tabAda, tabDiklik, ringkasanDibuka, sudahSiap }
    */
    async function siapkanHalamanImpacted() {
      var laporan = {
        tabAda: false,
        tabDiklik: false,
        ringkasanDibuka: false,
        sudahSiap: false,
      };

      // Kalau tabelnya sudah terlihat, tidak perlu ada sentuhan sama sekali.
      var cepat = cariKonteksTabel();
      if (cepat && cepat.identity && cepat.terlihat) {
        laporan.tabAda = true;
        laporan.sudahSiap = true;
        return laporan;
      }

      var docs = semuaDokumen();
      var i, tab;

      // 1. Cari dan aktifkan tab.
      for (i = 0; i < docs.length; i++) {
        tab = cariTabImpacted(docs[i].doc);
        if (!tab) continue;
        laporan.tabAda = true;
        if (!tab.aktif) {
          klikElemen(tab.el);
          laporan.tabDiklik = true;
          log("Klik tab Impacted Service secara otomatis.");
        }
        break;
      }

      if (laporan.tabDiklik) {
        // Beri waktu tab memuat isi lewat AJAX.
        await jeda(900);
      }

      // 2. Buka sub-section ringkasan bila masih tertutup.
      docs = semuaDokumen();
      for (i = 0; i < docs.length; i++) {
        if (bukaRingkasanImpacted(docs[i].doc)) {
          laporan.ringkasanDibuka = true;
          log("Membuka bagian Impacted Service Summary.");
          await jeda(700);
          break;
        }
      }

      laporan.sudahSiap = !!cariKonteksTabel();
      return laporan;
    }

    // ==================== BACA BARIS ====================

    // Indeks kolom yang benar-benar data: buang checkbox & kolom internal.
    function indeksTerlihat(headerTr) {
      var idx = [];
      if (!headerTr) return idx;
      var ths = headerTr.querySelectorAll("th");
      for (var i = 0; i < ths.length; i++) {
        var th = ths[i];
        var kelas = (th.className || "") + " " + (th.getAttribute("class") || "");
        var isCheckbox =
          /select_checkbox/i.test(kelas) || !!th.querySelector('input[type="checkbox"]');
        var isParentId = /column_C_PARENT_ID/i.test(kelas);
        if (isCheckbox || isParentId) continue;
        idx.push(i);
      }
      return idx;
    }

    function bacaHeader(container) {
      var headTr = barisHeader(container);
      if (!headTr) return [];
      var idx = indeksTerlihat(headTr);
      var hasil = [];
      var ths = headTr.querySelectorAll("th");
      for (var i = 0; i < idx.length; i++) {
        var th = ths[idx[i]];
        if (!th) continue;
        var nama = rapiNilai(th.innerText || th.textContent).toUpperCase() || "KOLOM_" + (i + 1);
        //.some header muncul dua kali, misalnya PHONE NUMBER -> label unik
        var n = 2;
        while (hasil.indexOf(nama) >= 0) nama = nama + "_" + n++;
        hasil.push(nama);
      }
      return hasil;
    }

    /*
    * Ambil seluruh <td> pada satu baris.
    *
    * Pagar akurasi: bila jumlah kolom baris tidak sama dengan header, posisi
    * semua kolom setelahnya bisa meleset dan file jadi berisi data di kolom
    * yang salah. Karena itu:
    *   - baris yang lebih pendek  -> diisi kosong di ujung, isinya tetap disimpan
    *   - baris yang lebih panjang -> dibuang & dicatat, karena tidak diketahui
    *                                 kolom mana yang sebenarnya berlebih
    *
    * Balikan: { rows,Pendek, Panjang, contoh }
    */
    function bacaBaris(container, idxTerlihat, panjangHarapan) {
      var out = [];
      var pendek = 0;
      var panjang = 0;
      var contoh = [];

      var trs = container.querySelectorAll("tbody tr");
      if (trs.length === 0) trs = container.querySelectorAll("tr");

      for (var i = 0; i < trs.length; i++) {
        var tr = trs[i];
        if (tr.querySelector("th")) continue;
        var tds = tr.querySelectorAll("td");
        if (tds.length === 0) continue;

        var vals = [];
        for (var j = 0; j < tds.length; j++) {
          if (idxTerlihat && idxTerlihat.indexOf(j) === -1) continue;
          vals.push(rapiNilai(tds[j].textContent));
        }

        // Buang baris yang benar-benar kosong semua (baris spacer di bawah tabel).
        if (!vals.some(function (v) { return v !== ""; })) continue;

        if (panjangHarapan && vals.length !== panjangHarapan) {
          if (vals.length < panjangHarapan) {
            pendek++;
            while (vals.length < panjangHarapan) vals.push("");
            if (contoh.length < 5) {
              contoh.push("baris " + (i + 1) + " hanya " + vals.length + " kolom, diisi kosong");
            }
          } else {
            panjang++;
            if (contoh.length < 5) {
              contoh.push(
                "baris " + (i + 1) + " punya " + vals.length +
                  " kolom, melebihi " + panjangHarapan + " - dibuang"
              );
            }
            continue;
          }
        }
        out.push(vals);
      }

      return { rows: out, Pendek: pendek, Panjang: panjang, contoh: contoh };
    }

    // Parse baris dari HTML hasil fetch (halaman lain).
    function ekstrakDariHTML(html, base, panjangHarapan) {
      var doc = new DOMParser().parseFromString(html, "text/html");
      var ctx = cariKonteksDiDokumen(doc, base);
      if (!ctx) return { header: [], rows: [], Pendek: 0, Panjang: 0, contoh: [] };
      var headTr = barisHeader(ctx.tabel);
      var dibaca = bacaBaris(ctx.tabel, indeksTerlihat(headTr), panjangHarapan);
      return {
        header: bacaHeader(ctx.tabel),
        rows: dibaca.rows,
        Pendek: dibaca.Pendek,
        Panjang: dibaca.Panjang,
        contoh: dibaca.contoh,
      };
    }

    // ==================== KETERANGAN JUMLAH ====================

    /*
    * Ubah teks angka dari keterangan paginator menjadi integer.
    *
    * Pemisah ribuan harus dibuang lebih dulu. Caption Insera memuatnya begitu
    * total melewati 999: "1,910 items found" berarti 1910, bukan 910. Kalau
    * pemisah tidak dibuang, regex (\d+) hanya mengambil digit terakhir sebelum
    * koma dan totalnya jadi 910 - jumlah halaman ikut jadi 91, dan scan berhenti
    * di sana tanpa satu pun halaman yang "gagal", jadi file tidak pernah
    * ditandai _SEBAGIAN padahal separuh data tidak ikut terambil.
    *
    * Titik hanya dibuang kalau diikuti tepat tiga digit, supaya angka desimal
    * tidak ikut berubah.
    */
    function angkaTeks(teks) {
      var s = String(teks === undefined || teks === null ? "" : teks)
        .replace(/[\s\u00a0\u202f\u2009]+/g, "")
        .replace(/[.,](?=\d{3}(?!\d))/g, "");
      var m = /-?\d+/.exec(s);
      return m ? parseInt(m[0], 10) : 0;
    }

    /*
    * Baca "displaying A to B" dari keterangan paginator.
    *
    * Teksnya dipotong dulu, baru tiap angka dibersihkan dengan angkaTeks.
    * Regex \d+ tidak bisa dipakai langsung: pada "displaying 991 to 1,000"
    * angka kedua hanya terbaca "1", dan pada "displaying 1,891 to 1,900"
    * polanya tidak cocok sama sekali - padahal kedua bentuk itu muncul begitu
    * offset melewati 999, yaitu mulai halaman ke-100.
    */
    function rentangDisplaying(teks) {
      var m = /displaying\s+([\s\S]*?)(?:[.]|$)/i.exec(String(teks || ""));
      if (!m) return null;
      var m2 = /(.*?)\s+to\s+(.*)/i.exec(m[1]);
      if (!m2) return null;
      var dari = angkaTeks(m2[1]);
      var sampai = angkaTeks(m2[2]);
      if (dari <= 0 || sampai < dari) return null;
      return { dari: dari, sampai: sampai };
    }

    /*
    * Teks di bawah tabel pada Insera:
    *   "175 items found, displaying 1 to 10."
    * Menghasilkan { totalItems, pageSize, aktif }.
    */
    function bacaKeterangan(dokumen) {
      if (!dokumen || !dokumen.querySelector) {
        return { totalItems: 0, pageSize: 0, teks: "", aktif: false };
      }

      var RE_ITEM = /items?\s+found/i;
      var teks = "";
      var terlihat = null;

      /*
      * Caption di bawah tabel pada Insera berbunyi:
      *   "175 items found, displaying 1 to 10."
      *
      * Halaman bisa memuat lebih dari satu caption, karena beberapa tab punya
      * tabel masing-masing. Hanya caption yang terlihat yang boleh dipakai.
      *
      * Versi lama memakai querySelector("span.pagebanner") langsung, yaitu
      * caption PERTAMA di urutan dokumen. Setelah pindah halaman, caption
      * pertama biasanya milik tabel yang sudah disembunyikan, sehingga
      * "displaying" tidak pernah naik dan proses menganggap tidak maju.
      */
      var pilih = function (el) {
        var isi = (el.textContent || "").trim();
        if (!isi || !RE_ITEM.test(isi)) return false;
        if (isi.length >= 200) return false;
        if (!teks) teks = isi;
        if (!terlihat && terlihatDiLayar(el)) terlihat = isi;
        return true;
      };

      var banners = dokumen.querySelectorAll("span.pagebanner");
      for (var i = 0; i < banners.length; i++) pilih(banners[i]);

      // Class pagebanner tidak ada: cari elemen mana pun yang polanya sama.
      if (!terlihat) {
        var semua = dokumen.querySelectorAll("span,div,td");
        for (var j = 0; j < semua.length; j++) pilih(semua[j]);
      }

      if (terlihat) teks = terlihat;

      /*
      * Token angka di sini boleh memuat pemisah ribuan: "1,910 items found",
      * "1.910 items found", dan "1 910 items found" (space yang tidak terputus)
      * semuanya berarti 1910. Pola [0-9.,]* saja tidak cukup karena pemisah
      *ruang harus ikut masuk ke token, kalau tidak whole caption ditolak dan
      *cuma digit terakhir sebelum space yang terbaca - 910.
      */
      var mTotal = /([0-9][0-9.,\s\u00a0\u202f]*)\s*items?\s+found/i.exec(teks);
      var totalItems = mTotal ? angkaTeks(mTotal[1]) : 0;

      var pageSize = 0;
      var rentang = rentangDisplaying(teks);
      if (rentang) pageSize = rentang.sampai - rentang.dari + 1;

      if (!pageSize) {
        var mPs = /ps=(\d+)/i.exec(location.search);
        if (mPs) pageSize = angkaTeks(mPs[1]);
      }

      return {
        totalItems: totalItems,
        pageSize: pageSize,
        teks: rapiNilai(teks),
        aktif: totalItems > 0,
      };
    }
    // ==================== DETEKSI NO TIKET ====================

    /*
    * Nomor tiket hanya dipakai untuk nama file. Sumbernya diurutkan dari yang
    * paling dapat dipercaya ke yang paling rapuh:
    *   1. query param URL          - paling pasti milik tiket ini
    *   2. judul tab browser        - biasanya "Tiket INCxxx - ..."
    *   3. field bertanda tiket     - "Ticket No: INCxxx" di halaman detail
    *   4. link Incident di tabel
    *   5. seluruh teks halaman     - terakhir, bisa salah kalau halamannya
    *                                  daftar tiket, bukan detail satu tiket
    *
    * Balikan: { nomor, sumber }. Sumber ditampilkan di panel supaya user bisa
    * melihat nomor tiket mana yang terpakai, dan mengoreksinya bila salah.
    */

    function bersihkanTiket(teks) {
      return String(teks || "")
        .toUpperCase()
        .replace(/[-_ ]/g, "");
    }

    /*
    * Baris ringkasan di atas tulisan "Impacted Service".
    *
    * Contoh isi baris itu di halaman Insera:
    *
    *   INC53545268 | [SQM GAMAS] | AKSES | DISTRIBUSI | TIF-2| REG-2 | JAKUT
    *   | TPR | [ODC-TPR-FBJ, ODC-TPR-FBL] | Datek ODP Terdampak : [...] | 4 Jam
    *   | IboosterDIST_15811790723803
    *
    * Yang dipakai INC53545268 saja.
    *
    * Ciri pembeda dari penyebutan INC lain di halaman: token INC diikuti tanda
    * "|" dalam jarak dekat. Halaman detail tiket sering punya nomor incident
    * lain di tabel, breadcrumb, atau panel Expanded, dan semuanya itu akan
    * salah diambil kalau hanya "batas kata INC" jadi patokan. Pipe-nya yang
    * membedakan baris ringkasan dari penyebutan biasa.
    *
    * Panjang jendela 160 karakter dipilih longgar: baris ringkasan punya banyak
    * field setelah INC, tapi baris penyebutan lain tidak pernah diikuti pipe
    * dalam jarak segitu. Teks IboosterDIST_15811790723803 yang muncul di
    * ujung baris ringkasan tidak punya token INC, jadi tidak akan tertangkap.
    */
    var RE_INC_TOKEN = /\bINC[-_ ]?[0-9]{5,}\b/gi;
    var JENDELA_PIPE = 160;

    /*
    * Cari nomor INC yang diikuti tanda pipe, di dalam satu potongan teks.
    *
    * Mengembalikan string kosong kalau tidak ada. Politeks dipotong ke 8.000
    * karakter supaya tidak memindai seluruh DOM setiap kali dibutuhkan.
    */
    function incDenganPipe(teks, batas) {
      var t = rapiNilai(String(teks || "").slice(0, batas || 8000));
      if (!t) return "";
      var re = new RegExp(RE_INC_TOKEN.source, "gi");
      var m;
      while ((m = re.exec(t)) !== null) {
        var mulaiEkor = m.index + m[0].length;
        var ekor = t.slice(mulaiEkor, mulaiEkor + JENDELA_PIPE);
        var posisiPipe = ekor.indexOf("|");
        if (posisiPipe < 0) continue;
        /*
        * Kalau ada INC lain sebelum pipe-nya, token ini bukan pemilik baris
        * ringkasan - dia cuma penyebutan yang kebetulan punya pipe milik token
        * berikutnya. Tanpa pemeriksaan ini, teks "INC11111111 tanpa pipe lalu
        * INC22222222 | ..." akan mengembalikan INC11111111, padahal yang
        * baris ringkasan miliknya adalah yang kedua.
        */
        var lain = /INC[-_ ]?[0-9]{5,}\b/i.exec(ekor.slice(0, posisiPipe));
        if (lain) continue;
        return bersihkanTiket(m[0]);
      }
      return "";
    }

    /*
    * Kandidat elemen yang letaknya DI ATAS sebuah judul.
    *
    * Alur: dari judul itu sendiri naik ke induknya, lalu di tiap tingkat ambil
    * saudara sebelumnya. Baris ringkasan bisa jadi saudara langsung dari judul,
    * atau anak dari kotak yang sama, atau satu blok di atasnya - karena itu
    * tiga tingkat keduanya diperiksa.
    *
    * previousElementSibling dipakai, bukan previousSibling: yang terakhir bisa
    * berupa node teks atau whitespace yang tidak punya textContent.
    */
    function kandidatDiAtas(el) {
      var hasil = [];
      var node = el;
      for (var tingkat = 0; tingkat < 4 && node && node.parentElement; tingkat++) {
        var saudara = node.previousElementSibling;
        var diambil = 0;
        while (saudara && diambil < 4) {
          hasil.push(saudara);
          saudara = saudara.previousElementSibling;
          diambil++;
        }
        node = node.parentElement;
      }
      return hasil;
    }

    /*
    * Nomor INC dari baris ringkasan di atas "Impacted Service".
    *
    * Dicari di sekitar judulnya dulu, karena itu lokasi yang Anda lihat. Kalau
    * judulnya tidak ketemu - misalnya markup-nya berubah - baris yang sama tetap
    * dicari di seluruh dokumen dengan syarat pipe yang sama.
    *
    * Balikan: { nomor, sumber } atau null kalau tidak ada yang cocok.
    */
    function cariIncDariBarisRingkasan() {
      var docs = semuaDokumen();
      var judul, kandidat, nomor, i, j;

      // 1. Di sekitar judul Impacted Service, di dokumen yang sama.
      for (i = 0; i < docs.length; i++) {
        judul = cariElemenRingkasan(docs[i].doc) ||
          (cariTabImpacted(docs[i].doc) || {}).el ||
          null;
        if (!judul) continue;
        kandidat = kandidatDiAtas(judul);
        for (j = 0; j < kandidat.length; j++) {
          nomor = incDenganPipe(kandidat[j].textContent);
          if (nomor) {
            return {
              nomor: nomor,
              sumber: "baris ringkasan di atas Impacted Service",
            };
          }
        }
      }

      // 2. Seluruh dokumen, tetap dengan syarat pipe.
      for (i = 0; i < docs.length; i++) {
        if (!docs[i].doc.body) continue;
        nomor = incDenganPipe(docs[i].doc.body.textContent);
        if (nomor) return { nomor: nomor, sumber: "baris ringkasan di halaman" };
      }

      return null;
    }

    // Seluruh nilai query param URL, termasuk yang nama kuncinya berawalan
    // "d-<componentId>-" ala ICEfaces. forEach tidak boleh dipakai di sini
    // karena `return` di dalam callback tidak keluar dari loop-nya.
    function nilaiQueryUrl(url) {
      var hasil = [];
      try {
        var u = new URL(url, location.href);
        u.searchParams.forEach(function (nilai, kunci) {
          hasil.push({ kunci: kunci, nilai: nilai });
        });
      } catch (e) {
        /* abaikan */
      }
      return hasil;
    }

    function deteksiNoTiket(ctx) {
      var doc = ctx ? ctx.doc : document;
      var dariUrl = null;
      var dariJudul = null;
      var m, i, u;

      /*
      * 1. Baris ringkasan di atas tulisan "Impacted Service".
      *
      * Dicek lebih dulu, bukan lebih akhir, karena sumber inilah yang paling
      * dekat dengan data yang sedang diunduh: kalau halaman menampilkan
      * ringkasan tiketnya sendiri, INC di situ adalah INC yang benar untuk
      * tabel Impacted Service yang sedang dibaca.
      *
      * URL dan judul tab tetap ikut dibandingkan, bukan diabaikan. Kalau
      * berbeda, log memberi tahu - diam-diam memakai angka yang berbeda dari
      * yang terlihat di browser akan membuat nama file menyesatkan. Keduanya
      * hanya dikumpulkan di sini, TIDAK langsung dikembalikan: kalau sumber
      * mana pun return duluan, baris ringkasan tidak akan pernah dibaca.
      */
      var dariRingkasan = cariIncDariBarisRingkasan();

      // 2. query param URL, dokumen atas + iframe
      var urlDicoba = [location.href];
      if (ctx && ctx.base) urlDicoba.push(ctx.base);
      for (u = 0; u < urlDicoba.length && !dariUrl; u++) {
        var params = nilaiQueryUrl(urlDicoba[u]);
        for (i = 0; i < params.length && !dariUrl; i++) {
          m = RE_TIKET.exec(params[i].nilai);
          if (m) dariUrl = { nomor: bersihkanTiket(m[0]), sumber: "URL" };
        }
      }

      // 3. judul tab browser
      m = RE_TIKET.exec(document.title || "");
      if (m) dariJudul = { nomor: bersihkanTiket(m[0]), sumber: "judul tab" };

      // Bandingkan setiap penyebutan yang ketemu dengan baris ringkasan.
      if (dariRingkasan) {
        var lain = [dariUrl, dariJudul];
        for (i = 0; i < lain.length; i++) {
          if (!lain[i] || lain[i].nomor === dariRingkasan.nomor) continue;
          log(
            "INC dari baris ringkasan (" + dariRingkasan.nomor +
              ") berbeda dari penyebutan lain di halaman (" + lain[i].nomor +
              ", dari " + lain[i].sumber + "). Yang dipakai: " +
              dariRingkasan.nomor + "."
          );
        }
        return dariRingkasan;
      }

      if (dariUrl) return dariUrl;
      if (dariJudul) return dariJudul;

      // 4. field bertanda tiket: "Ticket No: INCxxx", "Nomor Tiket", "Incident"
      var RE_LABEL_TIKET =
        /(ticket|tiket|incident|problem|issue|report)\s*(no\.?|nomor|number|num|code|id)?\s*[:#=-]?/i;
      try {
        var kandidat = doc.querySelectorAll(
          "th, td, label, dt, dd, legend, caption, b, strong, h1, h2, h3, h4"
        );
        for (var j = 0; j < kandidat.length; j++) {
          var teks = (kandidat[j].textContent || "").replace(/\s+/g, " ").trim();
          if (!teks || teks.length > 160) continue;
          var m2 = RE_TIKET.exec(teks);
          if (!m2) continue;
          // Hanya percaya kalau ada kata penanda tiket, supaya tidak diambil
          // dari sel biasa yang kebetulan memuat nomor incident lain.
          if (RE_LABEL_TIKET.test(teks)) {
            return { nomor: bersihkanTiket(m2[0]), sumber: "field tiket" };
          }
        }
      } catch (e) {
        /* abaikan */
      }

      // 5. link Incident di area tabel / breadcrumb
      if (ctx) {
        try {
          var links = ctx.tabel.querySelectorAll("a");
          for (var k = 0; k < links.length; k++) {
            var t2 = (links[k].innerText || links[k].textContent || "").trim();
            var m3 = RE_TIKET.exec(t2);
            if (m3) return { nomor: bersihkanTiket(m3[0]), sumber: "link tabel" };
          }
        } catch (e) {
          /* abaikan */
        }
      }

      // 6. seluruh teks halaman - sumber paling rapuh, dipakai terakhir
      try {
        if (doc.body) {
          var m4 = RE_TIKET.exec(rapiNilai(doc.body.textContent).slice(0, 6000));
          if (m4) return { nomor: bersihkanTiket(m4[0]), sumber: "teks halaman" };
        }
      } catch (e) {
        /* abaikan */
      }

      return { nomor: "", sumber: "" };
    }

    function deteksiNoTiketAman() {
      try {
        return deteksiNoTiket(konteksCache);
      } catch (e) {
        return { nomor: "", sumber: "" };
      }
    }

    // ==================== PAGINATION: PANEN LINK ====================

    /*
    * Cara paling aman: bukan menebak URL, tapi memetik href yang benar-benar
    * ada di DOM. Jadi tidak perlu tahu component id ICEfaces Insera.
    * Hasil: peta { nomorHalaman: urlLengkap }
    */
    function panenLinkHalaman(ctx) {
      var peta = {};
      var sumber = ctx.doc.querySelectorAll(
        'a[title*="Go to page"], .pagelinks a, a[title*="page"]'
      );
      for (var i = 0; i < sumber.length; i++) {
        var a = sumber[i];
        var judul = a.getAttribute("title") || "";
        var href = a.getAttribute("href") || "";
        if (!href || href === "#" || /^javascript:/i.test(href)) continue;

        var n = 0;
        var m = /Go to page (\d+)/i.exec(judul);
        if (m) n = parseInt(m[1], 10);
        if (!n) {
          var t = (a.innerText || a.textContent || "").trim();
          if (/^\d+$/.test(t)) n = parseInt(t, 10);
        }
        if (!n) {
          var m2 = /[?&;](?:d-\d+-)?p=(\d+)/i.exec(href);
          if (m2) n = parseInt(m2[1], 10);
        }
        if (!n || n < 2) continue;
        if (!peta[n]) {
          // Satu link yang tidak bisa dibaca tidak boleh menghentikan pengambilan
          // ratusan baris, jadi dilewati saja.
          var penuh = hrefTAnchor(a, ctx.base);
          if (penuh) peta[n] = penuh;
        }
      }
      return peta;
    }

    // Link "next" untuk kasus manual, dipakai sebagai fallback terakhir.
    function linkBerikutnya(ctx) {
      var kandidat = ctx.doc.querySelectorAll("a");
      for (var i = 0; i < kandidat.length; i++) {
        var a = kandidat[i];
        var label =
          (a.getAttribute("title") || "") + " " + (a.innerText || a.textContent || "");
        if (!/\bnext\b|berikutnya|selanjut/i.test(label)) continue;
        var href = a.getAttribute("href") || "";
        if (href && href !== "#" && !/^javascript:/i.test(href)) {
          return hrefTAnchor(a, ctx.base);
        }
      }
      return null;
    }

    /*
    * Kembalikan tabel ke halaman 1 setelah mode klik selesai, supaya halaman
    * Insera milik user tidak ditinggalkan di halaman terakhir.
    * Dicoba lewat link "first page"/"<<", lalu reload halaman sebagai cadangan.
    */
    async function kembaliKeHalaman1(ctx, pageSize, panjangHarapan) {
      try {
        if (!ctx || !ctx.tabel) return;
        var sebelum = sidikHalaman(ctx, panjangHarapan);
        if (sebelum.dari <= 1) return; // sudah di halaman 1

        var selector =
          ".pagelinks a[title*='First'], .pagelinks a[title*='first']," +
          ".pagelinks a:first-child, a[title*='Go to first']";
        var awal = null;
        try {
          awal = ctx.doc.querySelector(selector);
        } catch (e) {
          awal = null;
        }
        if (!awal) return; // tidak ada tombol ke halaman 1, biarkan saja
        klikBerikut(awal);
        await tungguRender(ctx, sebelum, panjangHarapan, 8000);
        log("Tabel dikembalikan ke halaman 1.");
      } catch (e) {
        log("Gagal mengembalikan tabel ke halaman 1: " + e);
      }
    }

    // ==================== PAGINATION: JALUR CADANGAN (KLIK DOM) ====================

    /*
    * Jalur cepat di atas mengira paginasi bisa lewat URL. Untuk ICEfaces/ADF
    * itu sering salah: halaman berikutnya bukan GET bebas, tapi submit form
    * dengan view state. Jalur ini tidak menebak apa pun - ia mengklik tombol
    * "next" sungguhan di tabel, lalu membaca hasilnya dari DOM, persis seperti
    * yang dilakukan user.
    *
    * Konsekuensinya satu halaman per waktu (tidak bisa paralel), jadi
    * 500 halaman bisa memakan waktu 5-15 menit. Status dan Batalkan tetap jalan.
    */

    // Ambil nama class sebagai teks, aman untuk elemen SVG yang className-nya
    // bukan string.
    function kelasTeks(el) {
      try {
        if (!el.className) return "";
        if (typeof el.className === "string") return el.className;
        return el.getAttribute("class") || "";
      } catch (e) {
        return "";
      }
    }

    // Gabungan teks, title, class, dan atribut lain sebagai "label" kontrol.
    function labelKontrol(el) {
      var teks = ((el.innerText || el.textContent || "") + "").trim();
      var alt = "";
      try {
        var gambar = el.querySelector("img[alt]");
        if (gambar) alt = gambar.getAttribute("alt") || "";
      } catch (e) {
        /* abaikan */
      }
      return (
        teks + " " +
        (el.getAttribute("title") || "") + " " +
        (el.getAttribute("rel") || "") + " " +
        (el.getAttribute("aria-label") || "") + " " +
        (el.getAttribute("name") || "") + " " +
        (el.getAttribute("value") || "") + " " +
        alt + " " +
        kelasTeks(el) + " " +
        (el.id || "")
      ).toLowerCase();
    }

    /*
    * Petunjuk bahwa sebuah kontrol memang untuk navigasi halaman. Dipakai
    * untuk menyaring fallback yang terlalu longgar.
    */
    function adakahPetunjukNavigasi(el) {
      var RE_NAV = new RegExp(
        "(\\b(next|prev|previous|first|last|page|pages|halaman|berikutnya|selanjut|sebelum|akhir|awal|[0-9])\\b)" +
        "|[\\u00ab\\u00bb\\u2039\\u203a\\u2192\\u27a1]" +
        "|[<>=]{1,3}\\s*$",
        "i"
      );
      return RE_NAV.test(labelKontrol(el));
    }

    /*
    * Nilai sebuah kontrol: makin tinggi, makin cocok sebagai "halaman
    * berikutnya".
    *
    * Ini diperlukan karena beberapa paginator Insera menaruh tombol Last atau
    * nomor halaman di posisi terakhir. Versi lama memakai .pagelinks
    * a:last-child apa adanya, sehingga tombol Last ikut terambil dan diklik
    * terus-menerus: halaman tidak pernah maju, lalu proses dianggap macet.
    *
    * Nilai 0 berarti jangan dipakai sama sekali: nomor halaman, First, dan
    * Previous, karena ketiganya hanya bisa membuat halaman berjalan mundur.
    */
    function nilaiKandidatBerikut(el) {
      var teks = ((el.innerText || el.textContent || "") + "").trim();
      var label = labelKontrol(el);
      var panah = new RegExp("[\\u00bb\\u203a\\u2192\\u27a1]", "i");
      var panahUjung = new RegExp("[\\u00bb\\u203a\\u2192\\u27a1]\\s*$", "i");
      var mundur = new RegExp("\\bprev(ious)?\\b|sebelum|kembali|<{1,3}", "i");
      var keAwal = new RegExp("\\bfirst\\b|awal|halaman\\s+pertama", "i");
      var keAkhir = new RegExp("\\blast\\b|akhir|halaman\\s+terakhir", "i");
      var maju = new RegExp("\\b(next|go\\s+to\\s+next)\\b|berikutnya|selanjut|halaman\\s+berikut", "i");
      var gtUjung = new RegExp("[>]{2,}\\s*$", "i");
      var gtTunggal = new RegExp("^\\s*>\\s*$", "i");
      var RE_ANGKA = /^\d+$/;

      if (RE_ANGKA.test(teks)) return 0;
      if (mundur.test(label)) return 0;
      if (keAwal.test(label)) return 0;
      if (maju.test(label)) return 100;
      if (panahUjung.test(teks)) return 90;
      if (panah.test(label)) return 85;
      if (gtUjung.test(teks)) return 80;
      if (gtTunggal.test(teks)) return 75;
      //
      // Last melompat ke halaman paling akhir. Jaring pengaman terakhir: masih
      // berguna, tapi tidak pernah boleh dipilih sebelum tombol next.
      if (keAkhir.test(label)) return 8;

      return 50;
    }

    // Kandidat elemen "halaman berikutnya", diurutkan dari yang paling cocok.
    function kandidatBerikut(ctx) {
      var sel = [
        ".pagelinks a.next",
        ".pagelinks a[title*='Next']",
        ".pagelinks a[title*='next']",
        "a[title*='Next Page']",
        "a[title*='Go to next']",
        "a[title*='next page']",
        ".pagelinks a:last-child",
        ".pagination a",
        ".pagelinks a",
      ];
      var dok = ctx.doc;
      var found = [];
      var sudah = [];
      var i, j;

      var tambah = function (el, bonus, longgar) {
        if (!el || sudah.indexOf(el) !== -1) return;
        var kelas = kelasTeks(el);
        var mati =
          el.hasAttribute("disabled") ||
          el.getAttribute("aria-disabled") === "true" ||
          kelas.indexOf("disable") >= 0 ||
          kelas.indexOf("inactive") >= 0;
        if (mati) return;
        //
        // Kandidat dari fallback longgar harus punya petunjuk navigasi,
        // supaya link biasa di halaman tidak ikut dianggap tombol halaman
        // berikutnya.
        if (longgar && !adakahPetunjukNavigasi(el)) return;
        // Nilai dasar 0 berarti elemen ini tidak boleh dipakai sama sekali
        // (nomor halaman, First, Previous). Bonus selector tidak boleh
        // mengangkatnya, kalau tidak filter di bawah tidak pernah bekerja.
        var dasar = nilaiKandidatBerikut(el);
        if (dasar <= 0) return;
        sudah.push(el);
        found.push({ el: el, nilai: dasar + bonus });
      };

      var list;
      var bonus;
      for (i = 0; i < sel.length; i++) {
        try {
          list = dok.querySelectorAll(sel[i]);
          // Selector yang lebih spesifik diberi bonus kecil supaya menang dari
          // tebakan teks yang hanya sama kuatnya.
          bonus = (sel.length - i) * 2;
          for (j = 0; j < list.length; j++) tambah(list[j], bonus, false);
        } catch (e) {
          /* selector tidak berlaku, lanjut */
        }
      }

      /*
      * Fallback terakhir: link di dalam wadah pagination saja.
      *
      * Tanpa batas ini, querySelectorAll("a") menarik seluruh tautan di
      * halaman termasuk link tab. Mengklik link tab tidak memindahkan
      * halaman, jadi proses membuang dua kali percobaan lalu salah melapor
      * macet padahal tombol next-nya ada.
      */
      try {
        var wadah = dok.querySelectorAll(
          ".pagelinks, .pagination, .paginate, .pager, " +
          "[class*='paginat'], [id*='paginat'], [class*='pagelink'], [id*='pagelink']"
        );
        for (i = 0; i < wadah.length; i++) {
          var isi = wadah[i].querySelectorAll("a");
          for (j = 0; j < isi.length; j++) tambah(isi[j], 0, true);
        }
      } catch (e) {
        /* abaikan */
      }

      found.sort(function (a, b) { return b.nilai - a.nilai; });

      var hasil = [];
      for (i = 0; i < found.length; i++) {
        // Nilai 0: nomor halaman, First, Previous. Jangan dipakai.
        if (found[i].nilai <= 0) continue;
        hasil.push(found[i].el);
      }
      return hasil;
    }

    /*
    * Klik tombol berikutnya.
    *
    * Kalau button-nya dibungkus <a> atau <input type=submit>, yang diklik adalah
    * bagian dalam yang benar-benar memicu navigasi.
    *
    * Hanya SATU event click yang dikirim. Versi lama mengirim click di elemen
    * anak lalu el.click() lagi di elemen induknya, yang membuat Icefaces
    * menjalankan navigasi dua kali dan sempat melompati satu halaman.
    */
    function klikBerikut(el) {
      var target = el.querySelector("input[type='submit'], span") || el;
      var win = (el.ownerDocument && el.ownerDocument.defaultView) || window;
      try {
        var urutan = ["mousedown", "mouseup", "click"];
        for (var i = 0; i < urutan.length; i++) {
          target.dispatchEvent(
            new win.MouseEvent(urutan[i], { bubbles: true, cancelable: true, view: win })
          );
        }
        return true;
      } catch (e) {
        try {
          target.click();
          return true;
        } catch (e2) {
          return false;
        }
      }
    }

    // Tanda sidik halaman sekarang: keterangan displaying + isi baris pertama.
    function sidikHalaman(ctx, panjangHarapan) {
      var ket = bacaKeterangan(ctx.doc);
      var rentang = rentangDisplaying(ket.teks);
      var headTr = ctx.tabel ? barisHeader(ctx.tabel) : null;
      var baris = ctx.tabel
        ? bacaBaris(ctx.tabel, indeksTerlihat(headTr), panjangHarapan).rows
        : [];
      var pertama = baris.length ? String(baris[0][1] || "") : "";
      return {
        dari: rentang ? rentang.dari : -1,
        sampai: rentang ? rentang.sampai : -1,
        jumlahBaris: baris.length,
        pertama: pertama,
        teks: ket.teks,
      };
    }

    /*
    * Setelah halaman berikutnya dimuat, widget tidak selalu mengisi ulang tabel
    * yang sama. Kadang elemen <table>-nya diganti seluruhnya, sehingga ctx.tabel
    * menunjuk ke elemen yang sudah lepas dari dokumen dan sidik halaman tidak
    * akan pernah berubah.
    *
    * Karena itu tabel dicari ulang setiap kali halaman ditunggu. Konteks diupdate
    * di tempat supaya pemanggil ikut memakai tabel yang baru.
    */
    function segarkanKonteksTabel(ctx) {
      if (!ctx) return false;

      /*
      * Tabel yang sedang dipakai boleh dilewati kalau masih terpasang DAN masih
      * terlihat.
      *
      * Pemeriksaan ini dipanggil tiap 250 ms selama menunggu render, sementara
      * cariKonteksTabel() memindai seluruh tabel di semua dokumen. Pada widget
      * yang mengisi ulang tabel yang sama - kasus paling umum - hasilnya sudah
      * pasti sama, jadi pemindaian itu sia-sia.
      *
      * Syarat "terlihat" tidak boleh dihilangkan. Banyak paginator menyiapkan
      * tabel untuk semua halaman sekaligus lalu hanya menyembunyikan yang tidak
      * aktif. Di sana tabel yang benar BERGANTI setiap halaman, sementara tabel
      * yang lama tetap terpasang di dokumen.
      */
      if (ctx.tabel && ctx.tabel.isConnected) {
        try {
          if (terlihatDiLayar(ctx.tabel)) return false;
        } catch (e) {
          /* kalau tidak bisa dipastikan, cari ulang */
        }
      }

      try {
        var baru = cariKonteksTabel();
        if (baru && baru.identity && baru.tabel !== ctx.tabel) {
          ctx.tabel = baru.tabel;
          ctx.doc = baru.doc;
          ctx.base = baru.base;
          return true;
        }
      } catch (e) {
        /* abaikan, pakai tabel lama */
      }
      return false;
    }

    // Tunggu sampai keterangan atau baris pertama berubah (ajax selesai render).
    function tungguRender(ctx, sebelum, panjangHarapan, batasMs) {
      return new Promise(function (selesai) {
        var batas = batasMs || 20000;
        var mulai = Date.now();
        var cek = function () {
          // Widget bisa mengganti elemen tabel, jadi cari ulang tiap polling.
          segarkanKonteksTabel(ctx);
          var s = sidikHalaman(ctx, panjangHarapan);
          var berubah =
            s.dari !== sebelum.dari ||
            s.sampai !== sebelum.sampai ||
            s.pertama !== sebelum.pertama;
          if (berubah) return selesai({ ok: true, sekarang: s });
          if (sedangBatal) return selesai({ ok: false, alasan: "dibatalkan" });
          if (Date.now() - mulai > batas) {
            return selesai({ ok: false, alasan: "tidak berubah dalam " + Math.round(batas / 1000) + " detik" });
          }
          setTimeout(cek, 250);
        };
        cek();
      });
    }

    /*
    * Rangkai seluruh halaman dengan mengklik tombol next. Dipakai otomatis
    * kalau jalur URL tidak bisa diverifikasi.
    *
    * Dua kondisi selesai yang berbeda:
    *  - "sudahMentok": keterangan menunjukkan semua item sudah terbaca.
    *  - "tombolHabis": tombol next tidak ada atau nonaktif, padahal baris yang
    *    terkumpul masih kurang dari total. Hanya yang terakhir ini yang gagal.
    *
    * Versi lama menganggap keduanya sebagai kegagalan, sehingga tabel yang hanya
    * punya satu halaman selalu dilaporkan macet.
    *
    * Balikan: { ok, selesaiMentok, halamanDiproses, macetDi, alasan, semua }
    */
    async function rangkaiLewatKlik(ctx, header, totalItems, onProgres) {
      var panjangHarapan = header.length;
      var petaBaris = [];
      var macetDi = 0;
      var selesaiMentok = false;
      var alasan = "";

      var hitungBaris = function () {
        var n = 0;
        for (var i = 0; i < petaBaris.length; i++) {
          n += (petaBaris[i].rows || []).length;
        }
        return n;
      };

      /*
      * Posisi baris untuk ditampilkan ke panel.
      *
      * "displaying A to B" milik server adalah sumber yang paling akurat. Kalau
      * keterangannya tidak terbaca, posisi dihitung dari baris yang sudah
      * terkumpul supaya panel tetap menampilkan angka yang masuk akal.
      *
      * Penting: yang dikembalikan adalah NOMOR BARIS, bukan nomor halaman.
      * Sebelumnya onProgres mengirim nomor halaman sementara pemanggil membacanya
      * sebagai nomor baris, jadi panel menampilkan "baris 7 dari 175" untuk
      * halaman 7 yang sebenarnya baris 61-70, dan progress bar mentok di ~10%.
      */
      var posisiBaris = function (sidik, terkumpul, jumlahBaris) {
        var dari = sidik && sidik.dari > 0 ? sidik.dari : terkumpul + 1;
        var sampai = sidik && sidik.sampai > 0 ? sidik.sampai : dari + jumlahBaris - 1;
        return { dari: dari, sampai: sampai };
      };

      // Halaman 1 sudah ada di DOM.
      var awal = sidikHalaman(ctx, panjangHarapan);
      var barisSekarang = bacaBaris(ctx.tabel, indeksTerlihat(barisHeader(ctx.tabel)), panjangHarapan);
      petaBaris.push(barisSekarang);
      if (onProgres) onProgres(posisiBaris(awal, 0, barisSekarang.rows.length), totalItems, barisSekarang.rows.length);

      // Kalau total item sudah habis di halaman pertama, tidak ada yang perlu diklik.
      if (totalItems > 0 && awal.sampai >= totalItems) selesaiMentok = true;

      var sebelum = awal;
      var halamanSekarang = 1;

      /*
      * Kandidat yang sudah dicoba pada halaman ini, dihitung per penanda.
      * Kalau tombol yang diklik tidak memindahkan halaman, kandidat lain
      * dicoba sebelum menyerah. Tanpa ini satu tombol yang salah bisa membuat
      * proses berhenti padahal masih ada tombol yang benar.
      */
      var percobaan = {};
      var MAKS_PERCOBAAN = 2;
      var penandaKandidat = function (el) {
        try {
          return (
            (el.getAttribute("href") || "") + "|" +
            ((el.textContent || "") + "").trim().toLowerCase() + "|" +
            kelasTeks(el)
          );
        } catch (e) {
          return "";
        }
      };

      while (!sedangBatal && !selesaiMentok) {
        var terkumpul = hitungBaris();

        // Berhenti saat keterangan sudah menunjukkan baris terakhir.
        if (sebelum.sampai > 0 && totalItems > 0 && sebelum.sampai >= totalItems) {
          selesaiMentok = true;
          break;
        }

        var semuaKandidat = kandidatBerikut(ctx);
        var kandidat = [];
        for (var ki = 0; ki < semuaKandidat.length; ki++) {
          var tanda = penandaKandidat(semuaKandidat[ki]);
          if ((percobaan[tanda] || 0) >= MAKS_PERCOBAAN) continue;
          kandidat.push({ el: semuaKandidat[ki], tanda: tanda });
        }

        if (!kandidat.length) {
          // Tidak ada tombol next yang belum dicoba. Kalau datanya sudah cukup,
          // ini normal untuk tabel satu halaman. Kalau kurang, baru masalah.
          if (totalItems <= 0 || terkumpul >= totalItems) {
            selesaiMentok = true;
            break;
          }
          alasan =
            "tombol halaman berikutnya tidak ditemukan setelah " +
            terkumpul + " dari " + totalItems + " baris";
          macetDi = halamanSekarang + 1;
          break;
        }

        var dipilih = kandidat[0];
        percobaan[dipilih.tanda] = (percobaan[dipilih.tanda] || 0) + 1;
        klikBerikut(dipilih.el);

        var hasil = await tungguRender(ctx, sebelum, panjangHarapan);
        if (!hasil.ok) {
          if (hasil.alasan === "dibatalkan") {
            return {
              ok: false,
              dibatalkan: true,
              selesaiMentok: false,
              halamanDiproses: petaBaris.length,
              macetDi: halamanSekarang + 1,
              alasan: "dibatalkan",
              semua: petaBaris,
            };
          }
          // Render tidak berubah. Coba kandidat lain pada iterasi berikutnya,
          // jangan langsung menyerah selama masih ada sisa.
          alasan = hasil.alasan;
          continue;
        }

        // Render selesai, tapi pastikan benar-benar maju. Tanpa cek ini, tombol
        // First atau Previous akan membuat proses berjalan mundur terus.
        var maju = true;
        if (hasil.sekarang.sampai > 0 && sebelum.sampai > 0) {
          maju = hasil.sekarang.sampai > sebelum.sampai;
        }
        if (!maju) {
          alasan =
            "tombol '" + ((dipilih.el.textContent || "").trim() || "tanpa teks") +
            "' tidak memindahkan halaman ke depan";
          continue;
        }

        // Berhasil maju: hitungan percobaan direset untuk halaman berikutnya.
        percobaan = {};
        halamanSekarang++;
        var baru = bacaBaris(ctx.tabel, indeksTerlihat(barisHeader(ctx.tabel)), panjangHarapan);
        petaBaris.push(baru);
        sebelum = hasil.sekarang;
        if (onProgres) {
          onProgres(
            posisiBaris(sebelum, hitungBaris() - baru.rows.length, baru.rows.length),
            totalItems,
            baru.rows.length
          );
        }
      }

      return {
        ok: selesaiMentok && !macetDi,
        selesaiMentok: selesaiMentok,
        halamanDiproses: petaBaris.length,
        macetDi: macetDi,
        /*
        * Baris terakhir yang benar-benar terbaca. Ini yang dicari user saat
        * proses macet: "sampai baris berapa yang terbaca?" Previously the only
        * clue was a page number, which does not translate to a row position.
        */
        barisTercapai: sebelum && sebelum.sampai > 0 ? sebelum.sampai : 0,
        jumlahBarisTerbaca: hitungBaris(),
        alasan: alasan,
        semua: petaBaris,
      };
    }

    // Prefix komponen ICEfaces terdeteksi otomatis dari URL, jadi tidak perlu
    // menulis "d-5564009-" secara manual.
    function prefixIced(base) {
      var m = /d-\d+-/.exec(base || "");
      return m ? m[0] : "";
    }

    // Set query param, dengan menyunting prefix ICEfaces bila ada.
    function setParam(url, key, val) {
      try {
        var u = new URL(url, location.href);
        var pre = prefixIced(u.search);
        var buang = [];
        u.searchParams.forEach(function (v, k) {
          var dasar = k.indexOf(":") >= 0 ? k.slice(k.lastIndexOf(":") + 1) : k;
          if (dasar === key) buang.push(k);
        });
        buang.forEach(function (k) {
          u.searchParams.delete(k);
        });
        u.searchParams.set(pre + key, String(val));
        return u.href;
      } catch (e) {
        // URL-nya tidak terbaca. Kembalikan apa adanya supaya pemanggil tetap
        // bisa mencoba, dan tidak ikut gagal.
        return url;
      }
    }

    function urlHalaman(base, n, ps) {
      var url = ps ? setParam(base, "ps", ps) : base;
      return setParam(url, "p", n);
    }

    /*
    * Apakah URL ini sudah punya query param dengan nama dasar "key"?
    *
    * Pemakaiannya berlawanan dengan setParam: param sudah ada di href hasil
    * panen, dan menimpanya justru merusak. Param paginasi ADF/JSF bisa
    * berprefiks (mis. "d-1234-p") dan hanya bentuk berprefiks itu yang dibaca
    * server, sehingga "ps" polos hasil setParam akan diabaikan begitu saja.
    */
    function punyaParam(url, key) {
      try {
        var ketemu = false;
        new URL(url, location.href).searchParams.forEach(function (v, k) {
          var dasar = k.indexOf(":") >= 0 ? k.slice(k.lastIndexOf(":") + 1) : k;
          if (dasar === key) ketemu = true;
        });
        return ketemu;
      } catch (e) {
        return false;
      }
    }

    /*
    * Paginasi ADF/ICEfaces tidak selalu bisa dibangun dari base dokumen. Param
    * "p" hanya dipahami kalau ada param state milik framework: prefix
    * "d-1234-", loop state, view state. Base dokumen tidak punya satu pun itu,
    * sehingga permintaan dengan "p=9" diabaikan diam-diam dan server mengirim
    * ulang halaman 1.
    *
    * Karena itu halaman yang tidak punya link di DOM dibangun dari link yang
    * sudah terbukti jalan: hanya nomor halaman yang diganti, semua param lain
    * ikut terbawa.
    */
    function nomorTemplate(petaLink, target) {
      var nomor = Object.keys(petaLink);
      if (!nomor.length) return 0;
      var terbaik = parseInt(nomor[0], 10);
      var jarak = Math.abs(terbaik - target);
      for (var i = 1; i < nomor.length; i++) {
        var n = parseInt(nomor[i], 10);
        var d = Math.abs(n - target);
        if (d < jarak) {
          jarak = d;
          terbaik = n;
        }
      }
      return terbaik;
    }

    function urlDariTemplate(petaLink, target, base, pageSize) {
      var t = nomorTemplate(petaLink, target);
      if (t) {
        // Link yang dipakai jadi template: pakai setParam supaya bentuk
        // berprefiks param "p" ikut terjaga.
        return setParam(petaLink[t], "p", target);
      }
      return urlHalaman(base, target, pageSize);
    }

    // ==================== VERIFIKASI PEMINDAHAN HALAMAN ====================

    /*
    * Risiko terbesar: URL halaman N ternyata tidak berpindah halaman, lalu
    * data yang sama terambil hundreds kali dan jumlah baris jauh kurang dari
    * total. Karena itu SETIAP halaman dicek terhadap rentang baris yang
    * seharusnya, bukan cuma halaman 2.
    */
    function verifikasiHalaman(html, base, hopeDari, hopeSampai) {
      var ket = bacaKeterangan(new DOMParser().parseFromString(html, "text/html"));
      if (!ket.aktif) return { ok: false, alasan: "keterangan jumlah tidak terbaca" };
      var rentang = rentangDisplaying(ket.teks);
      if (!rentang) return { ok: false, alasan: "pola displaying tidak terbaca" };
      if (rentang.dari === hopeDari && rentang.sampai === hopeSampai) {
        return { ok: true, ket: ket };
      }
      return {
        ok: false,
        alasan:
          "berisi baris " + rentang.dari + "-" + rentang.sampai +
          ", seharusnya " + hopeDari + "-" + hopeSampai,
      };
    }

    // Rentang baris yang seharusnya berisi halaman ke-n.
    function rentangHalaman(nomor, pageSize, totalItems) {
      var dari = (nomor - 1) * pageSize + 1;
      var sampai = dari + pageSize - 1;
      // Halaman terakhir memang tidak penuh.
      if (totalItems > 0 && sampai > totalItems) sampai = totalItems;
      return { dari: dari, sampai: sampai };
    }

    /*
    * Ambil satu halaman lalu pastikan isinya benar-benar halaman tersebut.
    * Kalau HTTP gagal ATAU keterangan tidak cocok, coba lagi sampai 3 kali
    * (percobaan awal + MAKS_RETRY). Jeda dibuat memanjang tiap percobaan.
    *
    * Baris diekstrak di dalam fungsi ini, bukan dikembalikan sebagai HTML.
    * HTML satu halaman Insera adalah seluruh halaman - tab, script, header - jadi
    * ratusan KB sampai megabyte, bukan cuma tabel 10 barisnya. Kalau HTML
    * disimpan sampai semua halaman selesai, 1.000 halaman menahan ratusan MB
    * dan tab akan crash sebelum scan selesai. Yang dikembalikan hanya barisnya.
    */
    async function ambilHalamanTerverifikasi(
      url, nomor, pageSize, totalItems, base, onInfo, panjangHarapan
    ) {
      var r = rentangHalaman(nomor, pageSize, totalItems);
      var alasan = "";

      for (var percobaan = 1; percobaan <= MAKS_RETRY + 1; percobaan++) {
        if (sedangBatal) return { ok: false, alasan: "dibatalkan" };

        if (percobaan > 1) {
          var info = "Mengulang halaman " + nomor + " (percobaan " + percobaan + ")";
          log(info);
          if (onInfo) onInfo(info + "...");
          await jeda(400 * percobaan);
        }

        var res = await fetchHalaman(url);
        if (!res.ok) {
          alasan = "gagal diakses: " + res.err;
          continue;
        }
        var v = verifikasiHalaman(res.html, base, r.dari, r.sampai);
        if (v.ok) {
          var baca = ekstrakDariHTML(res.html, base, panjangHarapan);
          // res.html tidak lagi dirujuk di sini, jadi bisa dibuang sekarang.
          return {
            ok: true,
            rows: baca.rows,
            Pendek: baca.Pendek,
            Panjang: baca.Panjang,
            contoh: baca.contoh,
            percobaan: percobaan,
          };
        }
        alasan = v.alasan;
      }
      return { ok: false, alasan: alasan };
    }

    /*
    * Ambil banyak halaman sekaligus (worker paralel). Tiap halaman diverifikasi,
    * diekstrak, dan diulang sendiri oleh ambilHalamanTerverifikasi.
    *
    * mulaiHalaman adalah nomor halaman nyata milik urls[0]. Pemanggil sudah
    * membuang halaman 2 sebelum memanggil fungsi ini, jadi nomor halaman TIDAK
    * boleh lagi dihitung dari indeks array. Kalau dihitung dari indeks, halaman
    * 3 akan diperiksa terhadap rentang baris halaman 2 dan isinya yang benar
    * ikut dibuang.
    *
    * Balikan: { slot: [{ok, rows, nomor}] , gagal: [nomor halaman] }
    */
    async function ambilBanyakTerverifikasi(
      urls, pageSize, totalItems, base, onProgres, mulaiHalaman, panjangHarapan
    ) {
      var total = urls.length;
      var slot = new Array(total);
      var gagal = [];
      var berikut = 0;
      var aktif = 0;
      var selesai = 0;

      async function pekerja() {
        while (berikut < total && !sedangBatal) {
          var i = berikut++;
          aktif++;
          var nomor = mulaiHalaman + i;
          var hasil = await ambilHalamanTerverifikasi(
            urls[i], nomor, pageSize, totalItems, base,
            function (t) { if (onProgres) onProgres(selesai, total, t); },
            panjangHarapan
          );
          // Nomor disimpan di dalam slot, supaya pemanggil tidak perlu menebaknya
          // lagi dari indeks saat membaca hasil atau mencatat halaman gagal.
          hasil.nomor = nomor;
          slot[i] = hasil;
          if (!hasil.ok) {
            gagal.push(nomor);
            log("Halaman " + nomor + " gagal setelah " + (MAKS_RETRY + 1) + "x: " + hasil.alasan);
          }
          aktif--;
          selesai++;
          if (onProgres) onProgres(selesai, total, "");
        }
      }

      var pekerjaSemua = [];
      for (var w = 0; w < Math.min(koncAktif, total); w++) {
        pekerjaSemua.push(pekerja());
      }
      await Promise.all(pekerjaSemua);

      return { slot: slot, gagal: gagal };
    }

    // ==================== UPSAFE PAGE SIZE ====================

    /*
    * Halaman Impacted Services hanya 10 baris per halaman. Kalau server mendukung
    * parameter page size, menaikkan ke 100 membuat 500 halaman jadi 5 halaman.
    * Dicoba satu per satu, hanya dipakai bila keterangan benar-benar berubah.
    */
    async function cobaPageSizeNaik(base, pageSizeAwal, onInfo) {
      if (!pageSizeAwal) return { pageSize: pageSizeAwal, dipakai: false };
      for (var i = 0; i < PS_NAIK.length; i++) {
        var ps = PS_NAIK[i];
        if (ps <= pageSizeAwal) continue;
        if (sedangBatal) break;
        var res = await fetchHalaman(setParam(base, "ps", ps));
        if (!res.ok) continue;
        var ket = bacaKeterangan(new DOMParser().parseFromString(res.html, "text/html"));
        if (ket.pageSize >= ps && ket.totalItems > 0) {
          var info =
            "Page size dinaikkan " + pageSizeAwal + " menjadi " + ket.pageSize +
            " per halaman.";
          onInfo(info);
          return { pageSize: ket.pageSize, dipakai: true };
        }
      }
      return { pageSize: pageSizeAwal, dipakai: false };
    }

    // ==================== FETCH PARALEL ====================

    function fetchHalaman(url) {
      var full = new URL(url, location.href).href;
      return new Promise(function (resolve) {
        function ok(html) {
          resolve({ ok: true, url: full, html: html });
        }
        function gagal(err) {
          resolve({ ok: false, url: full, err: err });
        }
        if (typeof GM_xmlhttpRequest !== "undefined") {
          GM_xmlhttpRequest({
            method: "GET",
            url: full,
            onload: function (r) {
              if (r.status >= 200 && r.status < 300) ok(r.responseText);
              else gagal("HTTP " + r.status);
            },
            onerror: function (r) {
              gagal("onerror: " + (r && r.error));
            },
            ontimeout: function () {
              gagal("timeout");
            },
            timeout: TIMEOUT_MS,
          });
        } else {
          fetch(full, { credentials: "same-origin" })
            .then(function (r) {
              return r.text();
            })
            .then(ok)
            .catch(function (e) {
              gagal(String(e));
            });
        }
      });
    }

    // ==================== UI ====================

    var sedangBatal = false;
    var sedangJalan = false;
    var konteksCache = null;
    var ui = {};
    // Nomor tiket hasil deteksi otomatis, dipakai untuk nama file & sheet CATATAN.
    var tiketDipakai = "";
    // Tabel yang membuat panel tampil pada halaman tanpa elemen tab. Disimpan
    // supaya polling ringan bisa memastikan halaman ini masih relevan tanpa
    // perlu memindai semua tabel lagi.
    var tabelPendukung = null;
    // Jumlah request yang boleh jalan bersamaan untuk scan yang sedang berjalan.
    // Ditentukan dari jumlah item, bukan konstanta, supaya dataset besar
    // otomatis melambat tanpa perlu mengubah kode.
    var koncAktif = KONKURENSI_KECIL;

    function konkurensiUntuk(totalItems) {
      return totalItems >= AMBANG_KONKURENSI ? KONKURENSI_BESAR : KONKURENSI_KECIL;
    }

    var TOMBOL =
      "box-sizing:border-box;padding:0 16px;height:42px;display:flex;" +
      "align-items:center;justify-content:center;color:#fff;border:none;" +
      "border-radius:8px;cursor:pointer;font-size:13px;font-weight:bold;" +
      "font-family:inherit;box-shadow:0 2px 8px rgba(0,0,0,.25);";

    function setStatus(teks) {
      if (ui.status) ui.status.textContent = teks;
    }

    function setProgress(persen) {
      if (ui.bar) ui.bar.style.width = Math.max(0, Math.min(100, persen)) + "%";
    }

    /*
    * Hanya ada satu tombol. Selama proses berjalan tombolnya berubah jadi
    * "Batalkan" (merah) supaya tetap bisa dihentikan tanpa tombol kedua.
    */
    function aturTombol(busy) {
      if (!ui.btn) return;
      if (busy) {
        ui.btn.textContent = "Batalkan";
        ui.btn.style.background = "#c62828";
        ui.btn.disabled = false;
      } else {
        ui.btn.textContent = "Scan & Download XLS";
        ui.btn.style.background = "#2e7d32";
        ui.btn.disabled = false;
      }
      ui.btn.style.opacity = "1";
    }

    function toast(pesan, durasi) {
      durasi = durasi || 5000;
      var lama = document.getElementById("botimpact-toast");
      if (lama && lama.parentNode) lama.parentNode.removeChild(lama);
      var el = document.createElement("div");
      el.id = "botimpact-toast";
      el.textContent = pesan;
      el.style.cssText =
        "position:fixed;bottom:240px;right:20px;z-index:999999;" +
        "background:#263238;color:#fff;padding:12px 16px;border-radius:8px;" +
        "font-size:13px;font-weight:bold;max-width:340px;white-space:pre-wrap;" +
        "box-shadow:0 4px 12px rgba(0,0,0,.35);opacity:0;transition:opacity .25s;" +
        "font-family:inherit;line-height:1.45;";
      document.body.appendChild(el);
      el.offsetHeight;
      el.style.opacity = "1";
      setTimeout(function () {
        el.style.opacity = "0";
        setTimeout(function () {
          if (el.parentNode) el.parentNode.removeChild(el);
        }, 300);
      }, durasi);
    }

    function gagal(pesan) {
      log("GAGAL: " + pesan);
      setStatus("Gagal. Buka Console untuk detail.");
      toast("Gagal:\n" + pesan, 12000);
    }

    function pasangUI() {
      if (!document.body) {
        setTimeout(pasangUI, 200);
        return;
      }

      var lama = document.getElementById("botimpact-ui");
      if (lama && lama.parentNode) lama.parentNode.removeChild(lama);

      var wadah = document.createElement("div");
      wadah.id = "botimpact-ui";
      wadah.style.cssText =
        "position:fixed;right:20px;bottom:20px;z-index:99998;width:272px;" +
        "background:#fff;border-radius:10px;box-shadow:0 4px 18px rgba(0,0,0,.22);" +
        "border:1px solid #e0e0e0;padding:12px;font-family:inherit;";

      var judul = document.createElement("div");
      judul.textContent = "Bot Ambil Impacted Services";
      judul.style.cssText =
        "font-size:12px;font-weight:bold;color:#1565c0;margin-bottom:6px;";
      wadah.appendChild(judul);

      // Nomor tiket tampil sebagai teks baca-saja. Tidak ada kotak isian:
      // semuanya dari satu klik.
      ui.infoTiket = document.createElement("div");
      ui.infoTiket.style.cssText =
        "font-size:11px;color:#555;margin-bottom:9px;line-height:1.4;" +
        "word-break:break-word;";
      ui.infoTiket.textContent = "Nomor tiket: membaca halaman...";
      wadah.appendChild(ui.infoTiket);

      ui.btn = document.createElement("button");
      ui.btn.textContent = "Scan & Download XLS";
      ui.btn.style.cssText = TOMBOL + "width:100%;background:#2e7d32;";
      ui.btn.addEventListener("click", function () {
        if (sedangJalan) {
          sedangBatal = true;
          ui.btn.disabled = true;
          ui.btn.textContent = "Membatalkan...";
          setStatus("Membatalkan, tunggu proses yang sedang jalan...");
          return;
        }
        jalankan();
      });
      wadah.appendChild(ui.btn);

      var track = document.createElement("div");
      track.style.cssText =
        "height:6px;background:#e0e0e0;border-radius:3px;margin-top:10px;overflow:hidden;";
      ui.bar = document.createElement("div");
      ui.bar.style.cssText =
        "height:100%;width:0%;background:#1565c0;transition:width .2s;";
      track.appendChild(ui.bar);
      wadah.appendChild(track);

      ui.status = document.createElement("div");
      ui.status.textContent = "Siap. Klik Scan & Download XLS.";
      ui.status.style.cssText =
        "font-size:11px;color:#333;margin-top:7px;line-height:1.45;word-break:break-word;";
      wadah.appendChild(ui.status);

      document.body.appendChild(wadah);

      // Diagnosis tetap bisa dipanggil dari Console: botimpactCek()
      try {
        window.botimpactCek = cekStruktur;
      } catch (e) {
        /* abaikan */
      }

      setTimeout(function () {
        if (ui.infoTiket) {
          var t = deteksiNoTiketAman();
          tiketDipakai = t.nomor || "";
          ui.infoTiket.textContent = tiketDipakai
            ? "Tiket: " + tiketDipakai + " (dari " + t.sumber + ")"
            : "Tiket: tidak terdeteksi otomatis - nama file akan TANPA-TIKET.";
        }
      }, 900);
    }

    /*
    * Tampilkan panel bila ini halaman detail tiket yang punya tab "Impacted
    * Service".
    *
    * Panel sengaja tetap tampil walaupun tab-nya belum aktif. Kalau panel
    * disembunyikan sampai tabel terlihat, user tidak akan pernah punya tombol
    * untuk diklik, padahal script sekarang yang bisa mengklik tab itu sendiri.
    */
    function halamanPunyaImpacted() {
      var docs = semuaDokumen();
      for (var i = 0; i < docs.length; i++) {
        if (cariTabImpacted(docs[i].doc)) {
          tabelPendukung = null;
          return true;
        }
      }
      // Tanpa tab yang bisa dikenali, andalkan tabelnya.
      var ctx = cariKonteksTabel();
      tabelPendukung = ctx ? ctx.tabel : null;
      return !!ctx;
    }

    /*
    * Cek versi ringan untuk polling berkala.
    *
    * Dipakai ketika panel sudah tampil. Yang perlu dijawab hanya "apakah
    * halaman ini masih halaman tiket yang relevan", bukan "tabel mana yang
    * harus dipakai". Jadi tabel tidak dipindai sama sekali: cukup cari elemen
    * tab, atau lihat apakah tabel yang dulu dipakai masih terpasang.
    *
    * Tabel pendukung sengaja disimpan supaya panel tidak berkedip pada halaman
    * yang tidak punya elemen tab sama sekali.
    */
    function halamanPunyaImpactedRingan() {
      var docs = semuaDokumen();
      for (var i = 0; i < docs.length; i++) {
        if (cariTabImpacted(docs[i].doc)) return true;
      }
      if (tabelPendukung && tabelPendukung.isConnected) return true;
      tabelPendukung = null;
      return false;
    }

    function aturVisibilitas() {
      var wadah = document.getElementById("botimpact-ui");
      if (!wadah) return;
      var tampil = false;
      try {
        // Panel sudah tampil: pakai cek ringan. Panel tersembunyi: cek penuh.
        tampil = wadah.style.display !== "none"
          ? halamanPunyaImpactedRingan()
          : halamanPunyaImpacted();
      } catch (e) {
        tampil = false;
      }
      wadah.style.display = tampil ? "block" : "none";
    }

    // ==================== ORKESTRASI ====================

    /*
    * Gabung baris baru ke peta unik.
    *
    * Kuncinya SELURUH isi baris, bukan SERVICE ID.
    *
    * Kunci SERVICE ID membuang baris yang bedanya ada di kolom lain. Contoh nyata
    * dari log: 1.910 item menghasilkan 901 baris unik - 9 baris hilang karena
    * ServiceId sama dipakai entri berbeda, dan 1.009 baris hilang karena caption
    * "1,910" dibaca sebagai 910. Yang hilang diam-diam seperti itu tidak bisa
    * dikembalikan file mana pun, karena barisnya memang tidak pernah diambil.
    *
    * Yang bisa diperbaiki adalah sisi lain: dua baris dengan SERVICE ID sama tapi
    * kolom lain berbeda akan keduanya tersimpan, karena sidik kuncinya berbeda.
    * Yang tetap digabung hanya baris identik seluruhnya - itu kembar sungguhan,
    * dan memaksanya masuk dua kali hanya menambah baris tanpa informasi baru.
    *
    * Sidik baris disusun dengan panjang setiap sel ikut ditulis, jadi tidak ada
    * dua baris berbeda yang bisa menghasilkan sidik sama. Versi yang lebih
    * sederhana, join("\u0001"), bisa bertabrakan kalau ada sel yang teksnya
    * mengandung \u0001: baris ["a\u0001b","c"] dan ["a","b\u0001c"] akan
    * menghasilkan string yang persis sama. formatKey() tidak punya celah itu,
    * dan tidak perlu bergantung pada tebakan soal karakter mana yang mungkin
    * muncul di halaman Insera.
    */
    function formatKey(baris) {
      var bagian = [];
      for (var i = 0; i < baris.length; i++) {
        bagian.push(baris[i].length + ":" + baris[i]);
      }
      return bagian.join("|");
    }

    /*
    * Gabung baris yang identik seluruh isinya, bukan yang punya SERVICE ID sama.
    *
    * Balikan: jumlah baris yang benar-benar baru masuk peta. Baris identik yang
    * digabung dihitung di gabungBaris.kembar supaya bisa dilaporkan terpisah,
    * bukan dicampur dengan baris yang hilang. Penghitung itu di-nol-kan di awal
    * setiap scan, jadi dua scan di tab yang sama tidak saling menambah.
    */
    function gabungBaris(baru, peta) {
      var masuk = 0;
      for (var i = 0; i < baru.length; i++) {
        var k = formatKey(baru[i]);
        if (peta.has(k)) {
          gabungBaris.kembar++;
        } else {
          peta.set(k, baru[i]);
          masuk++;
        }
      }
      return masuk;
    }
    gabungBaris.kembar = 0;

    /*
    * Nama file: ImpactService_INC53545268_1910.xls
    *
    * Jumlah baris yang dicetak adalah yang benar-benar masuk file, bukan total
    * item di Insera. Untuk dataset yang bermasalah, kedua angka itu berbeda
    * jauh: nama ImpactService_INCxxx_901_SEBAGIAN.xls langsung menunjukkan
    * bahwa file-nya tidak lengkap, tanpa harus membuka filenya dulu.
    *
    * Endingan _SEBAGIAN tetap ada dan tidak bisa dihapus: itu satu-satunya
    * penanda di nama file bahwa isinya kurang.
    */
    function namaFile(parsial, jumlahBaris) {
      var tiket = (tiketDipakai || "").trim();
      tiket = amanUntukNamaBerkas(tiket.toUpperCase());
      if (!tiket) tiket = "TANPA-TIKET";
      return (
        "ImpactService_" + tiket + "_" + jumlahBaris +
        (parsial ? "_SEBAGIAN" : "") + ".xls"
      );
    }

    async function jalankan() {
      sedangBatal = false;
      sedangJalan = true;
      setProgress(0);
      setStatus("Menyiapkan...");
      aturTombol(true);

      // Akumulator anomali kolom, dikumpulkan dari seluruh halaman.
      var kolom = { Pendek: 0, Panjang: 0, contoh: [] };
      function rekamKolom(b) {
        kolom.Pendek += b.Pendek || 0;
        kolom.Panjang += b.Panjang || 0;
        for (var i = 0; i < (b.contoh || []).length && kolom.contoh.length < 8; i++) {
          kolom.contoh.push(b.contoh[i]);
        }
      }

      try {
        // 1. Pastikan tab "Impacted Service" aktif dan bagian ringkasannya terbuka.
        //    Script tidak lagi mengasumsikan user sudah mengklik manual.
        setStatus("Menyiapkan tab Impacted Service...");
        var siapkan = await siapkanHalamanImpacted();
        if (sedangBatal) return setStatus("Dibatalkan oleh user.");

        konteksCache = cariKonteksTabel();
        if (!konteksCache) {
          // Kumpulkan semua tabel yang ada supaya pesan error bisa menyebutkan
          // kolom apa saja yang sebenarnya ditemukan.
          var semuaTabel = [];
          var docs = semuaDokumen();
          for (var d = 0; d < docs.length; d++) {
            var daftar = docs[d].doc.querySelectorAll("table");
            for (var q = 0; q < daftar.length; q++) {
              if (tabelNavigasi(daftar[q])) continue;
              var h = bacaHeader(daftar[q]);
              if (!h.length) continue;
              semuaTabel.push({
                kolom: h,
                terlihat: terlihatDiLayar(daftar[q]),
              });
            }
          }

          var rincian = "";
          if (semuaTabel.length) {
            rincian = "\n\nTabel yang ada di halaman ini:\n";
            for (var s = 0; s < semuaTabel.length && s < 4; s++) {
              rincian +=
                (semuaTabel[s].terlihat ? "  - tampil: " : "  - tersembunyi: ") +
                semuaTabel[s].kolom.join(" | ") + "\n";
            }
          }

          var petunjuk = !siapkan.tabAda
            ? "\n\nCatatan: tab \"Impacted Service\" tidak ditemukan di halaman ini. " +
              "Pastikan Anda berada di halaman detail tiket Insera, bukan halaman daftar."
            : "";

          gagal(
            "Tabel Impacted Services tidak ditemukan.\n\n" +
              "Sudah dicoba otomatis: klik tab \"Impacted Service\" dan membuka " +
              "bagian \"Impacted Service Summary\".\n\n" +
              "Kolom yang dicari: " + HEADER_KUNCI.join(" + ") +
              ", minimal satu dari: " + HEADER_KHAMAT.join(", ") + "." +
              petunjuk +
              rincian +
              "\nCoba manual:\n" +
              "1. Klik tab Impacted Service.\n" +
              "2. Buka/expand bagian Impacted Service Summary.\n" +
              "3. Pastikan tabelnya benar-benar terlihat di layar.\n\n" +
              "Diagnosis: buka Console (F12) lalu ketik botimpactCek()"
          );
          return;
        }

        // 2. header + baris halaman aktif
        var base = konteksCache.base;
        var headTr = barisHeader(konteksCache.tabel);
        var header = bacaHeader(konteksCache.tabel);
        var nKolomHarapan = header.length;
        var bacaanAwal = bacaBaris(konteksCache.tabel, indeksTerlihat(headTr), nKolomHarapan);
        var barisAwal = bacaanAwal.rows;
        log("Header: " + header.join(" | "));
        log("Baris di halaman aktif: " + barisAwal.length);
        rekamKolom(bacaanAwal);
        if (bacaanAwal.Panjang || bacaanAwal.Pendek) {
          log(
            "Peringatan kolom: " + bacaanAwal.Pendek + " baris kurang kolom (diisi kosong), " +
              bacaanAwal.Panjang + " baris kelebihan kolom (dibuang). " +
              bacaanAwal.contoh.join("; ")
          );
        }

        if (barisAwal.length === 0) {
          gagal(
            "Tabel ada tapi tidak berisi data.\n\n" +
              "Kemungkinan:\n" +
              "1. Impacted Services untuk tiket ini memang kosong.\n" +
              "2. Data belum selesai dimuat, tunggu sebentar lalu klik ulang."
          );
          return;
        }

        // 3. keterangan jumlah
        var ket = bacaKeterangan(konteksCache.doc);
        log("Keterangan: " + ket.teks);

        // Nomor tiket dideteksi sekali di sini, dipakai nama file & sheet CATATAN.
        var deteksiTiket = deteksiNoTiketAman();
        tiketDipakai = deteksiTiket.nomor || "";
        if (tiketDipakai) {
          log("Nomor tiket: " + tiketDipakai + " (dari " + deteksiTiket.sumber + ")");
        } else {
          log("Nomor tiket tidak terdeteksi, nama file memakai TANPA-TIKET.");
        }

        var totalItems = ket.totalItems;
        var pageSize = ket.pageSize || barisAwal.length;
        // Nomor halaman yang tidak berhasil diambil. Dipakai untuk menghitung berapa
        // item yang hilang di gerbang akhir. Kunci anti-duplikat bukan lagi kolom
        // SERVICE ID, tapi seluruh isi baris - lihat gabungBaris().
        var halamanGagal = [];
        // Nomor halaman yang URL-nya tidak ada di DOM lalu dibangun dari link
        // lain, dan cara data diambil. Dua-duanya ditulis ke sheet CATATAN
        // supaya file yang tidak lengkap bisa dibaca sendiri tanpa console.
        var halamanDariTemplate = [];
        var caraAmbil = "1 halaman";

        // Peta unik: baris halaman aktif ikut masuk supaya tidak terambil dua kali.
        var peta = new Map();
        /*
        * kembar dihitung di properti fungsi, jadi harus di-nol-kan tiap scan.
        * Tanpa ini, scan kedua di tab yang sama akan menjumlahkan kembar dari
        * scan pertama dan melaporkan angka yang bukan milik berkas ini.
        */
        gabungBaris.kembar = 0;
        gabungBaris(barisAwal, peta);

        // 4. coba naikkan page size
        if (pageSize < 25 && totalItems > pageSize * 2) {
          setStatus("Mencoba naikkan baris per halaman...");
          var naik = await cobaPageSizeNaik(base, pageSize, function (info) {
            setStatus(info);
            log(info);
          });
          if (naik.dipakai) pageSize = naik.pageSize;
          if (sedangBatal) return setStatus("Dibatalkan oleh user.");
        }

        /*
        * Jumlah request yang jalan bersamaan ditentukan dari dataset, bukan
        * dari nilai tetap. Scan 10.000 baris berarti 1.000 halaman; dengan 8
        * permintaan sekaligus, session Insera bisa timeout di tengah jalan.
        */
        koncAktif = konkurensiUntuk(totalItems);
        log(
          "Konkurensi: " + koncAktif + " request sekaligus (" + totalItems +
            " item, ambang " + AMBANG_KONKURENSI + ")"
        );

        var totalHalaman = totalItems > 0 ? Math.ceil(totalItems / pageSize) : 1;
        var halamanTidakDiambil = 0;

        if (totalHalaman > MAKS_HALAMAN) {
          /*
          * Pemotongan batas HARUS remembered sampai gerbang akhir. Kalau tidak,
          * scan 25.000 baris akan diam-diam menghasilkan file 20.000 baris yang
          * tidak ditandai _SEBAGIAN - pola kehilangan data yang sama persis dengan
          * caption "1,910" yang terbaca 910.
          */
          halamanTidakDiambil = totalHalaman - MAKS_HALAMAN;
          totalHalaman = MAKS_HALAMAN;
          log(
            "PERINGATAN: " + (totalItems - (MAKS_HALAMAN * pageSize)) +
              " baris di luar batas " + MAKS_HALAMAN + " halaman TIDAK akan diambil."
          );
          setStatus(
            totalItems + " item / " + pageSize + " per halaman = " +
              Math.ceil(totalItems / pageSize) + " halaman. Batas script " +
              MAKS_HALAMAN + ", " + (halamanTidakDiambil * pageSize) +
              " baris terakhir tidak diambil."
          );
        }
        log(
          "Total halaman: " + totalHalaman + " (total item " + totalItems + ")" +
            (halamanTidakDiambil ? ", " + halamanTidakDiambil + " halaman di luar batas" : "")
        );

        // 5. kumpulkan URL halaman 2..N
        if (totalHalaman > 1) {
          caraAmbil = "request paralel per halaman";
          var petaLink = panenLinkHalaman(konteksCache);
          var jumlahLink = Object.keys(petaLink).length;
          var daftarUrl = [];
          halamanDariTemplate = [];
          for (var h = 2; h <= totalHalaman; h++) {
            var url = petaLink[h] || urlDariTemplate(petaLink, h, base, pageSize);
            if (!petaLink[h]) halamanDariTemplate.push(h);
            /*
            * Page size hanya dipasang kalau belum ada di URL. Param yang sudah
            * ada bisa berprefiks dan hanya bentuk itu yang dibaca server, jadi
            * menimpanya dengan bentuk polos akan diam-diam diabaikan.
            */
            if (pageSize && !punyaParam(url, "ps")) url = setParam(url, "ps", pageSize);
            daftarUrl.push(url);
          }
          log(
            "Pagination: " + jumlahLink + " link dari DOM, " +
            (jumlahLink ? "sisanya dibangun" : "semua dibangun") +
            " (prefix " + (prefixIced(base) || "tanpa prefix") + ")"
          );
          if (jumlahLink && halamanDariTemplate.length) {
            log(
              "Halaman tanpa link DOM, URL-nya diturunkan dari link yang " +
                "terdekat (mis. halaman " +
                nomorTemplate(petaLink, halamanDariTemplate[0]) + "): " +
                halamanDariTemplate.join(", ")
            );
          }

          // 6. verifikasi halaman 2 benar-benar berpindah halaman
          setStatus("Memverifikasi halaman 2...");
          var cek = await ambilHalamanTerverifikasi(
            daftarUrl[0], 2, pageSize, totalItems, base,
            function (t) { setStatus(t); }
          );

          if (!cek.ok) {
            /*
            * Jalur URL tidak bisa dipakai. Dua sebab yang mungkin:
            *   - session login habis (bukan soal paginasi)
            *   - paginasi butuh klik, bukan GET bebas
            * Dicoba jalur klik DOM. Kalau tombol next-nya memang tidak ada,
            * baru gagal dengan pesan.
            */
            log("Jalur URL tidak berhasil: " + cek.alasan + ". Mencoba jalur klik DOM.");
            setStatus("Navigasi URL tidak bisa. Mencoba klik halaman berikutnya...");

            var rangkai = await rangkaiLewatKlik(
              konteksCache, header, totalItems,
              function (pos, total, terkumpul) {
                var persenKlik = total > 0
                  ? Math.round((Math.min(pos.sampai, total) / total) * 100) : 0;
                setProgress(persenKlik);
                setStatus(
                  "Klik: baris " + pos.dari + "-" + pos.sampai + " dari " + total +
                    "  -  " + terkumpul + " baris terkumpul"
                );
              },
              nKolomHarapan
            );

            if (rangkai.dibatalkan || sedangBatal) {
              /*
              * Tabel harus dikembalikan ke halaman 1 di semua jalur keluar, bukan
              * hanya kalau berhasil. Kalau user menutup panel atau menekan stop
              * di tengah klik, tabel yang tertinggal di halaman 47 akan membuat
              * scan berikutnya membaca halaman yang salah sebagai halaman 1 - itu
              * sumber duplikat dan baris hilang yang sulit terlihat dari log.
              */
              await kembaliKeHalaman1(konteksCache, pageSize, nKolomHarapan);
              setStatus("Dibatalkan. Data parsial tidak diunduh.");
              return;
            }
            if (!rangkai.ok) {
              await kembaliKeHalaman1(konteksCache, pageSize, nKolomHarapan);
              gagal(
                "Paginasi tidak bisa dijalankan.\n\n" +
                  "Percobaan URL: " + cek.alasan + "\n" +
                  "Percobaan klik: " + (rangkai.alasan || "tidak berhasil") +
                  (rangkai.macetDi ? "\nMacet sekitar halaman " + rangkai.macetDi : "") +
                  (rangkai.barisTercapai
                    ? "\nBaris terakhir yang terbaca: " + rangkai.barisTercapai : "") +
                  "\n\nKemungkinan:\n" +
                  "1. Session login habis - refresh halaman lalu coba lagi.\n" +
                  "2. Tabel punya tombol navigasi dengan bentuk lain.\n\n" +
                  "Untuk diagnosis, buka Console lalu ketik: botimpactCek()"
              );
              return;
            }

            // Gabungkan semua halaman yang terbaca lewat klik.
            caraAmbil = "klik DOM satu per satu (URL tidak bisa dipakai)";
            for (var q = 0; q < rangkai.semua.length; q++) {
              var satuKlik = rangkai.semua[q];
              rekamKolom(satuKlik);
              gabungBaris(satuKlik.rows, peta);
            }
            log(
              "Jalur klik DOM selesai: " + rangkai.halamanDiproses +
                " halaman, " + peta.size + " baris unik terkumpul."
            );

            // Kembalikan tabel ke halaman 1 supaya halaman Insera seperti semula.
            setStatus("Mengembalikan tabel ke halaman 1...");
            await kembaliKeHalaman1(konteksCache, pageSize, nKolomHarapan);
            setProgress(100);
          } else {
            rekamKolom(cek);
            gabungBaris(cek.rows, peta);
            log(
              "Verifikasi OK: halaman 2 berisi baris " +
                rentangHalaman(2, pageSize, totalItems).dari + "-" +
                rentangHalaman(2, pageSize, totalItems).sampai + "."
            );

            // 7. sisa halaman: paralel, tiap halaman diverifikasi & diulang sendiri
            setStatus(
              "Mengambil " + (daftarUrl.length - 1) + " halaman lagi, " +
                koncAktif + " sekaligus..."
            );

            // daftarUrl[0] sudah dipakai untuk verifikasi halaman 2, jadi sisa
            // mulai dari halaman 3. Nomor halaman itu wajib diteruskan eksplisit
            // supaya verifikasi tidak memakai indeks array.
            var mulaiHalaman = 3;
            var banyak = await ambilBanyakTerverifikasi(
              daftarUrl.slice(1), pageSize, totalItems, base,
              function (sudah, total, ket) {
                setProgress(Math.round((sudah / total) * 100));
                if (ket) setStatus(ket);
                setStatus(
                  "Halaman " + (mulaiHalaman + sudah - 1) + "/" + totalHalaman +
                    "  -  " + peta.size + " baris terkumpul"
                );
              },
              mulaiHalaman,
              nKolomHarapan
            );

            if (sedangBatal) {
              setStatus("Dibatalkan. Data parsial tidak diunduh.");
              return;
            }

            // Baris sudah diekstrak di dalam worker, jadi yang perlu dilakukan
            // hanya menggabungkannya. HTML tiap halaman sudah dibuang.
            for (var t2 = 0; t2 < banyak.slot.length; t2++) {
              var satu = banyak.slot[t2];
              if (!satu || !satu.ok) continue;
              rekamKolom(satu);
              gabungBaris(satu.rows, peta);
            }
            halamanGagal = banyak.gagal;

            /*
            * Halaman yang masih gagal dicoba lagi satu per satu.
            *
            * Untuk tabel 1.000 halaman, fallback klik berarti 1.000 klik
            * berurutan - tidak realistis. Mengulang URL halaman yang gagal jauh
            * lebih murah dan tidak bergantung pada tombol navigasi. Putaran ini
            * sengaja jalan setelah semua halaman lain selesai, supaya tidak
            * berebut koneksi dengan halaman yang belum diambil.
            */
            for (var putar = 1; putar <= PUTARAN_ULANG && halamanGagal.length; putar++) {
              log(
                "Mengulang " + halamanGagal.length + " halaman yang gagal, putaran " +
                  putar + " dari " + PUTARAN_ULANG + "..."
              );
              setStatus(
                "Mengulang " + halamanGagal.length + " halaman gagal (putaran " +
                  putar + "/" + PUTARAN_ULANG + ")..."
              );
              /*
              * Peta nomor halaman -> URL diambil dari daftarUrl. daftarUrl
              * dimulai dari halaman 2, jadi daftarUrl[h-2] secara matematis
              * benar, tapi pemetaan eksplisit lebih aman kalau urutan pemanggil
              * berubah.
              */
              var nomorKeUrl = {};
              for (var pk = 0; pk < daftarUrl.length; pk++) {
                nomorKeUrl[pk + 2] = daftarUrl[pk];
              }
              var sisip = [];
              for (var g = 0; g < halamanGagal.length; g++) {
                var nomorGagal = halamanGagal[g];
                if (nomorKeUrl[nomorGagal]) {
                  sisip.push({ nomor: nomorGagal, url: nomorKeUrl[nomorGagal] });
                }
              }
              if (!sisip.length) break;
              var ulang = await ambilBanyakTerverifikasi(
                sisip.map(function (s) { return s.url; }),
                pageSize, totalItems, base, null, 0, nKolomHarapan
              );
              if (sedungBatal) {
                setStatus("Dibatalkan. Data parsial tidak diunduh.");
                return;
              }
              var masihGagal = [];
              for (var u = 0; u < ulang.slot.length; u++) {
                var slotU = ulang.slot[u];
                if (slotU && slotU.ok) {
                  rekamKolom(slotU);
                  gabungBaris(slotU.rows, peta);
                } else if (slotU && slotU.nomor) {
                  /*
                  * Nomor asli halaman yang gagal, bukan indeks 0..n-1. Ini yang
                  * ditulis ke sheet CATATAN dan yang menentukan berapa item
                  * hilang, jadi harus nomor halaman sebenarnya.
                  */
                  masihGagal.push(sisip[u].nomor);
                }
              }
              halamanGagal = masihGagal;
            }

            if (halamanGagal.length && halamanGagal.length > BATAS_LARANG_KLIK) {
              /*
              * Terlalu banyak halaman gagal untuk diklik satu per satu. Klik
              * berurutan untuk tabel besar akan berjalan sangat lama, dan file
              * sudah tetap akan ditandai sebagian. Jadi lebih baik berhenti di
              * sini dan menulis daftar halaman yang hilang ke sheet CATATAN.
              */
              log(
                "Sisa " + halamanGagal.length + " halaman gagal, melebihi batas " +
                  BATAS_LARANG_KLIK + " untuk navigasi klik. Lewati fallback klik."
              );
              caraAmbil = "request paralel + " + PUTARAN_ULANG +
                " putaran ulang; " + halamanGagal.length + " halaman tidak terbaca";
            } else if (halamanGagal.length) {
              log(
                "Halaman gagal lewat URL: " + halamanGagal.join(", ") +
                  ". Melengkapi lewat jalur klik DOM."
              );
              setStatus(
                "Sebagian halaman gagal lewat URL. Melengkapi lewat klik (" +
                  totalHalaman + " halaman, bisa makan waktu beberapa menit)..."
              );

              var ladenKlik = await rangkaiLewatKlik(
                konteksCache, header, totalItems,
                function (pos, total, terkumpul) {
                  setProgress(
                    total > 0
                      ? Math.round((Math.min(pos.sampai, total) / total) * 100) : 0
                  );
                  setStatus(
                    "Klik: baris " + pos.dari + "-" + pos.sampai + " dari " + total +
                      "  -  " + terkumpul + " baris terkumpul"
                  );
                },
                nKolomHarapan
              );

              if (ladenKlik.dibatalkan || sedangBatal) {
                await kembaliKeHalaman1(konteksCache, pageSize, nKolomHarapan);
                setStatus("Dibatalkan. Data parsial tidak diunduh.");
                return;
              }

              if (!ladenKlik.ok) {
                /*
                * Jalur klik juga gagal. Hasil jalur URL yang sudah terkumpul
                * tetap dipakai supaya file tetap ada, dan halaman yang gagal
                * dicatat di sheet CATATAN seperti sebelumnya.
                */
                log(
                  "Jalur klik juga tidak berhasil: " +
                    (ladenKlik.alasan || "tanpa alasan") +
                    ". Memakai hasil jalur URL."
                );
              } else {
                for (var q2 = 0; q2 < ladenKlik.semua.length; q2++) {
                  var satuKlik2 = ladenKlik.semua[q2];
                  rekamKolom(satuKlik2);
                  gabungBaris(satuKlik2.rows, peta);
                }
                // Jalur klik berjalan dari halaman 1 sampai habis, jadi tidak
                // ada lagi halaman yang belum dibaca.
                halamanGagal = [];
                caraAmbil = "request paralel + klik DOM untuk halaman yang gagal URL";
                log(
                  "Jalur klik DOM selesai: " + ladenKlik.halamanDiproses +
                    " halaman, " + peta.size + " baris unik terkumpul."
                );
                setStatus("Mengembalikan tabel ke halaman 1...");
                await kembaliKeHalaman1(konteksCache, pageSize, nKolomHarapan);
                setProgress(100);
              }
            }
          }
        } else {
          log("Hanya 1 halaman, tidak perlu fetch tambahan.");
        }

        /*
        * 8. Gerbang akhir.
        *
        * Jumlah baris di file bisa lebih sedikit dari total item, dan setiap
        * sebabnya harus diperlakukan berbeda karena 의미nya berbeda bagi user:
        *
        *   a) Semua halaman terbaca, tapi ada baris yang isinya identik
        *      -> itu kembar sungguhan. Tidak ada data hilang, file tetap bersih.
        *
        *   b) Ada halaman gagal, halaman di luar batas, atau baris dengan kolom
        *      berlebih yang dibuang
        *      -> benar-benar ada data yang tidak terbaca. File diberi akhiran
        *         _SEBAGIAN dan sheet CATATAN, supaya tidak pernah terlihat
        *         lengkap padahal isinya kurang.
        *
        * (a) dan (b) dulu disatukan jadi satu angka "SERVICE ID kembar" yang
        * dihitung dari totalItems - baris. Untuk kasus nyata di log, kode itu
        * melaporkan 1.910 item -> 901 baris sebagai "9 baris kembar, tidak ada
        * data hilang", padahal 1.009 baris tidak pernah diambil sama sekali.
        */
        var baris = Array.from(peta.values());
        var persen = totalItems > 0 ? Math.round((baris.length / totalItems) * 100) : 100;
        var adaHalamanGagal = halamanGagal.length > 0;
        var adaBarisDibuang = kolom.Panjang > 0;
        var adaYangTidakDiambil = halamanTidakDiambil > 0;

        /*
        * Baris kembar dihitung dari penggabungan, bukan dari totalItems - baris.
        *
        * Selisih itu sebelumnya dipakai untuk menyebut "kembar digabung", padahal
        * selisih yang sama muncul kalau halaman gagal atau ada baris di luar
        * batas. Untuk scan 1.910 item yang terpotong jadi 901 baris, kode lama
        * melaporkan "0 data hilang" padahal 1.009 baris hilang. Sekarang
        * baris kembar = jumlah baris identik yang benar-benar digabung, dan
        * setiap baris yang hilang dilaporkan sebagai sebab yang terpisah.
        */
var parsial = adaHalamanGagal || adaBarisDibuang || adaYangTidakDiambil;

        /*
        * Berapa item yang tidak masuk ke file, dan kenapa. Ini angka yang harus
        * dibaca user untuk tahu apakah file bisa dipakai.
        */
var itemTakMasuk = 0;
        var alasanTakMasuk = [];

        // Kejadian yang sudah punya penjelasan. Semua ini sudah tercermin di
        // itemTakMasuk, jadi harus dikurangi sebelum menghitung sisa yang tak
        // terduga - kalau tidak, baris yang sama akan dihitung dua kali.
        var sudahTerjelaskan = kolom.Panjang;

        if (adaYangTidakDiambil) {
          var barisPotong = halamanTidakDiambil * pageSize;
          // Halaman terakhir mungkin tidak penuh; jangan lebih dari sisa item.
          if (itemTakMasuk + barisPotong > totalItems) {
            barisPotong = Math.max(0, totalItems - itemTakMasuk);
          }
          itemTakMasuk += barisPotong;
          sudahTerjelaskan += barisPotong;
          alasanTakMasuk.push(
            halamanTidakDiambil + " halaman di luar batas " + MAKS_HALAMAN +
              " tidak diambil"
          );
        }
        if (adaHalamanGagal) {
          for (var hg = 0; hg < halamanGagal.length; hg++) {
            var rentangGagal = rentangHalaman(
              halamanGagal[hg], pageSize, totalItems
            );
            var jangkar = rentangGagal.sampai - rentangGagal.dari + 1;
            if (itemTakMasuk + jangkar > totalItems) {
              jangkar = Math.max(0, totalItems - itemTakMasuk);
            }
            itemTakMasuk += jangkar;
            sudahTerjelaskan += jangkar;
          }
          alasanTakMasuk.push(
            halamanGagal.length + " halaman gagal diakses: " +
              halamanGagal.join(", ")
          );
        }
        if (adaBarisDibuang) {
          alasanTakMasuk.push(
            kolom.Panjang + " baris punya kolom berlebih sehingga tidak bisa " +
              "dipakai di sheet ini"
          );
        }
        var kembarDigabung = gabungBaris.kembar;

        /*
        * Penghitung terakhir: baris yang hilang tanpa sebab yang diketahui.
        *
        * Ini yang menangkap kasus dari log. 1.910 item, 901 baris, 0 halaman
        * gagal, 0 halaman di luar batas - semua sebab yang tercatat bernilai nol,
        * tapi 1.009 baris tetap tidak ada di file. Penyebabnya ada di luar
        * pengamatan script: keterangan `1,910 items found` dibaca `910`, jadi
        * 100 halaman tidak pernah diambil dan tidak masuk daftar halaman gagal
        * karena memang tidak dicoba.
        *
        * Baris yang hilang seperti ini tidak boleh dilaporkan sebagai "tidak ada
        * data hilang". Kode lama persismelakukan itu, dan hasilnya file 901
        * baris tanpa `_SEBAGIAN` untuk dataset 1.910 item.
        *
        * Baris kembar identik dikurangi dulu, karena memang baris yang sama
        * dihitung dua kali oleh server dan itu bukan kehilangan data.
        */
        var selisih =
          totalItems - baris.length - kembarDigabung - sudahTerjelaskan;
        if (selisih > 0) {
          itemTakMasuk += selisih;
          parsial = true;
          alasanTakMasuk.push(
            selisih + " item tidak terbaca tanpa sebab yang diketahui. " +
              "Kemungkinan keterangan jumlah item salah dibaca, atau server " +
              "mengubah data saat halaman diambil. Ambil ulang, atau pakai " +
              "botimpactCek() untuk memeriksa keterangan paginator."
          );
        }

        /*
        * Percentage hanya ditampilkan kalau file-nya utuh. Untuk file sebagian,
        * "99%" dari 1.910 item padahal 1.009 baris hilang hanya mendorong user
        * salah percaya file sudah utuh, dan persentasenya tidak pernah mencapai
        * 100% - jadi lebih jujur ditulis sebagai jumlah item yang tidak terbaca.
        */
        log(
          "Ringkasan: " + baris.length + " baris dari " + totalItems + " item" +
            (parsial ? "" : " (" + persen + "%)") + ", " + header.length + " kolom, " +
            halamanGagal.length + " halaman gagal, " +
            kolom.Panjang + " baris dibuang, " +
            kembarDigabung + " baris identik digabung, " +
            (parsial ? itemTakMasuk + " item TIDAK masuk file" : "tidak ada data hilang")
        );

        // Isi sheet CATATAN. Ada kalau file ditandai sebagian.
        var catatan = [];
        if (parsial) {
          catatan.push(["WAKTU SCAN", stempelWaktu()]);
          catatan.push(["Nomor tiket", tiketDipakai]);
          /*
          * Dua angka ini harus dibaca bersamaan. "Total item" berasal dari
          * keterangan paginator, sedangkan "Baris unik di file ini" dihitung dari
          * baris yang benar-benar terkumpul. Bedanya adalah data yang hilang -
          * kecuali kalau baris kembar identik ikut dihitung, yang memang bukan
          * kehilangan.
          */
          catatan.push(["Total item di Insera", String(totalItems)]);
          catatan.push(["Baris unik di file ini", String(baris.length)]);
          catatan.push([
            "Baris terbaca (termasuk kembar)",
            String(baris.length + kembarDigabung),
          ]);
          catatan.push(["Halaman gagal", halamanGagal.join(", ") || "(tidak ada)"]);
          catatan.push([
            "Halaman di luar batas (tidak diambil)",
            halamanTidakDiambil ? halamanTidakDiambil + " halaman terakhir" : "(tidak ada)",
          ]);
          catatan.push([
            "Item tidak masuk file",
            itemTakMasuk + " item" + (alasanTakMasuk.length
              ? " - " + alasanTakMasuk.join("; ") : ""),
          ]);
          catatan.push([
            "Baris identik digabung (bukan data hilang)",
            kembarDigabung + " baris",
          ]);
          catatan.push([
            "Baris dibuang (kolom berlebih)",
            kolom.Panjang + " baris",
          ]);
          catatan.push([
            "Baris kurang kolom (diisi kosong)",
            kolom.Pendek + " baris",
          ]);
          catatan.push([
            "Cara diambil",
            caraAmbil,
          ]);
          if (halamanDariTemplate.length) {
            catatan.push([
              "Halaman dari URL template",
              halamanDariTemplate.length + " halaman: " +
                halamanDariTemplate.join(", "),
            ]);
          }
          if (kolom.contoh.length) {
            catatan.push(["Contoh anomali kolom", kolom.contoh.join(" | ")]);
          }
          catatan.push([
            "CATATAN",
            "File ini TIDAK lengkap. Data dari halaman yang gagal tidak ikut " +
              "tersimpan. Ambil ulang setelah beberapa saat.",
          ]);
        }

        setStatus("Menyusun file Excel...");
        setProgress(100);
        var nama = namaFile(parsial, baris.length);
        var xml = keSpreadsheetML(header, baris, catatan);
        unduhBerkas(nama, xml);

        if (parsial) {
          setStatus(
            "Selesai, tapi SEBAGIAN: " + baris.length + " dari " + totalItems +
              " item. Tidak masuk: " + itemTakMasuk + " item. " +
              (alasanTakMasuk.join("; ") || "") +
              " Buka sheet CATATAN."
          );
          toast(
            "Selesai! " + baris.length + " dari " + totalItems + " item.\n" +
              itemTakMasuk + " item tidak masuk file. File diberi nama " +
              "_SEBAGIAN dan ada sheet CATATAN.",
            11000
          );
        } else if (kembarDigabung > 0) {
          setStatus(
            "Selesai. " + baris.length + " baris dari " + totalItems +
              " item; " + kembarDigabung + " baris identik digabung."
          );
          toast(
            "Selesai! " + baris.length + " baris -> " + nama + "\n" +
              kembarDigabung + " baris identik digabung, tidak ada data hilang.",
            9000
          );
        } else {
          setStatus("Selesai. " + baris.length + " baris -> " + nama);
          toast(
            "Selesai! " + baris.length + " baris lengkap -> " + nama,
            9000
          );
        }
      } catch (e) {
        log("ERROR: " + (e && e.stack ? e.stack : e));
        gagal("Terjadi kesalahan: " + (e && e.message ? e.message : String(e)));
      } finally {
        sedangJalan = false;
        aturTombol(false);
      }
    }

    // ==================== SPREADSHEETML (.xls) ====================

    /*
    * Format .xls SpreadsheetML 2003 (XML). Dipilih karena:
    *   - dibuka langsung oleh Excel, LibreOffice, dan Google Sheets
    *   - tidak butuh pustaka eksternal (xlsx asli butuh library zip)
    *   - bisa menentukan lebar kolom, tebalkan header, dan kunci baris header
    *
    * PENTING: semua sel dikunci ss:Type="String". Kalau PHONE NUMBER
    * ("6281291870453" / "+628111032393") dikira angka, Excel mengubahnya jadi
    * scientific notation dan formatnya rusak.
    */
    function keSpreadsheetML(header, rows, catatan) {
      var kolom = header && header.length ? header : HEADER_CADANGAN;
      var nKolom = kolom.length;

      var lebar = "";
      for (var c = 0; c < nKolom; c++) {
        lebar += '<Column ss:Width="' + (LEBAR_KOLOM[kolom[c]] || LEBAR_DEFAULT) + '"/>';
      }

      var selHeader = "";
      for (var h = 0; h < nKolom; h++) {
        selHeader +=
          '<Cell ss:StyleID="hdr"><Data ss:Type="String">' +
          xmlEsc(kolom[h]) +
          "</Data></Cell>";
      }

      var isi = "";
      for (var r = 0; r < rows.length; r++) {
        var row = rows[r];
        var sel = "";
        for (var k = 0; k < nKolom; k++) {
          var v = row[k] !== undefined ? row[k] : "";
          sel +=
            v === ""
              ? "<Cell/>"
              : '<Cell ss:StyleID="txt"><Data ss:Type="String">' +
                xmlEsc(v) +
                "</Data></Cell>";
        }
        isi += "<Row>" + sel + "</Row>";
      }

      var dibuat = waktuSekarang().toISOString();

      /*
      * Sheet kedua "CATATAN", hanya dibuat bila ada catatan (file sebagian).
      * Dua kolom: label dan nilai, tanpa freeze/filter karena isinya pendek.
      */
      var sheetCatatan = "";
      if (catatan && catatan.length) {
        var barisCatatan = "";
        for (var c2 = 0; c2 < catatan.length; c2++) {
          var pasangan = catatan[c2];
          barisCatatan +=
            "<Row>" +
            '<Cell ss:StyleID="hdr"><Data ss:Type="String">' +
            xmlEsc(pasangan[0]) + "</Data></Cell>" +
            '<Cell ss:StyleID="txt"><Data ss:Type="String">' +
            xmlEsc(pasangan[1] === undefined ? "" : pasangan[1]) +
            "</Data></Cell>" +
            "</Row>";
        }
        sheetCatatan =
          '  <Worksheet ss:Name="CATATAN">\n' +
          '    <Table ss:ExpandedColumnCount="2" ss:ExpandedRowCount="' +
          catatan.length +
          '" x:FullColumns="1" x:FullRows="1" ss:DefaultRowHeight="14.5">\n' +
          '      <Column ss:Width="190.5"/><Column ss:Width="380.5"/>\n' +
          barisCatatan +
          "\n" +
          "    </Table>\n" +
          "  </Worksheet>\n";
      }
      // Urutan anak <Style> harus mengikuti skema: Alignment, Borders, Font,
      // Interior, NumberFormat, Protection. Jika <Border> diletakkan setelah
      // <Font>, Excel menolak membuka berkasnya.
      var border =
        "      <Borders>\n" +
        '        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D0D0D0"/>\n' +
        '        <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D0D0D0"/>\n' +
        '        <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D0D0D0"/>\n' +
        '        <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D0D0D0"/>\n' +
        "      </Borders>\n";

      return (
        '<?xml version="1.0"?>\n' +
        '<?mso-application progid="Excel.Sheet"?>\n' +
        '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"' +
        ' xmlns:o="urn:schemas-microsoft-com:office:office"' +
        ' xmlns:x="urn:schemas-microsoft-com:office:excel"' +
        ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"' +
        ' xmlns:html="http://www.w3.org/TR/REC-html40">\n' +
        '  <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">\n' +
        "    <Title>Impacted Services</Title>\n" +
        "    <Author>Bot Ambil Impacted Services</Author>\n" +
        "    <Created>" +
        xmlEsc(dibuat) +
        "</Created>\n" +
        "  </DocumentProperties>\n" +
        "  <Styles>\n" +
        '    <Style ss:ID="Default" ss:Name="Normal">\n' +
        '      <Alignment ss:Vertical="Bottom"/>\n' +
        "      <Borders/>\n" +
        '      <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>\n' +
        "      <Interior/>\n" +
        "      <NumberFormat/>\n" +
        "      <Protection/>\n" +
        "    </Style>\n" +
        '    <Style ss:ID="hdr">\n' +
        '      <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n' +
        border +
        '      <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Bold="1" ss:Color="#FFFFFF"/>\n' +
        '      <Interior ss:Color="#1565C0" ss:Pattern="Solid"/>\n' +
        "    </Style>\n" +
        '    <Style ss:ID="txt">\n' +
        '      <Alignment ss:Vertical="Top"/>\n' +
        border +
        '      <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>\n' +
        "      <Interior/>\n" +
        "    </Style>\n" +
        "  </Styles>\n" +
        '  <Worksheet ss:Name="Impacted Services">\n' +
        '    <Table ss:ExpandedColumnCount="' +
        nKolom +
        '" ss:ExpandedRowCount="' +
        (rows.length + 1) +
        '" x:FullColumns="1" x:FullRows="1" ss:DefaultRowHeight="14.5">\n' +
        "      " +
        lebar +
        "\n" +
        '      <Row ss:Height="24">' +
        selHeader +
        "</Row>\n" +
        "      " +
        isi +
        "\n" +
        "    </Table>\n" +
        '    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">\n' +
        '      <PageSetup><Layout x:Orientation="Landscape"/>' +
        '<PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/></PageSetup>\n' +
        "      <FreezePanes/>\n" +
        "      <FrozenNoSplit/>\n" +
        "      <SplitHorizontal>1</SplitHorizontal>\n" +
        "      <TopRowBottomPane>1</TopRowBottomPane>\n" +
        "      <ActivePane>2</ActivePane>\n" +
        "      <Panes><Pane><Number>3</Number></Pane><Pane><Number>2</Number></Pane></Panes>\n" +
        "    </WorksheetOptions>\n" +
        '    <AutoFilter x:Range="R1C1:R' +
        (rows.length + 1) +
        "C" +
        nKolom +
        '" xmlns="urn:schemas-microsoft-com:office:excel"/>\n' +
        "  </Worksheet>\n" +
        sheetCatatan +
        "</Workbook>"
      );
    }

    // ==================== UNDUH BERKAS ====================

    function unduhBerkas(nama, isi) {
      // BOM UTF-8: Excel lebih dulu membaca SpreadsheetML sebagai teks, BOM
      // memastikan ia mengira berkas ini berformat Windows-1252.
      var blob = new Blob(["\uFEFF", isi], {
        type: "application/vnd.ms-excel;charset=utf-8",
      });
      var url = URL.createObjectURL(blob);
      var a = document.createElement("a");
      a.href = url;
      a.download = nama;
      a.style.display = "none";
      document.body.appendChild(a);
      a.click();
      setTimeout(function () {
        if (a.parentNode) a.parentNode.removeChild(a);
        URL.revokeObjectURL(url);
      }, 2000);
      log("Berkas diunduh: " + nama + " (" + Math.round(isi.length / 1024) + " KB)");
    }

    // ==================== CEK STRUKTUR ====================

    // Salin teks ke clipboard. Clipboard API butuh izin, jadi ada jalur manual
    // sebagai cadangan.
    function copy(teks) {
      try {
        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(teks);
          return true;
        }
      } catch (e) {
        /* lanjut ke cara manual */
      }
      try {
        var a = document.createElement("textarea");
        a.value = teks;
        a.setAttribute("readonly", "");
        a.style.cssText = "position:fixed;left:-9999px;top:0;opacity:0";
        document.body.appendChild(a);
        a.select();
        var ok = document.execCommand && document.execCommand("copy");
        if (a.parentNode) a.parentNode.removeChild(a);
        return !!ok;
      } catch (e) {
        return false;
      }
    }

    // Ringkasan semua tabel di halaman: dipakai untuk diagnosis tab yang salah.
    function DaftarTabelRingkas() {
      var hasil = [];
      var docs = semuaDokumen();
      for (var i = 0; i < docs.length; i++) {
        var doc = docs[i].doc;
        var list;
        try {
          list = doc.querySelectorAll("table");
        } catch (e) {
          continue;
        }
        for (var j = 0; j < list.length; j++) {
          if (tabelNavigasi(list[j])) continue;
          var h = bacaHeader(list[j]);
          if (!h.length) continue;
          hasil.push({
            kolom: h,
            terlihat: terlihatDiLayar(list[j]),
            impact: adakahTabelImpact(list[j]),
            diIframe: i > 0,
          });
        }
      }
      return hasil;
    }

    // Diagnosis lengkap: laporan disalin ke clipboard lalu ditampilkan di Console.
    function cekStruktur() {
      var laporan = {};
      try {
        laporan.url_pakai = location.href;
        laporan.judul_halaman = document.title;
        laporan.jumlah_iframe = document.querySelectorAll("iframe").length;

        /*
        * src_atribut adalah nilai mentah dari atribut src, sedangkan base_setelah
        *_normalisasi adalah bentuk yang benar-benar dipakai script. Selisih
        * di antara keduanya menjelaskan langsung kenapa URL lama gagal dibaca.
        */
        laporan.iframe_src = [];
        var daftarFrame = document.querySelectorAll("iframe");
        for (var fi = 0; fi < daftarFrame.length; fi++) {
          var mentah = (daftarFrame[fi].getAttribute("src") || "").trim();
          laporan.iframe_src.push({
            src_atribut: mentah || "(tidak ada)",
            base_setelah_normalisasi: absolutURL(mentah, location.href) || "(tidak bisa)",
          });
        }

        laporan.header_yang_dicari = HEADER_KUNCI.concat(HEADER_KHAMAT);

        // Status tab dan sub-section, supaya jelas masalahnya di tab atau di tabel.
        var docs = semuaDokumen();
        laporan.tab_impacted_service = "TIDAK DITEMUKAN";
        for (var d = 0; d < docs.length; d++) {
          var tab = cariTabImpacted(docs[d].doc);
          if (tab) {
            laporan.tab_impacted_service = "ditemukan";
            laporan.tab_sudah_aktif = !!tab.aktif;
            laporan.tab_tag = tab.el.tagName;
            laporan.tab_class = tab.el.getAttribute("class") || "(tanpa class)";
            laporan.tab_teks = teksElement(tab.el);
            break;
          }
        }
        laporan.ringkasan_ditemukan = !!cariElemenRingkasan(document);

        var semuaTabel = DaftarTabelRingkas();
        laporan.tabel_yang_ditemukan = semuaTabel.map(function (t, i) {
          return {
            no: i + 1,
            kolom: t.kolom.join(" | "),
            terlihat: t.terlihat,
            tabel_impact: t.impact,
            diIframe: t.diIframe,
          };
        });

        var ctx = cariKonteksTabel(true);
        if (!ctx) {
          laporan.hasil = "TABEL IMPACTED SERVICES TIDAK DITEMUKAN";
          laporan.saran =
            "Klik tab Impacted Service, buka bagian Impacted Service Summary, " +
            "lalu jalankan cekStruktur() lagi.";
        } else {
          laporan.hasil = ctx.identity
            ? "TABEL COCOK (Impacted Services)"
            : "TIDAK ADA TABEL IMPACTED SERVICES";
          laporan.tabel_terpakai_id = ctx.tabel.id || "(tanpa id)";
          laporan.tabel_terpakai_class = ctx.tabel.className || "(tanpa class)";
          laporan.tabel_terpakai_terlihat = ctx.terlihat;
          laporan.url_dasar = ctx.base;

          if (ctx.identity) {
            laporan.header = bacaHeader(ctx.tabel);
            laporan.jumlah_kolom = laporan.header.length;
            laporan.baris_di_dom = bacaBaris(
              ctx.tabel,
              indeksTerlihat(barisHeader(ctx.tabel)),
              laporan.jumlah_kolom
            ).rows.length;

            var ket = bacaKeterangan(ctx.doc);
            laporan.keterangan = ket.teks;
            laporan.total_item = ket.totalItems;
            laporan.page_size = ket.pageSize;
            if (ket.totalItems > 0 && ket.pageSize > 0) {
              laporan.total_halaman = Math.ceil(ket.totalItems / ket.pageSize);
            }

            var peta = panenLinkHalaman(ctx);
            var kunciLink = Object.keys(peta);
            laporan.jumlah_link_halaman = kunciLink.length;
            laporan.contoh_link = kunciLink.slice(0, 5).map(function (k) {
              return "hal " + k + " -> " + String(peta[k]).slice(0, 130);
            });
            laporan.prefix_icefaces = prefixIced(ctx.base) || "(tidak ada di URL)";
            laporan.link_berikutnya = linkBerikutnya(ctx);
            laporan.kandidat_tombol_berikut = kandidatBerikut(ctx).length;
          }
          laporan.no_tiket = deteksiNoTiket(ctx);
        }
      } catch (e) {
        laporan.hasil = "ERROR: " + (e && e.message ? e.message : String(e));
      }

      var teks = JSON.stringify(laporan, null, 2);
      log("=== DIAGNOSIS (botimpactCek) ===\n" + teks);

      if (copy(teks)) {
        toast("Laporan diagnosis sudah tersalin ke clipboard.\nPaste ke chat agar bisa dibaca.", 10000);
      } else {
        toast(
          "Laporan sudah tampil di Console.\nSalin manual dari Console: " +
            (laporan.hasil || ""),
          10000
        );
      }

      setStatus("Diagnosis selesai, lihat clipboard atau Console.");
      return laporan;
    }

    // ==================== PASANG ====================

    function mulaiSemua() {
      pasangUI();
      aturVisibilitas();
      setInterval(aturVisibilitas, 1500);
      log("Bot Ambil Impacted Services aktif di " + location.href);
    }

    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", mulaiSemua);
    } else {
      mulaiSemua();
    }
  })();
