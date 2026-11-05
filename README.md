# Layanan-surat-sungai-jaga
<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Sistem Layanan Surat Administrasi Desa Sungai Jaga A</title>
  <!-- Bootstrap 5 CSS -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <!-- FontAwesome Icons -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <!-- SheetJS (xlsx) -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  
  <style>
    :root {
      --primary-color: #0d6efd;
      --paper-width: 210mm;
      --paper-min-height: 297mm;
    }

    body {
      background-color: #f4f6f9;
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }

    .no-print {
      display: block;
    }

    /* Paper Styling */
    .paper-container {
      display: flex;
      justify-content: center;
      padding: 20px 0;
    }

    .paper {
      width: var(--paper-width);
      min-height: var(--paper-min-height);
      padding: 20mm 20mm 20mm 20mm;
      background: white;
      box-shadow: 0 0 15px rgba(0,0,0,0.15);
      color: #000;
      font-family: 'Bookman Old Style', 'Times New Roman', Times, serif;
      font-size: 11pt;
      line-height: 1.5;
      position: relative;
      box-sizing: border-box;
    }

    /* Header Kop Surat */
    .header-kop {
      text-align: center;
      border-bottom: 3px double #000;
      padding-bottom: 8px;
      margin-bottom: 20px;
    }

    .header-kop h3 {
      font-size: 13pt;
      font-weight: bold;
      margin: 0;
      text-transform: uppercase;
    }

    .header-kop h2 {
      font-size: 15pt;
      font-weight: bold;
      margin: 0;
      text-transform: uppercase;
    }

    .header-kop p {
      font-size: 9pt;
      margin: 2px 0 0 0;
      font-style: italic;
    }

    /* Judul Surat */
    .surat-title {
      text-align: center;
      margin-bottom: 20px;
    }

    .surat-title h4 {
      font-size: 12pt;
      font-weight: bold;
      text-decoration: underline;
      margin: 0;
      text-transform: uppercase;
    }

    .surat-title p {
      font-size: 10pt;
      margin: 2px 0 0 0;
    }

    /* Table Surat */
    .table-surat {
      width: 100%;
      margin-bottom: 15px;
      border-collapse: collapse;
    }

    .table-surat td {
      padding: 2px 0;
      vertical-align: top;
    }

    .table-surat td.label-col {
      width: 30%;
    }

    .table-surat td.colon-col {
      width: 3%;
    }

    /* TTD Box */
    .ttd-container {
      width: 100%;
      margin-top: 30px;
      display: flex;
      justify-content: flex-end;
    }

    .ttd-box {
      width: 50%;
      text-align: center;
      float: right;
    }

    .ttd-space {
      height: 70px;
    }

    /* Autocomplete List */
    .autocomplete-items {
      position: absolute;
      border: 1px solid #d4d4d4;
      border-bottom: none;
      border-top: none;
      z-index: 99;
      top: 100%;
      left: 0;
      right: 0;
      max-height: 200px;
      overflow-y: auto;
      background-color: #fff;
    }

    .autocomplete-items div {
      padding: 8px 12px;
      cursor: pointer;
      border-bottom: 1px solid #d4d4d4;
      font-size: 0.875rem;
    }

    .autocomplete-items div:hover {
      background-color: #e9ecef;
    }

    /* Print Rules */
    @media print {
      .no-print {
        display: none !important;
      }

      body {
        background: white;
        margin: 0;
        padding: 0;
      }

      .paper-container {
        padding: 0;
      }

      .paper {
        box-shadow: none;
        width: 100%;
        min-height: auto;
        padding: 0;
        margin: 0;
      }

      @page {
        size: A4;
        margin: 20mm 15mm 20mm 20mm;
      }
    }
  </style>
</head>
<body>

<!-- Navbar & Role Switcher -->
<nav class="navbar navbar-dark bg-dark no-print">
  <div class="container-fluid px-4">
    <span class="navbar-brand mb-0 h1">
      <i class="fa-solid fa-building-columns me-2 text-warning"></i> Pelayanan Surat Desa Sungai Jaga A
    </span>
    <div class="d-flex align-items-center">
      <span class="text-light me-2 small">Mode Akses:</span>
      <select id="roleSelect" class="form-select form-select-sm bg-secondary text-white border-0 fw-bold" onchange="switchRole()">
        <option value="admin" selected>Admin (Perangkat Desa)</option>
        <option value="user">User (Warga Pemohon)</option>
      </select>
    </div>
  </div>
</nav>

<div class="container-fluid no-print py-3">
  <div class="row">
    <!-- Panel Form Input Kiri -->
    <div class="col-lg-5 col-md-12 mb-4">
      <div class="card shadow-sm border-0">
        <div class="card-header bg-primary text-white d-flex justify-content-between align-items-center">
          <h6 class="card-title m-0 fw-bold" id="panelTitle">
            <i class="fa-solid fa-sliders me-2"></i>Panel Kontrol Admin
          </h6>
          <button class="btn btn-sm btn-light text-primary fw-bold" onclick="window.print()">
            <i class="fa-solid fa-print me-1"></i>Cetak Surat
          </button>
        </div>
        <div class="card-body">
          
          <!-- KHUSUS ADMIN: Import Excel grid_agtkeluarga_4.xls -->
          <div id="adminImportBox" class="mb-3 p-3 bg-light rounded border">
            <label class="form-label fw-bold text-dark small mb-1">
              <i class="fa-solid fa-file-excel text-success me-1"></i> Import Data OpenSID (`grid_agtkeluarga_4.xls`)
            </label>
            <input type="file" id="excelFile" class="form-control form-control-sm" accept=".xls,.xlsx" onchange="importExcel(event)">
            <div id="importStatus" class="form-text mt-1 text-muted small">Khusus Admin: Unggah data penduduk untuk pencarian instan.</div>
          </div>

          <!-- Form Pemohon -->
          <h6 class="fw-bold border-bottom pb-2 text-primary small"><i class="fa-solid fa-user me-2"></i>Data Pemohon</h6>
          
          <!-- KHUSUS ADMIN: Pencarian Instan Real-Time Autocomplete -->
          <div id="adminSearchBox" class="mb-3 position-relative">
            <label class="form-label small mb-1 fw-bold text-success"><i class="fa-solid fa-magnifying-glass me-1"></i>Pencarian Instan (Ketik NIK / Nama)</label>
            <input type="text" id="searchInput" class="form-control form-control-sm border-success" placeholder="Ketik NIK atau Nama warga..." oninput="handleSearchInput(this.value)" autocomplete="off">
            <div id="autocompleteList" class="autocomplete-items shadow-sm d-none"></div>
          </div>

          <div class="row g-2 mb-2">
            <div class="col-md-6">
              <label class="form-label small mb-1">NIK</label>
              <input type="text" id="nik" class="form-control form-control-sm" placeholder="6107..." oninput="updatePreview()">
            </div>
            <div class="col-md-6">
              <label class="form-label small mb-1">Nama Lengkap</label>
              <input type="text" id="nama" class="form-control form-control-sm" placeholder="Nama Pemohon" oninput="updatePreview()">
            </div>
          </div>

          <div class="row g-2 mb-2">
            <div class="col-md-6">
              <label class="form-label small mb-1">Tempat Lahir</label>
              <input type="text" id="tempatLahir" class="form-control form-control-sm" placeholder="Sungai Raya" oninput="updatePreview()">
            </div>
            <div class="col-md-6">
              <label class="form-label small mb-1">Tanggal Lahir</label>
              <input type="date" id="tglLahir" class="form-control form-control-sm" onchange="updatePreview()">
            </div>
          </div>

          <div class="row g-2 mb-3">
            <div class="col-md-6">
              <label class="form-label small mb-1">Pekerjaan</label>
              <input type="text" id="pekerjaan" class="form-control form-control-sm" placeholder="Petani / Swasta" oninput="updatePreview()">
            </div>
            <div class="col-md-6">
              <label class="form-label small mb-1">Alamat</label>
              <input type="text" id="alamat" class="form-control form-control-sm" value="Desa Sungai Jaga A" oninput="updatePreview()">
            </div>
          </div>

          <!-- Detail Surat & 30 Templat -->
          <h6 class="fw-bold border-bottom pb-2 text-primary small"><i class="fa-solid fa-file-lines me-2"></i>Detail Surat (30 Templat)</h6>
          
          <div class="mb-2">
            <label class="form-label small mb-1">Pilih Templat Surat</label>
            <select id="jenisSurat" class="form-select form-select-sm" onchange="updatePreview()">
              <optgroup label="Pemerintahan & Kependudukan">
                <option value="SURAT KETERANGAN PENGANTAR" selected>Surat Keterangan Pengantar</option>
                <option value="SURAT KETERANGAN BEDA NAMA">Surat Keterangan Beda Nama</option>
                <option value="SURAT KETERANGAN DOMISILI">Surat Keterangan Domisili</option>
                <option value="SURAT KETERANGAN DOMISILI USAHA">Surat Keterangan Domisili Usaha</option>
                <option value="SURAT KETERANGAN PINDAH">Surat Keterangan Pindah</option>
                <option value="SURAT KETERANGAN PENGANTAR KTP">Surat Keterangan Pengantar KTP</option>
                <option value="SURAT KETERANGAN KELAHIRAN">Surat Keterangan Kelahiran</option>
                <option value="SURAT KETERANGAN KEMATIAN">Surat Keterangan Kematian</option>
                <option value="SURAT KETERANGAN KEHILANGAN">Surat Keterangan Kehilangan</option>
                <option value="SURAT KETERANGAN KELAKUAN BAIK">Surat Keterangan Kelakuan Baik (SKCK)</option>
                <option value="SURAT KETERANGAN PENGHASILAN">Surat Keterangan Penghasilan</option>
                <option value="SURAT KETERANGAN TIDAK MAMPU">Surat Keterangan Tidak Mampu (SKTM)</option>
              </optgroup>
              <optgroup label="Pelayanan & Kesejahteraan">
                <option value="SURAT KETERANGAN UNTUK NIKAH">Surat Keterangan Untuk Nikah</option>
                <option value="SURAT KETERANGAN BELUM MENIKAH">Surat Keterangan Belum Menikah</option>
                <option value="SURAT KETERANGAN DUDA / JANDA">Surat Keterangan Duda / Janda</option>
                <option value="SURAT PENGANTAR BPJS">Surat Pengantar BPJS</option>
                <option value="SURAT KETERANGAN YATIM / PIATU">Surat Keterangan Yatim / Piatu</option>
                <option value="SURAT KETERANGAN BEASISWA">Surat Keterangan Beasiswa</option>
                <option value="SURAT KETERANGAN BEPERGIAN">Surat Keterangan Bepergian / Surat Jalan</option>
                <option value="SURAT IZIN KERAMAIAN">Surat Izin Keramaian</option>
              </optgroup>
              <optgroup label="Pertanahan & Ekonomi">
                <option value="SURAT PERNYATAAN TANAH">Surat Pernyataan Tanah</option>
                <option value="SURAT PENYERAHAN TANAH">Surat Penyerahan / Jual Beli Tanah</option>
                <option value="SURAT KETERANGAN TANAH TIDAK SENGKETA">Surat Keterangan Tanah Tidak Sengketa</option>
                <option value="SURAT KETERANGAN USAHA">Surat Keterangan Usaha (SKU)</option>
                <option value="SURAT KETERANGAN KEPEMILIKAN">Surat Keterangan Kepemilikan</option>
                <option value="SURAT PENGANTAR IMB / PBG">Surat Pengantar IMB / PBG</option>
              </optgroup>
              <optgroup label="Kelembagaan & Internal">
                <option value="SURAT TUGAS">Surat Tugas / SPPD</option>
                <option value="SURAT KEPUTUSAN KEPALA DESA">Surat Keputusan (SK) Kepala Desa</option>
                <option value="SURAT UNDANGAN">Surat Undangan Rapat</option>
                <option value="SURAT EDARAN">Surat Edaran Kepala Desa</option>
              </optgroup>
            </select>
          </div>

          <div class="row g-2 mb-2">
            <div class="col-md-12">
              <label class="form-label small mb-1">Nomor Surat</label>
              <input type="text" id="noSurat" class="form-control form-control-sm" value="140 / 042 / Pem-2026" oninput="updatePreview()">
            </div>
          </div>

          <div class="mb-3">
            <label class="form-label small mb-1">Maksud / Keperluan</label>
            <textarea id="keperluan" class="form-control form-control-sm" rows="3" oninput="updatePreview()">Bahwa nama tersebut di atas adalah benar-benar warga Desa Sungai Jaga A, Kecamatan Sungai Raya, Kabupaten Bengkayang dan bermaksud untuk mengurus keperluan Administrasi.</textarea>
          </div>

          <!-- Opsi TTD Penandatangan -->
          <div id="adminTtdBox">
            <h6 class="fw-bold border-bottom pb-2 text-primary small"><i class="fa-solid fa-signature me-2"></i>Opsi Penandatangan (TTD)</h6>
            
            <div class="mb-2">
              <label class="form-label small mb-1">Skema Penandatanganan</label>
              <select id="tipeTtd" class="form-select form-select-sm" onchange="toggleTtdFields()">
                <option value="kades">Kepala Desa (Langsung)</option>
                <option value="an" selected>an. Kepala Desa (Mandat Standar)</option>
                <option value="ub">an. Kepala Desa u.b. (Mandat Bertingkat)</option>
              </select>
            </div>

            <div class="p-2 border rounded bg-white mb-2" id="boxPenandatangan1">
              <div class="fw-bold small text-secondary mb-1" id="lblPejabat1">Pejabat Penandatangan / Pelaksana Mandat</div>
              <div class="row g-2">
                <div class="col-md-6">
                  <input type="text" id="jabatan1" class="form-control form-control-sm" value="Kasi Pelayanan dan Kesejahteraan" oninput="updatePreview()">
                </div>
                <div class="col-md-6">
                  <input type="text" id="nama1" class="form-control form-control-sm" value="JAKA SWARA" oninput="updatePreview()">
                </div>
              </div>
            </div>

            <div class="p-2 border rounded bg-white mb-3 d-none" id="boxPenandatangan2">
              <div class="fw-bold small text-secondary mb-1">Pejabat Atas Nama Pertama (an.)</div>
              <div class="row g-2">
                <div class="col-md-6">
                  <input type="text" id="jabatan2" class="form-control form-control-sm" value="Sekretaris Desa" oninput="updatePreview()">
                </div>
                <div class="col-md-6">
                  <input type="text" id="nama2" class="form-control form-control-sm" value="NAMA SEKDES" oninput="updatePreview()">
                </div>
              </div>
            </div>
          </div>

        </div>
      </div>
    </div>

    <!-- Live Preview Kertas Surat A4 -->
    <div class="col-lg-7 col-md-12">
      <div class="paper-container">
        <div class="paper" id="paperPrint">
          
          <!-- Kop Surat Resmi Desa -->
          <div class="header-kop">
            <h3>PEMERINTAH KABUPATEN BENGKAYANG</h3>
            <h3>KECAMATAN SUNGAI RAYA</h3>
            <h2>PEMERINTAH DESA SUNGAI JAGA A</h2>
            <p>Jalan Raya Sungai Jaga A, Kec. Sungai Raya, Kab. Bengkayang, Kalimantan Barat 79252</p>
          </div>

          <!-- Judul Dokumen -->
          <div class="surat-title">
            <h4 id="prevJenisSurat">SURAT KETERANGAN PENGANTAR</h4>
            <p>Nomor: <span id="prevNoSurat">140 / 042 / Pem-2026</span></p>
          </div>

          <!-- Paragraf Pembuka -->
          <p>Yang bertanda tangan di bawah ini Pemerintah Desa Sungai Jaga A, Kecamatan Sungai Raya, Kabupaten Bengkayang, menerangkan dengan sebenarnya bahwa:</p>

          <!-- Tabel Data Pemohon -->
          <table class="table-surat">
            <tr>
              <td class="label-col">Nama Lengkap</td>
              <td class="colon-col">:</td>
              <td><b id="prevNama">[ NAMA PEMOHON ]</b></td>
            </tr>
            <tr>
              <td class="label-col">NIK</td>
              <td class="colon-col">:</td>
              <td id="prevNik">[ NIK PEMOHON ]</td>
            </tr>
            <tr>
              <td class="label-col">Tempat, Tgl. Lahir</td>
              <td class="colon-col">:</td>
              <td id="prevTtl">[ Tempat, Tanggal Lahir ]</td>
            </tr>
            <tr>
              <td class="label-col">Pekerjaan</td>
              <td class="colon-col">:</td>
              <td id="prevPekerjaan">[ Pekerjaan ]</td>
            </tr>
            <tr>
              <td class="label-col">Alamat</td>
              <td class="colon-col">:</td>
              <td id="prevAlamat">Desa Sungai Jaga A</td>
            </tr>
          </table>

          <!-- Keperluan -->
          <p id="prevKeperluan">Bahwa nama tersebut di atas adalah benar-benar warga Desa Sungai Jaga A, Kecamatan Sungai Raya, Kabupaten Bengkayang dan bermaksud untuk mengurus keperluan Administrasi.</p>

          <!-- Penutup -->
          <p>Demikian Surat Keterangan ini dibuat dengan sebenarnya agar dapat dipergunakan sebagaimana mestinya.</p>

          <!-- TTD Box -->
          <div class="ttd-container">
            <div class="ttd-box">
              <p class="mb-0">Sungai Jaga A, <span id="prevTanggalCetak"></span></p>
              
              <div id="prevBlokTtd">
                <!-- Injected via JS -->
              </div>

              <div class="ttd-space"></div>
              
              <p class="mb-0"><u class="fw-bold" id="prevNamaTtd">JAKA SWARA</u></p>
            </div>
          </div>

        </div>
      </div>
    </div>
  </div>
</div>

<script>
  let dataPendudukList = [];

  window.onload = function() {
    const today = new Date();
    const options = { day: 'numeric', month: 'long', year: 'numeric' };
    document.getElementById('prevTanggalCetak').innerText = today.toLocaleDateString('id-ID', options);
    
    switchRole();
    toggleTtdFields();
    updatePreview();
  };

  // Switch Role Admin vs User
  function switchRole() {
    const role = document.getElementById('roleSelect').value;
    const adminImportBox = document.getElementById('adminImportBox');
    const adminSearchBox = document.getElementById('adminSearchBox');
    const adminTtdBox = document.getElementById('adminTtdBox');
    const panelTitle = document.getElementById('panelTitle');

    if (role === 'admin') {
      adminImportBox.classList.remove('d-none');
      adminSearchBox.classList.remove('d-none');
      adminTtdBox.classList.remove('d-none');
      panelTitle.innerHTML = `<i class="fa-solid fa-sliders me-2"></i>Panel Kontrol Admin`;
    } else {
      // Sembunyikan Impor Excel & Search Penduduk dari User/Warga
      adminImportBox.classList.add('d-none');
      adminSearchBox.classList.add('d-none');
      adminTtdBox.classList.add('d-none');
      panelTitle.innerHTML = `<i class="fa-solid fa-paper-plane me-2"></i>Form Pengajuan Warga`;
    }
  }

  // TTD Switcher
  function toggleTtdFields() {
    const tipe = document.getElementById('tipeTtd').value;
    const box2 = document.getElementById('boxPenandatangan2');
    const lbl1 = document.getElementById('lblPejabat1');

    if (tipe === 'ub') {
      box2.classList.remove('d-none');
      lbl1.innerText = "Pejabat Eksekutor Penandatangan (u.b.)";
    } else {
      box2.classList.add('d-none');
      lbl1.innerText = "Pejabat Penandatangan / Pelaksana Mandat";
    }
    updatePreview();
  }

  // Live Update Preview Surat
  function updatePreview() {
    document.getElementById('prevJenisSurat').innerText = document.getElementById('jenisSurat').value;
    document.getElementById('prevNoSurat').innerText = document.getElementById('noSurat').value || '-';
    document.getElementById('prevNama').innerText = (document.getElementById('nama').value || '[ NAMA PEM