document.addEventListener('DOMContentLoaded', async () => {
  const tbody = document.querySelector('#attendanceTable tbody');
  const filterSelect = document.getElementById('kelasFilter');
  const jumlahDataEl = document.createElement('div');
  const bersihkanBtn = document.createElement('button');
  const hapusTerpilihBtn = document.createElement('button');
  const downloadCsvBtn = document.createElement('button');
  const downloadExcelBtn = document.createElement('button');

  jumlahDataEl.className = 'jumlah-data';
  bersihkanBtn.textContent = 'Bersihkan Data Lama (>30 Hari)';
  bersihkanBtn.className = 'clean-btn';
  hapusTerpilihBtn.textContent = 'Hapus Data Terpilih';
  hapusTerpilihBtn.className = 'hapus-btn';
  downloadCsvBtn.textContent = 'Download CSV';
  downloadExcelBtn.textContent = 'Download Excel';

  const tableSection = document.querySelector('section');
  tableSection.appendChild(jumlahDataEl);
  tableSection.appendChild(bersihkanBtn);
  tableSection.appendChild(hapusTerpilihBtn);
  tableSection.appendChild(downloadCsvBtn);
  tableSection.appendChild(downloadExcelBtn);

  let semuaData = [];
  let dataTerfilter = []; // Memperbaiki variabel yang tidak didefinisikan

  function updateJumlahData(count) {
    jumlahDataEl.innerText = `Menampilkan ${count} data presensi`;
  }

  function renderTabel(data) {
    tbody.innerHTML = '';
    data.forEach((record, index) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><input type="checkbox" class="checkbox-hapus" data-index="${index}"></td>
        <td>${record.nama}</td>
        <td>${record.kelas}</td>
        <td>${record.status}</td>
        <td>${record.tanggal} (${record.hari})<br>${record.waktu}</td>
        <td><img src="${record.foto}" width="100"></td>
      `;
      tbody.appendChild(tr);
    });
    updateJumlahData(data.length);
  }

  function filterDataKelas() {
    const kelasDipilih = filterSelect.value;
    if (kelasDipilih === 'semua') {
      dataTerfilter = semuaData;
    } else {
      dataTerfilter = semuaData.filter(d => d.kelas === kelasDipilih);
    }
    renderTabel(dataTerfilter);

    // (Opsional) tampilkan nama kelas aktif di halaman
    const label = document.getElementById('kelasAktifLabel');
    if (label) {
      label.textContent = (kelasDipilih === 'semua') ? 'Semua Kelas' : `Kelas: ${kelasDipilih}`;
    }
  }

  function bersihkanDataLama(data) {
    const sekarang = new Date();
    const dataAwal = data.length;
    const dataBaru = data.filter(record => {
      const recordDateTime = new Date(`${record.tanggal} ${record.waktu}`);
      const selisihHari = (sekarang - recordDateTime) / (1000 * 60 * 60 * 24);
      return selisihHari <= 30;
    });

    const jumlahDihapus = dataAwal - dataBaru.length;
    if (jumlahDihapus > 0) {
      alert(`${jumlahDihapus} data lama berhasil dihapus.`);
    } else {
      alert('Tidak ada data yang perlu dihapus.');
    }
    return dataBaru;
  }

  async function loadData() {
    try {
      const res = await fetch('http://localhost:5000/api/absen');
      const data = await res.json();
      semuaData = data;
      dataTerfilter = data; // Inisialisasi data terfilter
      renderTabel(semuaData);
    } catch (e) {
      console.error('Gagal load data', e);
      alert('Gagal mengambil data dari server!');
    }
  }

  // Event listeners
  filterSelect.addEventListener('change', filterDataKelas);

  bersihkanBtn.addEventListener('click', () => {
    semuaData = bersihkanDataLama(semuaData);
    dataTerfilter = semuaData; // Update data terfilter juga
    renderTabel(semuaData);
  });

  hapusTerpilihBtn.addEventListener('click', () => {
    const checkboxList = document.querySelectorAll('.checkbox-hapus');
    const indexTerpilih = [];

    checkboxList.forEach(cb => {
      if (cb.checked) {
        indexTerpilih.push(parseInt(cb.getAttribute('data-index')));
      }
    });

    if (indexTerpilih.length === 0) {
      alert('Pilih dulu data yang mau dihapus.');
      return;
    }

    if (!confirm(`Yakin mau hapus ${indexTerpilih.length} data terpilih?`)) return;

    semuaData = semuaData.filter((_, idx) => !indexTerpilih.includes(idx));
    filterDataKelas(); // Filter ulang data setelah menghapus
  });

  downloadCsvBtn.addEventListener('click', () => {
    const csvContent = "data:text/csv;charset=utf-8," +
      ["Nama,Kelas,Status,Tanggal,Hari,Waktu"].join(",") + "\n" +
      dataTerfilter.map(e => `${e.nama},${e.kelas},${e.status},${e.tanggal},${e.hari},${e.waktu}`).join("\n");

    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    const kelasNama = filterSelect.value.replace(/\s+/g, "_");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `presensi_${kelasNama}.csv`);
    document.body.appendChild(link);
    link.click();
  });

  downloadExcelBtn.addEventListener('click', () => {
    // Tanggal untuk nama file
    const today = new Date();
    const dateStr = today.toLocaleDateString('id-ID').replace(/\//g, '-');
    
    // Judul laporan dan info kelas
    const kelasDipilih = filterSelect.value;
    const kelasNama = kelasDipilih === 'semua' ? 'Semua_Kelas' : kelasDipilih.replace(/\s+/g, "_");
    const reportTitle = `LAPORAN PRESENSI SISWA ${kelasDipilih === 'semua' ? 'SEMUA KELAS' : kelasDipilih.toUpperCase()}`;
    
    // Data untuk Excel
    const header = ['No', 'Nama', 'Kelas', 'Status', 'Tanggal', 'Hari', 'Waktu'];
    const data = dataTerfilter.map((e, idx) => [idx + 1, e.nama, e.kelas, e.status, e.tanggal, e.hari, e.waktu]);
    
    // Buat workbook dan worksheet
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([]);
    
    // Tambahkan judul dan informasi laporan
    XLSX.utils.sheet_add_aoa(ws, [[reportTitle]], { origin: "A1" });
    XLSX.utils.sheet_add_aoa(ws, [[`Tanggal Export: ${today.toLocaleDateString('id-ID')} ${today.toLocaleTimeString('id-ID')}`]], { origin: "A2" });
    XLSX.utils.sheet_add_aoa(ws, [[`Total Data: ${dataTerfilter.length} presensi`]], { origin: "A3" });
    XLSX.utils.sheet_add_aoa(ws, [[""]], { origin: "A4" }); // Baris kosong
    
    // Tambahkan header dan data utama
    XLSX.utils.sheet_add_aoa(ws, [header], { origin: "A5" });
    XLSX.utils.sheet_add_aoa(ws, data, { origin: "A6" });
    
    // Atur lebar kolom
    const columnWidths = [
      { wch: 5 },  // No
      { wch: 25 }, // Nama
      { wch: 15 }, // Kelas
      { wch: 12 }, // Status
      { wch: 12 }, // Tanggal
      { wch: 10 }, // Hari
      { wch: 10 }  // Waktu
    ];
    ws['!cols'] = columnWidths;
    
    // Definisikan styles untuk workbook
    // 1. Tema warna kustom
    const colorTheme = {
      headerBg: "1F4E78",     // Biru tua untuk header
      headerText: "FFFFFF",   // Putih untuk teks header
      titleBg: "D5E3F0",      // Background judul light blue
      altRowBg: "EBF1F8",     // Baris bergantian light blue
      borderColor: "4472C4",  // Warna border biru
      tableBorder: "000000"   // Warna border tabel hitam
    };
    
    // 2. Style untuk sel-sel berbeda
    const styles = {
      title: { 
        font: { bold: true, sz: 16, color: { rgb: "000000" } },
        fill: { fgColor: { rgb: colorTheme.titleBg } },
        alignment: { horizontal: "center", vertical: "center" },
        border: {
          top: { style: "medium", color: { rgb: colorTheme.tableBorder } },
          left: { style: "medium", color: { rgb: colorTheme.tableBorder } },
          right: { style: "medium", color: { rgb: colorTheme.tableBorder } },
          bottom: { style: "thin", color: { rgb: colorTheme.tableBorder } }
        }
      },
      info: { 
        font: { sz: 11 },
        alignment: { horizontal: "left" },
        border: {
          left: { style: "medium", color: { rgb: colorTheme.tableBorder } },
          right: { style: "medium", color: { rgb: colorTheme.tableBorder } }
        }
      },
      lastInfo: {
        font: { sz: 11 },
        alignment: { horizontal: "left" },
        border: {
          left: { style: "medium", color: { rgb: colorTheme.tableBorder } },
          right: { style: "medium", color: { rgb: colorTheme.tableBorder } },
          bottom: { style: "thin", color: { rgb: colorTheme.tableBorder } }
        }
      },
      header: { 
        font: { bold: true, color: { rgb: colorTheme.headerText }, sz: 12 }, 
        fill: { fgColor: { rgb: colorTheme.headerBg } },
        alignment: { horizontal: "center", vertical: "center", wrapText: true },
        border: {
          top: { style: "medium", color: { rgb: colorTheme.tableBorder } },
          left: { style: "thin", color: { rgb: colorTheme.tableBorder } },
          right: { style: "thin", color: { rgb: colorTheme.tableBorder } },
          bottom: { style: "medium", color: { rgb: colorTheme.tableBorder } }
        }
      },
      cell: {
        alignment: { vertical: "center" },
        border: {
          left: { style: "thin", color: { rgb: colorTheme.tableBorder } },
          right: { style: "thin", color: { rgb: colorTheme.tableBorder } },
          bottom: { style: "thin", color: { rgb: colorTheme.tableBorder } }
        }
      },
      cellAlt: {
        alignment: { vertical: "center" },
        fill: { fgColor: { rgb: colorTheme.altRowBg } },
        border: {
          left: { style: "thin", color: { rgb: colorTheme.tableBorder } },
          right: { style: "thin", color: { rgb: colorTheme.tableBorder } },
          bottom: { style: "thin", color: { rgb: colorTheme.tableBorder } }
        }
      },
      lastRow: {
        alignment: { vertical: "center" },
        border: {
          left: { style: "thin", color: { rgb: colorTheme.tableBorder } },
          right: { style: "thin", color: { rgb: colorTheme.tableBorder } },
          bottom: { style: "medium", color: { rgb: colorTheme.tableBorder } }
        }
      },
      lastRowAlt: {
        alignment: { vertical: "center" },
        fill: { fgColor: { rgb: colorTheme.altRowBg } },
        border: {
          left: { style: "thin", color: { rgb: colorTheme.tableBorder } },
          right: { style: "thin", color: { rgb: colorTheme.tableBorder } },
          bottom: { style: "medium", color: { rgb: colorTheme.tableBorder } }
        }
      }
    };
    
    // Terapkan style untuk title dan info
    const headerRange = XLSX.utils.decode_range(ws['!ref']);
    const lastDataRow = 5 + data.length;
    
    // Style untuk judul & informasi
    for (let C = 0; C <= 6; ++C) {
      // Judul 
      const titleAddress = XLSX.utils.encode_cell({ r: 0, c: C });
      if (ws[titleAddress]) ws[titleAddress].s = styles.title;
      
      // Info baris 2
      const info1Address = XLSX.utils.encode_cell({ r: 1, c: C });
      if (ws[info1Address]) ws[info1Address].s = styles.info;
      
      // Info baris 3 (terakhir)
      const info2Address = XLSX.utils.encode_cell({ r: 2, c: C });
      if (ws[info2Address]) ws[info2Address].s = styles.lastInfo;
    }
    
    // Style untuk baris header dan data
    for (let C = 0; C <= headerRange.e.c; ++C) {
      // Header
      const headerAddress = XLSX.utils.encode_cell({ r: 4, c: C });
      if (ws[headerAddress]) ws[headerAddress].s = styles.header;
      
      // Data rows
      for (let R = 5; R <= headerRange.e.r; ++R) {
        const isLastRow = (R === headerRange.e.r);
        const isAltRow = R % 2 === 0;
        const dataAddress = XLSX.utils.encode_cell({ r: R, c: C });
        
        // Pilih style berdasarkan posisi dan jenis baris
        let cellStyle;
        if (isLastRow) {
          cellStyle = isAltRow ? styles.lastRowAlt : styles.lastRow;
        } else {
          cellStyle = isAltRow ? styles.cellAlt : styles.cell;
        }
        
        // Terapkan style dan spesifik alignment untuk kolom-kolom tertentu
        if (ws[dataAddress]) {
          ws[dataAddress].s = { ...cellStyle };
          
          // Center alignment untuk kolom No, Status, Hari
          if (C === 0 || C === 3 || C === 5) {
            ws[dataAddress].s.alignment = { ...ws[dataAddress].s.alignment, horizontal: "center" };
          } 
          // Left alignment untuk nama dan kelas
          else if (C === 1 || C === 2) {
            ws[dataAddress].s.alignment = { ...ws[dataAddress].s.alignment, horizontal: "left" };
          }
          // Center alignment untuk tanggal dan waktu
          else {
            ws[dataAddress].s.alignment = { ...ws[dataAddress].s.alignment, horizontal: "center" };
          }
        }
      }
    }
    
    // Gabungkan cell untuk judul
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }, // Merge judul (A1:G1)
    ];
    
    // Tambahkan worksheet ke workbook dan setting tabel resmi Excel
    XLSX.utils.book_append_sheet(wb, ws, 'Presensi');
    
    // Menambahkan format tabel Excel yang resmi (tidak didukung di SheetJS core)
    // Kita mengaturnya dengan styling manual
    
    // Simpan file
    XLSX.writeFile(wb, `presensi_${kelasNama}_${dateStr}.xlsx`);
  });

  await loadData(); // load data awal

  // Logout otomatis setelah 30 menit
  setTimeout(() => {
    alert('Sesi Anda habis. Silakan login ulang.');
    localStorage.removeItem('isLoggedIn');
    window.location.href = 'login.html';
  }, 30 * 60 * 1000); // 30 menit
});