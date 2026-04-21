// Ambil elemen dari HTML
const btnGenerate = document.getElementById('generateLaporan');
const fileInput = document.getElementById('uploadSipp');
const statusMsg = document.getElementById('statusMessage');

// Fungsi pembantu untuk nampilin status (sukses/loading/error)
function showStatus(message, type) {
    statusMsg.textContent = message;
    statusMsg.classList.remove('hidden', 'bg-emerald-100', 'text-emerald-800', 'bg-red-100', 'text-red-800', 'bg-blue-100', 'text-blue-800');
    
    if (type === 'success') statusMsg.classList.add('bg-emerald-100', 'text-emerald-800');
    else if (type === 'error') statusMsg.classList.add('bg-red-100', 'text-red-800');
    else statusMsg.classList.add('bg-blue-100', 'text-blue-800');
}

// Event listener saat tombol Generate diklik
btnGenerate.addEventListener('click', async () => {
    // 1. Cek apakah ada file yang diupload
    if (fileInput.files.length === 0) {
        showStatus('Bro, upload dulu file mentahan SIPP-nya!', 'error');
        return;
    }

    const file = fileInput.files[0];
    showStatus('Sedang memproses data SIPP...', 'loading');

    try {
        // 2. Baca file Excel mentah pakai FileReader dan ExcelJS
        const arrayBuffer = await file.arrayBuffer();
        const rawWorkbook = new ExcelJS.Workbook();
        await rawWorkbook.xlsx.load(arrayBuffer);

        // Ambil sheet pertama dari file SIPP
        const rawSheet = rawWorkbook.worksheets[0]; 

        // 3. Buat Workbook & Worksheet baru untuk Laporan PRO-TAMA
        const newWorkbook = new ExcelJS.Workbook();
        const newSheet = newWorkbook.addWorksheet('Laporan Bulanan');

        // Setup Header Laporan (Bisa lo modif besok sesuai format PTA)
        newSheet.columns = [
            { header: 'No', key: 'no', width: 5 },
            { header: 'Nomor Perkara', key: 'nomor_perkara', width: 25 },
            { header: 'Susunan Majelis / Hakim', key: 'hakim', width: 40 },
            { header: 'Status Perkara', key: 'status', width: 20 },
            { header: 'Keterangan', key: 'ket', width: 20 }
        ];

        // 4. Looping data dari file SIPP mentah
        // Asumsi baris 1 adalah header SIPP, jadi kita mulai dari baris 2
        let rowIndex = 1;
        rawSheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) { 
                
                // --- BAGIAN INI YANG BESOK LO SESUAIKAN DENGAN SIPP ---
                // Asumsi: 
                // Kolom B (2) = Nomor Perkara
                // Kolom C (3) = Hakim Ketua / Hakim Tunggal
                // Kolom D (4) = Hakim Anggota 1
                // Kolom E (5) = Hakim Anggota 2
                
                let noPerkara = row.getCell(2).value;
                let hakimKetua = row.getCell(3).value;
                let hakimAnggota1 = row.getCell(4).value;
                let hakimAnggota2 = row.getCell(5).value;
                
                // --- LOGIKA SMART FORMATTING (TUNGGAL VS MAJELIS) ---
                let susunanHakim = "";
                
                // Cek jika Hakim Anggota 1 kosong (Berarti Hakim Tunggal seperti Dispensasi Kawin)
                if (!hakimAnggota1 || hakimAnggota1 === "") {
                    susunanHakim = "Hakim Tunggal: " + hakimKetua;
                } else {
                    // Jika ada anggota, format jadi Majelis
                    susunanHakim = "Ketua: " + hakimKetua + "\n" + 
                                   "Anggota 1: " + hakimAnggota1 + "\n" + 
                                   "Anggota 2: " + hakimAnggota2;
                }

                // Masukkan data yang sudah dirapikan ke sheet baru
                newSheet.addRow({
                    no: rowIndex,
                    nomor_perkara: noPerkara,
                    hakim: susunanHakim,
                    status: 'Belum Putus', // Besok bisa ditarik dari rumus/kolom SIPP
                    ket: '-'
                });
                rowIndex++;
            }
        });

        // 5. Styling & Rapihin Tabel (Biar elegan)
        newSheet.eachRow((row, rowNumber) => {
            row.eachCell((cell) => {
                // Kasih border tiap cell
                cell.border = {
                    top: {style:'thin'}, left: {style:'thin'},
                    bottom: {style:'thin'}, right: {style:'thin'}
                };
                
                // Alignment biar teks rapi di tengah/atas, apalagi ada text wrap buat Majelis
                cell.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
                
                // Header (Baris 1) dibikin tebal dan rata tengah
                if (rowNumber === 1) {
                    cell.font = { bold: true };
                    cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFD3D3D3' } // Warna abu-abu elegan untuk header
                    };
                }
            });
        });

        // 6. Export dan Download File Jadinya
        const buffer = await newWorkbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = 'Laporan_PRO-TAMA_SIPP.xlsx';
        a.click();
        
        window.URL.revokeObjectURL(url);
        showStatus('Laporan berhasil di-generate! Cek folder Download.', 'success');

    } catch (error) {
        console.error("Error processing Excel: ", error);
        showStatus('Gagal memproses file. Pastikan formatnya benar.', 'error');
    }
});
