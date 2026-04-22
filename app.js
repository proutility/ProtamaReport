// =========================================================================
// PRO-TAMA REPORT (AUTOMASI SIPP) - APP.JS (V3 - FINAL JELAS MERGE 3 BARIS)
// =========================================================================

const btnGenerate = document.getElementById('generateLaporan');
const fileInput = document.getElementById('uploadSipp');
const statusMsg = document.getElementById('statusMessage');

function showStatus(message, type) {
    statusMsg.textContent = message;
    statusMsg.classList.remove('hidden', 'bg-emerald-100', 'text-emerald-800', 'bg-red-100', 'text-red-800', 'bg-blue-100', 'text-blue-800');
    if (type === 'success') statusMsg.classList.add('bg-emerald-100', 'text-emerald-800');
    else if (type === 'error') statusMsg.classList.add('bg-red-100', 'text-red-800');
    else statusMsg.classList.add('bg-blue-100', 'text-blue-800');
}

function safeText(cell) {
    if (!cell || cell.value === null || cell.value === undefined) return "";
    if (typeof cell.value === 'object') {
        if (cell.value.richText) return cell.value.richText.map(rt => rt.text).join('').trim();
        if (cell.value.result !== undefined) return String(cell.value.result).trim();
    }
    return String(cell.value).trim();
}

function getRawDateStr(cell) {
    if (!cell || !cell.value) return "";
    let val = cell.value;
    if (typeof val === 'object' && val.result !== undefined) val = val.result;
    
    if (val instanceof Date) {
        const d = val.getDate().toString().padStart(2, '0');
        const m = (val.getMonth() + 1).toString().padStart(2, '0');
        const y = val.getFullYear();
        return `${d}/${m}/${y}`;
    }
    return String(val).trim();
}

btnGenerate.addEventListener('click', async () => {
    if (fileInput.files.length === 0) {
        showStatus('Pilih file SIPP dulu bro!', 'error');
        return;
    }

    const file = fileInput.files[0];
    showStatus('Sedang merakit LIPA 1 format 3 baris...', 'loading');

    try {
        const arrayBuffer = await file.arrayBuffer();
        const rawWorkbook = new ExcelJS.Workbook();
        await rawWorkbook.xlsx.load(arrayBuffer);
        const rawSheet = rawWorkbook.worksheets[0]; 

        const newWorkbook = new ExcelJS.Workbook();
        const newSheet = newWorkbook.addWorksheet('LIPA 1');

        // Setup Header (Kolom 14 Standar LIPA 1)
        newSheet.columns = [
            { header: 'No', width: 5 }, { header: 'No Perkara', width: 25 },
            { header: 'Tgl Penerimaan', width: 15 }, { header: 'Nama Majelis Hakim', width: 40 },
            { header: 'Tgl PMH', width: 12 }, { header: 'Tgl PHS', width: 12 },
            { header: 'Tgl Sidang I', width: 12 }, { header: 'Tgl Diputus', width: 12 },
            { header: 'Jenis Putusan', width: 15 }, { header: 'Sisa Belum Dibagi', width: 18 },
            { header: 'Sisa Belum Diputus', width: 25 }, { header: 'Keterangan', width: 12 }
        ];

        let currentRow = 8; // Data mulai di baris 8 (setelah Kop & Header)
        let sisaLalu_G = 0, tambahIni_G = 0, putusIni_G = 0;
        const bulanLaporan = "3"; // Hardcode Maret

        // Membaca data SIPP asli
        rawSheet.eachRow((row, rowNumber) => {
            // Asumsi No Perkara SIPP di Kolom B (CELL 2)
            let strNoPerkara = safeText(row.getCell(2)); 

            if (strNoPerkara !== "" && strNoPerkara.includes("PA.")) {
                
                let startMerge = currentRow;
                let endMerge = currentRow + 2; // Setiap perkara dikasih jatah 3 baris

                // --- 1. TULIS DATA UTAMA KE BARIS PERTAMA DARI JATAH 3 BARIS ---
                // Data ini nanti bakal di-merge vertikal
                newSheet.getRow(startMerge).values = [
                    (currentRow - 8) / 3 + 1, // Penomoran otomatis No 1, 2, 3...
                    strNoPerkara, // No Perkara
                    getRawDateStr(row.getCell(6)), // Tgl Masuk (misal Kolom F)
                    safeText(row.getCell(3)), // BARIS 1 NAMA HAKIM: Hakim Ketua (misal Kolom C)
                    getRawDateStr(row.getCell(7)), // Tgl PMH (Kolom G)
                    getRawDateStr(row.getCell(8)), // Tgl PHS (Kolom H)
                    getRawDateStr(row.getCell(9)), // Tgl Sidang I (Kolom I)
                    getRawDateStr(row.getCell(10)), // Tgl Putus (Kolom J)
                    safeText(row.getCell(11)), // Jenis Putusan (Kolom K)
                    "", // Belum Dibagi (Kosong)
                    (getRawDateStr(row.getCell(10)) === "" ? strNoPerkara : ""), // Sisa Belum Diputus
                    "-" // Ket
                ];

                // --- 2. TULIS NAMA HAKIM ANGGOTA DI BARIS 2 & 3 ---
                // Kita cuma nulis di kolom ke-4 (Kolom D: Nama Majelis Hakim)
                newSheet.getCell(`D${startMerge + 1}`).value = safeText(row.getCell(4)); // Baris 2: Anggota 1 (misal Kolom D SIPP)
                newSheet.getCell(`D${startMerge + 2}`).value = safeText(row.getCell(5)); // Baris 3: Anggota 2 (misal Kolom E SIPP)

                // --- 3. LOGIKA MERGE VERTIKAL 3 BARIS ---
                // Kita gabung sel setinggi 3 baris untuk semua kolom KECUALI Kolom Nama Hakim
                // Kolom indices: 1(No), 2(No Perkara), 3(Tgl Masuk), 5(Tgl PMH)... sampai 12(Ket)
                const columnsToMerge = [1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12];
                
                columnsToMerge.forEach(colIndex => {
                    newSheet.mergeCells(startMerge, colIndex, endMerge, colIndex);
                });

                // --- 4. STYLING BIAR MIRIP FOTO 2 ---
                // Kita kasih border ke semua 3 baris dan bikin rata tengah-tengah
                for(let r = startMerge; r <= endMerge; r++) {
                    for(let c = 1; c <= 12; c++) {
                        const cell = newSheet.getCell(r, c);
                        cell.border = {
                            top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
                        };
                        // Biar data di sel yang di-merge posisinya ada di tengah kotak
                        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                    }
                }
                
                // Khusus kolom Nama Hakim kita bikin rata Kiri-Tengah biar rapi dibaca
                newSheet.getCell(`D${startMerge}`).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
                newSheet.getCell(`D${startMerge+1}`).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
                newSheet.getCell(`D${startMerge+2}`).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

                currentRow += 3; // Lompat ke baris perkara berikutnya (karena tadi udah pake 3 baris)
            }
        });

        // --- 5. DOWNLOAD FINISHED FILE ---
        const buffer = await newWorkbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'LIPA_1_SERANG_FORMAT_ASLI.xlsx';
        a.click();
        
        showStatus('Eksekusi Selesai Bro! Formatnya plekkketiplek sama Foto 2!', 'success');

    } catch (error) {
        console.error("Error: ", error);
        showStatus('Gagal memproses. Pastikan file SIPP lo udah di "Save As .xlsx" ya bro!', 'error');
    }
});
