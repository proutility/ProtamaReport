// =========================================================================
// PRO-TAMA REPORT (AUTOMASI SIPP) - APP.JS (V4 - FORMAT LIPA 1 100% ASLI)
// Develop by: Ilham Nur Pratama (PA Serang)
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

// Fungsi nyuci teks biasa
function safeText(cell) {
    if (!cell || cell.value === null || cell.value === undefined) return "";
    if (typeof cell.value === 'object') {
        if (cell.value.richText) return cell.value.richText.map(rt => rt.text).join('').trim();
        if (cell.value.result !== undefined) return String(cell.value.result).trim();
    }
    return String(cell.value).trim();
}

// Fungsi nyuci Tanggal biar nggak ada yang hilang/kosong
function getRawDateStr(cell) {
    if (!cell || !cell.value) return "";
    let val = cell.value;
    if (typeof val === 'object' && val.result !== undefined) val = val.result;
    
    // Kalau dia format Date asli dari Excel
    if (val instanceof Date) {
        const d = val.getDate().toString().padStart(2, '0');
        const m = (val.getMonth() + 1).toString().padStart(2, '0');
        const y = val.getFullYear();
        return `${d}/${m}/${y}`;
    }
    // Kalau dia format teks (string) dari SIPP, kembalikan apa adanya
    return String(val).trim();
}

// Fungsi ngambil bulan buat ngitung Sisa (Rekap Dapur)
function getMonthFromCell(cell) {
    if (!cell || !cell.value) return "0";
    let val = cell.value;
    if (typeof val === 'object' && val.result !== undefined) val = val.result;
    
    if (val instanceof Date) return String(val.getMonth() + 1);
    if (typeof val === 'string') {
        let parts = val.split('/');
        if (parts.length === 3) return String(parseInt(parts[1], 10)); 
        let partsDash = val.split('-');
        if (partsDash.length === 3) return String(parseInt(partsDash[1], 10)); 
    }
    return "0";
}

btnGenerate.addEventListener('click', async () => {
    if (fileInput.files.length === 0) {
        showStatus('Pilih file SIPP dulu bro!', 'error');
        return;
    }

    const file = fileInput.files[0];
    showStatus('Sedang merakit LIPA 1 Format Resmi...', 'loading');

    try {
        const arrayBuffer = await file.arrayBuffer();
        const rawWorkbook = new ExcelJS.Workbook();
        await rawWorkbook.xlsx.load(arrayBuffer);
        const rawSheet = rawWorkbook.worksheets[0]; 

        const newWorkbook = new ExcelJS.Workbook();
        const newSheet = newWorkbook.addWorksheet('LIPA 1');

        // --- 1. SETUP LEBAR 14 KOLOM ---
        newSheet.getColumn(1).width = 5;   // A: No
        newSheet.getColumn(2).width = 25;  // B: No Perkara
        newSheet.getColumn(3).width = 15;  // C: Kode Perkara
        newSheet.getColumn(4).width = 35;  // D: Nama Hakim
        newSheet.getColumn(5).width = 25;  // E: Nama PP
        newSheet.getColumn(6).width = 12;  // F: Penerimaan
        newSheet.getColumn(7).width = 12;  // G: PMH
        newSheet.getColumn(8).width = 12;  // H: PHS
        newSheet.getColumn(9).width = 12;  // I: Sidang I
        newSheet.getColumn(10).width = 12; // J: Diputus
        newSheet.getColumn(11).width = 15; // K: Jenis Putusan
        newSheet.getColumn(12).width = 12; // L: Belum Dibagi
        newSheet.getColumn(13).width = 15; // M: Belum Diputus
        newSheet.getColumn(14).width = 10; // N: Ket

        // --- 2. BIKIN KOP SURAT (BARIS 1-3) ---
        newSheet.mergeCells('A1:N1'); newSheet.getCell('A1').value = "LAPORAN KEADAAN PERKARA";
        newSheet.mergeCells('A2:N2'); newSheet.getCell('A2').value = "PENGADILAN AGAMA SERANG";
        newSheet.mergeCells('A3:N3'); newSheet.getCell('A3').value = "BULAN MARET 2026"; // Bisa dibikin dinamis nanti
        for(let r=1; r<=3; r++) {
            newSheet.getCell(`A${r}`).font = { bold: true, size: 12 };
            newSheet.getCell(`A${r}`).alignment = { horizontal: 'center' };
        }

        // --- 3. BIKIN HEADER TABEL RUMIT (BARIS 5-7) ---
        newSheet.mergeCells('A5:A6'); newSheet.getCell('A5').value = "No";
        newSheet.mergeCells('B5:B6'); newSheet.getCell('B5').value = "No Perkara";
        newSheet.mergeCells('C5:C6'); newSheet.getCell('C5').value = "Kode Perkara";
        newSheet.mergeCells('D5:D6'); newSheet.getCell('D5').value = "Nama Majelis Hakim";
        newSheet.mergeCells('E5:E6'); newSheet.getCell('E5').value = "Nama PP";
        newSheet.mergeCells('F5:J5'); newSheet.getCell('F5').value = "Tanggal";
        newSheet.getCell('F6').value = "Penerimaan";
        newSheet.getCell('G6').value = "PMH";
        newSheet.getCell('H6').value = "PHS";
        newSheet.getCell('I6').value = "Sidang I";
        newSheet.getCell('J6').value = "Diputus";
        newSheet.mergeCells('K5:K6'); newSheet.getCell('K5').value = "Jenis Putusan";
        newSheet.mergeCells('L5:M5'); newSheet.getCell('L5').value = "Sisa Akhir Bulan";
        newSheet.getCell('L6').value = "Belum Dibagi";
        newSheet.getCell('M6').value = "Belum Diputus";
        newSheet.mergeCells('N5:N6'); newSheet.getCell('N5').value = "Ket";

        // Angka 1-14 di baris 7
        for(let i=1; i<=14; i++) newSheet.getCell(7, i).value = i;

        // Styling Header
        for(let r=5; r<=7; r++) {
            for(let c=1; c<=14; c++) {
                let cell = newSheet.getCell(r, c);
                cell.font = { bold: true };
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
                cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
            }
        }

        // --- 4. BACA DATA DARI SIPP MENTAH ---
        const gugatanCases = [];
        const permohonanCases = [];
        let currentCase = null;
        
        let sisaLalu_G = 0, sisaLalu_P = 0, tambahIni_G = 0, tambahIni_P = 0, putusIni_G = 0, putusIni_P = 0;
        const statusSah = ["Dikabulkan", "Ditolak", "Gugur", "Tidak Dapat Diterima", "Dicabut", "Perdamaian", "Digugurkan", "Dicoret dari Register"];
        let bulanLaporan = "3"; // Maret

        rawSheet.eachRow((row, rowNumber) => {
            let strNoPerkara = safeText(row.getCell(2)); // Kolom B
            let namaHakim = safeText(row.getCell(4));    // Kolom D

            if (strNoPerkara !== "" && strNoPerkara.includes("PA.")) {
                // Simpan perkara sebelumnya ke array yang tepat
                if (currentCase) {
                    if (currentCase.noPerkara.includes("Pdt.G")) gugatanCases.push(currentCase);
                    else permohonanCases.push(currentCase);
                }
                
                // Buat keranjang baru lengkap 14 kolom
                currentCase = {
                    noPerkara: strNoPerkara,
                    kodePerkara: safeText(row.getCell(3)), // C
                    hakimKetua: namaHakim,
                    hakimAnggota1: "",
                    hakimAnggota2: "",
                    namaPP: safeText(row.getCell(5)),      // E
                    tglMasuk: getRawDateStr(row.getCell(6)), // F
                    tglPMH: getRawDateStr(row.getCell(7)),   // G
                    tglPHS: getRawDateStr(row.getCell(8)),   // H
                    tglSidang: getRawDateStr(row.getCell(9)),// I
                    tglPutus: getRawDateStr(row.getCell(10)),// J
                    statusPutusan: safeText(row.getCell(11)),// K
                    // Rekap Dapur (tersembunyi)
                    blnMasuk: getMonthFromCell(row.getCell(6)),
                    blnPutus: getMonthFromCell(row.getCell(10))
                };
            } else if (currentCase && namaHakim !== "") {
                if (currentCase.hakimAnggota1 === "") currentCase.hakimAnggota1 = namaHakim;
                else if (currentCase.hakimAnggota2 === "") currentCase.hakimAnggota2 = namaHakim;
            }
        });
        if (currentCase) {
            if (currentCase.noPerkara.includes("Pdt.G")) gugatanCases.push(currentCase);
            else permohonanCases.push(currentCase);
        }

        // --- 5. TULIS KE EXCEL (KUNCI MATI 3 BARIS) ---
        let currentRow = 8;
        
        // Fungsi pembantu buat nulis per blok (Gugatan / Permohonan)
        const writeCases = (casesArray, title) => {
            // Tulis Header Kategori (GUGATAN / PERMOHONAN)
            newSheet.mergeCells(`A${currentRow}:N${currentRow}`);
            let titleCell = newSheet.getCell(`A${currentRow}`);
            titleCell.value = title;
            titleCell.font = { bold: true };
            titleCell.alignment = { horizontal: 'center' };
            for(let c=1; c<=14; c++) newSheet.getCell(currentRow, c).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
            currentRow++;

            // Looping data
            casesArray.forEach((perkara, index) => {
                let isPutus = (perkara.tglPutus !== "");
                
                // BARIS 1: Data Inti
                newSheet.getCell(`A${currentRow}`).value = index + 1;
                newSheet.getCell(`B${currentRow}`).value = perkara.noPerkara;
                newSheet.getCell(`C${currentRow}`).value = perkara.kodePerkara;
                newSheet.getCell(`D${currentRow}`).value = perkara.hakimKetua;
                newSheet.getCell(`E${currentRow}`).value = perkara.namaPP;
                newSheet.getCell(`F${currentRow}`).value = perkara.tglMasuk;
                newSheet.getCell(`G${currentRow}`).value = perkara.tglPMH;
                newSheet.getCell(`H${currentRow}`).value = perkara.tglPHS;
                newSheet.getCell(`I${currentRow}`).value = perkara.tglSidang;
                newSheet.getCell(`J${currentRow}`).value = perkara.tglPutus;
                newSheet.getCell(`K${currentRow}`).value = perkara.statusPutusan;
                newSheet.getCell(`L${currentRow}`).value = ""; // Belum dibagi selalu kosong
                newSheet.getCell(`M${currentRow}`).value = isPutus ? "" : perkara.noPerkara; // Belum diputus
                newSheet.getCell(`N${currentRow}`).value = ""; // Ket

                // BARIS 2 & 3: Hakim Anggota (Wajib ada biar Hakim Tunggal sekalipun kotak-nya tetap 3 baris)
                newSheet.getCell(`D${currentRow+1}`).value = perkara.hakimAnggota1;
                newSheet.getCell(`D${currentRow+2}`).value = perkara.hakimAnggota2;

                // MERGE VERTIKAL 3 BARIS (Kecuali Kolom D / Hakim)
                const colsToMerge = [1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14];
                colsToMerge.forEach(col => {
                    newSheet.mergeCells(currentRow, col, currentRow + 2, col);
                });

                // STYLING BORDER & RATA TENGAH UNTUK 3 BARIS INI
                for(let r=currentRow; r<=currentRow+2; r++) {
                    for(let c=1; c<=14; c++) {
                        let cell = newSheet.getCell(r, c);
                        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                        if (c !== 4) cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                        else cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }; // Hakim rata kiri
                    }
                }

                // Hitung Rekap Dapur
                let jenis = perkara.noPerkara.includes("Pdt.G") ? "G" : "P";
                if (jenis === "G") {
                    if (perkara.blnMasuk !== bulanLaporan) sisaLalu_G++;
                    if (perkara.blnMasuk === bulanLaporan) tambahIni_G++;
                    if (perkara.blnPutus === bulanLaporan && statusSah.includes(perkara.statusPutusan)) putusIni_G++;
                } else {
                    if (perkara.blnMasuk !== bulanLaporan) sisaLalu_P++;
                    if (perkara.blnMasuk === bulanLaporan) tambahIni_P++;
                    if (perkara.blnPutus === bulanLaporan && statusSah.includes(perkara.statusPutusan)) putusIni_P++;
                }

                currentRow += 3; // Lompat 3 baris untuk perkara selanjutnya
            });
        };

        // Eksekusi nulis tabel
        writeCases(gugatanCases, "GUGATAN");
        writeCases(permohonanCases, "PERMOHONAN");

        // --- 6. TABEL REKAPITULASI BAWAH ---
        let sisaAkhir_G = sisaLalu_G + tambahIni_G - putusIni_G;
        let sisaAkhir_P = sisaLalu_P + tambahIni_P - putusIni_P;

        currentRow++; // Spasi 1 baris kosong
        newSheet.getCell(`B${currentRow}`).value = "REKAPITULASI PERKARA";
        newSheet.getCell(`C${currentRow}`).value = "Sisa Lalu";
        newSheet.getCell(`D${currentRow}`).value = "Masuk Ini";
        newSheet.getCell(`E${currentRow}`).value = "Putus Ini";
        newSheet.getCell(`F${currentRow}`).value = "Sisa Akhir";
        
        newSheet.getCell(`B${currentRow+1}`).value = "GUGATAN (G)";
        newSheet.getCell(`C${currentRow+1}`).value = sisaLalu_G;
        newSheet.getCell(`D${currentRow+1}`).value = tambahIni_G;
        newSheet.getCell(`E${currentRow+1}`).value = putusIni_G;
        newSheet.getCell(`F${currentRow+1}`).value = sisaAkhir_G;

        newSheet.getCell(`B${currentRow+2}`).value = "PERMOHONAN (P)";
        newSheet.getCell(`C${currentRow+2}`).value = sisaLalu_P;
        newSheet.getCell(`D${currentRow+2}`).value = tambahIni_P;
        newSheet.getCell(`E${currentRow+2}`).value = putusIni_P;
        newSheet.getCell(`F${currentRow+2}`).value = sisaAkhir_P;

        // Styling tabel rekap
        for(let r=currentRow; r<=currentRow+2; r++) {
            for(let c=2; c<=6; c++) {
                let cell = newSheet.getCell(r, c);
                cell.font = { bold: true };
                cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                cell.alignment = { horizontal: 'center' };
                if (c === 2) cell.alignment = { horizontal: 'left' };
            }
        }

        // --- 7. DOWNLOAD FIX NATIVE ---
        const buffer = await newWorkbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'LIPA_1_PRO_TAMA_Final_Lengkap.xlsx';
        document.body.appendChild(a); 
        a.click(); 
        document.body.removeChild(a); 
        window.URL.revokeObjectURL(url); 

        showStatus('Mantap bro! LIPA 1 100% Asli Selesai Di-generate!', 'success');

    } catch (error) {
        console.error(error);
        showStatus('Gagal, cek console log.', 'error');
    }
});
