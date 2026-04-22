// =========================================================================
// PRO-TAMA REPORT (AUTOMASI SIPP) - APP.JS (V4.2 - FIX FORMASI HAKIM 3 BARIS)
// Develop by: Ilham Nur Pratama (PA Serang)
// =========================================================================

const btnGenerate = document.getElementById('generateLaporan');
const fileInput = document.getElementById('uploadSipp');
const statusMsg = document.getElementById('statusMessage');

function showStatus(message, type) {
    statusMsg.innerHTML = message;
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

function getMonthFromCell(cell) {
    if (!cell || !cell.value) return "0";
    let val = cell.value;
    if (typeof val === 'object' && val.result !== undefined) val = val.result;
    if (val instanceof Date) return String(val.getMonth() + 1);
    if (typeof val === 'string') {
        let parts = val.split('/');
        if (parts.length === 3) return String(parseInt(parts[1], 10)); 
    }
    return "0";
}

btnGenerate.addEventListener('click', async () => {
    if (fileInput.files.length === 0) {
        showStatus('Pilih file SIPP dulu bro!', 'error');
        return;
    }

    const file = fileInput.files[0];
    showStatus('Sedang merakit LIPA 1 dengan Formasi Hakim 3 Baris...', 'loading');

    try {
        const arrayBuffer = await file.arrayBuffer();
        const rawWorkbook = new ExcelJS.Workbook();
        await rawWorkbook.xlsx.load(arrayBuffer);
        const rawSheet = rawWorkbook.worksheets[0]; 

        const newWorkbook = new ExcelJS.Workbook();
        const newSheet = newWorkbook.addWorksheet('LIPA 1');

        // Setup Kolom (14 Kolom Standar LIPA 1)
        newSheet.columns = [
            { header: 'No', width: 5 }, { header: 'No Perkara', width: 25 },
            { header: 'Kode', width: 15 }, { header: 'Majelis Hakim', width: 40 },
            { header: 'PP', width: 25 }, { header: 'Penerimaan', width: 12 },
            { header: 'PMH', width: 12 }, { header: 'PHS', width: 12 },
            { header: 'Sidang I', width: 12 }, { header: 'Diputus', width: 12 },
            { header: 'Status', width: 15 }, { header: 'Belum Bagi', width: 12 },
            { header: 'Belum Putus', width: 15 }, { header: 'Ket', width: 10 }
        ];

        // Kop Surat
        newSheet.mergeCells('A1:N1'); newSheet.getCell('A1').value = "LAPORAN KEADAAN PERKARA";
        newSheet.mergeCells('A2:N2'); newSheet.getCell('A2').value = "PENGADILAN AGAMA SERANG";
        newSheet.mergeCells('A3:N3'); newSheet.getCell('A3').value = "BULAN MARET 2026";
        for(let r=1; r<=3; r++) {
            newSheet.getCell(`A${r}`).font = { bold: true, size: 12 };
            newSheet.getCell(`A${r}`).alignment = { horizontal: 'center' };
        }

        // Header Tabel (Baris 5-7)
        newSheet.mergeCells('A5:A6'); newSheet.getCell('A5').value = "No";
        newSheet.mergeCells('B5:B6'); newSheet.getCell('B5').value = "No Perkara";
        newSheet.mergeCells('D5:D6'); newSheet.getCell('D5').value = "Nama Majelis Hakim";
        newSheet.mergeCells('F5:J5'); newSheet.getCell('F5').value = "Tanggal";
        newSheet.getCell('F6').value = "Penerimaan"; newSheet.getCell('J6').value = "Diputus";
        newSheet.mergeCells('L5:M5'); newSheet.getCell('L5').value = "Sisa Akhir";
        for(let i=1; i<=14; i++) newSheet.getCell(7, i).value = i;

        const gugatanCases = [];
        const permohonanCases = [];
        let currentCase = null;
        let sisaLalu_G = 0, tambahIni_G = 0, putusIni_G = 0;
        let sisaLalu_P = 0, tambahIni_P = 0, putusIni_P = 0;
        const statusSah = ["Dikabulkan", "Ditolak", "Gugur", "Tidak Dapat Diterima", "Dicabut", "Perdamaian", "Digugurkan", "Dicoret dari Register"];
        let bulanLaporan = "3";

        // Baca Data
        rawSheet.eachRow((row, rowNumber) => {
            let strNoPerkara = safeText(row.getCell(2)); 
            let namaHakim = safeText(row.getCell(4));    

            if (strNoPerkara !== "" && strNoPerkara.includes("PA.")) {
                if (currentCase) {
                    if (currentCase.noPerkara.includes("Pdt.G")) gugatanCases.push(currentCase);
                    else permohonanCases.push(currentCase);
                }
                currentCase = {
                    noPerkara: strNoPerkara,
                    kodePerkara: safeText(row.getCell(3)), 
                    hakimKetua: namaHakim,
                    hakimAnggota1: "",
                    hakimAnggota2: "",
                    namaPP: safeText(row.getCell(5)),      
                    tglMasuk: getRawDateStr(row.getCell(6)), 
                    tglPutus: getRawDateStr(row.getCell(10)),
                    statusPutusan: safeText(row.getCell(11)),
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

        let currentRow = 8;
        const writeCases = (casesArray, title) => {
            newSheet.mergeCells(`A${currentRow}:N${currentRow}`);
            newSheet.getCell(`A${currentRow}`).value = title;
            newSheet.getCell(`A${currentRow}`).font = { bold: true };
            newSheet.getCell(`A${currentRow}`).alignment = { horizontal: 'center' };
            currentRow++;

            casesArray.forEach((perkara, index) => {
                // ISI BARIS 1 (Data Utama + Hakim Ketua)
                newSheet.getCell(`A${currentRow}`).value = index + 1;
                newSheet.getCell(`B${currentRow}`).value = perkara.noPerkara;
                newSheet.getCell(`D${currentRow}`).value = perkara.hakimKetua; // Baris 1 Ketua
                newSheet.getCell(`F${currentRow}`).value = perkara.tglMasuk;
                newSheet.getCell(`J${currentRow}`).value = perkara.tglPutus;
                newSheet.getCell(`K${currentRow}`).value = perkara.statusPutusan;
                newSheet.getCell(`M${currentRow}`).value = (perkara.tglPutus === "") ? perkara.noPerkara : "";

                // ISI BARIS 2 & 3 (Khusus Hakim Anggota)
                newSheet.getCell(`D${currentRow+1}`).value = perkara.hakimAnggota1; // Baris 2 Anggota 1
                newSheet.getCell(`D${currentRow+2}`).value = perkara.hakimAnggota2; // Baris 3 Anggota 2

                // MERGE KOLOM NON-HAKIM (1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14)
                [1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14].forEach(col => {
                    newSheet.mergeCells(currentRow, col, currentRow + 2, col);
                });

                // STYLING SEMUA SEL PERKARA
                for(let r=currentRow; r<=currentRow+2; r++) {
                    for(let c=1; c<=14; c++) {
                        let cell = newSheet.getCell(r, c);
                        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                        cell.alignment = { vertical: 'middle', wrapText: true };
                        if (c === 4) cell.alignment.horizontal = 'left'; // Hakim rata kiri
                        else cell.alignment.horizontal = 'center'; // Sisanya tengah
                    }
                }

                // Rekap
                let jenis = perkara.noPerkara.includes("Pdt.G") ? "G" : "P";
                if (jenis === "G") {
                    if (perkara.blnMasuk !== bulanLaporan) sisaLalu_G++;
                    else tambahIni_G++;
                    if (perkara.blnPutus === bulanLaporan && statusSah.includes(perkara.statusPutusan)) putusIni_G++;
                } else {
                    if (perkara.blnMasuk !== bulanLaporan) sisaLalu_P++;
                    else tambahIni_P++;
                    if (perkara.blnPutus === bulanLaporan && statusSah.includes(perkara.statusPutusan)) putusIni_P++;
                }
                currentRow += 3; 
            });
        };

        writeCases(gugatanCases, "GUGATAN");
        writeCases(permohonanCases, "PERMOHONAN");

        // Download Final
        const buffer = await newWorkbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = 'LIPA_1_SERANG_FIX_HAKIM.xlsx';
        a.click();
        
        showStatus(`<b>Mantap bro!</b> Laporan LIPA 1 Formasi Hakim 3 Baris Selesai. <br> <a href="${url}" download="LIPA_1_SERANG_FIX_HAKIM.xlsx" style="color:blue;">Klik Sini Jika Tidak Terdownload Otomatis</a>`, 'success');

    } catch (error) {
        console.error(error);
        showStatus('Error, pastikan file SIPP sudah di Save As .xlsx ya bro.', 'error');
    }
});
