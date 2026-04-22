// =========================================================================
// PRO-TAMA REPORT (AUTOMASI SIPP) - APP.JS (V7 - FINAL MASTER LIPA 1)
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

function cleanHakim(name) {
    if (!name) return "";
    // Menghilangkan awalan "Hakim Ketua:", "Hakim Anggota:", dll jika ada dari SIPP
    return name.replace(/Hakim Ketua:|Hakim Anggota \d:|Hakim Anggota:/gi, '').trim();
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
    if (val instanceof Date) return val.toLocaleDateString('id-ID');
    return String(val).trim();
}

btnGenerate.addEventListener('click', async () => {
    if (fileInput.files.length === 0) {
        showStatus('Pilih file SIPP dulu bro!', 'error');
        return;
    }

    const file = fileInput.files[0];
    showStatus('Sedang merakit LIPA 1 Master (3 Baris Berjenjang)...', 'loading');

    try {
        const arrayBuffer = await file.arrayBuffer();
        const rawWorkbook = new ExcelJS.Workbook();
        await rawWorkbook.xlsx.load(arrayBuffer);
        const rawSheet = rawWorkbook.worksheets[0]; 

        const newWorkbook = new ExcelJS.Workbook();
        const newSheet = newWorkbook.addWorksheet('LIPA 1');

        // Setup 14 Kolom Standar PA Serang
        newSheet.columns = [
            { header: 'No', width: 5 },              // 1 (A)
            { header: 'No Perkara', width: 30 },     // 2 (B)
            { header: 'Kode Perkara', width: 15 },   // 3 (C)
            { header: 'Nama Majelis Hakim', width: 45 }, // 4 (D) -> INI YANG TIDAK DI-MERGE
            { header: 'Nama PP', width: 30 },        // 5 (E)
            { header: 'Penerimaan', width: 12 },     // 6 (F)
            { header: 'PMH', width: 12 },            // 7 (G)
            { header: 'PHS', width: 12 },            // 8 (H)
            { header: 'Sidang I', width: 12 },       // 9 (I)
            { header: 'Diputus', width: 12 },        // 10 (J)
            { header: 'Jenis Putusan', width: 18 },  // 11 (K)
            { header: 'Belum Bagi', width: 12 },     // 12 (L)
            { header: 'Belum Putus', width: 30 },    // 13 (M)
            { header: 'Ket', width: 10 }             // 14 (N)
        ];

        // Kop Surat PA Serang
        newSheet.mergeCells('A1:N1'); newSheet.getCell('A1').value = "LAPORAN KEADAAN PERKARA PENGADILAN AGAMA SERANG";
        newSheet.getCell('A1').font = { bold: true, size: 14 };
        newSheet.getCell('A1').alignment = { horizontal: 'center' };

        // Data Processing
        const gugatan = [];
        const permohonan = [];
        let currentCase = null;

        rawSheet.eachRow((row, rowNumber) => {
            let noPerk = safeText(row.getCell(2)); 
            let hakim = cleanHakim(safeText(row.getCell(4)));

            if (noPerk !== "" && noPerk.includes("/")) {
                if (currentCase) {
                    if (currentCase.noPerk.includes("Pdt.G")) gugatan.push(currentCase);
                    else permohonan.push(currentCase);
                }
                currentCase = {
                    noPerk: noPerk,
                    kode: safeText(row.getCell(3)),
                    h1: hakim, h2: "", h3: "",
                    pp: safeText(row.getCell(5)),
                    dates: [getRawDateStr(row.getCell(6)), getRawDateStr(row.getCell(7)), getRawDateStr(row.getCell(8)), getRawDateStr(row.getCell(9)), getRawDateStr(row.getCell(10))],
                    stat: safeText(row.getCell(11))
                };
            } else if (currentCase && hakim !== "") {
                if (currentCase.h2 === "") currentCase.h2 = hakim;
                else if (currentCase.h3 === "") currentCase.h3 = hakim;
            }
        });
        if (currentCase) {
            if (currentCase.noPerk.includes("Pdt.G")) gugatan.push(currentCase);
            else permohonan.push(currentCase);
        }

        let currentRow = 8;
        const writeData = (data, title) => {
            newSheet.mergeCells(`A${currentRow}:N${currentRow}`);
            let titleCell = newSheet.getCell(`A${currentRow}`);
            titleCell.value = title;
            titleCell.font = { bold: true };
            titleCell.alignment = { horizontal: 'center' };
            currentRow++;

            data.forEach((item, index) => {
                let start = currentRow;
                let end = currentRow + 2; // FIX JATAH 3 BARIS

                // Baris 1: Semua data inti + Hakim Ketua
                newSheet.getRow(start).values = [
                    index + 1, item.noPerk, item.kode, item.h1, item.pp,
                    item.dates[0], item.dates[1], item.dates[2], item.dates[3], item.dates[4],
                    item.stat, "", (item.dates[4] === "" ? item.noPerk : ""), ""
                ];

                // Baris 2 & 3: Hanya untuk Hakim Anggota di Kolom D
                newSheet.getCell(`D${start + 1}`).value = item.h2;
                newSheet.getCell(`D${start + 2}`).value = item.h3;

                // MERGE SEMUA KOLOM VERTIKAL KECUALI KOLOM D (URUTAN 4)
                [1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14].forEach(col => {
                    newSheet.mergeCells(start, col, end, col);
                });

                // STYLING BORDER & ALIGNMENT
                for (let r = start; r <= end; r++) {
                    for (let c = 1; c <= 14; c++) {
                        let cell = newSheet.getCell(r, c);
                        cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
                        cell.alignment = { vertical: 'middle', wrapText: true };
                        // Kolom Hakim Rata Kiri, Sisanya Tengah
                        cell.alignment.horizontal = (c === 4 ? 'left' : 'center');
                    }
                }
                currentRow += 3; 
            });
        };

        writeData(gugatan, "GUGATAN");
        writeData(permohonan, "PERMOHONAN");

        // --- DOWNLOAD ---
        const buffer = await newWorkbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = 'LIPA_1_SERANG_MASTER_V7.xlsx';
        a.click();
        
        showStatus(`<b>Berhasil!</b> LIPA 1 sudah rapi 3 baris berjenjang. <br> <a href="${url}" download="LIPA_1_SERANG_MASTER_V7.xlsx" style="color:blue;">Klik Manual Jika Tidak Download</a>`, 'success');

    } catch (error) {
        console.error(error);
        showStatus('Gagal, pastikan file SIPP benar.', 'error');
    }
});
