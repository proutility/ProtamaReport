// =========================================================================
// PRO-TAMA REPORT (AUTOMASI SIPP) - APP.JS (V6 - STANDAR MA SERANG)
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
    if (val instanceof Date) return val.toLocaleDateString('id-ID');
    return String(val).trim();
}

btnGenerate.addEventListener('click', async () => {
    if (fileInput.files.length === 0) {
        showStatus('Pilih file SIPP dulu bro!', 'error');
        return;
    }

    const file = fileInput.files[0];
    showStatus('Sedang merakit LIPA 1 (Struktur 3 Baris Fix)...', 'loading');

    try {
        const arrayBuffer = await file.arrayBuffer();
        const rawWorkbook = new ExcelJS.Workbook();
        await rawWorkbook.xlsx.load(arrayBuffer);
        const rawSheet = rawWorkbook.worksheets[0]; 

        const newWorkbook = new ExcelJS.Workbook();
        const newSheet = newWorkbook.addWorksheet('LIPA 1');

        // --- 1. SETUP 14 KOLOM ---
        newSheet.columns = [
            { header: 'No', width: 5 },              // 1 (A)
            { header: 'No Perkara', width: 25 },     // 2 (B)
            { header: 'Kode Perkara', width: 20 },   // 3 (C) -> KOLOM SENDIRI
            { header: 'Majelis Hakim', width: 40 },  // 4 (D) -> KOLOM SENDIRI (3 BARIS)
            { header: 'PP', width: 25 },             // 5 (E) -> KOLOM SENDIRI
            { header: 'Penerimaan', width: 12 },     // 6 (F)
            { header: 'PMH', width: 12 },            // 7 (G)
            { header: 'PHS', width: 12 },            // 8 (H)
            { header: 'Sidang I', width: 12 },       // 9 (I)
            { header: 'Diputus', width: 12 },        // 10 (J)
            { header: 'Jenis Putusan', width: 18 },  // 11 (K)
            { header: 'Belum Bagi', width: 12 },     // 12 (L)
            { header: 'Belum Putus', width: 25 },    // 13 (M)
            { header: 'Ket', width: 10 }             // 14 (N)
        ];

        // Header Atas (Kop)
        newSheet.mergeCells('A1:N1'); newSheet.getCell('A1').value = "LAPORAN KEADAAN PERKARA PENGADILAN AGAMA SERANG";
        newSheet.getCell('A1').font = { bold: true, size: 14 };
        newSheet.getCell('A1').alignment = { horizontal: 'center' };

        // --- 2. PROSES DATA SIPP ---
        const gugatan = [];
        const permohonan = [];
        let currentCase = null;

        rawSheet.eachRow((row, rowNumber) => {
            let noPerk = safeText(row.getCell(2)); // Kolom B
            let hakim = safeText(row.getCell(4));  // Kolom D

            if (noPerk !== "" && noPerk.includes("/")) {
                if (currentCase) {
                    if (currentCase.noPerk.includes("Pdt.G")) gugatan.push(currentCase);
                    else permohonan.push(currentCase);
                }
                currentCase = {
                    noPerk: noPerk,
                    kode: safeText(row.getCell(3)), // Kolom C SIPP (Jenis Perkara)
                    h1: hakim, h2: "", h3: "",
                    pp: safeText(row.getCell(5)),   // Kolom E SIPP (PP)
                    tgl: [getRawDateStr(row.getCell(6)), getRawDateStr(row.getCell(7)), getRawDateStr(row.getCell(8)), getRawDateStr(row.getCell(9)), getRawDateStr(row.getCell(10))],
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

        // --- 3. PENULISAN KE EXCEL (FIX 3 BARIS PER PERKARA) ---
        let currentRow = 8;
        const writeData = (data, sectionTitle) => {
            newSheet.mergeCells(`A${currentRow}:N${currentRow}`);
            newSheet.getCell(`A${currentRow}`).value = sectionTitle;
            newSheet.getCell(`A${currentRow}`).font = { bold: true };
            newSheet.getCell(`A${currentRow}`).alignment = { horizontal: 'center' };
            currentRow++;

            data.forEach((item, index) => {
                let start = currentRow;
                let end = currentRow + 2; // Paksa jatah 3 baris

                // Baris 1: Isi Semua Data Utama
                newSheet.getRow(start).values = [
                    index + 1, item.noPerk, item.kode, item.h1, item.pp,
                    item.tgl[0], item.tgl[1], item.tgl[2], item.tgl[3], item.tgl[4],
                    item.stat, "", (item.tgl[4] === "" ? item.noPerk : ""), ""
                ];

                // Baris 2 & 3: Hanya untuk Hakim Anggota
                newSheet.getCell(`D${start + 1}`).value = item.h2;
                newSheet.getCell(`D${start + 2}`).value = item.h3;

                // MERGE SEMUA KOLOM KECUALI KOLOM 4 (MAJELIS HAKIM)
                [1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14].forEach(col => {
                    newSheet.mergeCells(start, col, end, col);
                });

                // STYLING SEMUA BARIS (Start sampai End)
                for (let r = start; r <= end; r++) {
                    for (let c = 1; c <= 14; c++) {
                        let cell = newSheet.getCell(r, c);
                        cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
                        cell.alignment = { vertical: 'middle', wrapText: true };
                        // Hakim rata kiri, selain itu tengah
                        cell.alignment.horizontal = (c === 4 ? 'left' : 'center');
                    }
                }
                currentRow += 3; // Selalu lompat 3 baris
            });
        };

        writeData(gugatan, "GUGATAN");
        writeData(permohonan, "PERMOHONAN");

        // --- 4. DOWNLOAD ---
        const buffer = await newWorkbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = 'LIPA_1_FIX_SERANG_14KOLOM.xlsx';
        a.click();
        
        showStatus('<b>Sukses Bro!</b> LIPA 1 sudah rapi 14 kolom & formasi 3 baris.', 'success');

    } catch (error) {
        console.error(error);
        showStatus('Error pas baca SIPP. Cek filenya udah .xlsx asli belum?', 'error');
    }
});
