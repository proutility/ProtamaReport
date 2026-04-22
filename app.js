// =========================================================================
// PRO-TAMA REPORT (AUTOMASI SIPP) - APP.JS (V8 - FINAL SAKLEK 3 BARIS)
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

function cleanName(name) {
    if (!name) return "";
    return name.replace(/Hakim Ketua:|Hakim Anggota \d:|Hakim Anggota:|Panitera Pengganti:/gi, '').trim();
}

function safeText(cell) {
    if (!cell || cell.value === null) return "";
    if (typeof cell.value === 'object') {
        if (cell.value.richText) return cell.value.richText.map(rt => rt.text).join('').trim();
        if (cell.value.result !== undefined) return String(cell.value.result).trim();
    }
    return String(cell.value).trim();
}

function getRawDate(cell) {
    if (!cell || !cell.value) return "";
    let val = cell.value;
    if (val instanceof Date) return val.toLocaleDateString('id-ID');
    return String(val).trim();
}

btnGenerate.addEventListener('click', async () => {
    if (fileInput.files.length === 0) {
        showStatus('Pilih file SIPP dulu bro!', 'error');
        return;
    }

    const file = fileInput.files[0];
    showStatus('Sedang merakit LIPA 1 (3 Baris Urut)...', 'loading');

    try {
        const arrayBuffer = await file.arrayBuffer();
        const rawWorkbook = new ExcelJS.Workbook();
        await rawWorkbook.xlsx.load(arrayBuffer);
        const rawSheet = rawWorkbook.worksheets[0]; 

        const newWorkbook = new ExcelJS.Workbook();
        const newSheet = newWorkbook.addWorksheet('LIPA 1');

        // Setup 14 Kolom (A sampai N)
        newSheet.columns = [
            { header: 'No', width: 5 }, { header: 'No Perkara', width: 28 },
            { header: 'Kode Perkara', width: 15 }, { header: 'Nama Majelis Hakim', width: 45 },
            { header: 'Nama PP', width: 30 }, { header: 'Penerimaan', width: 12 },
            { header: 'PMH', width: 12 }, { header: 'PHS', width: 12 },
            { header: 'Sidang I', width: 12 }, { header: 'Diputus', width: 12 },
            { header: 'Jenis Putusan', width: 18 }, { header: 'Belum Bagi', width: 12 },
            { header: 'Belum Putus', width: 30 }, { header: 'Ket', width: 10 }
        ];

        // Kop Surat
        newSheet.mergeCells('A1:N1'); newSheet.getCell('A1').value = "LAPORAN KEADAAN PERKARA PENGADILAN AGAMA SERANG";
        newSheet.getCell('A1').font = { bold: true, size: 14 };
        newSheet.getCell('A1').alignment = { horizontal: 'center' };

        const gugatan = [];
        const permohonan = [];
        let current = null;

        // Baca SIPP
        rawSheet.eachRow((row) => {
            let noPerk = safeText(row.getCell(2)); 
            let hakim = cleanName(safeText(row.getCell(4)));

            if (noPerk !== "" && noPerk.includes("/")) {
                if (current) {
                    if (current.noPerk.includes("Pdt.G")) gugatan.push(current);
                    else permohonan.push(current);
                }
                current = {
                    noPerk: noPerk, kode: safeText(row.getCell(3)),
                    h1: hakim, h2: "", h3: "",
                    pp: cleanName(safeText(row.getCell(5))),
                    dates: [getRawDate(row.getCell(6)), getRawDate(row.getCell(7)), getRawDate(row.getCell(8)), getRawDate(row.getCell(9)), getRawDate(row.getCell(10))],
                    stat: safeText(row.getCell(11))
                };
            } else if (current && hakim !== "") {
                if (current.h2 === "") current.h2 = hakim;
                else if (current.h3 === "") current.h3 = hakim;
            }
        });
        if (current) {
            if (current.noPerk.includes("Pdt.G")) gugatan.push(current);
            else permohonan.push(current);
        }

        let currentRow = 8;
        const writeData = (data, title) => {
            newSheet.mergeCells(`A${currentRow}:N${currentRow}`);
            let tCell = newSheet.getCell(`A${currentRow}`);
            tCell.value = title; tCell.font = { bold: true }; tCell.alignment = { horizontal: 'center' };
            currentRow++;

            data.forEach((item, index) => {
                let isMajelis = item.h2 !== "";
                let rowCount = isMajelis ? 3 : 1;
                let start = currentRow;
                let end = start + (rowCount - 1);

                // Baris 1: Semua Data + Hakim Ketua
                newSheet.getRow(start).values = [
                    index + 1, item.noPerk, item.kode, item.h1, item.pp,
                    item.dates[0], item.dates[1], item.dates[2], item.dates[3], item.dates[4],
                    item.stat, "", (item.dates[4] === "" ? item.noPerk : ""), ""
                ];

                if (isMajelis) {
                    // Baris 2 & 3: Hanya kolom Hakim (Kolom 4 / D)
                    newSheet.getCell(`D${start + 1}`).value = item.h2;
                    newSheet.getCell(`D${start + 2}`).value = item.h3;

                    // MERGE VERTIKAL selain Kolom Hakim
                    [1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14].forEach(c => {
                        newSheet.mergeCells(start, c, end, c);
                    });
                }

                // STYLING SEMUA BARIS
                for (let r = start; r <= end; r++) {
                    for (let c = 1; c <= 14; c++) {
                        let cell = newSheet.getCell(r, c);
                        cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
                        cell.alignment = { vertical: 'middle', wrapText: true };
                        cell.alignment.horizontal = (c === 4 ? 'left' : 'center');
                    }
                }
                currentRow += rowCount;
            });
        };

        writeData(gugatan, "GUGATAN");
        writeData(permohonan, "PERMOHONAN");

        const buffer = await newWorkbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = 'LIPA_1_SERANG_V8_FINAL.xlsx';
        a.click();
        
        showStatus('<b>Berhasil Bro!</b> Formasi hakim sudah urut 3 baris ke bawah.', 'success');

    } catch (error) {
        console.error(error);
        showStatus('Gagal, cek format file SIPP lo.', 'error');
    }
});
