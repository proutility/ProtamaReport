// =========================================================================
// PRO-TAMA REPORT (AUTOMASI SIPP) - APP.JS (V3 - Konsistensi 3 Baris)
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

function safeDate(cell) {
    if (!cell || cell.value === null || cell.value === undefined || cell.value === "") return null;
    let val = cell.value;
    if (typeof val === 'object' && val.result !== undefined) val = val.result;
    let parsedDate = new Date(val);
    return !isNaN(parsedDate.getTime()) ? parsedDate : null;
}

btnGenerate.addEventListener('click', async () => {
    if (fileInput.files.length === 0) {
        showStatus('Pilih file dulu bro!', 'error');
        return;
    }

    const file = fileInput.files[0];
    showStatus('Sedang merakit laporan 3 baris...', 'loading');

    try {
        const arrayBuffer = await file.arrayBuffer();
        const rawWorkbook = new ExcelJS.Workbook();
        await rawWorkbook.xlsx.load(arrayBuffer);
        const rawSheet = rawWorkbook.worksheets[0]; 

        const newWorkbook = new ExcelJS.Workbook();
        const newSheet = newWorkbook.addWorksheet('LIPA 1');

        // Setup Header
        newSheet.columns = [
            { header: 'No', width: 5 },
            { header: 'Nomor Perkara', width: 25 },
            { header: 'Nama Majelis Hakim', width: 40 },
            { header: 'Tgl Penerimaan', width: 15 },
            { header: 'Tgl Putus', width: 15 },
            { header: 'Status Putusan', width: 20 },
            { header: 'Belum Diputus', width: 25 },
            { header: 'Lama Proses', width: 15 }
        ];

        let sisaLalu_G = 0, sisaLalu_P = 0, tambahIni_G = 0, tambahIni_P = 0, putusIni_G = 0, putusIni_P = 0;
        const statusSah = ["Dikabulkan", "Ditolak", "Gugur", "Tidak Dapat Diterima", "Dicabut", "Perdamaian", "Digugurkan", "Dicoret dari Register"];
        let bulanLaporan = "3"; 

        const casesData = [];
        let currentCase = null;

        // Membaca data SIPP
        rawSheet.eachRow((row, rowNumber) => {
            let strNoPerkara = safeText(row.getCell(2)); 
            let namaHakim = safeText(row.getCell(4));    

            if (strNoPerkara !== "" && strNoPerkara.includes("PA.")) {
                if (currentCase) casesData.push(currentCase);
                currentCase = {
                    noPerkara: strNoPerkara,
                    hakimKetua: namaHakim,
                    hakimAnggota1: "",
                    hakimAnggota2: "",
                    tglMasuk: safeDate(row.getCell(6)),
                    tglPutus: safeDate(row.getCell(10)),
                    statusPutusan: safeText(row.getCell(11))
                };
            } else if (currentCase && namaHakim !== "") {
                if (currentCase.hakimAnggota1 === "") currentCase.hakimAnggota1 = namaHakim;
                else if (currentCase.hakimAnggota2 === "") currentCase.hakimAnggota2 = namaHakim;
            }
        });
        if (currentCase) casesData.push(currentCase);

        // Menulis ke sheet baru dengan format 3 baris
        let currentRow = 2; // Mulai setelah header (baris 1)
        casesData.forEach((perkara, index) => {
            let startRow = currentRow;
            let endRow = currentRow + 2;

            // Isi Baris 1
            newSheet.getRow(startRow).values = [
                index + 1,
                perkara.noPerkara,
                perkara.hakimKetua,
                perkara.tglMasuk ? perkara.tglMasuk.toLocaleDateString('id-ID') : "",
                perkara.tglPutus ? perkara.tglPutus.toLocaleDateString('id-ID') : "",
                perkara.statusPutusan,
                !perkara.tglPutus ? perkara.noPerkara : "-",
                perkara.tglPutus && perkara.tglMasuk ? Math.ceil((perkara.tglPutus - perkara.tglMasuk)/(1000*3600*24)) + " hari" : "-"
            ];

            // Isi Baris 2 & 3 untuk Nama Hakim
            newSheet.getRow(startRow + 1).getCell(3).value = perkara.hakimAnggota1;
            newSheet.getRow(startRow + 2).getCell(3).value = perkara.hakimAnggota2;

            // Merge Kolom Vertikal (Kecuali kolom Nama Hakim)
            [1, 2, 4, 5, 6, 7, 8].forEach(col => {
                newSheet.mergeCells(startRow, col, endRow, col);
            });

            // Rekapitulasi (Logika tetap sama)
            let jenis = perkara.noPerkara.includes("Pdt.G") ? "G" : "P";
            let blnMasuk = perkara.tglMasuk ? String(perkara.tglMasuk.getMonth() + 1) : "0";
            let blnPutus = perkara.tglPutus ? String(perkara.tglPutus.getMonth() + 1) : "0";

            if (jenis === "G") {
                if (blnMasuk !== bulanLaporan) sisaLalu_G++;
                if (blnMasuk === bulanLaporan) tambahIni_G++;
                if (blnPutus === bulanLaporan && statusSah.includes(perkara.statusPutusan)) putusIni_G++;
            } else {
                if (blnMasuk !== bulanLaporan) sisaLalu_P++;
                if (blnMasuk === bulanLaporan) tambahIni_P++;
                if (blnPutus === bulanLaporan && statusSah.includes(perkara.statusPutusan)) putusIni_P++;
            }

            currentRow += 3;
        });

        // Styling Border & Alignment
        newSheet.eachRow((row) => {
            row.eachCell((cell) => {
                cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            });
        });

        // Download
        const buffer = await newWorkbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        saveAs(blob, 'LIPA_1_PRO_TAMA.xlsx');
        showStatus('Laporan Selesai! Cek formatnya bro.', 'success');

    } catch (error) {
        console.error(error);
        showStatus('Ada error di kodingan, cek console.', 'error');
    }
});
