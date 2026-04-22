// =========================================================================
// PRO-TAMA REPORT (AUTOMASI SIPP) - APP.JS (V5.1 - FULL HYBRID + REKAP)
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
    showStatus('Sedang merakit LIPA 1 (Auto-Hybrid Row)...', 'loading');

    try {
        const arrayBuffer = await file.arrayBuffer();
        const rawWorkbook = new ExcelJS.Workbook();
        await rawWorkbook.xlsx.load(arrayBuffer);
        const rawSheet = rawWorkbook.worksheets[0]; 

        const newWorkbook = new ExcelJS.Workbook();
        const newSheet = newWorkbook.addWorksheet('LIPA 1');

        // --- 1. SETUP KOLOM ---
        newSheet.columns = [
            { header: 'No', width: 5 }, { header: 'No Perkara', width: 25 },
            { header: 'Kode Perkara', width: 20 }, { header: 'Nama Hakim', width: 40 },
            { header: 'Nama PP', width: 25 }, { header: 'Penerimaan', width: 12 },
            { header: 'PMH', width: 12 }, { header: 'PHS', width: 12 },
            { header: 'Sidang I', width: 12 }, { header: 'Diputus', width: 12 },
            { header: 'Jenis Putusan', width: 18 }, { header: 'Belum Bagi', width: 12 },
            { header: 'Belum Putus', width: 25 }, { header: 'Ket', width: 10 }
        ];

        // --- 2. KOP SURAT ---
        newSheet.mergeCells('A1:N1'); newSheet.getCell('A1').value = "LAPORAN KEADAAN PERKARA";
        newSheet.mergeCells('A2:N2'); newSheet.getCell('A2').value = "PENGADILAN AGAMA SERANG";
        newSheet.mergeCells('A3:N3'); newSheet.getCell('A3').value = "BULAN MARET 2026";
        for(let r=1; r<=3; r++) {
            newSheet.getCell(`A${r}`).font = { bold: true, size: 12 };
            newSheet.getCell(`A${r}`).alignment = { horizontal: 'center' };
        }

        // --- 3. HEADER TABEL ---
        newSheet.mergeCells('A5:A6'); newSheet.getCell('A5').value = "No";
        newSheet.mergeCells('B5:B6'); newSheet.getCell('B5').value = "No Perkara";
        newSheet.mergeCells('C5:C6'); newSheet.getCell('C5').value = "Kode Perkara";
        newSheet.mergeCells('D5:D6'); newSheet.getCell('D5').value = "Nama Majelis Hakim";
        newSheet.mergeCells('E5:E6'); newSheet.getCell('E5').value = "Nama PP";
        newSheet.mergeCells('F5:J5'); newSheet.getCell('F5').value = "Tanggal";
        newSheet.getCell('F6').value = "Penerimaan"; newSheet.getCell('G6').value = "PMH";
        newSheet.getCell('H6').value = "PHS"; newSheet.getCell('I6').value = "Sidang I";
        newSheet.getCell('J6').value = "Diputus";
        newSheet.mergeCells('K5:K6'); newSheet.getCell('K5').value = "Jenis Putusan";
        newSheet.mergeCells('L5:M5'); newSheet.getCell('L5').value = "Sisa Akhir Bulan";
        newSheet.getCell('L6').value = "Belum Dibagi"; newSheet.getCell('M6').value = "Belum Diputus";
        newSheet.mergeCells('N5:N6'); newSheet.getCell('N5').value = "Ket";

        for(let i=1; i<=14; i++) newSheet.getCell(7, i).value = i;

        for(let r=5; r<=7; r++) {
            for(let c=1; c<=14; c++) {
                let cell = newSheet.getCell(r, c);
                cell.font = { bold: true };
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
                cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
            }
        }

        // --- 4. BACA DATA SIPP ---
        const gugatan = [];
        const permohonan = [];
        let currentCase = null;
        let sisaLalu_G = 0, tambahIni_G = 0, putusIni_G = 0;
        let sisaLalu_P = 0, tambahIni_P = 0, putusIni_P = 0;
        const statusSah = ["Dikabulkan", "Ditolak", "Gugur", "Tidak Dapat Diterima", "Dicabut", "Perdamaian", "Digugurkan", "Dicoret dari Register"];
        let bulanLaporan = "3"; 

        rawSheet.eachRow((row, rowNumber) => {
            let strNoPerkara = safeText(row.getCell(2)); 
            let namaHakim = safeText(row.getCell(4));    

            if (strNoPerkara !== "" && strNoPerkara.includes("PA.")) {
                if (currentCase) {
                    if (currentCase.noPerkara.includes("Pdt.G")) gugatan.push(currentCase);
                    else permohonan.push(currentCase);
                }
                currentCase = {
                    noPerkara: strNoPerkara,
                    kodePerkara: safeText(row.getCell(3)), 
                    h1: namaHakim, h2: "", h3: "",
                    pp: safeText(row.getCell(5)),      
                    tglMasuk: getRawDateStr(row.getCell(6)), tglPMH: getRawDateStr(row.getCell(7)),   
                    tglPHS: getRawDateStr(row.getCell(8)), tglSidang: getRawDateStr(row.getCell(9)),
                    tglPutus: getRawDateStr(row.getCell(10)),
                    stat: safeText(row.getCell(11)),
                    blnMasuk: getMonthFromCell(row.getCell(6)),
                    blnPutus: getMonthFromCell(row.getCell(10))
                };
            } else if (currentCase && namaHakim !== "") {
                if (currentCase.h2 === "") currentCase.h2 = namaHakim;
                else if (currentCase.h3 === "") currentCase.h3 = namaHakim;
            }
        });
        if (currentCase) {
            if (currentCase.noPerkara.includes("Pdt.G")) gugatan.push(currentCase);
            else permohonan.push(currentCase);
        }

        // --- 5. TULIS DATA (HYBRID 1 BARIS / 3 BARIS) ---
        let currentRow = 8;
        const writeSection = (dataArray, title) => {
            newSheet.mergeCells(`A${currentRow}:N${currentRow}`);
            let titleCell = newSheet.getCell(`A${currentRow}`);
            titleCell.value = title;
            titleCell.font = { bold: true };
            titleCell.alignment = { horizontal: 'center' };
            for(let c=1; c<=14; c++) newSheet.getCell(currentRow, c).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
            currentRow++;

            dataArray.forEach((item, index) => {
                let isMajelis = item.h2 !== ""; 
                let rowCount = isMajelis ? 3 : 1; 
                let start = currentRow;
                
                // Isi Baris 1
                newSheet.getRow(start).values = [
                    index + 1, item.noPerkara, item.kodePerkara, item.h1, item.pp,
                    item.tglMasuk, item.tglPMH, item.tglPHS, item.tglSidang, item.tglPutus,
                    item.stat, "", (item.tglPutus === "" ? item.noPerkara : ""), ""
                ];

                // Kalau Majelis, isi Baris 2 dan 3 lalu Merge
                if (isMajelis) {
                    newSheet.getCell(`D${start+1}`).value = item.h2;
                    newSheet.getCell(`D${start+2}`).value = item.h3;

                    [1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14].forEach(col => {
                        newSheet.mergeCells(start, col, start + 2, col);
                    });
                }

                // Styling Kotak
                for(let r=start; r<start+rowCount; r++) {
                    for(let c=1; c<=14; c++) {
                        let cell = newSheet.getCell(r, c);
                        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                        cell.alignment = { vertical: 'middle', wrapText: true };
                        if (c === 4) cell.alignment.horizontal = 'left'; // Hakim rata kiri
                        else cell.alignment.horizontal = 'center'; // Sisanya rata tengah
                    }
                }

                // Hitung Rekap
                let jenis = item.noPerkara.includes("Pdt.G") ? "G" : "P";
                if (jenis === "G") {
                    if (item.blnMasuk !== bulanLaporan) sisaLalu_G++;
                    else tambahIni_G++;
                    if (item.blnPutus === bulanLaporan && statusSah.includes(item.stat)) putusIni_G++;
                } else {
                    if (item.blnMasuk !== bulanLaporan) sisaLalu_P++;
                    else tambahIni_P++;
                    if (item.blnPutus === bulanLaporan && statusSah.includes(item.stat)) putusIni_P++;
                }

                currentRow += rowCount; 
            });
        };

        writeSection(gugatan, "GUGATAN");
        writeSection(permohonan, "PERMOHONAN");

        // --- 6. TABEL REKAPITULASI BAWAH ---
        let sisaAkhir_G = sisaLalu_G + tambahIni_G - putusIni_G;
        let sisaAkhir_P = sisaLalu_P + tambahIni_P - putusIni_P;

        currentRow++; 
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

        for(let r=currentRow; r<=currentRow+2; r++) {
            for(let c=2; c<=6; c++) {
                let cell = newSheet.getCell(r, c);
                cell.font = { bold: true };
                cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                cell.alignment = { horizontal: 'center' };
                if (c === 2) cell.alignment = { horizontal: 'left' };
            }
        }

        // --- 7. DOWNLOAD ---
        const buffer = await newWorkbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = 'LIPA_1_SERANG_FINAL_PRO.xlsx';
        document.body.appendChild(a); 
        a.click(); 
        
        setTimeout(() => { document.body.removeChild(a); }, 2000);

        showStatus(`
            <div style="text-align:center;">
                <span style="font-weight:bold;">Mantap bro! LIPA 1 Selesai Di-generate!</span><br>
                <span style="font-size:0.9em;">Kalau file tidak ter-download otomatis, </span>
                <a href="${url}" download="LIPA_1_SERANG_FINAL_PRO.xlsx" style="color:blue; text-decoration:underline; font-weight:bold;">Klik Di Sini Untuk Download Manual</a>
            </div>
        `, 'success');

    } catch (error) {
        console.error(error);
        showStatus('Gagal, cek console log.', 'error');
    }
});
