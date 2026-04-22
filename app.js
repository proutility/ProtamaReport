// =========================================================================
// PRO-TAMA REPORT (AUTOMASI SIPP) - APP.JS
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

btnGenerate.addEventListener('click', async () => {
    if (fileInput.files.length === 0) {
        showStatus('Bro, upload dulu file mentahan SIPP-nya!', 'error');
        return;
    }

    const file = fileInput.files[0];
    showStatus('Sedang memproses membaca data SIPP secara vertikal...', 'loading');

    try {
        const arrayBuffer = await file.arrayBuffer();
        const rawWorkbook = new ExcelJS.Workbook();
        await rawWorkbook.xlsx.load(arrayBuffer);
        const rawSheet = rawWorkbook.worksheets[0]; 

        const newWorkbook = new ExcelJS.Workbook();
        const newSheet = newWorkbook.addWorksheet('LIPA 1 - PRO TAMA');

        newSheet.columns = [
            { header: 'No', key: 'no', width: 5 },
            { header: 'Nomor Perkara', key: 'nomor_perkara', width: 25 },
            { header: 'Susunan Majelis / Hakim', key: 'hakim', width: 40 },
            { header: 'Tgl Penerimaan', key: 'tgl_masuk', width: 15 },
            { header: 'Tgl Putus', key: 'tgl_putus', width: 15 },
            { header: 'Status Putusan', key: 'status', width: 20 },
            { header: 'Sisa Belum Diputus', key: 'sisa_belum', width: 25 },
            { header: 'Lama Proses', key: 'lama_proses', width: 15 },
            { header: 'Keterangan', key: 'ket', width: 15 }
        ];

        let sisaLalu_G = 0, sisaLalu_P = 0;
        let tambahIni_G = 0, tambahIni_P = 0;
        let putusIni_G = 0, putusIni_P = 0;
        const statusSah = ["Dikabulkan", "Ditolak", "Gugur", "Tidak Dapat Diterima", "Dicabut", "Perdamaian", "Digugurkan", "Dicoret dari Register"];
        let bulanLaporan = 3; // Hardcode Maret dulu

        // --- LOGIKA PEMBACAAN VERTIKAL ---
        const casesData = [];
        let currentCase = null;

        rawSheet.eachRow((row, rowNumber) => {
            // Asumsi baris 1-7 itu kop surat/header. Kita mulai baca data
            // Cek apakah ada Nomor Perkara
            let rawNoPerkara = row.getCell(2).value; // Kolom B
            let strNoPerkara = rawNoPerkara ? rawNoPerkara.toString().trim() : "";
            
            let namaHakim = row.getCell(4).value ? row.getCell(4).value.toString().trim() : ""; // Kolom D

            if (strNoPerkara !== "" && strNoPerkara.includes("PA.")) {
                // Kalo currentCase udah ada isinya, masukin ke lemari (casesData)
                if (currentCase) {
                    casesData.push(currentCase);
                }
                
                // Bikin keranjang perkara baru
                currentCase = {
                    noPerkara: strNoPerkara,
                    hakimKetua: namaHakim,
                    hakimAnggota1: "",
                    hakimAnggota2: "",
                    tglMasuk: row.getCell(6).value, // Kolom F
                    tglPutus: row.getCell(10).value, // Kolom J
                    statusPutusan: row.getCell(11).value ? row.getCell(11).value.toString().trim() : "" // Kolom K
                };
            } else {
                // Kalo strNoPerkara kosong, tapi currentCase ada, berarti ini baris Anggota
                if (currentCase && namaHakim !== "") {
                    if (currentCase.hakimAnggota1 === "") {
                        currentCase.hakimAnggota1 = namaHakim;
                    } else if (currentCase.hakimAnggota2 === "") {
                        currentCase.hakimAnggota2 = namaHakim;
                    }
                }
            }
        });
        // Push perkara terakhir setelah looping selesai
        if (currentCase) casesData.push(currentCase);

        // --- PENGOLAHAN DATA & REKAP ---
        let rowIndex = 1;
        casesData.forEach(perkara => {
            // 1. Setup Hakim Tunggal vs Majelis
            let susunanHakim = "";
            if (perkara.hakimAnggota1 === "") {
                susunanHakim = "Hakim Tunggal:\n" + perkara.hakimKetua;
            } else {
                susunanHakim = "Ketua: " + perkara.hakimKetua + "\nAnggota 1: " + perkara.hakimAnggota1 + "\nAnggota 2: " + perkara.hakimAnggota2;
            }

            // 2. Setup Sisa & Lama Proses
            let sisaBelumDiputus = "-";
            let lamaProses = "-";

            if (!perkara.tglPutus || perkara.tglPutus === "") {
                sisaBelumDiputus = perkara.noPerkara;
            } else {
                let dateMasuk = new Date(perkara.tglMasuk);
                let datePutus = new Date(perkara.tglPutus);
                let hariProses = Math.ceil((datePutus.getTime() - dateMasuk.getTime()) / (1000 * 3600 * 24)); 
                lamaProses = hariProses + " hari";
            }

            // 3. Rekapitulasi
            let jenisPerkara = perkara.noPerkara.includes("Pdt.G") ? "G" : "P";
            let col_AB = perkara.tglMasuk ? (new Date(perkara.tglMasuk).getMonth() + 1).toString() : "0";
            let col_AC = perkara.tglPutus ? (new Date(perkara.tglPutus).getMonth() + 1).toString() : "0";
            let col_AE = (col_AB === bulanLaporan.toString()) ? "1" : "0";

            if (jenisPerkara === "G") {
                if (col_AE === "0") sisaLalu_G++;
                if (col_AE === "1") tambahIni_G++;
                if (col_AC === bulanLaporan.toString() && statusSah.includes(perkara.statusPutusan)) putusIni_G++;
            } else if (jenisPerkara === "P") {
                if (col_AE === "0") sisaLalu_P++;
                if (col_AE === "1") tambahIni_P++;
                if (col_AC === bulanLaporan.toString() && statusSah.includes(perkara.statusPutusan)) putusIni_P++;
            }

            // 4. Masukin ke Excel Baru
            let textTglMasuk = perkara.tglMasuk ? new Date(perkara.tglMasuk).toLocaleDateString('id-ID') : "";
            let textTglPutus = perkara.tglPutus ? new Date(perkara.tglPutus).toLocaleDateString('id-ID') : "";

            newSheet.addRow({
                no: rowIndex,
                nomor_perkara: perkara.noPerkara,
                hakim: susunanHakim,
                tgl_masuk: textTglMasuk,
                tgl_putus: textTglPutus,
                status: perkara.statusPutusan,
                sisa_belum: sisaBelumDiputus,
                lama_proses: lamaProses,
                ket: "-"
            });
            rowIndex++;
        });

        // --- BIKIN TABEL REKAP DI BAWAH ---
        let sisaAkhir_G = sisaLalu_G + tambahIni_G - putusIni_G;
        let sisaAkhir_P = sisaLalu_P + tambahIni_P - putusIni_P;

        newSheet.addRow([]);
        newSheet.addRow(['', 'REKAPITULASI PERKARA', 'Sisa Bulan Lalu', 'Tambah Bulan Ini', 'Putus Bulan Ini', 'Sisa Akhir']);
        newSheet.addRow(['', 'GUGATAN (G)', sisaLalu_G, tambahIni_G, putusIni_G, sisaAkhir_G]);
        newSheet.addRow(['', 'PERMOHONAN (P)', sisaLalu_P, tambahIni_P, putusIni_P, sisaAkhir_P]);

        // --- STYLING BORDER ---
        newSheet.eachRow((row, rowNumber) => {
            row.eachCell((cell) => {
                if(cell.value !== null && cell.value !== "") {
                    cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                }
                cell.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
            });
            if (rowNumber === 1) { row.font = { bold: true }; row.alignment = { vertical: 'middle', horizontal: 'center' }; }
        });

        let rekapHeaderRow = newSheet.lastRow - 2; 
        newSheet.getRow(rekapHeaderRow).font = { bold: true };
        newSheet.getRow(rekapHeaderRow + 1).font = { bold: true };
        newSheet.getRow(rekapHeaderRow + 2).font = { bold: true };

        const buffer = await newWorkbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'PRO_TAMA_LIPA1_Vertikal.xlsx';
        a.click();
        window.URL.revokeObjectURL(url);
        
        showStatus('Eksekusi Selesai Bro! LIPA 1 siap dikirim!', 'success');

    } catch (error) {
        console.error("Error: ", error);
        showStatus('Gagal memproses. Cek console log.', 'error');
    }
});
