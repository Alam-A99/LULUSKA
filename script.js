let excelData = [];
let currentStudent = null;

// Inisialisasi jsPDF
const { jsPDF } = window.jspdf;

// Load data Excel saat halaman dimuat
window.onload = function() {
    loadExcelData();
};

// Load data dari file Excel
async function loadExcelData() {
    try {
        const url = 'data/kelulusan.xlsx';
        const response = await fetch(url);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        excelData = XLSX.utils.sheet_to_json(firstSheet);
        console.log('Data loaded:', excelData);
    } catch (error) {
        console.error('Error loading Excel data:', error);
        alert('Gagal memuat data kelulusan. Silakan coba lagi.');
    }
}

// Cari siswa
function searchStudent() {
    const searchInput = document.getElementById('searchInput').value.trim();
    
    if (!searchInput) {
        alert('Masukkan NISN atau nama siswa');
        return;
    }

    if (excelData.length === 0) {
        alert('Data belum dimuat. Silakan tunggu sebentar.');
        return;
    }

    const student = excelData.find(item => 
        item.NISN == searchInput || 
        item.Nama.toLowerCase().includes(searchInput.toLowerCase())
    );

    if (student) {
        displayStudentResult(student);
        currentStudent = student;
    } else {
        alert('Data siswa tidak ditemukan');
        document.getElementById('resultSection').style.display = 'none';
    }
}

// Tampilkan hasil pencarian
function displayStudentResult(student) {
    const resultSection = document.getElementById('resultSection');
    const studentDetails = document.getElementById('studentDetails');
    const downloadBtn = document.getElementById('downloadBtn');

    studentDetails.innerHTML = `
        <div class="detail-item">
            <strong>NISN:</strong> ${student.NISN}
        </div>
        <div class="detail-item">
            <strong>Nama:</strong> ${student.Nama}
        </div>
        <div class="detail-item">
            <strong>Kelas:</strong> ${student.Kelas}
        </div>
        <div class="detail-item">
            <strong>Jurusan:</strong> ${student.Jurusan}
        </div>
        <div class="detail-item">
            <strong>Status Kelulusan:</strong> 
            <span class="${student.Status === 'LULUS' ? 'status-lulus' : 'status-tidak-lulus'}">
                ${student.Status}
            </span>
        </div>
        ${student.Status === 'LULUS' ? `
        <div class="detail-item">
            <strong>Keterangan:</strong> Selamat! Anda dinyatakan LULUS
        </div>
        ` : `
        <div class="detail-item">
            <strong>Keterangan:</strong> Mohon maaf, Anda belum LULUS
        </div>
        `}
    `;

    // Tampilkan tombol download hanya untuk yang lulus
    if (student.Status === 'LULUS') {
        downloadBtn.style.display = 'inline-block';
    } else {
        downloadBtn.style.display = 'none';
    }

    resultSection.style.display = 'block';
    document.getElementById('allDataSection').style.display = 'none';
}

// Download bukti kelulusan
function downloadBukti() {
    if (!currentStudent || currentStudent.Status !== 'LULUS') {
        alert('Tidak dapat mengunduh bukti kelulusan');
        return;
    }

    const doc = new jsPDF();
    
    // Header
    doc.setFontSize(18);
    doc.text('BUKTI KELULUSAN', 105, 20, null, null, 'center');
    
    doc.setFontSize(14);
    doc.text('SMK Negeri 1 Jakarta', 105, 30, null, null, 'center');
    doc.text('Tahun Ajaran 2023/2024', 105, 37, null, null, 'center');
    
    // Garis pemisah
    doc.line(20, 45, 190, 45);
    
    // Informasi siswa
    doc.setFontSize(12);
    let yPos = 60;
    
    doc.text(`NISN: ${currentStudent.NISN}`, 30, yPos);
    yPos += 10;
    doc.text(`Nama: ${currentStudent.Nama}`, 30, yPos);
    yPos += 10;
    doc.text(`Kelas: ${currentStudent.Kelas}`, 30, yPos);
    yPos += 10;
    doc.text(`Jurusan: ${currentStudent.Jurusan}`, 30, yPos);
    yPos += 15;
    
    // Status kelulusan
    doc.setFontSize(14);
    doc.setTextColor(0, 128, 0);
    doc.text('DINYATAKAN LULUS', 105, yPos, null, null, 'center');
    doc.setTextColor(0, 0, 0);
    
    yPos += 20;
    doc.setFontSize(10);
    doc.text('Bukti ini diterbitkan secara resmi oleh sekolah', 105, yPos, null, null, 'center');
    
    yPos += 10;
    const tanggal = new Date().toLocaleDateString('id-ID');
    doc.text(`Tanggal: ${tanggal}`, 105, yPos, null, null, 'center');
    
    // Tanda tangan
    yPos += 30;
    doc.text('Kepala Sekolah,', 150, yPos);
    yPos += 25;
    doc.text('(____________________)', 150, yPos);
    doc.text('Dr. Budi Santoso, M.Pd', 150, yPos + 5);
    
    // Simpan file
    doc.save(`bukti_kelulusan_${currentStudent.NISN}.pdf`);
}

// Load semua data
function loadAllData() {
    if (excelData.length === 0) {
        alert('Data belum dimuat. Silakan tunggu sebentar.');
        return;
    }

    const tableBody = document.getElementById('tableBody');
    tableBody.innerHTML = '';

    excelData.forEach(student => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${student.NISN}</td>
            <td>${student.Nama}</td>
            <td>${student.Kelas}</td>
            <td class="${student.Status === 'LULUS' ? 'status-lulus-table' : 'status-tidak-lulus-table'}">
                ${student.Status}
            </td>
            <td>
                ${student.Status === 'LULUS' ? 
                    `<button class="download-btn-table" onclick="downloadBuktiTable('${student.NISN}')">
                        Download
                    </button>` : 
                    '-'
                }
            </td>
        `;
        tableBody.appendChild(row);
    });

    document.getElementById('allDataSection').style.display = 'block';
    document.getElementById('resultSection').style.display = 'none';
}

// Download bukti dari tabel
function downloadBuktiTable(nisn) {
    const student = excelData.find(item => item.NISN == nisn);
    if (student && student.Status === 'LULUS') {
        currentStudent = student;
        downloadBukti();
    } else {
        alert('Siswa tidak lulus atau data tidak ditemukan');
    }
}

// Fungsi pencarian dengan Enter
document.getElementById('searchInput').addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
        searchStudent();
    }
});
