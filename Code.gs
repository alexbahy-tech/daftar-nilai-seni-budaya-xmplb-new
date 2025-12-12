// File: Code.gs

// Global variables
const SPREADSHEET_ID = '1QwWR5UTUeRyKz3JNH5LbuUW6u80qa3AuC5mz0tJRBvE'; // <== GANTI DENGAN ID GOOGLE SHEET ANDA
const SHEET_NAME = 'Sheet1';
const SETTINGS_SHEET_NAME = 'Settings';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sistem Daftar Nilai - ' + getClassAndSubject().className)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ===============================================
// === UTILITY FUNCTIONS (Settings, Init, Data) ===
// ===============================================

/**
 * Get settings data from the 'Settings' sheet.
 */
function getClassAndSubject() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);

  if (!settingsSheet) {
    // Initialize settings if the sheet does not exist
    settingsSheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    const defaults = [
      ['className', 'Kelas X Teknik Jaringan Komputer dan Telekomunikasi 2'],
      ['subjectName', 'Bahasa Inggris'],
      ['schoolName', 'SMK Negeri 1'],
      ['teacherName', ''],
      ['semester', 'Ganjil'],
      ['academicYear', '2025/2026']
    ];
    settingsSheet.getRange(1, 1, defaults.length, 2).setValues(defaults);
  }

  const settingsRange = settingsSheet.getRange('A1:B6').getValues();
  const settings = {};
  settingsRange.forEach(row => {
    if (row[0]) settings[row[0]] = row[1] || '';
  });

  return {
    className: settings.className || 'Kelas X TKJ 2',
    subjectName: settings.subjectName || 'Bahasa Inggris',
    schoolName: settings.schoolName || 'SMK Negeri 1',
    teacherName: settings.teacherName || '',
    semester: settings.semester || 'Ganjil',
    academicYear: settings.academicYear || '2025/2026'
  };
}

/**
 * Update settings in the 'Settings' sheet.
 */
function updateSettings(settingsData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);

    if (!settingsSheet) {
      settingsSheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    }

    const updates = [
      ['className', settingsData.className],
      ['subjectName', settingsData.subjectName],
      ['schoolName', settingsData.schoolName],
      ['teacherName', settingsData.teacherName],
      ['semester', settingsData.semester],
      ['academicYear', settingsData.academicYear]
    ];

    settingsSheet.getRange(1, 1, updates.length, 2).setValues(updates);

    return { success: true, message: "Pengaturan berhasil disimpan!" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Initialize main sheet headers if not exists
 */
function initializeSheetHeaders() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Check if headers exist
  const lastRow = sheet.getLastRow();
  if (lastRow === 0 || !sheet.getRange(1, 1).getValue()) {
    const headers = [
      'No', 'Nama', 'NIS', 'NISN', 'L/P',
      'F1', 'F2', 'F3', 'F4', 'F5', 'Rata-rata Formatif',
      'S1', 'S2', 'S3', 'Rata-rata Sumatif',
      'Nilai Akhir', 'Nilai Rapor'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f1f3f4');
  }
}

/**
 * Get all valid student rows, filter empty rows, and sort them by name.
 */
function getAllStudentData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  initializeSheetHeaders();

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  // Get all data from row 2 onwards (17 columns total)
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 17); 
  const values = dataRange.getValues();

  // Filter out empty rows and create student objects
  let students = [];
  values.forEach((row, index) => {
    // Check if row has valid student data (nama must exist and not empty)
    if (row[1] && row[1].toString().trim() !== '') {
      students.push({
        originalRowIndex: index + 2, // Original position in sheet
        nama: row[1],
        nis: row[2],
        nisn: row[3],
        jk: row[4],
        formatif: {
          f1: row[5] || 0, f2: row[6] || 0, f3: row[7] || 0, f4: row[8] || 0, f5: row[9] || 0,
          rataRata: row[10] || 0
        },
        sumatif: {
          s1: row[11] || 0, s2: row[12] || 0, s3: row[13] || 0,
          rataRata: row[14] || 0
        },
        akhirSemester: row[15] || 0,
        nilaiRapor: row[16] || 0
      });
    }
  });

  // Sort students by name to ensure consistent ordering in the app
  students.sort((a, b) => a.nama.localeCompare(b.nama));

  // Add sequential numbering and new row index (which is consistent with the sheet after reorganization)
  return students.map((student, index) => ({
    ...student,
    no: index + 1,
    rowIndex: index + 2 // New sequential row position (for grade saving)
  }));
}

/**
 * Reorganize sheet data to ensure proper sorting and remove gaps.
 */
function reorganizeSheetData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  const students = getAllStudentData();

  // Clear existing data (except headers)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 17).clear();
  }

  // Prepare new organized data
  const newData = students.map((student, index) => [
    index + 1, // Sequential number
    student.nama,
    student.nis,
    student.nisn,
    student.jk,
    student.formatif.f1, student.formatif.f2, student.formatif.f3, student.formatif.f4, student.formatif.f5,
    student.formatif.rataRata,
    student.sumatif.s1, student.sumatif.s2, student.sumatif.s3,
    student.sumatif.rataRata,
    student.akhirSemester,
    student.nilaiRapor
  ]);

  // Write organized data back to sheet
  if (newData.length > 0) {
    sheet.getRange(2, 1, newData.length, 17).setValues(newData);
  }
}

/**
 * Update sequential numbering after deletion.
 */
function updateSequentialNumbering(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const nameColumn = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  let studentNumber = 1;

  for (let i = 0; i < nameColumn.length; i++) {
    const name = nameColumn[i][0];
    if (name && name.toString().trim() !== '') {
      sheet.getRange(i + 2, 1).setValue(studentNumber);
      studentNumber++;
    }
  }
}


// ============================================
// === STUDENT MANAGEMENT FUNCTIONS (CRUD) ===
// ============================================

/**
 * Validate student data for duplicates.
 */
function validateStudentData(newStudents, existingStudents = null) {
  const errors = [];
  if (!existingStudents) {
    existingStudents = getAllStudentData();
  }

  const newNIS = new Set();
  const newNISN = new Set();

  newStudents.forEach((student, index) => {
    // Check required fields
    if (!student.nama || !student.nis || !student.nisn || !student.jk) {
      errors.push(`Siswa ke-${index + 1}: Semua field wajib diisi!`);
      return;
    }

    // Trim whitespace
    student.nama = student.nama.trim();
    student.nis = student.nis.trim();
    student.nisn = student.nisn.trim();

    // Check duplicates within new students batch
    if (newNIS.has(student.nis)) {
      errors.push(`Siswa ke-${index + 1}: NIS ${student.nis} duplikat dalam input!`);
    } else {
      newNIS.add(student.nis);
    }

    if (newNISN.has(student.nisn)) {
      errors.push(`Siswa ke-${index + 1}: NISN ${student.nisn} duplikat dalam input!`);
    } else {
      newNISN.add(student.nisn);
    }

    // Check duplicates with existing students
    const nisExists = existingStudents.some(existing => 
        existing.nis.toString() === student.nis &&
        existing.originalRowIndex !== parseInt(student.rowIndex) // Exclude self if editing
    );
    const nisnExists = existingStudents.some(existing => 
        existing.nisn.toString() === student.nisn &&
        existing.originalRowIndex !== parseInt(student.rowIndex) // Exclude self if editing
    );

    if (nisExists) {
      errors.push(`Siswa ke-${index + 1}: NIS ${student.nis} sudah ada!`);
    }

    if (nisnExists) {
      errors.push(`Siswa ke-${index + 1}: NISN ${student.nisn} sudah ada!`);
    }
  });

  return errors;
}

/**
 * Add a single new student, then resort and rewrite all data.
 */
function addNewStudent(studentData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    // Validate single student
    const validationErrors = validateStudentData([studentData]);
    if (validationErrors.length > 0) {
      throw new Error(validationErrors[0]);
    }

    // Get current students and add new one with default grades
    const existingStudents = getAllStudentData();
    const newStudentDefaults = {
      nama: studentData.nama.trim(), nis: studentData.nis.trim(), nisn: studentData.nisn.trim(), jk: studentData.jk,
      formatif: { f1: 0, f2: 0, f3: 0, f4: 0, f5: 0, rataRata: 0 },
      sumatif: { s1: 0, s2: 0, s3: 0, rataRata: 0 },
      akhirSemester: 0, nilaiRapor: 0
    };
    
    // Perform full reorganization
    const allStudents = [...existingStudents.map(s => ({...s, originalRowIndex: null})), newStudentDefaults];
    allStudents.sort((a, b) => a.nama.localeCompare(b.nama));
    
    // Clear and rewrite all data
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 17).clear();
    }

    const newData = allStudents.map((student, index) => [
      index + 1, student.nama, student.nis, student.nisn, student.jk,
      student.formatif.f1, student.formatif.f2, student.formatif.f3, student.formatif.f4, student.formatif.f5, student.formatif.rataRata,
      student.sumatif.s1, student.sumatif.s2, student.sumatif.s3, student.sumatif.rataRata,
      student.akhirSemester, student.nilaiRapor
    ]);

    if (newData.length > 0) {
      sheet.getRange(2, 1, newData.length, 17).setValues(newData);
    }

    return { success: true, message: `Siswa ${studentData.nama} berhasil ditambahkan!` };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Add multiple students, then resort and rewrite all data.
 */
function addMultipleStudents(studentsData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    const existingStudents = getAllStudentData();

    // Validate all student data against current and new students
    const validationErrors = validateStudentData(studentsData, existingStudents);
    if (validationErrors.length > 0) {
      throw new Error(validationErrors.join('\n'));
    }

    // Prepare new students with default grades
    const newStudentsDefaults = studentsData.map(student => ({
      nama: student.nama.trim(), nis: student.nis.trim(), nisn: student.nisn.trim(), jk: student.jk,
      formatif: { f1: 0, f2: 0, f3: 0, f4: 0, f5: 0, rataRata: 0 },
      sumatif: { s1: 0, s2: 0, s3: 0, rataRata: 0 },
      akhirSemester: 0, nilaiRapor: 0
    }));

    // Combine existing and new students, then sort
    const allStudents = [...existingStudents.map(s => ({...s, originalRowIndex: null})), ...newStudentsDefaults];
    allStudents.sort((a, b) => a.nama.localeCompare(b.nama));

    // Clear and rewrite all data
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 17).clear();
    }

    const newData = allStudents.map((student, index) => [
      index + 1, student.nama, student.nis, student.nisn, student.jk,
      student.formatif.f1, student.formatif.f2, student.formatif.f3, student.formatif.f4, student.formatif.f5, student.formatif.rataRata,
      student.sumatif.s1, student.sumatif.s2, student.sumatif.s3, student.sumatif.rataRata,
      student.akhirSemester, student.nilaiRapor
    ]);

    if (newData.length > 0) {
      sheet.getRange(2, 1, newData.length, 17).setValues(newData);
    }

    return {
      success: true,
      message: `${studentsData.length} siswa berhasil ditambahkan!`,
      count: studentsData.length
    };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Update student data (Name, NIS, NISN, JK).
 */
function updateStudentData(studentData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    // Use current rowIndex to exclude self from duplicate check
    const tempStudentData = { ...studentData, rowIndex: studentData.rowIndex }; 

    // Validate against all data (excluding itself)
    const validationErrors = validateStudentData([tempStudentData]);
    if (validationErrors.length > 0) {
      throw new Error(validationErrors[0]);
    }

    // Get the original student data before update (to check if name changed)
    const existingData = getAllStudentData();
    const currentStudent = existingData.find(student => student.originalRowIndex === parseInt(studentData.rowIndex));

    if (!currentStudent) {
      throw new Error('Data siswa tidak ditemukan!');
    }

    // Update the student data in original row index
    sheet.getRange(parseInt(studentData.rowIndex), 2, 1, 4).setValues([[
      studentData.nama.trim(), studentData.nis.trim(), studentData.nisn.trim(), studentData.jk
    ]]);

    // Reorganize sheet if name changed (affects sorting)
    if (currentStudent.nama !== studentData.nama.trim()) {
      reorganizeSheetData();
    }

    return { success: true, message: `Data siswa berhasil diperbarui!` };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Delete a student row by index and update numbering.
 */
function deleteStudent(rowIndex) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    const studentName = sheet.getRange(rowIndex, 2).getValue();

    if (!studentName || studentName.toString().trim() === '') {
      throw new Error('Data siswa tidak ditemukan!');
    }

    // Delete the row directly
    sheet.deleteRow(rowIndex);

    // Update sequential numbering after deletion
    updateSequentialNumbering(sheet);

    return { success: true, message: `Siswa ${studentName} berhasil dihapus!` };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}


// ======================================
// === GRADE MANAGEMENT FUNCTIONS (C/U) ===
// ======================================

/**
 * Save grades for Formatif, Sumatif, or Akhir.
 */
function saveGrades(gradeData, type) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    // Get current organized student data (for correct row index)
    const currentStudents = getAllStudentData();
    if (currentStudents.length !== gradeData.length) {
      throw new Error('Data tidak sinkron. Silakan refresh halaman.');
    }

    gradeData.forEach((student, index) => {
      const rowIndex = currentStudents[index].rowIndex; // Use the sequential row index

      if (type === 'formatif') {
        const values = [
          parseFloat(student.formatif.f1) || 0,
          parseFloat(student.formatif.f2) || 0,
          parseFloat(student.formatif.f3) || 0,
          parseFloat(student.formatif.f4) || 0,
          parseFloat(student.formatif.f5) || 0
        ];
        sheet.getRange(rowIndex, 6, 1, 5).setValues([values]); // F1 to F5

        const validValues = values.filter(v => v > 0);
        const average = validValues.length > 0 ? (validValues.reduce((sum, val) => sum + val, 0) / validValues.length).toFixed(1) : 0;
        sheet.getRange(rowIndex, 11).setValue(average); // Rata-rata Formatif

      } else if (type === 'sumatif') {
        const values = [
          parseFloat(student.sumatif.s1) || 0,
          parseFloat(student.sumatif.s2) || 0,
          parseFloat(student.sumatif.s3) || 0
        ];
        sheet.getRange(rowIndex, 12, 1, 3).setValues([values]); // S1 to S3

        const validValues = values.filter(v => v > 0);
        const average = validValues.length > 0 ? (validValues.reduce((sum, val) => sum + val, 0) / validValues.length).toFixed(1) : 0;
        sheet.getRange(rowIndex, 15).setValue(average); // Rata-rata Sumatif

      } else if (type === 'akhir') {
        const finalScore = parseFloat(student.akhirSemester) || 0;
        sheet.getRange(rowIndex, 16).setValue(finalScore); // Nilai Akhir

        // Calculate rapor score: 60% Sumatif Avg + 40% Nilai Akhir
        const sumatifAvg = sheet.getRange(rowIndex, 15).getValue() || 0;
        const nilaiRapor = (sumatifAvg * 0.6) + (finalScore * 0.4);
        sheet.getRange(rowIndex, 17).setValue(nilaiRapor.toFixed(1)); // Nilai Rapor
      }
    });

    return { success: true, message: `Data ${type} berhasil disimpan!` };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Wrapper functions for GAS remote calls
function saveFormatifData(formatifData) { return saveGrades(formatifData, 'formatif'); }
function saveSumatifData(sumatifData) { return saveGrades(sumatifData, 'sumatif'); }
function saveAkhirSemesterData(akhirData) { return saveGrades(akhirData, 'akhir'); }

/**
 * Reset all grades (clear all grade columns).
 */
function resetAllGrades() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    const students = getAllStudentData();
    if (students.length === 0) {
      return { success: true, message: 'Tidak ada data nilai yang direset.' };
    }

    // Clear grade columns (F1 to Nilai Rapor: Column 6 to 17)
    sheet.getRange(2, 6, students.length, 12).clearContent();

    return { success: true, message: 'Semua nilai berhasil direset!' };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}


// ===================================
// === REPORT & EXPORT FUNCTIONS ===
// ===================================

/**
 * Function to get HTML for printing/previewing data
 */
function getPreviewHTML() {
  const settings = getClassAndSubject();
  const students = getAllStudentData();

  let tableRows = '';
  if (students.length === 0) {
    tableRows = '<tr><td colspan="17">Tidak ada data siswa.</td></tr>';
  } else {
    tableRows = students.map(student => `
      <tr>
        <td>${student.no}</td>
        <td class="nama">${student.nama}</td>
        <td>${student.nis || '-'}</td>
        <td>${student.nisn || '-'}</td>
        <td>${student.jk || '-'}</td>
        <td>${student.formatif.f1 || '-'}</td>
        <td>${student.formatif.f2 || '-'}</td>
        <td>${student.formatif.f3 || '-'}</td>
        <td>${student.formatif.f4 || '-'}</td>
        <td>${student.formatif.f5 || '-'}</td>
        <td>${student.formatif.rataRata ? parseFloat(student.formatif.rataRata).toFixed(1) : '-'}</td>
        <td>${student.sumatif.s1 || '-'}</td>
        <td>${student.sumatif.s2 || '-'}</td>
        <td>${student.sumatif.s3 || '-'}</td>
        <td>${student.sumatif.rataRata ? parseFloat(student.sumatif.rataRata).toFixed(1) : '-'}</td>
        <td>${student.akhirSemester || '-'}</td>
        <td>${student.nilaiRapor ? parseFloat(student.nilaiRapor).toFixed(1) : '-'}</td>
      </tr>
    `).join('');
  }


  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>Cetak Daftar Nilai</title>
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Arial, sans-serif; padding: 20px; }
        .header { text-align: center; margin-bottom: 30px; }
        h1 { color: #333; margin-bottom: 10px; font-size: 18px; }
        .info { color: #666; margin-bottom: 5px; font-size: 12px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { padding: 6px; text-align: center; border: 1px solid #ddd; font-size: 11px; }
        th { background-color: #f5f5f5; font-weight: bold; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        .nama { text-align: left !important; }
        
        @media print {
            body { padding: 0; margin: 0; }
            .header { margin-bottom: 20px; }
            h1 { font-size: 16px; }
            .info { font-size: 10px; }
            table { font-size: 10px; }
            th, td { padding: 4px; }
        }
      </style>
    </head>
    <body>
      <div class="header">
        <h1>DAFTAR NILAI - ${settings.subjectName.toUpperCase()}</h1>
        <p class="info">Kelas: ${settings.className} | Tahun Ajaran: ${settings.academicYear} | Semester: ${settings.semester}</p>
        <p class="info">Guru Mata Pelajaran: ${settings.teacherName || '-'}</p>
        <p class="info">Sekolah: ${settings.schoolName}</p>
      </div>

      <table>
        <thead>
          <tr>
            <th rowspan="2">No</th>
            <th rowspan="2" style="width: 150px;">Nama</th>
            <th rowspan="2">NIS</th>
            <th rowspan="2">NISN</th>
            <th rowspan="2">L/P</th>
            <th colspan="6">Nilai Formatif</th>
            <th colspan="4">Nilai Sumatif</th>
            <th rowspan="2">Nilai Akhir</th>
            <th rowspan="2">Nilai Rapor</th>
          </tr>
          <tr>
            <th>F1</th>
            <th>F2</th>
            <th>F3</th>
            <th>F4</th>
            <th>F5</th>
            <th>Rata-rata</th>
            <th>S1</th>
            <th>S2</th>
            <th>S3</th>
            <th>Rata-rata</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows}
        </tbody>
      </table>
      
      <p style="margin-top: 20px; font-size: 10px; color: #999;">Dicetak pada: ${new Date().toLocaleString()}</p>
    </body>
    </html>
  `;
  
  return html;
}

/**
 * Download data as Excel (using the Google Sheet file itself).
 */
function downloadExcel() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const fileName = `Nilai_${getClassAndSubject().subjectName.replace(/\s/g, '_')}_${getClassAndSubject().className.replace(/\s/g, '_')}.xlsx`;

  // Get the spreadsheet blob
  const blob = ss.getBlob().copyBlob();
  blob.setName(fileName);
  
  // Convert blob to base64 string for client-side download
  const base64Content = Utilities.base64Encode(blob.getBytes());
  
  return {
    success: true,
    fileName: fileName,
    content: base64Content
  };
}