import express from "express";
import cors from "cors";
import path from "path";
import fs from "fs";
import { z } from "zod";

const app = express();
const PORT = 3000;

// Enable CORS and JSON / plain text body parsing
app.use(cors());
app.use(express.json());
app.use(express.text({ type: "*/*" }));

// Zod Schemas for Validation
export const studentNameSchema = z.string({
  message: "Nama siswa harus berupa string."
}).trim().min(1, "Nama siswa tidak boleh kosong.");

export const gradeValueSchema = z.union([
  z.number().min(0, "Nilai tidak boleh kurang dari 0.").max(100, "Nilai tidak boleh lebih dari 100."),
  z.string().transform((val, ctx) => {
    if (val.trim() === "") return undefined;
    const num = Number(val);
    if (isNaN(num)) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "Nilai harus berupa angka yang valid."
      });
      return z.NEVER;
    }
    if (num < 0 || num > 100) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "Nilai harus berada antara 0 dan 100."
      });
      return z.NEVER;
    }
    return num;
  }),
  z.null(),
  z.undefined()
]);

export const studentSchema = z.object({
  NISN: z.string().trim().min(1, "NISN wajib diisi."),
  Nama: studentNameSchema,
  Kelas: z.string().optional(),
  NoOrangTua: z.string().optional(),
  Keterangan: z.string().optional()
});

const GRADE_FIELDS = [
  "Nilai Harian/Tugas",
  "Nilai TP1",
  "TP2",
  "TP3",
  "TP4",
  "TP5",
  "Nilai LM1",
  "LM2",
  "LM3",
  "LM4",
  "LM5",
  "Nilai Formatif",
  "Nilai Sumatif",
  "Nilai Sikap",
  "Nilai Kehadiran"
];

function validateRowData(sheetName: string, rowData: Record<string, any>, isUpdate = false): { success: boolean; error?: string } {
  // Validate student name if field 'Nama' is present or required
  if (sheetName === "Daftar Siswa") {
    if (!isUpdate || rowData.Nama !== undefined) {
      const nameValidation = studentNameSchema.safeParse(rowData.Nama);
      if (!nameValidation.success) {
        return {
          success: false,
          error: nameValidation.error.issues.map((i) => i.message).join(", ")
        };
      }
    }
  } else {
    // If 'Nama' is passed in grade sheet or other sheets, validate it if present
    if (rowData.Nama !== undefined && rowData.Nama !== "") {
      const nameValidation = studentNameSchema.safeParse(rowData.Nama);
      if (!nameValidation.success) {
        return {
          success: false,
          error: nameValidation.error.issues.map((i) => i.message).join(", ")
        };
      }
    }
  }

  // Validate grades if present in rowData
  for (const field of GRADE_FIELDS) {
    if (rowData[field] !== undefined && rowData[field] !== null && rowData[field] !== "") {
      const gradeValidation = gradeValueSchema.safeParse(rowData[field]);
      if (!gradeValidation.success) {
        return {
          success: false,
          error: `Validasi nilai pada '${field}' gagal: ` + gradeValidation.error.issues.map((i) => i.message).join(", ")
        };
      }
    }
  }

  return { success: true };
}

// DB Schema
const DB_SCHEMA: Record<string, string[]> = {
  "Daftar Siswa": ["No", "NISN", "Nama", "Kelas", "NoOrangTua", "Keterangan"],
  "Daftar Nilai Matematika": ["No", "NISN", "Nama", "Nilai Harian/Tugas", "Nilai TP1", "TP2", "TP3", "TP4", "TP5", "Nilai LM1", "LM2", "LM3", "LM4", "LM5", "Nilai Formatif", "Nilai Sumatif", "Nilai Sikap", "Nilai Kehadiran", "Keterangan"],
  "Daftar Nilai B.Indonesia": ["No", "NISN", "Nama", "Nilai Harian/Tugas", "Nilai TP1", "TP2", "TP3", "TP4", "TP5", "Nilai LM1", "LM2", "LM3", "LM4", "LM5", "Nilai Formatif", "Nilai Sumatif", "Nilai Sikap", "Nilai Kehadiran", "Keterangan"],
  "IPAS": ["No", "NISN", "Nama", "Nilai Harian/Tugas", "Nilai TP1", "TP2", "TP3", "TP4", "TP5", "Nilai LM1", "LM2", "LM3", "LM4", "LM5", "Nilai Formatif", "Nilai Sumatif", "Nilai Sikap", "Nilai Kehadiran", "Keterangan"],
  "Pendidikan Pancasila": ["No", "NISN", "Nama", "Nilai Harian/Tugas", "Nilai TP1", "TP2", "TP3", "TP4", "TP5", "Nilai LM1", "LM2", "LM3", "LM4", "LM5", "Nilai Formatif", "Nilai Sumatif", "Nilai Sikap", "Nilai Kehadiran", "Keterangan"],
  "Seni": ["No", "NISN", "Nama", "Nilai Harian/Tugas", "Nilai TP1", "TP2", "TP3", "TP4", "TP5", "Nilai LM1", "LM2", "LM3", "LM4", "LM5", "Nilai Formatif", "Nilai Sumatif", "Nilai Sikap", "Nilai Kehadiran", "Keterangan"],
  "B Sunda": ["No", "NISN", "Nama", "Nilai Harian/Tugas", "Nilai TP1", "TP2", "TP3", "TP4", "TP5", "Nilai LM1", "LM2", "LM3", "LM4", "LM5", "Nilai Formatif", "Nilai Sumatif", "Nilai Sikap", "Nilai Kehadiran", "Keterangan"],
  "Kokurikuler": ["No", "NISN", "Nama", "Nilai Harian/Tugas", "Nilai TP1", "TP2", "TP3", "TP4", "TP5", "Nilai LM1", "LM2", "LM3", "LM4", "LM5", "Nilai Formatif", "Nilai Sumatif", "Nilai Sikap", "Nilai Kehadiran", "Keterangan"],
  "Catatan Wali Kelas": ["No", "NISN", "Nama", "Kelas", "NoOrangTua", "Catatan"]
};

const DATA_FILE = process.env.VERCEL
  ? path.join("/tmp", "db_data.json")
  : path.join(process.cwd(), "db_data.json");

type DBStore = Record<string, Record<string, any>[]>;

function loadDatabase(): DBStore {
  try {
    if (fs.existsSync(DATA_FILE)) {
      const fileData = fs.readFileSync(DATA_FILE, "utf-8");
      return JSON.parse(fileData);
    }
    const localSeed = path.join(process.cwd(), "db_data.json");
    if (fs.existsSync(localSeed)) {
      const fileData = fs.readFileSync(localSeed, "utf-8");
      return JSON.parse(fileData);
    }
  } catch (err) {
    console.error("Error reading database file:", err);
  }

  // Initial seed data if file doesn't exist
  const initialDb: DBStore = {};
  for (const sheetName in DB_SCHEMA) {
    initialDb[sheetName] = [];
  }

  // Seed Students
  const sampleStudents = [
    { No: 1, NISN: "1001", Nama: "Ahmad Rizky", Kelas: "VI-A", NoOrangTua: "081234567890", Keterangan: "Ketua Kelas" },
    { No: 2, NISN: "1002", Nama: "Siti Nurhaliza", Kelas: "VI-A", NoOrangTua: "081234567891", Keterangan: "Aktif Berorganisasi" },
    { No: 3, NISN: "1003", Nama: "Budi Santoso", Kelas: "VI-A", NoOrangTua: "081234567892", Keterangan: "Berprestasi Olahraga" },
    { No: 4, NISN: "1004", Nama: "Dewi Lestari", Kelas: "VI-A", NoOrangTua: "081234567893", Keterangan: "Juara Puisi" },
    { No: 5, NISN: "1005", Nama: "Eko Prasetyo", Kelas: "VI-A", NoOrangTua: "081234567894", Keterangan: "Anggota Pramuka" }
  ];
  initialDb["Daftar Siswa"] = sampleStudents;

  // Seed Grades
  const subjects = [
    "Daftar Nilai Matematika",
    "Daftar Nilai B.Indonesia",
    "IPAS",
    "Pendidikan Pancasila",
    "Seni",
    "B Sunda",
    "Kokurikuler"
  ];

  subjects.forEach((subj) => {
    initialDb[subj] = sampleStudents.map((s, idx) => ({
      No: idx + 1,
      NISN: s.NISN,
      Nama: s.Nama,
      "Nilai Harian/Tugas": 80 + (idx % 3) * 5,
      "Nilai TP1": 85,
      TP2: 82,
      TP3: 88,
      TP4: 80,
      TP5: 84,
      "Nilai LM1": 83,
      LM2: 85,
      LM3: 80,
      LM4: 86,
      LM5: 82,
      "Nilai Formatif": 85,
      "Nilai Sumatif": 84,
      "Nilai Sikap": 88,
      "Nilai Kehadiran": 95,
      Keterangan: "Sangat baik"
    }));
  });

  // Seed Notes
  initialDb["Catatan Wali Kelas"] = [
    { No: 1, NISN: "1001", Nama: "Ahmad Rizky", Kelas: "VI-A", NoOrangTua: "081234567890", Catatan: "Ananda Ahmad menunjukkan kepemimpinan yang sangat baik sebagai ketua kelas. Pertahankan kebiasaan belajar yang teratur." },
    { No: 2, NISN: "1002", Nama: "Siti Nurhaliza", Kelas: "VI-A", NoOrangTua: "081234567891", Catatan: "Siti sangat aktif dalam diskusi kelas dan tugas kelompok. Perlu ditingkatkan ketelitian saat mengerjakan soal matematika." },
    { No: 3, NISN: "1003", Nama: "Budi Santoso", Kelas: "VI-A", NoOrangTua: "081234567892", Catatan: "Budi memiliki potensi luar biasa. Disiplin latihan fisik sangat baik, disarankan menambah porsi belajar mandiri di rumah." }
  ];

  saveDatabase(initialDb);
  return initialDb;
}

function saveDatabase(db: DBStore) {
  try {
    fs.writeFileSync(DATA_FILE, JSON.stringify(db, null, 2), "utf-8");
  } catch (err) {
    console.error("Error saving database file:", err);
  }
}

let db = loadDatabase();

function initializeDatabase() {
  const log: string[] = [];
  for (const sheetName in DB_SCHEMA) {
    if (!db[sheetName]) {
      db[sheetName] = [];
      log.push(`Sheet '${sheetName}' berhasil dibuat.`);
    } else {
      log.push(`Sheet '${sheetName}' sudah ada dan sesuai.`);
    }
  }
  saveDatabase(db);
  return { success: true, message: "Inisialisasi database selesai.", log };
}

function readData(sheetName: string): { success: boolean; error?: string; data?: Record<string, any>[] } {
  if (!db[sheetName]) {
    return { success: false, error: `Sheet '${sheetName}' tidak ditemukan.` };
  }
  const data = db[sheetName].map((row, idx) => ({ ...row, _rowIndex: idx + 2 }));
  return { success: true, data };
}

function createData(sheetName: string, rowData: Record<string, any>) {
  if (!db[sheetName]) {
    return { success: false, error: `Sheet '${sheetName}' tidak ditemukan.` };
  }

  const validation = validateRowData(sheetName, rowData, false);
  if (!validation.success) {
    return { success: false, error: validation.error };
  }

  const sheetData = db[sheetName];
  if (rowData.NISN) {
    const exists = sheetData.some(
      (r) => String(r.NISN).trim() === String(rowData.NISN).trim()
    );
    if (exists) {
      return { success: false, error: `NISN ${rowData.NISN} sudah terdaftar di '${sheetName}'.` };
    }
  }

  let maxNo = 0;
  sheetData.forEach((r) => {
    const val = parseInt(r.No, 10);
    if (!isNaN(val) && val > maxNo) maxNo = val;
  });
  rowData["No"] = maxNo + 1;

  sheetData.push(rowData);
  saveDatabase(db);
  return { success: true, message: `Data berhasil ditambahkan ke ${sheetName}`, data: rowData };
}

function updateData(sheetName: string, rowData: Record<string, any>) {
  if (!db[sheetName]) {
    return { success: false, error: `Sheet '${sheetName}' tidak ditemukan.` };
  }

  const validation = validateRowData(sheetName, rowData, true);
  if (!validation.success) {
    return { success: false, error: validation.error };
  }

  const sheetData = db[sheetName];
  let index = -1;

  if (rowData.NISN) {
    index = sheetData.findIndex(
      (r) => String(r.NISN).trim() === String(rowData.NISN).trim()
    );
  }

  if (index === -1 && rowData.No) {
    index = sheetData.findIndex(
      (r) => String(r.No).trim() === String(rowData.No).trim()
    );
  }

  if (index === -1) {
    return { success: false, error: `Data tidak ditemukan di '${sheetName}' untuk di-update.` };
  }

  sheetData[index] = { ...sheetData[index], ...rowData };
  saveDatabase(db);
  return { success: true, message: `Data berhasil diperbarui di ${sheetName}` };
}

function deleteData(sheetName: string, key: any) {
  if (!db[sheetName]) {
    return { success: false, error: `Sheet '${sheetName}' tidak ditemukan.` };
  }

  const sheetData = db[sheetName];
  const initialLength = sheetData.length;

  db[sheetName] = sheetData.filter(
    (r) =>
      String(r.NISN).trim() !== String(key).trim() &&
      String(r.No).trim() !== String(key).trim()
  );

  if (db[sheetName].length === initialLength) {
    return { success: false, error: "Data tidak ditemukan untuk dihapus." };
  }

  saveDatabase(db);
  return { success: true, message: `Data berhasil dihapus dari ${sheetName}` };
}

function getSiswaList() {
  const res = readData("Daftar Siswa");
  if (!res.success || !res.data) return res;

  const list = res.data.map((s: any) => ({
    NISN: s.NISN,
    Nama: s.Nama,
    Kelas: s.Kelas,
    NoOrangTua: s.NoOrangTua
  }));
  return { success: true, data: list };
}

function getStudentSummary(nisn: string) {
  const siswaRes = readData("Daftar Siswa");
  if (!siswaRes.success || !siswaRes.data) return siswaRes;

  const student = siswaRes.data.find(
    (s: any) => String(s.NISN).trim() === String(nisn).trim()
  );

  if (!student) {
    return { success: false, error: `Siswa dengan NISN ${nisn} tidak ditemukan.` };
  }

  const subjects = [
    "Daftar Nilai Matematika",
    "Daftar Nilai B.Indonesia",
    "IPAS",
    "Pendidikan Pancasila",
    "Seni",
    "B Sunda",
    "Kokurikuler"
  ];

  const grades: Record<string, any> = {};
  let totalSikaps = 0;
  let totalKehadirans = 0;
  let totalCognitives = 0;
  let subjectScoreCount = 0;
  let sikapCount = 0;
  let kehadiranCount = 0;

  subjects.forEach((subject) => {
    const res = readData(subject);
    const subName = subject.replace("Daftar Nilai ", "");

    grades[subName] = null;
    if (res.success && res.data) {
      const row = res.data.find(
        (r: any) => String(r.NISN).trim() === String(nisn).trim()
      );

      if (row) {
        const harian = parseFloat(row["Nilai Harian/Tugas"]) || 0;
        const tp1 = parseFloat(row["Nilai TP1"]) || 0;
        const tp2 = parseFloat(row["TP2"]) || 0;
        const tp3 = parseFloat(row["TP3"]) || 0;
        const tp4 = parseFloat(row["TP4"]) || 0;
        const tp5 = parseFloat(row["TP5"]) || 0;
        const lm1 = parseFloat(row["Nilai LM1"]) || 0;
        const lm2 = parseFloat(row["LM2"]) || 0;
        const lm3 = parseFloat(row["LM3"]) || 0;
        const lm4 = parseFloat(row["LM4"]) || 0;
        const lm5 = parseFloat(row["LM5"]) || 0;
        const formatif = parseFloat(row["Nilai Formatif"]) || 0;
        const sumatif = parseFloat(row["Nilai Sumatif"]) || 0;
        const sikap = parseFloat(row["Nilai Sikap"]) || 0;
        const kehadiran = parseFloat(row["Nilai Kehadiran"]) || 0;

        let tpCount = 0, tpSum = 0;
        [tp1, tp2, tp3, tp4, tp5].forEach((v) => { if (v > 0) { tpSum += v; tpCount++; } });
        const tpAvg = tpCount > 0 ? tpSum / tpCount : 0;

        let lmCount = 0, lmSum = 0;
        [lm1, lm2, lm3, lm4, lm5].forEach((v) => { if (v > 0) { lmSum += v; lmCount++; } });
        const lmAvg = lmCount > 0 ? lmSum / lmCount : 0;

        const cognitValues: number[] = [];
        if (harian > 0) cognitValues.push(harian);
        if (tpAvg > 0) cognitValues.push(tpAvg);
        if (lmAvg > 0) cognitValues.push(lmAvg);
        if (formatif > 0) cognitValues.push(formatif);
        if (sumatif > 0) cognitValues.push(sumatif);

        let avgCognitive = 0;
        if (cognitValues.length > 0) {
          avgCognitive = cognitValues.reduce((a, b) => a + b, 0) / cognitValues.length;
        }

        grades[subName] = {
          harian,
          tpAvg: tpAvg.toFixed(1),
          lmAvg: lmAvg.toFixed(1),
          formatif,
          sumatif,
          sikap,
          kehadiran,
          rataRata: avgCognitive.toFixed(1),
          keterangan: row["Keterangan"] || "-"
        };

        if (avgCognitive > 0) {
          totalCognitives += avgCognitive;
          subjectScoreCount++;
        }
        if (sikap > 0) {
          totalSikaps += sikap;
          sikapCount++;
        }
        if (kehadiran > 0) {
          totalKehadirans += kehadiran;
          kehadiranCount++;
        }
      }
    }
  });

  const catatanRes = readData("Catatan Wali Kelas");
  let catatan = "-";
  if (catatanRes.success && catatanRes.data) {
    const row = catatanRes.data.find(
      (r: any) => String(r.NISN).trim() === String(nisn).trim()
    );
    if (row && row.Catatan) {
      catatan = row.Catatan;
    }
  }

  const classAverage = subjectScoreCount > 0 ? (totalCognitives / subjectScoreCount).toFixed(1) : "0";
  const sikapAverage = sikapCount > 0 ? (totalSikaps / sikapCount).toFixed(1) : "0";
  const kehadiranAverage = kehadiranCount > 0 ? (totalKehadirans / kehadiranCount).toFixed(1) : "0";

  return {
    success: true,
    data: {
      siswa: student,
      nilai: grades,
      catatanWaliKelas: catatan,
      ringkasan: {
        rataRataKognitif: classAverage,
        rataRataSikap: sikapAverage,
        rataRataKehadiran: kehadiranAverage
      }
    }
  };
}

function getDashboardStats() {
  const siswaRes = readData("Daftar Siswa");
  if (!siswaRes.success || !siswaRes.data) return siswaRes;

  const totalSiswa = siswaRes.data.length;

  const subjects = [
    "Daftar Nilai Matematika",
    "Daftar Nilai B.Indonesia",
    "IPAS",
    "Pendidikan Pancasila",
    "Seni",
    "B Sunda",
    "Kokurikuler"
  ];

  let totalCognitivesSum = 0;
  let totalScoreCount = 0;
  const subjectAverages: Record<string, string> = {};

  subjects.forEach((subject) => {
    const res = readData(subject);
    const subName = subject.replace("Daftar Nilai ", "");

    let subSum = 0;
    let subCount = 0;

    if (res.success && res.data) {
      res.data.forEach((row: any) => {
        const harian = parseFloat(row["Nilai Harian/Tugas"]) || 0;
        const formatif = parseFloat(row["Nilai Formatif"]) || 0;
        const sumatif = parseFloat(row["Nilai Sumatif"]) || 0;

        let tpCount = 0, tpSum = 0;
        ["Nilai TP1", "TP2", "TP3", "TP4", "TP5"].forEach((col) => {
          const v = parseFloat(row[col]) || 0;
          if (v > 0) { tpSum += v; tpCount++; }
        });
        const tpAvg = tpCount > 0 ? tpSum / tpCount : 0;

        let lmCount = 0, lmSum = 0;
        ["Nilai LM1", "LM2", "LM3", "LM4", "LM5"].forEach((col) => {
          const v = parseFloat(row[col]) || 0;
          if (v > 0) { lmSum += v; lmCount++; }
        });
        const lmAvg = lmCount > 0 ? lmSum / lmCount : 0;

        const cognitiveTerms: number[] = [];
        if (harian > 0) cognitiveTerms.push(harian);
        if (tpAvg > 0) cognitiveTerms.push(tpAvg);
        if (lmAvg > 0) cognitiveTerms.push(lmAvg);
        if (formatif > 0) cognitiveTerms.push(formatif);
        if (sumatif > 0) cognitiveTerms.push(sumatif);

        if (cognitiveTerms.length > 0) {
          const avgCognitive = cognitiveTerms.reduce((a, b) => a + b, 0) / cognitiveTerms.length;
          subSum += avgCognitive;
          subCount++;
        }
      });
    }

    const avgScore = subCount > 0 ? subSum / subCount : 0;
    subjectAverages[subName] = avgScore.toFixed(1);

    if (avgScore > 0) {
      totalCognitivesSum += avgScore;
      totalScoreCount++;
    }
  });

  const classGeneralAverage = totalScoreCount > 0 ? (totalCognitivesSum / totalScoreCount).toFixed(1) : "0";

  const catatanRes = readData("Catatan Wali Kelas");
  let totalCatatan = 0;
  if (catatanRes.success && catatanRes.data) {
    catatanRes.data.forEach((row: any) => {
      if (row.Catatan && row.Catatan.trim() !== "" && row.Catatan.trim() !== "-") {
        totalCatatan++;
      }
    });
  }

  return {
    success: true,
    data: {
      totalSiswa,
      rataRataKelas: classGeneralAverage,
      totalBimbingan: totalCatatan,
      grafikRataRata: subjectAverages
    }
  };
}

function executeAction(request: any) {
  try {
    const action = request.action;

    if (action === "initDb") {
      return initializeDatabase();
    }

    switch (action) {
      case "read":
        return readData(request.sheetName);
      case "create":
        return createData(request.sheetName, request.data);
      case "update":
        return updateData(request.sheetName, request.data);
      case "delete":
        return deleteData(
          request.sheetName,
          request.nisn || request.no || (request.data && (request.data.NISN || request.data.No))
        );
      case "getSiswaList":
        return getSiswaList();
      case "getStudentSummary":
        return getStudentSummary(request.nisn);
      case "getDashboardStats":
        return getDashboardStats();
      default:
        return { success: false, error: `Aksi tidak dikenal: ${action}` };
    }
  } catch (error: any) {
    return { success: false, error: error.toString() };
  }
}

// Helper to safely parse incoming request body
function parseRequestBody(req: express.Request): any {
  if (typeof req.body === "string" && req.body.trim()) {
    try {
      return JSON.parse(req.body);
    } catch {
      return {};
    }
  } else if (typeof req.body === "object" && req.body !== null) {
    return req.body;
  }
  return {};
}

// Router handler for /exec or /api/exec or /api
const handleRequest = (req: express.Request, res: express.Response) => {
  const bodyData = parseRequestBody(req);
  const request = { ...req.query, ...bodyData };
  const result = executeAction(request);
  res.json(result);
};

app.all("/exec", handleRequest);
app.all("/api/exec", handleRequest);
app.all("/api", handleRequest);

// REST API Endpoints for Student Records (/api/students & /api/siswa)
const studentRouter = express.Router();

// GET /api/students - Get all student records
studentRouter.get("/", (_req, res) => {
  const result = readData("Daftar Siswa");
  if (!result.success) {
    res.status(400).json(result);
    return;
  }
  res.json(result);
});

// GET /api/students/:nisn - Get a specific student by NISN (including summary & grades)
studentRouter.get("/:nisn", (req, res) => {
  const nisn = req.params.nisn;
  const summary = getStudentSummary(nisn);
  if (!summary.success) {
    res.status(404).json(summary);
    return;
  }
  res.json(summary);
});

// POST /api/students - Create a new student record
studentRouter.post("/", (req, res) => {
  const body = parseRequestBody(req);
  const studentData = {
    NISN: body.NISN || body.nisn || "",
    Nama: body.Nama || body.nama || "",
    Kelas: body.Kelas || body.kelas || "",
    NoOrangTua: body.NoOrangTua || body.noOrangTua || body.phone || "",
    Keterangan: body.Keterangan || body.keterangan || ""
  };

  if (!studentData.NISN || !studentData.Nama) {
    res.status(400).json({
      success: false,
      error: "NISN dan Nama siswa wajib diisi."
    });
    return;
  }

  const result = createData("Daftar Siswa", studentData);
  if (!result.success) {
    res.status(400).json(result);
    return;
  }
  res.status(201).json(result);
});

// PUT /api/students/:nisn or /api/students - Update an existing student record
studentRouter.put("/:nisn?", (req, res) => {
  const body = parseRequestBody(req);
  const nisn = req.params.nisn || body.NISN || body.nisn;

  if (!nisn) {
    res.status(400).json({ success: false, error: "NISN diperlukan untuk update data siswa." });
    return;
  }

  const updatePayload: Record<string, any> = {
    NISN: String(nisn).trim()
  };

  if (body.Nama !== undefined || body.nama !== undefined) updatePayload.Nama = body.Nama ?? body.nama;
  if (body.Kelas !== undefined || body.kelas !== undefined) updatePayload.Kelas = body.Kelas ?? body.kelas;
  if (body.NoOrangTua !== undefined || body.noOrangTua !== undefined) updatePayload.NoOrangTua = body.NoOrangTua ?? body.noOrangTua;
  if (body.Keterangan !== undefined || body.keterangan !== undefined) updatePayload.Keterangan = body.Keterangan ?? body.keterangan;

  const result = updateData("Daftar Siswa", updatePayload);
  if (!result.success) {
    res.status(400).json(result);
    return;
  }
  res.json(result);
});

// DELETE /api/students/:nisn - Delete a student record
studentRouter.delete("/:nisn?", (req, res) => {
  const body = parseRequestBody(req);
  const nisn = req.params.nisn || req.query.nisn || body.NISN || body.nisn;

  if (!nisn) {
    res.status(400).json({ success: false, error: "NISN diperlukan untuk hapus data siswa." });
    return;
  }

  const result = deleteData("Daftar Siswa", nisn);
  if (!result.success) {
    res.status(404).json(result);
    return;
  }
  res.json(result);
});

app.use("/api/students", studentRouter);
app.use("/api/siswa", studentRouter);

// Serve static HTML index at root
app.get("/", (_req, res) => {
  const indexFile = fs.existsSync(path.join(process.cwd(), "index.html"))
    ? path.join(process.cwd(), "index.html")
    : path.join(process.cwd(), "Index.html");
  res.sendFile(indexFile);
});

export default app;

if (!process.env.VERCEL) {
  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server is running on http://0.0.0.0:${PORT}`);
  });
}
