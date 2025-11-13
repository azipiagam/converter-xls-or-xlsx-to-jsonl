#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

// Path tetap untuk memindahkan file yang sudah di-convert
const DONE_INPUT_PATH = "E:\\Azi\\NetBackup\\Done - Input";

// Konstanta untuk batasan file
const MAX_FILE_SIZE_MB = 99;
const MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024;

// Fungsi untuk clean up header names
function cleanHeader(header) {
  if (typeof header !== "string") {
    header = String(header);
  }

  return header
    .toLowerCase() // Convert ke lowercase
    .replace(/[^a-z0-9]+/g, "_") // Ganti semua non-alphanumeric dengan underscore
    .replace(/^_+|_+$/g, "") // Hapus underscore di awal dan akhir
    .replace(/_+/g, "_"); // Ganti multiple underscore dengan single underscore
}

// Fungsi untuk mengkonversi Excel date number ke Date object
function excelDateToJSDate(excelDate) {
  // Excel menggunakan 1900-01-01 sebagai tanggal dasar (serial number 1)
  // Tapi Excel salah menganggap 1900 adalah tahun kabisat
  const excelEpoch = new Date(1899, 11, 30); // 30 Desember 1899
  return new Date(excelEpoch.getTime() + excelDate * 24 * 60 * 60 * 1000);
}

// Fungsi untuk mengecek apakah nilai adalah Excel date
function isExcelDate(value) {
  // Cek apakah value adalah number dan dalam range yang wajar untuk tanggal
  if (typeof value === "number" && value > 1 && value < 2958466) {
    // 1900-01-01 to 9999-12-31
    return true;
  }
  return false;
}

// Mapping bulan Indonesia ke angka
const indonesianMonths = {
  jan: 1,
  januari: 1,
  feb: 2,
  februari: 2,
  mar: 3,
  maret: 3,
  apr: 4,
  april: 4,
  mei: 5,
  may: 5,
  jun: 6,
  juni: 6,
  jul: 7,
  juli: 7,
  agu: 8,
  agustus: 8,
  sep: 9,
  september: 9,
  okt: 10,
  oktober: 10,
  nov: 11,
  november: 11,
  des: 12,
  desember: 12,
  dec: 12,
};

// Fungsi untuk mengecek tipe data temporal
function detectTemporalType(value) {
  if (typeof value !== "string") return null;

  const trimmed = value.trim();

  // Pattern untuk format Indonesia: "01 Agu 2025 23:51"
  const indonesianDateTimePattern =
    /^(\d{1,2})\s+(\w+)\s+(\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/i;

  // Pattern untuk waktu saja: "23:51" atau "23:51:30"
  const timeOnlyPattern = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/;

  // Pattern untuk tanggal standar
  const datePatterns = [
    { pattern: /^\d{4}-\d{1,2}-\d{1,2}$/, type: "date" }, // YYYY-MM-DD
    { pattern: /^\d{1,2}\/\d{1,2}\/\d{4}$/, type: "date" }, // MM/DD/YYYY atau DD/MM/YYYY
    { pattern: /^\d{1,2}-\d{1,2}-\d{4}$/, type: "date" }, // MM-DD-YYYY atau DD-MM-YYYY
    { pattern: /^\d{1,2}\.\d{1,2}\.\d{4}$/, type: "date" }, // DD.MM.YYYY
    { pattern: /^\d{4}\/\d{1,2}\/\d{1,2}$/, type: "date" }, // YYYY/MM/DD
    // NEW: Tambahan untuk format 2-digit year
    { pattern: /^\d{1,2}\/\d{1,2}\/\d{2}$/, type: "date" }, // MM/DD/YY atau DD/MM/YY
    { pattern: /^\d{1,2}-\d{1,2}-\d{2}$/, type: "date" }, // MM-DD-YY atau DD-MM-YY
    { pattern: /^\d{1,2}\.\d{1,2}\.\d{2}$/, type: "date" }, // DD.MM.YY
  ];

  // Pattern untuk datetime standar
  const datetimePatterns = [
    {
      pattern: /^\d{4}-\d{1,2}-\d{1,2}\s+\d{1,2}:\d{1,2}(?::\d{1,2})?/,
      type: "datetime",
    },
    {
      pattern: /^\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{1,2}(?::\d{1,2})?/,
      type: "datetime",
    },
    {
      pattern: /^\d{1,2}-\d{1,2}-\d{4}\s+\d{1,2}:\d{1,2}(?::\d{1,2})?/,
      type: "datetime",
    },
    // NEW: Tambahan untuk format 2-digit year datetime
    {
      pattern: /^\d{1,2}\/\d{1,2}\/\d{2}\s+\d{1,2}:\d{1,2}(?::\d{1,2})?/,
      type: "datetime",
    },
    {
      pattern: /^\d{1,2}-\d{1,2}-\d{2}\s+\d{1,2}:\d{1,2}(?::\d{1,2})?/,
      type: "datetime",
    },
  ];

  // Cek format Indonesia
  if (indonesianDateTimePattern.test(trimmed)) {
    const match = trimmed.match(indonesianDateTimePattern);
    if (match[4]) {
      // Ada komponen waktu
      return "datetime";
    } else {
      return "date";
    }
  }

  // Cek waktu saja
  if (timeOnlyPattern.test(trimmed)) {
    return "time";
  }

  // Cek datetime patterns
  for (const { pattern, type } of datetimePatterns) {
    if (pattern.test(trimmed)) {
      return type;
    }
  }

  // Cek date patterns
  for (const { pattern, type } of datePatterns) {
    if (pattern.test(trimmed)) {
      return type;
    }
  }

  return null;
}

// NEW: Fungsi untuk mengkonversi 2-digit year menjadi 4-digit year
function convert2DigitYear(year) {
  const currentYear = new Date().getFullYear();
  const currentCentury = Math.floor(currentYear / 100) * 100;
  const currentYearInCentury = currentYear % 100;

  // Rule: Jika year <= current year dalam century ini, anggap century ini
  // Jika year > current year dalam century ini, anggap century sebelumnya
  if (year <= currentYearInCentury) {
    return currentCentury + year;
  } else {
    return currentCentury - 100 + year;
  }
}

// Fungsi untuk parse format Indonesia
function parseIndonesianDate(value) {
  const indonesianPattern =
    /^(\d{1,2})\s+(\w+)\s+(\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/i;
  const match = value.match(indonesianPattern);

  if (!match) return null;

  const day = parseInt(match[1]);
  const monthName = match[2].toLowerCase();
  const year = parseInt(match[3]);
  const hour = match[4] ? parseInt(match[4]) : 0;
  const minute = match[5] ? parseInt(match[5]) : 0;
  const second = match[6] ? parseInt(match[6]) : 0;

  const monthNumber = indonesianMonths[monthName];
  if (!monthNumber) return null;

  return new Date(year, monthNumber - 1, day, hour, minute, second);
}

// ENHANCED: Fungsi untuk mengkonversi berbagai format tanggal ke format standar
function convertToStandardDate(value) {
  try {
    let dateObj;
    let formatType;

    // Jika value adalah Excel date number
    if (isExcelDate(value)) {
      dateObj = excelDateToJSDate(value);
      formatType = "datetime"; // Default untuk Excel date
    }
    // Jika value adalah string, deteksi tipe temporal
    else if (typeof value === "string") {
      formatType = detectTemporalType(value);

      if (!formatType) {
        return value; // Bukan format temporal, kembalikan nilai asli
      }

      const trimmed = value.trim();

      // Handle format Indonesia
      const indonesianDate = parseIndonesianDate(trimmed);
      if (indonesianDate) {
        dateObj = indonesianDate;
      }
      // Handle time only format
      else if (formatType === "time") {
        const timeMatch = trimmed.match(/^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
        if (timeMatch) {
          const hour = parseInt(timeMatch[1]);
          const minute = parseInt(timeMatch[2]);
          const second = timeMatch[3] ? parseInt(timeMatch[3]) : 0;

          // Return formatted time only
          return `${String(hour).padStart(2, "0")}:${String(minute).padStart(
            2,
            "0"
          )}:${String(second).padStart(2, "0")}`;
        }
      }
      // Handle standard date formats
      else {
        // NEW: Enhanced parsing untuk 2-digit year
        const parts = trimmed.split(/[\/\-\.\s]/);

        // Cek jika ada komponen waktu (untuk datetime)
        const hasTime = /\d{1,2}:\d{1,2}/.test(trimmed);
        let timePart = null;

        if (hasTime) {
          const timeMatch = trimmed.match(/(\d{1,2}:\d{1,2}(?::\d{1,2})?)/);
          if (timeMatch) {
            timePart = timeMatch[1];
          }
        }

        if (parts.length >= 3) {
          let day, month, year;

          // Parse date parts
          const part1 = parseInt(parts[0]);
          const part2 = parseInt(parts[1]);
          let part3 = parseInt(parts[2]);

          // NEW: Handle 2-digit year conversion
          if (part3 < 100) {
            part3 = convert2DigitYear(part3);
            console.log(`🔄 Konversi 2-digit year: ${parts[2]} → ${part3}`);
          }

          // Determine date format based on values
          if (part3 > 31) {
            // part3 is year
            year = part3;

            // Determine if MM/DD/YYYY or DD/MM/YYYY
            if (part1 > 12) {
              // part1 must be day (DD/MM/YYYY)
              day = part1;
              month = part2;
              console.log(`📅 Format detected: DD/MM/YYYY → ${day}/${month}/${year}`);
            } else if (part2 > 12) {
              // part2 must be day (MM/DD/YYYY)
              month = part1;
              day = part2;
              console.log(`📅 Format detected: MM/DD/YYYY → ${month}/${day}/${year}`);
            } else {
              // FIXED: Ambiguous case - DEFAULT TO MM/DD/YYYY (US FORMAT)
              month = part1;
              day = part2;
              console.log(`📅 Format ambiguous, using US format: MM/DD/YYYY → ${month}/${day}/${year}`);
            }
          } else {
            // Try original parsing logic as fallback
            dateObj = new Date(trimmed);
            if (isNaN(dateObj.getTime())) {
              // Fallback dengan asumsi DD/MM/YY
              day = part1;
              month = part2;
              year = part3;
            }
          }

          if (day !== undefined && month !== undefined && year !== undefined) {
            dateObj = new Date(year, month - 1, day);

            // Add time if present
            if (timePart && hasTime) {
              const timeMatch = timePart.match(
                /(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/
              );
              if (timeMatch) {
                const hour = parseInt(timeMatch[1]);
                const minute = parseInt(timeMatch[2]);
                const second = timeMatch[3] ? parseInt(timeMatch[3]) : 0;
                dateObj.setHours(hour, minute, second);
              }
            }
          }
        }

        // Fallback to standard Date parsing if custom parsing failed
        if (!dateObj || isNaN(dateObj.getTime())) {
          dateObj = new Date(trimmed);
        }
      }
    }
    // Jika value sudah berupa Date object
    else if (value instanceof Date) {
      dateObj = value;
      formatType = "datetime"; // Default
    } else {
      return value; // Bukan tanggal, kembalikan nilai asli
    }

    // Handle time only - sudah di-handle di atas
    if (formatType === "time") {
      return value; // Sudah di-return di atas
    }

    // Cek apakah dateObj valid
    if (!dateObj || isNaN(dateObj.getTime())) {
      return value; // Jika tidak valid, kembalikan nilai asli
    }

    // Format berdasarkan tipe
    const year = dateObj.getFullYear();
    const month = String(dateObj.getMonth() + 1).padStart(2, "0");
    const day = String(dateObj.getDate()).padStart(2, "0");
    const hours = String(dateObj.getHours()).padStart(2, "0");
    const minutes = String(dateObj.getMinutes()).padStart(2, "0");
    const seconds = String(dateObj.getSeconds()).padStart(2, "0");

    if (formatType === "date") {
      // ENHANCED: Cek apakah ada komponen waktu yang tidak nol
      if (
        dateObj.getHours() === 0 &&
        dateObj.getMinutes() === 0 &&
        dateObj.getSeconds() === 0
      ) {
        return `${year}-${month}-${day}`;
      } else {
        return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
      }
    } else {
      // datetime format
      return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
    }
  } catch (error) {
    console.warn(
      `⚠️  Gagal konversi tanggal untuk nilai: ${value}`,
      error.message
    );
    return value; // Kembalikan nilai asli jika gagal
  }
}

// ===== UPDATE FUNGSI INI (support negative currency) =====
// Fungsi untuk parse nominal mata uang ke float (termasuk negative values)
function parseCurrencyToFloat(value) {
  if (typeof value !== "string") return value;

  const trimmed = value.trim();
  
  // Pattern untuk berbagai format mata uang (dengan support minus di depan atau di tengah):
  // -Rp5,793,200.00 atau Rp-5,793,200.00 atau -$5,793,200.00
  // Support format:
  // 1. Minus di depan currency: -Rp5,793,200.00 atau -$5,793,200.00
  // 2. Minus setelah currency: Rp-5,793,200.00 atau $-5,793,200.00
  // 3. Minus di dalam kurung: (Rp5,793,200.00) atau ($5,793,200.00)
  
  // Pattern 1 & 2: Minus di depan atau setelah currency symbol
  const currencyPattern = /^-?([A-Z]{3}|Rp\.?|[$€£¥₹₽₩฿RM]|S\$)\s?-?([\d,]+(?:\.\d+)?)$/i;
  
  // Pattern 3: Format kurung untuk negative (accounting format)
  const accountingPattern = /^\(([A-Z]{3}|Rp\.?|[$€£¥₹₽₩฿RM]|S\$)?\s?([\d,]+(?:\.\d+)?)\)$/i;
  
  // Pattern untuk angka dengan koma sebagai thousand separator (tanpa symbol)
  const numberWithCommaPattern = /^-?[\d,]+\.\d{2}$/;
  
  // Pattern untuk format Eropa (titik sebagai thousand separator, koma sebagai decimal)
  // Contoh: -5.793.200,00 atau (5.793.200,00)
  const europeanPattern = /^-?([A-Z]{3}|Rp\.?|[$€£¥₹₽₩฿RM]|S\$)?\s?-?([\d.]+,\d{2})$/i;
  const europeanAccountingPattern = /^\(([A-Z]{3}|Rp\.?|[$€£¥₹₽₩฿RM]|S\$)?\s?([\d.]+,\d{2})\)$/i;
  
  // Check untuk format accounting (kurung)
  if (accountingPattern.test(trimmed)) {
    const match = trimmed.match(accountingPattern);
    if (match) {
      const currencySymbol = match[1] || "";
      const numberPart = match[2];
      const floatValue = parseFloat(numberPart.replace(/,/g, ""));
      
      if (!isNaN(floatValue)) {
        // Negative karena dalam kurung
        const rounded = Math.round(floatValue * -100) / 100;
        console.log(`💰 Konversi ${currencySymbol || "Currency"} (negative/accounting): "${trimmed}" → ${rounded}`);
        return rounded;
      }
    }
  }
  
  // Check untuk format standard dengan currency symbol
  if (currencyPattern.test(trimmed)) {
    const match = trimmed.match(currencyPattern);
    if (match) {
      const currencySymbol = match[1];
      const numberPart = match[2];
      
      // Cek apakah ada minus di string asli
      const isNegative = trimmed.startsWith("-") || trimmed.includes("--") || /[A-Z$€£¥₹₽₩฿]-/.test(trimmed);
      
      // Hapus koma (thousand separator) dan parse ke float
      let floatValue = parseFloat(numberPart.replace(/,/g, ""));
      
      if (!isNaN(floatValue)) {
        // Apply negative jika perlu
        if (isNegative) {
          floatValue = -Math.abs(floatValue);
        }
        
        // Round ke 2 decimal places
        const rounded = Math.round(floatValue * 100) / 100;
        console.log(`💰 Konversi ${currencySymbol}${isNegative ? " (negative)" : ""}: "${trimmed}" → ${rounded}`);
        return rounded;
      }
    }
  }
  
  // Handle format Eropa dengan kurung (accounting)
  if (europeanAccountingPattern.test(trimmed)) {
    const match = trimmed.match(europeanAccountingPattern);
    if (match) {
      const currencySymbol = match[1] || "EUR";
      const numberPart = match[2];
      const normalized = numberPart.replace(/\./g, "").replace(",", ".");
      const floatValue = parseFloat(normalized);
      
      if (!isNaN(floatValue)) {
        const rounded = Math.round(floatValue * -100) / 100; // Negative
        console.log(`💰 Konversi ${currencySymbol} (EU format, negative): "${trimmed}" → ${rounded}`);
        return rounded;
      }
    }
  }
  
  // Handle format Eropa standard
  if (europeanPattern.test(trimmed)) {
    const match = trimmed.match(europeanPattern);
    if (match) {
      const currencySymbol = match[1] || "EUR";
      const numberPart = match[2];
      
      // Cek minus
      const isNegative = trimmed.startsWith("-") || /[A-Z$€£¥₹₽₩฿]-/.test(trimmed);
      
      // Hapus titik (thousand separator) dan ganti koma dengan titik (decimal)
      const normalized = numberPart.replace(/\./g, "").replace(",", ".");
      let floatValue = parseFloat(normalized);
      
      if (!isNaN(floatValue)) {
        if (isNegative) {
          floatValue = -Math.abs(floatValue);
        }
        
        const rounded = Math.round(floatValue * 100) / 100;
        console.log(`💰 Konversi ${currencySymbol} (EU format${isNegative ? ", negative" : ""}): "${trimmed}" → ${rounded}`);
        return rounded;
      }
    }
  }
  
  // Cek juga format angka biasa dengan koma (tanpa symbol)
  if (numberWithCommaPattern.test(trimmed)) {
    const isNegative = trimmed.startsWith("-");
    const floatValue = parseFloat(trimmed.replace(/,/g, ""));
    if (!isNaN(floatValue)) {
      const rounded = Math.round(floatValue * 100) / 100;
      console.log(`💰 Konversi angka${isNegative ? " (negative)" : ""}: "${trimmed}" → ${rounded}`);
      return rounded;
    }
  }
  
  return value; // Kembalikan nilai asli jika bukan format currency
}
// ===== END FUNGSI UPDATED =====

// Fungsi untuk memproses dan mengkonversi data
function processRowData(row) {
  const processedRow = {};

  Object.keys(row).forEach((key) => {
    const cleanKey = cleanHeader(key);
    let originalValue = row[key];

    // Handle NULL untuk empty values
    if (
      originalValue === undefined ||
      originalValue === null ||
      originalValue === "" ||
      (typeof originalValue === "string" && originalValue.trim() === "")
    ) {
      processedRow[cleanKey] = null;
      return;
    }

    // Handle special case untuk "--" atau dash variants
    if (typeof originalValue === "string") {
      const trimmedValue = originalValue.trim();
      if (
        trimmedValue === "--" ||
        trimmedValue === "—" ||
        trimmedValue === "−"
      ) {
        processedRow[cleanKey] = null;
        console.log(
          `🔄 Konversi dash: "${originalValue}" → null di kolom "${cleanKey}"`
        );
        return;
      }
    }

    // ===== UBAH NAMA FUNGSI DARI parseRupiahToFloat JADI parseCurrencyToFloat =====
    const currencyParsed = parseCurrencyToFloat(originalValue);
    if (currencyParsed !== originalValue) {
      // Jika berhasil di-parse sebagai currency, langsung assign
      processedRow[cleanKey] = currencyParsed;
      return;
    }
    // ===== END =====

    // Konversi tanggal jika terdeteksi
    const convertedValue = convertToStandardDate(originalValue);

    // Log jika ada konversi tanggal (untuk debugging)
    if (convertedValue !== originalValue) {
      const temporalType = detectTemporalType(originalValue);
      if (temporalType || isExcelDate(originalValue)) {
        console.log(
          `📅 Konversi ${
            temporalType || "excel-date"
          }: "${originalValue}" → "${convertedValue}" di kolom "${cleanKey}"`
        );
      }
    }

    processedRow[cleanKey] = convertedValue;
  });

  return processedRow;
}

// Fungsi untuk mendapatkan ukuran file dalam bytes
function getFileSizeInBytes(filePath) {
  try {
    const stats = fs.statSync(filePath);
    return stats.size;
  } catch (error) {
    return 0;
  }
}

// Fungsi untuk format ukuran file
function formatFileSize(bytes) {
  if (bytes === 0) return "0 Bytes";
  const k = 1024;
  const sizes = ["Bytes", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
}

// NEW: Fungsi untuk split JSONL file jika terlalu besar
function splitLargeJsonlFile(filePath, maxSizeBytes = MAX_FILE_SIZE_BYTES) {
  try {
    const stats = fs.statSync(filePath);

    if (stats.size <= maxSizeBytes) {
      console.log(
        `📏 File ${path.basename(filePath)} (${formatFileSize(
          stats.size
        )}) tidak perlu di-split`
      );
      return [filePath]; // File tidak perlu di-split
    }

    console.log(
      `📏 File ${path.basename(filePath)} (${formatFileSize(
        stats.size
      )}) melebihi batas ${formatFileSize(maxSizeBytes)}, akan di-split...`
    );

    const fileContent = fs.readFileSync(filePath, "utf8");
    const lines = fileContent.split("\n").filter((line) => line.trim());

    console.log(`📊 Total records dalam file: ${lines.length}`);

    const baseName = path.basename(filePath, ".jsonl");
    const dir = path.dirname(filePath);
    const splitFiles = [];

    let currentBatch = [];
    let currentSize = 0;
    let partNumber = 1;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const lineSize = Buffer.byteLength(line + "\n", "utf8");

      // Jika menambah line ini akan melebihi batas dan batch sudah ada isinya
      if (currentSize + lineSize > maxSizeBytes && currentBatch.length > 0) {
        // Simpan batch saat ini
        const partFileName = `${baseName}_part_${String(partNumber).padStart(
          3,
          "0"
        )}.jsonl`;
        const partFilePath = path.join(dir, partFileName);

        fs.writeFileSync(partFilePath, currentBatch.join("\n"), "utf8");
        splitFiles.push(partFilePath);

        const partSize = Buffer.byteLength(currentBatch.join("\n"), "utf8");
        console.log(
          `💾 Part ${partNumber} disimpan: ${partFileName} (${formatFileSize(
            partSize
          )}, ${currentBatch.length} records)`
        );

        // Reset untuk part berikutnya
        currentBatch = [];
        currentSize = 0;
        partNumber++;
      }

      // Tambahkan line ke batch saat ini
      currentBatch.push(line);
      currentSize += lineSize;
    }

    // Simpan batch terakhir jika ada
    if (currentBatch.length > 0) {
      const partFileName = `${baseName}_part_${String(partNumber).padStart(
        3,
        "0"
      )}.jsonl`;
      const partFilePath = path.join(dir, partFileName);

      fs.writeFileSync(partFilePath, currentBatch.join("\n"), "utf8");
      splitFiles.push(partFilePath);

      const partSize = Buffer.byteLength(currentBatch.join("\n"), "utf8");
      console.log(
        `💾 Part ${partNumber} disimpan: ${partFileName} (${formatFileSize(
          partSize
        )}, ${currentBatch.length} records)`
      );
    }

    // Hapus file asli setelah split berhasil
    try {
      fs.unlinkSync(filePath);
      console.log(`🗑️  File asli dihapus: ${path.basename(filePath)}`);
    } catch (error) {
      console.warn(`⚠️  Gagal hapus file asli: ${error.message}`);
    }

    console.log(`✅ Split selesai: ${splitFiles.length} part files dibuat`);
    return splitFiles;
  } catch (error) {
    console.error(`❌ Error splitting file ${filePath}:`, error.message);
    return [filePath]; // Return original file jika gagal split
  }
}

// Fungsi untuk memindahkan file ke Done - Input
function moveFileToProcessed(filePath) {
  try {
    // Pastikan folder Done - Input ada
    if (!fs.existsSync(DONE_INPUT_PATH)) {
      fs.mkdirSync(DONE_INPUT_PATH, { recursive: true });
      console.log(`📁 Folder Done - Input dibuat: ${DONE_INPUT_PATH}`);
    }

    const fileName = path.basename(filePath);
    const destinationPath = path.join(DONE_INPUT_PATH, fileName);

    // Cek jika file sudah ada di destination
    if (fs.existsSync(destinationPath)) {
      const baseName = path.basename(fileName, path.extname(fileName));
      const ext = path.extname(fileName);
      const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
      const newFileName = `${baseName}_${timestamp}${ext}`;
      const newDestinationPath = path.join(DONE_INPUT_PATH, newFileName);

      fs.renameSync(filePath, newDestinationPath);
      console.log(
        `📦 File dipindahkan ke: ${newDestinationPath} (renamed karena duplikat)`
      );
      return newDestinationPath;
    } else {
      fs.renameSync(filePath, destinationPath);
      console.log(`📦 File dipindahkan ke: ${destinationPath}`);
      return destinationPath;
    }
  } catch (error) {
    console.error(`❌ Gagal memindahkan file ${filePath}:`, error.message);
    return null;
  }
}

// ENHANCED: Fungsi untuk convert single file dengan auto-split
function convertXlsToJsonl(
  inputPath,
  outputPath = null,
  moveAfterConvert = true
) {
  try {
    // Normalize input path
    inputPath = path.resolve(inputPath);

    // Baca file Excel
    console.log(`📖 Membaca file: ${inputPath}`);
    const workbook = XLSX.readFile(inputPath);

    // Ambil sheet pertama (atau bisa dimodif untuk pilih sheet)
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Convert ke JSON dengan opsi untuk mempertahankan tipe data
    const rawJsonData = XLSX.utils.sheet_to_json(worksheet, {
      raw: false, // Jangan convert semua ke string
      dateNF: "yyyy-mm-dd", // Format tanggal default
      cellDates: true, // Parse tanggal sebagai Date object
    });

    console.log(`🔄 Memproses ${rawJsonData.length} baris data...`);

    // Proses setiap row untuk konversi tanggal dan clean headers
    const jsonData = rawJsonData.map((row, index) => {
      try {
        return processRowData(row);
      } catch (error) {
        console.warn(`⚠️  Error memproses baris ${index + 1}:`, error.message);
        // Fallback: clean headers saja tanpa konversi tanggal
        const cleanedRow = {};
        Object.keys(row).forEach((key) => {
          const cleanKey = cleanHeader(key);
          cleanedRow[cleanKey] = row[key];
        });
        return cleanedRow;
      }
    });

    // Convert array ke JSON Lines format (satu object per line)
    const jsonlData = jsonData.map((row) => JSON.stringify(row)).join("\n");

    // Tentukan output path
    if (!outputPath) {
      const inputName = path.basename(inputPath, path.extname(inputPath));

      // Cek apakah ada folder output di working directory
      const outputDir = path.join(process.cwd(), "output");
      if (fs.existsSync(outputDir) && fs.statSync(outputDir).isDirectory()) {
        outputPath = path.join(outputDir, `${inputName}.jsonl`);
      } else {
        // Fallback ke folder yang sama dengan input
        const inputDir = path.dirname(inputPath);
        outputPath = path.join(inputDir, `${inputName}.jsonl`);
      }
    } else {
      // Handle jika outputPath adalah folder, bukan file
      outputPath = path.resolve(outputPath);

      if (fs.existsSync(outputPath) && fs.statSync(outputPath).isDirectory()) {
        // Jika outputPath adalah folder, buat nama file
        const inputName = path.basename(inputPath, path.extname(inputPath));
        outputPath = path.join(outputPath, `${inputName}.jsonl`);
      } else if (
        !outputPath.endsWith(".jsonl") &&
        !outputPath.endsWith(".json")
      ) {
        // Jika tidak ada ekstensi dan folder belum ada, anggap sebagai folder
        const inputName = path.basename(inputPath, path.extname(inputPath));
        // Buat folder jika belum ada
        fs.mkdirSync(outputPath, { recursive: true });
        outputPath = path.join(outputPath, `${inputName}.jsonl`);
      }
    }

    // Pastikan output directory ada
    const outputDir = path.dirname(outputPath);
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
      console.log(`📁 Folder output dibuat: ${outputDir}`);
    }

    // Tulis file JSON Lines
    fs.writeFileSync(outputPath, jsonlData, "utf8");
    console.log(`✅ Berhasil convert: ${outputPath}`);
    console.log(`📊 Total records: ${jsonData.length}`);

    // NEW: Check dan split file jika terlalu besar
    const finalFiles = splitLargeJsonlFile(outputPath);

    // Pindahkan file Excel ke Done - Input jika conversion berhasil
    if (moveAfterConvert) {
      const movedPath = moveFileToProcessed(inputPath);
      if (movedPath) {
        return { success: true, outputPath: finalFiles, movedPath, jsonData };
      } else {
        return {
          success: true,
          outputPath: finalFiles,
          movedPath: null,
          jsonData,
        };
      }
    }

    return { success: true, outputPath: finalFiles, jsonData };
  } catch (error) {
    console.error(`❌ Error converting ${inputPath}:`, error.message);
    return { success: false, error: error.message };
  }
}

// ENHANCED: Fungsi untuk merge JSONL files dengan batasan ukuran dan auto-split
function mergeJsonlFiles(jsonlFiles, outputDir, baseName = "merged") {
  try {
    console.log(`\n🔄 Memulai proses merge ${jsonlFiles.length} file JSONL...`);
    console.log(`📏 Batasan ukuran per file: ${MAX_FILE_SIZE_MB}MB`);

    let currentBatch = [];
    let currentSize = 0;
    let batchNumber = 1;
    let allMergedFiles = []; // Array untuk semua file hasil merge dan split
    let totalRecords = 0;

    for (const filePath of jsonlFiles) {
      console.log(`📖 Membaca: ${path.basename(filePath)}`);

      const fileContent = fs.readFileSync(filePath, "utf8");
      const fileSize = Buffer.byteLength(fileContent, "utf8");

      console.log(`   Ukuran: ${formatFileSize(fileSize)}`);

      // Jika file ini akan membuat batch melebihi batas, simpan batch saat ini
      if (
        currentSize + fileSize > MAX_FILE_SIZE_BYTES &&
        currentBatch.length > 0
      ) {
        const mergedFilePath = saveBatch(
          currentBatch,
          outputDir,
          baseName,
          batchNumber,
          currentSize
        );

        // NEW: Check dan split jika hasil merge terlalu besar
        const splitResults = splitLargeJsonlFile(mergedFilePath);
        allMergedFiles.push(...splitResults);

        // Reset untuk batch baru
        currentBatch = [];
        currentSize = 0;
        batchNumber++;
      }

      // Tambahkan file ke batch saat ini
      currentBatch.push(fileContent);
      currentSize += fileSize;

      // Hitung jumlah records
      const recordCount = fileContent
        .split("\n")
        .filter((line) => line.trim()).length;
      totalRecords += recordCount;

      console.log(`   Records: ${recordCount}`);
    }

    // Simpan batch terakhir jika ada
    if (currentBatch.length > 0) {
      const mergedFilePath = saveBatch(
        currentBatch,
        outputDir,
        baseName,
        batchNumber,
        currentSize
      );

      // NEW: Check dan split jika hasil merge terlalu besar
      const splitResults = splitLargeJsonlFile(mergedFilePath);
      allMergedFiles.push(...splitResults);
    }

    // Hapus file JSONL individual setelah merge berhasil
    console.log(`\n🗑️  Menghapus file JSONL individual...`);
    let deletedCount = 0;
    for (const filePath of jsonlFiles) {
      try {
        fs.unlinkSync(filePath);
        deletedCount++;
        console.log(`   ✅ Dihapus: ${path.basename(filePath)}`);
      } catch (error) {
        console.warn(
          `   ⚠️  Gagal hapus: ${path.basename(filePath)} - ${error.message}`
        );
      }
    }

    console.log(`\n🎉 Merge selesai!`);
    console.log(`📊 Summary:`);
    console.log(`   Total file merged: ${allMergedFiles.length}`);
    console.log(`   Total records: ${totalRecords.toLocaleString()}`);
    console.log(
      `   File individual dihapus: ${deletedCount}/${jsonlFiles.length}`
    );

    console.log(`\n📁 File hasil merge (termasuk split):`);
    allMergedFiles.forEach((file, index) => {
      const fileSize = getFileSizeInBytes(file);
      console.log(`   ${index + 1}. ${path.basename(file)}`);
      console.log(`      Ukuran: ${formatFileSize(fileSize)}`);
    });

    return {
      success: true,
      mergedFiles: allMergedFiles,
      totalRecords,
      deletedCount,
    };
  } catch (error) {
    console.error(`❌ Error merging files:`, error.message);
    return { success: false, error: error.message };
  }
}

// Fungsi helper untuk menyimpan batch
function saveBatch(
  batchContent,
  outputDir,
  baseName,
  batchNumber,
  currentSize
) {
  const paddedNumber = String(batchNumber).padStart(3, "0");
  const mergedFileName = `${baseName}_${paddedNumber}.jsonl`;
  const mergedFilePath = path.join(outputDir, mergedFileName);

  const mergedContent = batchContent.join("\n");
  fs.writeFileSync(mergedFilePath, mergedContent, "utf8");

  const recordCount = mergedContent
    .split("\n")
    .filter((line) => line.trim()).length;

  console.log(`💾 Batch ${batchNumber} disimpan: ${mergedFileName}`);
  console.log(`   Ukuran: ${formatFileSize(currentSize)}`);
  console.log(`   Records: ${recordCount.toLocaleString()}`);
  console.log(`   File count: ${batchContent.length}`);

  return mergedFilePath;
}

// ENHANCED: Fungsi untuk convert multiple files di folder dengan merge dan auto-split
function convertFolderWithMerge(
  folderPath,
  customOutputDir = null,
  moveAfterConvert = true
) {
  try {
    const files = fs.readdirSync(folderPath);
    const xlsFiles = files.filter(
      (file) =>
        path.extname(file).toLowerCase() === ".xls" ||
        path.extname(file).toLowerCase() === ".xlsx"
    );

    if (xlsFiles.length === 0) {
      console.log("❌ Tidak ada file .xls/.xlsx ditemukan di folder ini");
      return;
    }

    // Sort files numerically if they follow pattern like 1.xlsx, 2.xlsx, etc.
    xlsFiles.sort((a, b) => {
      const aNum = parseInt(path.basename(a, path.extname(a)));
      const bNum = parseInt(path.basename(b, path.extname(b)));
      if (!isNaN(aNum) && !isNaN(bNum)) {
        return aNum - bNum;
      }
      return a.localeCompare(b);
    });

    console.log(`📁 Ditemukan ${xlsFiles.length} file Excel di: ${folderPath}`);

    // Tentukan output directory
    let outputDir;
    if (customOutputDir) {
      outputDir = path.resolve(customOutputDir);
    } else {
      // Default ke folder "Output" di dalam input folder
      outputDir = path.join(folderPath, "Output");
    }

    // Cek atau buat folder output
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
      console.log(`📁 Folder output dibuat: ${outputDir}`);
    }

    console.log(`📤 Output akan disimpan ke: ${outputDir}`);
    if (moveAfterConvert) {
      console.log(
        `📦 File Excel yang berhasil akan dipindahkan ke: ${DONE_INPUT_PATH}`
      );
    }

    // Buat folder temp untuk JSONL individual
    const tempDir = path.join(outputDir, ".temp");
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }

    let successCount = 0;
    let movedCount = 0;
    const results = [];
    const jsonlFiles = [];

    console.log(`\n🔄 === TAHAP 1: CONVERT EXCEL KE JSONL ===`);

    // Convert semua file Excel ke JSONL individual
    xlsFiles.forEach((file, index) => {
      const inputPath = path.join(folderPath, file);
      const fileName = path.basename(file, path.extname(file));
      const tempOutputPath = path.join(tempDir, `${fileName}.jsonl`);

      console.log(`\n📄 Processing ${index + 1}/${xlsFiles.length}: ${file}`);
      const result = convertXlsToJsonl(
        inputPath,
        tempOutputPath,
        moveAfterConvert
      );
      results.push({ file, result });

      if (result.success) {
        successCount++;
        // NEW: Handle jika outputPath adalah array (hasil split)
        if (Array.isArray(result.outputPath)) {
          jsonlFiles.push(...result.outputPath);
        } else {
          jsonlFiles.push(result.outputPath);
        }
        if (result.movedPath) {
          movedCount++;
        }
      }
    });

    console.log(
      `\n✅ Tahap 1 selesai: ${successCount}/${xlsFiles.length} file berhasil diconvert`
    );
    console.log(`📊 Total JSONL files (termasuk split): ${jsonlFiles.length}`);

    if (jsonlFiles.length === 0) {
      console.log(
        "❌ Tidak ada file JSONL yang berhasil dibuat untuk di-merge"
      );
      // Hapus temp folder
      try {
        fs.rmSync(tempDir, { recursive: true, force: true });
      } catch (error) {
        console.warn(`⚠️  Gagal hapus temp folder: ${error.message}`);
      }
      return;
    }

    // Merge JSONL files
    console.log(`\n🔄 === TAHAP 2: MERGE JSONL FILES ===`);
    const folderBaseName = path.basename(folderPath);
    const mergeResult = mergeJsonlFiles(jsonlFiles, outputDir, folderBaseName);

    // Hapus temp folder
    try {
      fs.rmSync(tempDir, { recursive: true, force: true });
      console.log(`🗑️  Temp folder dihapus`);
    } catch (error) {
      console.warn(`⚠️  Gagal hapus temp folder: ${error.message}`);
    }

    if (mergeResult.success) {
      console.log(`\n🎉 PROSES SELESAI!`);
      console.log(`📊 Summary total:`);
      console.log(
        `   Excel files processed: ${successCount}/${xlsFiles.length}`
      );
      console.log(`   Final merged files: ${mergeResult.mergedFiles.length}`);
      console.log(
        `   Total records: ${mergeResult.totalRecords.toLocaleString()}`
      );

      if (moveAfterConvert) {
        console.log(`   Excel files moved: ${movedCount}/${successCount}`);
      }

      // Show failed moves if any
      const failedMoves = results.filter(
        (r) => r.result.success && r.result.movedPath === null
      );
      if (failedMoves.length > 0) {
        console.log(
          `⚠️  ${failedMoves.length} file Excel berhasil diconvert tapi gagal dipindahkan:`
        );
        failedMoves.forEach((fm) => console.log(`     - ${fm.file}`));
      }
    } else {
      console.error(`❌ Merge gagal: ${mergeResult.error}`);
    }
  } catch (error) {
    console.error("❌ Error processing folder:", error.message);
  }
}

// ENHANCED: Fungsi untuk convert multiple files di folder tanpa merge (legacy) dengan auto-split
function convertFolder(
  folderPath,
  customOutputDir = null,
  moveAfterConvert = true
) {
  try {
    const files = fs.readdirSync(folderPath);
    const xlsFiles = files.filter(
      (file) =>
        path.extname(file).toLowerCase() === ".xls" ||
        path.extname(file).toLowerCase() === ".xlsx"
    );

    if (xlsFiles.length === 0) {
      console.log("❌ Tidak ada file .xls/.xlsx ditemukan di folder ini");
      return;
    }

    console.log(`📁 Ditemukan ${xlsFiles.length} file Excel di: ${folderPath}`);

    // Tentukan output directory
    let outputDir;
    if (customOutputDir) {
      outputDir = path.resolve(customOutputDir);
    } else {
      // Default ke folder output di project
      outputDir = path.join(process.cwd(), "output");
    }

    // Cek atau buat folder output
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
      console.log(`📁 Folder output dibuat: ${outputDir}`);
    }

    console.log(`📤 Output akan disimpan ke: ${outputDir}`);
    if (moveAfterConvert) {
      console.log(
        `📦 File yang berhasil akan dipindahkan ke: ${DONE_INPUT_PATH}`
      );
    }

    let successCount = 0;
    let movedCount = 0;
    let totalFinalFiles = 0; // NEW: Track total files termasuk split
    const results = [];

    xlsFiles.forEach((file) => {
      const inputPath = path.join(folderPath, file);
      const fileName = path.basename(file, path.extname(file));
      const outputPath = path.join(outputDir, `${fileName}.jsonl`);

      console.log(`\n🔄 Processing: ${file}`);
      const result = convertXlsToJsonl(inputPath, outputPath, moveAfterConvert);
      results.push({ file, result });

      if (result.success) {
        successCount++;
        // NEW: Handle jika outputPath adalah array (hasil split)
        if (Array.isArray(result.outputPath)) {
          totalFinalFiles += result.outputPath.length;
          console.log(
            `   📊 File ini di-split menjadi ${result.outputPath.length} part(s)`
          );
        } else {
          totalFinalFiles += 1;
        }

        if (result.movedPath) {
          movedCount++;
        }
      }
    });

    console.log(
      `\n🎉 Selesai! ${successCount}/${xlsFiles.length} file berhasil diconvert`
    );
    console.log(`📊 Total JSONL files (termasuk split): ${totalFinalFiles}`);

    if (moveAfterConvert) {
      console.log(
        `📦 ${movedCount}/${successCount} file berhasil dipindahkan ke Done - Input`
      );
    }

    // Show summary of any failed moves
    const failedMoves = results.filter(
      (r) => r.result.success && r.result.movedPath === null
    );
    if (failedMoves.length > 0) {
      console.log(
        `⚠️  ${failedMoves.length} file berhasil diconvert tapi gagal dipindahkan:`
      );
      failedMoves.forEach((fm) => console.log(`   - ${fm.file}`));
    }
  } catch (error) {
    console.error("❌ Error membaca folder:", error.message);
  }
}

// Main function
function main() {
  const args = process.argv.slice(2);

  if (args.length === 0) {
    console.log(`
🔄 XLS to JSON Lines Converter with Smart Merger & Auto-Split (Max ${MAX_FILE_SIZE_MB}MB per file)

NEW FEATURES:
  📅 ENHANCED DATE SUPPORT: Support untuk format mm/dd/yy (2-digit year)
  ✂️  AUTO-SPLIT: File JSONL yang >99MB otomatis di-split menjadi beberapa part

Cara pakai:
  node converter.js <input-folder>                       # Convert & merge semua file di folder
  node converter.js <input-folder> <output-folder>       # Convert & merge dengan custom output folder
  node converter.js <file.xls>                           # Convert single file (with auto-split)
  node converter.js <file.xls> <output.jsonl>            # Convert single file dengan custom output
  node converter.js --no-merge <input-folder>            # Convert tanpa merge (individual JSONL)
  node converter.js --no-move <input-folder>             # Convert & merge tanpa memindahkan Excel files
  node converter.js --help                               # Show help

Contoh:
  node converter.js "C:\\Users\\PIAGAM\\Downloads\\Juli"
  node converter.js "C:\\Users\\PIAGAM\\Downloads\\Juli" "C:\\Users\\PIAGAM\\Downloads\\Juli\\jsonl"
  node converter.js data.xlsx
  node converter.js --no-merge "C:\\Data\\Excel\\"
  node converter.js --no-move "C:\\Data\\Excel\\"

NEW DATE FORMATS SUPPORTED:
  📅 01/15/25 → 2025-01-15 (mm/dd/yy)
  📅 15/01/25 → 2025-01-15 (dd/mm/yy) 
  📅 01-15-25 → 2025-01-15 (mm-dd-yy)
  📅 15.01.25 → 2025-01-15 (dd.mm.yy)
  📅 Plus semua format lama yang sudah didukung

AUTO-SPLIT FEATURES:
  ✂️  Single file >99MB otomatis di-split jadi multiple parts
  ✂️  Merged files >99MB juga otomatis di-split
  📋 Part files diberi nama: filename_part_001.jsonl, filename_part_002.jsonl, dst.
  📊 Split berdasarkan jumlah records untuk memastikan integritas JSON Lines

Features:
  📦 MERGE: Menggabungkan multiple JSONL dengan batasan ${MAX_FILE_SIZE_MB}MB per file
  📅 AUTO-CONVERT: Tanggal ke format yyyy-mm-dd hh:mm:ss (termasuk 2-digit year)
  ✂️  AUTO-SPLIT: File >99MB otomatis di-split menjadi parts
  📁 DYNAMIC PATH: Input dan output folder yang fleksibel
  🏷️  AUTO-NAMING: File merged menggunakan nama folder + nomor urut
  📂 DEFAULT OUTPUT: Jika tidak ada output path, buat folder "Output" di input folder
  📦 AUTO-MOVE: File Excel yang berhasil dipindahkan ke: ${DONE_INPUT_PATH}
        `);
    return;
  }

  if (args[0] === "--help" || args[0] === "-h") {
    console.log(`
🔄 XLS to JSON Lines Converter with Smart Merger & Auto-Split - Help

NEW FEATURES:
  📅 ENHANCED DATE SUPPORT: 
     - Otomatis konversi 2-digit year (25 → 2025, 95 → 1995)
     - Support format: mm/dd/yy, dd/mm/yy, mm-dd-yy, dd.mm.yy
     - Logic: year ≤ current year → current century, year > current year → previous century

  ✂️  AUTO-SPLIT FEATURE:
     - File JSONL >99MB otomatis di-split menjadi multiple parts
     - Berlaku untuk single file conversion dan hasil merge
     - Part files: filename_part_001.jsonl, filename_part_002.jsonl, dst.
     - Split berdasarkan line count untuk menjaga integritas JSON Lines format

MODES:
  1. FOLDER MODE (with merge & auto-split): node converter.js <input-folder> [output-folder]
     - Convert semua Excel files di folder
     - Merge hasil JSONL dengan batasan ${MAX_FILE_SIZE_MB}MB per file
     - Auto-split jika hasil merge >99MB
     - Default output: <input-folder>/Output/

  2. SINGLE FILE MODE (with auto-split): node converter.js <file.xlsx> [output.jsonl]
     - Convert single file
     - Auto-split jika hasil >99MB
     - Enhanced date format support

  3. NO-MERGE MODE (with auto-split): node converter.js --no-merge <input-folder> [output-folder]
     - Convert folder tapi tanpa merge (individual JSONL files)
     - Individual files tetap di-split jika >99MB

  4. NO-MOVE MODE: node converter.js --no-move <input-folder> [output-folder]
     - Convert & merge tapi tidak pindahkan Excel files ke Done-Input

FLAGS:
  --no-merge    : Convert folder tanpa merge (individual JSONL files + auto-split)
  --no-move     : Tidak memindahkan Excel files setelah conversion
  --help, -h    : Show this help

DATE FORMAT EXAMPLES:
  INPUT FORMATS:
    📅 Excel dates (serial numbers)
    📅 01/15/2025 → 2025-01-15
    📅 01/15/25   → 2025-01-15 (NEW!)
    📅 15/01/25   → 2025-01-15 (NEW!)
    📅 01-15-25   → 2025-01-15 (NEW!)
    📅 15.01.25   → 2025-01-15 (NEW!)
    📅 01 Agu 2025 23:51 → 2025-08-01 23:51:00
    📅 2025-01-15 14:30:00 → 2025-01-15 14:30:00

AUTO-SPLIT EXAMPLES:
  📂 Single File:
     large_data.xlsx (contains data that creates 150MB JSONL)
     ↓
     large_data_part_001.jsonl (99MB)
     large_data_part_002.jsonl (51MB)

  📂 Merged Files:
     multiple files → merged_001.jsonl (120MB)
     ↓
     merged_001_part_001.jsonl (99MB)
     merged_001_part_002.jsonl (21MB)

EXAMPLES:
  # Enhanced conversion dengan auto-split
  node converter.js "C:\\Users\\PIAGAM\\Downloads\\Juli"
  
  # Single large file dengan auto-split
  node converter.js "C:\\Data\\large_report.xlsx"
  
  # No merge tapi tetap auto-split per file
  node converter.js --no-merge "C:\\Excel\\Data\\" "D:\\Individual\\"

SUPPORTED FORMATS:
  📁 Input: .xls, .xlsx files
  📄 Output: .jsonl (JSON Lines format) with auto-split
  📅 Date formats: All previous + new 2-digit year support
  🔤 Headers: Auto-clean (lowercase + underscore)

2-DIGIT YEAR CONVERSION LOGIC:
  Current year: 2025
  - 25 → 2025 (same century)
  - 24 → 2024 (same century)  
  - 26 → 1926 (previous century, because 26 > 25)
  - 95 → 1995 (previous century)
  - 00 → 2000 (same century)
        `);
    return;
  }

  // Check for flags
  let noMerge = false;
  let moveAfterConvert = true;
  let actualArgs = [...args];

  if (args[0] === "--no-merge") {
    noMerge = true;
    actualArgs = args.slice(1);
    console.log(
      "🚫 Mode: Tidak akan merge file JSONL (individual files + auto-split)"
    );
  } else if (args[0] === "--no-move") {
    moveAfterConvert = false;
    actualArgs = args.slice(1);
    console.log(
      "🚫 Mode: Excel files tidak akan dipindahkan setelah conversion"
    );
  }

  if (actualArgs.length === 0) {
    console.error("❌ Setelah flag, harus ada input file/folder");
    return;
  }

  const inputPath = actualArgs[0];
  const outputPath = actualArgs[1];

  // Cek apakah path ada
  if (!fs.existsSync(inputPath)) {
    console.error("❌ File atau folder tidak ditemukan:", inputPath);
    return;
  }

  // Normalize paths
  const resolvedInputPath = path.resolve(inputPath);
  const stats = fs.statSync(resolvedInputPath);

  if (stats.isDirectory()) {
    // Convert semua file di folder
    let outputFolder = outputPath;

    // Jika tidak ada output path, buat folder "Output" di dalam input folder
    if (!outputFolder) {
      outputFolder = path.join(resolvedInputPath, "Output");
      console.log(
        `📁 Output path tidak disebutkan, akan menggunakan: ${outputFolder}`
      );
    }

    if (noMerge) {
      // Mode tanpa merge (legacy behavior) + auto-split
      convertFolder(resolvedInputPath, outputFolder, moveAfterConvert);
    } else {
      // Mode dengan merge (new behavior) + auto-split
      convertFolderWithMerge(resolvedInputPath, outputFolder, moveAfterConvert);
    }
  } else if (stats.isFile()) {
    // Convert single file (dengan auto-split)
    const ext = path.extname(resolvedInputPath).toLowerCase();
    if (ext === ".xls" || ext === ".xlsx") {
      const result = convertXlsToJsonl(
        resolvedInputPath,
        outputPath,
        moveAfterConvert
      );
      if (result.success && moveAfterConvert && !result.movedPath) {
        console.log("⚠️  File berhasil diconvert tapi gagal dipindahkan");
      }

      // NEW: Show split results untuk single file
      if (result.success && Array.isArray(result.outputPath)) {
        console.log(
          `✂️  File di-split menjadi ${result.outputPath.length} part(s):`
        );
        result.outputPath.forEach((file, index) => {
          const fileSize = getFileSizeInBytes(file);
          console.log(
            `   ${index + 1}. ${path.basename(file)} (${formatFileSize(
              fileSize
            )})`
          );
        });
      }
    } else {
      console.error("❌ File harus berformat .xls atau .xlsx");
    }
  }
}

// Jalankan program
if (require.main === module) {
  main();
}
