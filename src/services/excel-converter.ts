
import * as XLSX from 'xlsx';
import { format, parse, isValid, getHours, getMinutes, getSeconds } from 'date-fns';

/**
 * Converts a CSV file content (as string) to XLS format (as Buffer).
 * Applies transformations similar to the provided VBA script.
 * Handles potential date/time format variations.
 *
 * @param csvData The CSV data as a string.
 * @returns A promise that resolves to the XLS data as a Buffer.
 */
export async function convertCsvToXls(csvData: string): Promise<Buffer> {
  // 1. Parse CSV Data (using semicolon delimiter)
  const workbook = XLSX.read(csvData, { type: 'string', cellDates: false, FS: ';' });
  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  const jsonData: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', rawNumbers: false }); // Use rawNumbers: false to get strings initially

  if (!jsonData || jsonData.length === 0) {
    throw new Error("CSV data is empty or invalid.");
  }

  const originalHeaders = jsonData[0].map(h => String(h)); // Ensure headers are strings
  const dataRows = jsonData.slice(1);

  // 2. Locate "Date/Time Played" and "Duration" columns (case-insensitive)
  const dtColIndex = originalHeaders.findIndex(h => h.toLowerCase() === "date/time played");
  const durColIndex = originalHeaders.findIndex(h => h.toLowerCase() === "duration");

  if (dtColIndex === -1) {
    console.warn("Column 'Date/Time Played' not found. Skipping date/time split.");
  }
   if (durColIndex === -1) {
    console.warn("Column 'Duration' not found. Skipping duration formatting.");
  }

  let newHeaders = [...originalHeaders];
  const processedDataRows: (string | number | Date | null)[][] = []; // Store processed data

  // Rename "Date/Time Played" to "Date Played"
  if (dtColIndex !== -1) {
      newHeaders[dtColIndex] = "Date Played";
  }

  // 3. Insert new "Time Played" column header if "Date/Time Played" exists
  let timeColIndex = -1;
  if (dtColIndex !== -1) {
    timeColIndex = dtColIndex + 1;
    newHeaders.splice(timeColIndex, 0, "Time Played");
  }

  // 4. Process each data row
  dataRows.forEach(row => {
    // Create a new array for the processed row, ensuring correct length if Time Played is added
    const processedRow: (string | number | Date | null)[] = [];
    let originalRowIndex = 0;
    for (let i = 0; i < newHeaders.length; i++) {
        if (i === timeColIndex) {
            processedRow.push(null); // Placeholder for Time Played initially
        } else if (i === dtColIndex && dtColIndex !== -1) {
             processedRow.push(row[originalRowIndex] ?? null);
             originalRowIndex++;
        } else {
             processedRow.push(row[originalRowIndex] ?? null);
             originalRowIndex++;
        }
    }


    if (dtColIndex !== -1 && timeColIndex !== -1) {
      const fullText = processedRow[dtColIndex]?.toString() || ''; // Get value from the correct index in processedRow
      const splitPos = fullText.indexOf(" ");

      let datePart: Date | string | null = null;
      let timePart: Date | string | null = null;

      if (splitPos > 0) {
        const dateStr = fullText.substring(0, splitPos).trim();
        const timeStr = fullText.substring(splitPos + 1).trim();

        // --- Date Parsing ---
        const dateFormats = ['M/d/yyyy', 'MM/dd/yyyy', 'yyyy-MM-dd', 'dd.MM.yyyy', 'd.M.yyyy'];
        for (const fmt of dateFormats) {
          const parsed = parse(dateStr, fmt, new Date());
          if (isValid(parsed) && parsed.getFullYear() > 1900) {
                datePart = parsed;
                break;
          }
        }
        if (datePart === null) {
            console.warn(`Could not parse date string: "${dateStr}". Keeping original.`);
            datePart = dateStr;
        }

        // --- Time Parsing ---
        const baseDateForTimeParsing = new Date(1970, 0, 1);
        const timeFormats = ['h:mm:ss a', 'hh:mm:ss a', 'H:mm:ss', 'HH:mm:ss', 'h:mm a', 'H:mm'];
        for (const fmt of timeFormats) {
          const parsed = parse(`${format(baseDateForTimeParsing, 'yyyy-MM-dd')} ${timeStr}`, `yyyy-MM-dd ${fmt}`, new Date());
          if (isValid(parsed)) {
              timePart = parsed;
              break;
          }
        }
         if (timePart === null && /^\d{1,2}:\d{1,2}:\d{1,2}$/.test(timeStr)) {
            const [h, m, s] = timeStr.split(':').map(Number);
            if (!isNaN(h) && !isNaN(m) && !isNaN(s)) {
                const tempDate = new Date(baseDateForTimeParsing);
                tempDate.setHours(h, m, s, 0);
                if (isValid(tempDate)) {
                    timePart = tempDate;
                }
            }
        }

        if (timePart === null) {
          console.warn(`Could not parse time string: "${timeStr}". Setting time column to empty.`);
          timePart = '';
        }

      } else {
        console.warn(`Could not split date/time string: "${fullText}". Keeping original date, empty time.`);
         const dateFormats = ['M/d/yyyy', 'MM/dd/yyyy', 'yyyy-MM-dd', 'dd.MM.yyyy', 'd.M.yyyy'];
         let parsedAsDateOnly : Date | null = null;
         for (const fmt of dateFormats) {
            const parsed = parse(fullText, fmt, new Date());
            if (isValid(parsed) && parsed.getFullYear() > 1900) {
               parsedAsDateOnly = parsed;
               break;
            }
         }
         datePart = parsedAsDateOnly || fullText;
         timePart = '';
      }

      // Update the processed row values in the correct positions
      processedRow[dtColIndex] = datePart; // Update Date Played column
      processedRow[timeColIndex] = timePart; // Insert Time Played value
    }

    processedDataRows.push(processedRow);
  });


  // 5. Prepare data for new worksheet
  // Pass the processed data which now includes the split columns
  const dataForSheet = [newHeaders, ...processedDataRows];
  const newWs = XLSX.utils.aoa_to_sheet(dataForSheet, { cellDates: true }); // Use cellDates: true to handle Date objects

  // --- Post-processing and Formatting ---
  const range = XLSX.utils.decode_range(newWs['!ref'] || 'A1:A1');

  for (let R = range.s.r + 1; R <= range.e.r; ++R) { // Start from row 1 (data)

    // Format Date Played column
    if (dtColIndex !== -1) {
      const dateCellRef = XLSX.utils.encode_cell({ r: R, c: dtColIndex });
      const dateCell = newWs[dateCellRef];
      // Check if aoa_to_sheet correctly converted it to a number (Excel serial date)
      if (dateCell && typeof dateCell.v === 'number' && dateCell.t === 'n') {
         // It's already an Excel serial date number, just apply the desired format.
         dateCell.z = 'm/d/yyyy'; // Apply the specific date format
      } else if (dateCell && dateCell.v instanceof Date && isValid(dateCell.v)) {
         // Fallback: If it's still a Date object (less likely with cellDates:true),
         // ensure type is 'n' and apply format. write with cellDates:false should handle this.
         console.warn("Date cell was still a Date object after aoa_to_sheet, applying format. Value:", dateCell.v);
         dateCell.t = 'n';
         dateCell.z = 'm/d/yyyy';
      } else if (dateCell && typeof dateCell.v === 'string') {
          // If it remained a string (parsing failed earlier), ensure type is 's'
          dateCell.t = 's';
      } else if (dateCell) {
          // Log unexpected types
          console.warn(`Unexpected type in Date Played column at row ${R+1}:`, typeof dateCell.v, dateCell.v);
          if (!dateCell.t) dateCell.t = 's'; // Default to string if type is missing
      }
    }


    // Format Time Played column
    if (timeColIndex !== -1) {
        const timeCellRef = XLSX.utils.encode_cell({ r: R, c: timeColIndex });
        const timeCell = newWs[timeCellRef];
        // Check if aoa_to_sheet handled it (might be number if it was a JS Date)
        if (timeCell && typeof timeCell.v === 'number' && timeCell.t === 'n') {
            // If aoa_to_sheet converted it to Excel serial time
             timeCell.z = 'hh:mm:ss'; // Apply time format
        } else if (timeCell && timeCell.v instanceof Date && isValid(timeCell.v)) {
            // If it's still a JS Date object from parsing
            const hours = getHours(timeCell.v);
            const minutes = getMinutes(timeCell.v);
            const seconds = getSeconds(timeCell.v);
            const excelTime = (hours * 3600 + minutes * 60 + seconds) / (24 * 60 * 60);
            timeCell.t = 'n';
            timeCell.v = excelTime;
            timeCell.z = 'hh:mm:ss';
        } else if (timeCell && typeof timeCell.v === 'string' && timeCell.v === '') {
            // Handle empty string case explicitly
            timeCell.t = 's';
            timeCell.v = '';
        } else if (timeCell) {
            // Fallback for unexpected types or invalid dates
            console.warn(`Unexpected type or invalid value in Time Played column at row ${R+1}:`, typeof timeCell.v, timeCell.v);
            timeCell.t = 's'; // Keep original string value if conversion wasn't possible
        }
    }


    // Format Duration column
    if (durColIndex !== -1) {
      // Calculate actual column index for Duration *after* potentially inserting Time Played
       let actualDurColIndex = durColIndex;
       if (timeColIndex !== -1 && durColIndex >= dtColIndex) {
           actualDurColIndex = durColIndex + 1;
       }

      const durCellRef = XLSX.utils.encode_cell({ r: R, c: actualDurColIndex });
      const durCell = newWs[durCellRef];

      if (durCell && typeof durCell.v === 'string') {
          const durationStr = durCell.v;
          const parts = durationStr.split(':').map(Number);
          let totalSeconds = 0;
          if (parts.length === 2 && !isNaN(parts[0]) && !isNaN(parts[1])) {
              totalSeconds = parts[0] * 60 + parts[1];
          } else if (parts.length === 3 && !isNaN(parts[0]) && !isNaN(parts[1]) && !isNaN(parts[2])) {
              totalSeconds = parts[0] * 3600 + parts[1] * 60 + parts[2];
          } else {
              console.warn(`Could not parse duration string: "${durationStr}" at row ${R+1}. Keeping original.`);
              durCell.t = 's';
              continue;
          }

           const excelTime = totalSeconds / (24 * 60 * 60);
           durCell.t = 'n';
           durCell.v = excelTime;
           durCell.z = totalSeconds >= 3600 ? '[h]:mm:ss' : 'mm:ss';
       } else if (durCell && typeof durCell.v === 'number') {
           // If it was already a number (e.g., from rawNumbers: true or prior conversion)
           durCell.t = 'n';
           durCell.z = durCell.v * 24 >= 1 ? '[h]:mm:ss' : 'mm:ss';
       } else if (durCell) {
            console.warn(`Unexpected type in Duration column at row ${R+1}:`, typeof durCell.v, durCell.v);
            if (!durCell.t) durCell.t = 's'; // Default to string
       }
    }
  }

  // Set column widths (optional but good for readability)
  newWs['!cols'] = newHeaders.map((header, C) => {
      let width = 10; // Default width
      const lowerHeader = header.toLowerCase();
      if (lowerHeader === "date played") width = 12;
      if (lowerHeader === "time played") width = 10;
      if (lowerHeader === "duration") width = 8;
      // Add more specific widths based on header name if needed
      return { wch: width };
  });


  // 6. Create new workbook and write to XLS buffer
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, newWs, sheetName);

  // Write to buffer (XLS format - BIFF8)
  // cellDates: false is crucial here so xlsx uses the number/string values and formats *we* set,
  // rather than potentially re-interpreting Date objects differently during the write process.
  const xlsBuffer = XLSX.write(newWorkbook, { bookType: 'xls', type: 'buffer', cellDates: false });

  return xlsBuffer;
}
