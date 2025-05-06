
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

  const headers = jsonData[0].map(h => String(h)); // Ensure headers are strings
  const dataRows = jsonData.slice(1);

  // 2. Locate "Date/Time Played" and "Duration" columns (case-insensitive)
  const dtColIndex = headers.findIndex(h => h.toLowerCase() === "date/time played");
  const durColIndex = headers.findIndex(h => h.toLowerCase() === "duration");

  if (dtColIndex === -1) {
    console.warn("Column 'Date/Time Played' not found. Skipping date/time split.");
  }
   if (durColIndex === -1) {
    console.warn("Column 'Duration' not found. Skipping duration formatting.");
  }

  let newHeaders = [...headers];
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
    const processedRow: (string | number | Date | null)[] = [...row]; // Start with original row data

    if (dtColIndex !== -1 && timeColIndex !== -1) {
      const fullText = row[dtColIndex]?.toString() || '';
      const splitPos = fullText.indexOf(" ");

      let datePart: Date | string | null = null;
      let timePart: Date | string | null = null;

      if (splitPos > 0) {
        const dateStr = fullText.substring(0, splitPos).trim();
        const timeStr = fullText.substring(splitPos + 1).trim();

        // --- Date Parsing ---
        // Try common date formats
        const dateFormats = ['M/d/yyyy', 'MM/dd/yyyy', 'yyyy-MM-dd', 'dd.MM.yyyy', 'd.M.yyyy'];
        for (const fmt of dateFormats) {
          const parsed = parse(dateStr, fmt, new Date());
          if (isValid(parsed)) {
            // Check if the parsed year is reasonable (e.g., not 1899 which can happen with just time)
            if (parsed.getFullYear() > 1900) {
                datePart = parsed;
                break;
            }
          }
        }
        if (datePart === null) {
            // Fallback if no format matched
            console.warn(`Could not parse date string: "${dateStr}". Keeping original.`);
            datePart = dateStr; // Keep original string if parsing fails
        }

        // --- Time Parsing ---
        // Create a consistent base date for reliable time parsing
        const baseDateForTimeParsing = new Date(1970, 0, 1); // Use epoch start or similar fixed date
        const timeFormats = ['h:mm:ss a', 'hh:mm:ss a', 'H:mm:ss', 'HH:mm:ss', 'h:mm a', 'H:mm'];
        for (const fmt of timeFormats) {
          // Combine with a known date part to help parsing
          const parsed = parse(`${format(baseDateForTimeParsing, 'yyyy-MM-dd')} ${timeStr}`, `yyyy-MM-dd ${fmt}`, new Date());
          if (isValid(parsed)) {
              timePart = parsed;
              break;
          }
        }
         if (timePart === null && /^\d{1,2}:\d{1,2}:\d{1,2}$/.test(timeStr)) {
            // Handle cases like '0:05:30' which might fail strict parsing
            const [h, m, s] = timeStr.split(':').map(Number);
            if (!isNaN(h) && !isNaN(m) && !isNaN(s)) {
                // Use the base date but set the time components
                const tempDate = new Date(baseDateForTimeParsing);
                tempDate.setHours(h, m, s, 0);
                if (isValid(tempDate)) {
                    timePart = tempDate;
                }
            }
        }

        if (timePart === null) {
          console.warn(`Could not parse time string: "${timeStr}". Setting time column to empty.`);
          timePart = ''; // Set time column to empty string if parsing fails
        }

      } else {
        // Handle cases where there's no space or invalid format
        console.warn(`Could not split date/time string: "${fullText}". Keeping original date, empty time.`);
        // Attempt to parse the whole string as just a date
         const dateFormats = ['M/d/yyyy', 'MM/dd/yyyy', 'yyyy-MM-dd', 'dd.MM.yyyy', 'd.M.yyyy'];
         let parsedAsDateOnly : Date | null = null;
         for (const fmt of dateFormats) {
            const parsed = parse(fullText, fmt, new Date());
            if (isValid(parsed) && parsed.getFullYear() > 1900) {
               parsedAsDateOnly = parsed;
               break;
            }
         }
         datePart = parsedAsDateOnly || fullText; // Keep original value in date column if parsing fails
         timePart = '';      // Set time column to empty
      }

      // Update the processed row values
      processedRow[dtColIndex] = datePart; // Update Date Played column
      processedRow.splice(timeColIndex, 0, timePart); // Insert Time Played value
    }

    processedDataRows.push(processedRow);
  });


  // 5. Prepare data for new worksheet
  const dataForSheet = [newHeaders, ...processedDataRows];
  const newWs = XLSX.utils.aoa_to_sheet(dataForSheet, { cellDates: true }); // Use cellDates: true initially

  // --- Post-processing and Formatting ---
  const range = XLSX.utils.decode_range(newWs['!ref'] || 'A1:A1');

  for (let R = range.s.r + 1; R <= range.e.r; ++R) { // Start from row 1 (data)
    // Format Date Played column
    if (dtColIndex !== -1) {
      const dateCellRef = XLSX.utils.encode_cell({ r: R, c: dtColIndex });
      const dateCell = newWs[dateCellRef];
      if (dateCell && dateCell.v instanceof Date && isValid(dateCell.v)) {
        // Ensure it's treated as a number (Excel date serial) and apply format
        dateCell.t = 'n'; // Important: Set type to number for date formatting
        dateCell.z = 'm/d/yyyy'; // Or 'yyyy-mm-dd' or other preferred format
        // dateCell.v = XLSX.SSF.parse_date_code(format(dateCell.v, 'M/d/yyyy')).v; // Convert Date to Excel serial date
      } else if (dateCell) {
        dateCell.t = 's'; // Ensure it's explicitly string if not a valid date object
      }
    }

    // Format Time Played column
    if (timeColIndex !== -1) {
        const timeCellRef = XLSX.utils.encode_cell({ r: R, c: timeColIndex });
        const timeCell = newWs[timeCellRef];
        if (timeCell && timeCell.v instanceof Date && isValid(timeCell.v)) {
            const hours = getHours(timeCell.v);
            const minutes = getMinutes(timeCell.v);
            const seconds = getSeconds(timeCell.v);
            // Calculate Excel serial time (fraction of a day)
            const excelTime = (hours * 3600 + minutes * 60 + seconds) / (24 * 60 * 60);
            timeCell.t = 'n'; // Set type to number for time formatting
            timeCell.v = excelTime; // Set the calculated serial time value
            timeCell.z = 'hh:mm:ss'; // Apply time format
        } else if (timeCell && typeof timeCell.v === 'string' && timeCell.v === '') {
            // Handle empty string case explicitly if needed
            timeCell.t = 's';
            timeCell.v = '';
        } else if (timeCell) {
             // Fallback for unexpected types or invalid dates that might slip through
            timeCell.t = 's';
        }
    }


    // Format Duration column
    if (durColIndex !== -1) {
      // Adjust actual column index based on whether Time Played was inserted
      const actualDurColIndex = timeColIndex !== -1 && durColIndex >= dtColIndex ? durColIndex + 1 : durColIndex;
      const durCellRef = XLSX.utils.encode_cell({ r: R, c: actualDurColIndex });
      const durCell = newWs[durCellRef];

      if (durCell && typeof durCell.v === 'string') {
          const durationStr = durCell.v;
          // Assuming duration is like "m:ss", "mm:ss", "h:mm:ss" etc.
          const parts = durationStr.split(':').map(Number);
          let totalSeconds = 0;
          if (parts.length === 2 && !isNaN(parts[0]) && !isNaN(parts[1])) {
              // mm:ss
              totalSeconds = parts[0] * 60 + parts[1];
          } else if (parts.length === 3 && !isNaN(parts[0]) && !isNaN(parts[1]) && !isNaN(parts[2])) {
              // hh:mm:ss
              totalSeconds = parts[0] * 3600 + parts[1] * 60 + parts[2];
          } else {
              console.warn(`Could not parse duration string: "${durationStr}". Keeping original.`);
              durCell.t = 's'; // Mark as string if parsing fails
              continue; // Skip formatting for this cell
          }

           // Convert total seconds to Excel serial time format (fraction of a day)
           const excelTime = totalSeconds / (24 * 60 * 60);
           durCell.t = 'n';
           durCell.v = excelTime;
           // Choose format based on whether duration could exceed an hour
           durCell.z = totalSeconds >= 3600 ? '[h]:mm:ss' : 'mm:ss';
       } else if (durCell && typeof durCell.v === 'number') {
           // If it was already a number (e.g., rawNumbers: true), ensure format
           durCell.t = 'n';
           // Determine format based on the numeric value (fraction of a day)
           durCell.z = durCell.v * 24 >= 1 ? '[h]:mm:ss' : 'mm:ss';
       }
    }
  }

  // Set column widths (optional but good for readability)
  newWs['!cols'] = newHeaders.map((_, C) => {
    let width = 10; // Default width
    if (C === dtColIndex) width = 12; // Date Played
    if (C === timeColIndex) width = 10; // Time Played
    if (durColIndex !== -1) {
         const actualDurColIndex = timeColIndex !== -1 && durColIndex >= dtColIndex ? durColIndex + 1 : durColIndex;
         if (C === actualDurColIndex) width = 8; // Duration
    }
    // Add widths for other columns if needed
    return { wch: width };
  });


  // 6. Create new workbook and write to XLS buffer
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, newWs, sheetName);

  // Write to buffer (XLS format - BIFF8)
  // cellDates: false is needed here so xlsx uses the number/string values we set, not re-interpreting Date objects
  const xlsBuffer = XLSX.write(newWorkbook, { bookType: 'xls', type: 'buffer', cellDates: false });

  return xlsBuffer;
}
