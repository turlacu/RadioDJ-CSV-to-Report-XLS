import * as XLSX from 'xlsx';
import { format, parse } from 'date-fns';

/**
 * Converts a CSV file content (as string) to XLS format (as Buffer).
 * Applies transformations similar to the provided VBA script.
 *
 * @param csvData The CSV data as a string.
 * @returns A promise that resolves to the XLS data as a Buffer.
 */
export async function convertCsvToXls(csvData: string): Promise<Buffer> {
  // 1. Parse CSV Data (using semicolon delimiter)
  // Assume the first row is the header
  const workbook = XLSX.read(csvData, { type: 'string', cellDates: false, FS: ';' });
  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  const jsonData: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });

  if (!jsonData || jsonData.length === 0) {
    throw new Error("CSV data is empty or invalid.");
  }

  const headers = jsonData[0];
  const dataRows = jsonData.slice(1);

  // 2. Locate "Date/Time Played" and "Duration" columns
  const dtColIndex = headers.findIndex(h => h === "Date/Time Played");
  const durColIndex = headers.findIndex(h => h === "Duration");

  if (dtColIndex === -1) {
    console.warn("Column 'Date/Time Played' not found. Skipping date/time split.");
  }
   if (durColIndex === -1) {
    console.warn("Column 'Duration' not found. Skipping duration formatting.");
  }

  let newHeaders = [...headers];
  let newDataRows = [...dataRows];

  // 3. Insert new "Time Played" column if "Date/Time Played" exists
  if (dtColIndex !== -1) {
    newHeaders.splice(dtColIndex + 1, 0, "Time Played");

    // 4. Split "Date/Time Played", format, and populate new columns
    newDataRows = dataRows.map(row => {
      const newRow = [...row]; // Create a copy to avoid modifying the original
      const fullText = row[dtColIndex]?.toString() || '';
      const splitPos = fullText.indexOf(" ");

      let datePart: Date | string = '';
      let timePart: Date | string = '';

      if (splitPos > 0) {
        const dateStr = fullText.substring(0, splitPos);
        const timeStr = fullText.substring(splitPos + 1);

        // Attempt to parse date and time - adjust formats as needed for locale
        // Using flexible parsing, fallback to string if parse fails
        try {
          // Try common date formats (adjust as needed)
          datePart = parse(dateStr, 'M/d/yyyy', new Date()); // Example: US format
          if (isNaN(datePart.getTime())) {
             datePart = parse(dateStr, 'dd.MM.yyyy', new Date()); // Example: European format
             if (isNaN(datePart.getTime())) datePart = dateStr; // Fallback to string
          }
        } catch {
          datePart = dateStr; // Fallback
        }

        try {
          // Try common time formats
          timePart = parse(timeStr, 'h:mm:ss a', new Date()); // Example: 1:23:45 PM
          if (isNaN(timePart.getTime())) {
            timePart = parse(timeStr, 'HH:mm:ss', new Date()); // Example: 13:23:45
             if (isNaN(timePart.getTime())) timePart = timeStr; // Fallback to string
          }
        } catch {
          timePart = timeStr; // Fallback
        }

      } else {
        // Handle cases where there's no space (maybe just date or just time)
        // Or just keep the original value if no space
        datePart = fullText;
        timePart = ''; // Leave time empty
      }

      // Update the row values
      newRow[dtColIndex] = datePart; // Update original Date/Time col with just Date
      newRow.splice(dtColIndex + 1, 0, timePart); // Insert the Time part

      return newRow;
    });
  }


  // 5. Prepare data for new worksheet, including formatting
  const processedData = [newHeaders, ...newDataRows];
  const newWs = XLSX.utils.aoa_to_sheet(processedData);

  // Apply formatting
  const columnFormats: { [col: number]: string } = {};

  if (dtColIndex !== -1) {
    // Format Date column (e.g., Short Date - format depends on Excel locale)
     newWs[`!cols`] = newWs[`!cols`] || [];
     newWs[`!cols`][dtColIndex] = { wch: 12 }; // Adjust width if needed
     processedData.forEach((row, r) => {
        if (r > 0 && row[dtColIndex] instanceof Date) {
            const cellRef = XLSX.utils.encode_cell({ r: r, c: dtColIndex });
            newWs[cellRef].t = 'd'; // Mark as date type
            newWs[cellRef].z = 'm/d/yyyy'; // Apply date format - adjust as needed
        }
     });

    // Format Time Played column (e.g., hh:mm:ss)
    const timeColIndex = dtColIndex + 1;
    newWs[`!cols`][timeColIndex] = { wch: 10 }; // Adjust width
     processedData.forEach((row, r) => {
        if (r > 0 && row[timeColIndex] instanceof Date) {
            const cellRef = XLSX.utils.encode_cell({ r: r, c: timeColIndex });
            newWs[cellRef].t = 'n'; // Mark as number type for time
            newWs[cellRef].z = 'hh:mm:ss'; // Apply time format
            // Convert Date object to Excel serial time format
            newWs[cellRef].v = XLSX.SSF.parse_date_code(format(row[timeColIndex] as Date, 'HH:mm:ss')).v;
        }
     });
  }

  if (durColIndex !== -1) {
    // Format Duration column (e.g., mm:ss)
    // Adjust the actual column index based on whether Time Played was inserted
    const actualDurColIndex = dtColIndex !== -1 && durColIndex > dtColIndex ? durColIndex + 1 : durColIndex;
     newWs[`!cols`] = newWs[`!cols`] || [];
     newWs[`!cols`][actualDurColIndex] = { wch: 8 }; // Adjust width
     processedData.forEach((row, r) => {
       if (r > 0 && typeof row[actualDurColIndex] === 'string') {
          const cellRef = XLSX.utils.encode_cell({ r: r, c: actualDurColIndex });
          // Assuming duration is like "0:30" or "1:25"
          const parts = (row[actualDurColIndex] as string).split(':');
          if (parts.length === 2) {
             const minutes = parseInt(parts[0], 10);
             const seconds = parseInt(parts[1], 10);
             if (!isNaN(minutes) && !isNaN(seconds)) {
                 // Convert mm:ss to Excel serial time format (fraction of a day)
                 const excelTime = (minutes * 60 + seconds) / (24 * 60 * 60);
                 newWs[cellRef] = { t: 'n', v: excelTime, z: 'mm:ss' };
             }
          }
       }
     });
  }


  // 6. Create new workbook and write to XLS buffer
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, newWs, sheetName);

  // Write to buffer (XLS format - BIFF8)
  const xlsBuffer = XLSX.write(newWorkbook, { bookType: 'xls', type: 'buffer' });

  return xlsBuffer;
}
