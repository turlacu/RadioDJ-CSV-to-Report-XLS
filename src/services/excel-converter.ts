import * as XLSX from 'xlsx';
import { format, parse, isValid, getHours, getMinutes, getSeconds } from 'date-fns';

/**
 * Converts a CSV file content (as string) to XLS format (as Buffer).
 * Splits "Date/Time Played" into separate date and time,
 * splits "Duration" into minutes and seconds,
 * renames columns to Romanian headers in uppercase,
 * inserts blank columns,
 * reorders columns,
 * and applies font size 10 to all cells.
 *
 * @param csvData The CSV data as a string.
 * @returns A promise that resolves to the XLS data as a Buffer.
 */
export async function convertCsvToXls(csvData: string): Promise<Buffer> {
  // 1. Parse CSV Data (semicolon delimiter)
  const workbook = XLSX.read(csvData, { type: 'string', FS: ';' });
  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as any[][];
  if (!raw.length) throw new Error('CSV data is empty or invalid.');

  // 2. Trim and index original headers
  const originalHeaders = raw[0].map(h => String(h).trim());
  const dataRows = raw.slice(1);
  const idxMap: Record<string, number> = {};
  originalHeaders.forEach((h, i) => { idxMap[h.toLowerCase()] = i; });

  const dtI = idxMap['date/time played'] ?? -1;
  const durI = idxMap['duration'] ?? -1;
  const artistI = idxMap['artist'] ?? -1;
  const titleI = idxMap['title played'] ?? idxMap['title'] ?? -1;
  const albumI = idxMap['album'] ?? -1;
  const composerI = idxMap['composer'] ?? -1;
  const yearI = idxMap['year'] ?? -1;
  const pubI = idxMap['publisher'] ?? -1;
  const copyI = idxMap['copyright'] ?? -1;

  // 3. Process each row into structured object
  const processed = dataRows.map(row => {
    let dateVal: Date | string = '';
    let timeVal: Date | string = '';
    if (dtI !== -1) {
      const full = String(row[dtI]).trim();
      const [dateStr, timeStr = ''] = full.includes(' ') ? full.split(/ (.+)/) : [full];
      const dateFmts = ['M/d/yyyy','MM/dd/yyyy','yyyy-MM-dd','dd.MM.yyyy','d.M.yyyy'];
      for (const fmt of dateFmts) {
        const d = parse(dateStr, fmt, new Date());
        if (isValid(d) && d.getFullYear() > 1900) { dateVal = d; break; }
      }
      if (!dateVal) dateVal = dateStr;
      const timeFmts = ['h:mm:ss a','hh:mm:ss a','H:mm:ss','HH:mm:ss','h:mm a','H:mm'];
      const base = format(new Date(1970, 0, 1), 'yyyy-MM-dd');
      for (const fmt of timeFmts) {
        const t = parse(`${base} ${timeStr}`, `yyyy-MM-dd ${fmt}`, new Date());
        if (isValid(t)) { timeVal = t; break; }
      }
      if (!timeVal && /^\d{1,2}:\d{1,2}(:\d{1,2})?$/.test(timeStr)) {
        const [h,m,s=0] = timeStr.split(':').map(Number);
        const tmp = new Date(1970, 0, 1);
        tmp.setHours(h, m, s, 0);
        if (isValid(tmp)) timeVal = tmp;
      }
    }
    let mins: number | string = '';
    let secs: number | string = '';
    if (durI !== -1) {
      const txt = String(row[durI]).trim();
      const parts = txt.split(':').map(Number);
      if (parts.length === 2 && parts.every(p => !isNaN(p))) {
        [mins, secs] = parts;
      } else if (parts.length === 3 && parts.every(p => !isNaN(p))) {
        mins = parts[0] * 60 + parts[1];
        secs = parts[2];
      }
    }
    return {
      dateVal,
      timeVal,
      mins,
      secs,
      title: titleI !== -1 ? row[titleI] : '',
      copy: copyI !== -1 ? row[copyI] : '',
      composer: composerI !== -1 ? row[composerI] : '',
      artist: artistI !== -1 ? row[artistI] : '',
      album: albumI !== -1 ? row[albumI] : '',
      pub: pubI !== -1 ? row[pubI] : '',
      year: yearI !== -1 ? row[yearI] : ''
    };
  });

  // 4. Define final columns
  const columnMapping = [
    { header: 'DATA DIFUZARII', key: 'dateVal' },
    { header: 'NUMELE EMISIUNII', key: null },
    { header: 'ORA DIFUZARII', key: 'timeVal' },
    { header: 'MINUTE DIFUZATE', key: 'mins' },
    { header: 'SECUNDE DIFUZATE', key: 'secs' },
    { header: 'TITLUL PIESEI', key: 'title' },
    { header: 'AUTOR MUZICA', key: 'copy' },
    { header: 'AUTOR TEXT', key: 'composer' },
    { header: 'ARTIST', key: 'artist' },
    { header: 'ORCHESTRA, FORMATIE, GRUP', key: null },
    { header: 'NR. DE ARTISTI', key: null },
    { header: 'ALBUM', key: 'album' },
    { header: 'NUMAR CATALOG', key: null },
    { header: 'LABEL', key: null },
    { header: 'PRODUCATOR', key: 'pub' },
    { header: 'TARA', key: null },
    { header: 'ANUL INREGISTRARII', key: 'year' },
    { header: 'TIPUL INREGISTRARII', key: null }
  ];

  // 5. Build sheet data
  const aoa = [
    columnMapping.map(c => c.header),
    ...processed.map(r => columnMapping.map(c => c.key ? (r as any)[c.key] : ''))
  ];
  const newWs = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });

  // 6. Style and type conversion
  const epoch = Date.UTC(1899, 11, 30);
  const range = XLSX.utils.decode_range(newWs['!ref']!);
  const idx: Record<string, number> = {};
  columnMapping.forEach((c, i) => idx[c.header] = i);

  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const ref = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = newWs[ref];
      if (!cell) continue;
      // apply font size 10
      cell.s = cell.s || {};
      cell.s.font = { ...(cell.s.font || {}), sz: 10 };

      // header row no further conversion
      if (R === range.s.r) continue;

      const header = aoa[0][C];
      // Date
      if (header === 'DATA DIFUZARII' && cell.v instanceof Date) {
        const dt: Date = cell.v;
        cell.v = (Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate()) - epoch) / 86400000;
        cell.t = 'n';
        cell.z = 'mm/dd/yyyy';
      }
      // Time
      if (header === 'ORA DIFUZARII' && cell.v instanceof Date) {
        const d: Date = cell.v;
        const h = getHours(d), m = getMinutes(d), s = getSeconds(d);
        cell.v = (h * 3600 + m * 60 + s) / 86400;
        cell.t = 'n';
        cell.z = 'hh:mm:ss';
      }
      // Numbers
      if (['MINUTE DIFUZATE','SECUNDE DIFUZATE'].includes(header) && typeof cell.v === 'number') {
        cell.t = 'n';
        cell.z = '0';
      }
    }
  }

  // 7. Set column widths
  newWs['!cols'] = [
    {wch:14}, // A
    {wch:8},  // B
    {wch:14}, // C
    {wch:6.5},// D
    {wch:6.5},// E
    {wch:48}, // F
    {wch:8},  // G
    {wch:8},  // H
    {wch:48}, // I
    {wch:8},  // J
    {wch:8},  // K
    {wch:48}, // L
    {wch:8},  // M
    {wch:8},  // N
    {wch:8},  // O
    {wch:8},  // P
    {wch:8},  // Q
    {wch:8}   // R
  ];

  // 8. Write workbook
  const newWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWb, newWs, sheetName);
  return XLSX.write(newWb, { bookType: 'xls', type: 'buffer', cellDates: true });
}
