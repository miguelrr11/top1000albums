// @ts-nocheck
/**
 * Script TypeScript para Excel - Automatización de listado de álbumes
 *
 * Este script:
 * 1. Busca todas las celdas que comienzan con '*' (marcador de álbumes)
 * 2. Para cada álbum, extrae las notas de las canciones siguientes
 * 3. Calcula estadísticas: media, desviación típica, notas ≥10, etc.
 * 4. Lista los álbumes ordenados por media en columnas a la derecha
 *
 * INSTRUCCIONES DE USO:
 * 1. Marca cada título de álbum con '*' al inicio (ejemplo: *TOOL - Lateralus)
 * 2. Asegúrate de que el formato sea: artista - álbum (con guion '-')
 * 3. Las canciones deben estar en filas consecutivas debajo del título
 * 4. Las notas van en la columna inmediatamente a la derecha del nombre de canción
 * 5. Los interludios son canciones sin nota (celda vacía o sin número válido)
 * 6. Copia y pega este código en Excel: Automatizar > Nuevo script
 *
 * NOTA: El comentario @ts-nocheck es solo para evitar errores en editores locales.
 *       Excel proporciona automáticamente las definiciones de ExcelScript.
 *
 * ──────────────────────────────────────────────────────────────────────────────
 * CÓMO AÑADIR UNA NUEVA COLUMNA A LA TABLA DE ÁLBUMES
 * ──────────────────────────────────────────────────────────────────────────────
 * Solo necesitas hacer DOS cambios:
 *
 * 1. Añadir el campo a la interfaz AlbumInfo (al principio de main):
 *       dateOfReview: string;
 *
 * 2. Añadir una entrada al array COLUMNS (sección "COLUMN DEFINITIONS"):
 *       {
 *         header: 'Fecha Reseña',     // Texto del encabezado
 *         property: 'dateOfReview',   // Nombre del campo en AlbumInfo
 *         persistent: true,           // true = el usuario lo rellena manualmente en Excel
 *         align: 'center',            // Alineación opcional
 *       }
 *
 * ¡Eso es todo! Headers, filas de datos, formato y persistencia son automáticos.
 * ──────────────────────────────────────────────────────────────────────────────
 */

function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const usedRange = sheet.getUsedRange();

  if (!usedRange) {
    console.log("No hay datos en la hoja");
    return;
  }

  // =================== INTERFACES ===================

  interface AlbumInfo {
    titulo: string;
    artista: string;
    album: string;
    media: number;
    mediana: number;
    desviacionTipica: number;
    notasMayoresIgual10: number;
    totalCanciones: number;
    interludios: number;
    fila: number;
    genero: string;
    num105: number;
    thirdEyeScore: number;
    year: number;
    duration: string;
    durationMinutes: number;
    dateOfReview: string;
    dateOfReviewTimestamp: number;
  }

  interface ArtistaStats {
    artista: string;
    numAlbumes: number;
    media: number;
    thirdEyeScore: number;
    mediana: number;
    desviacionTipica: number;
    notasMayoresIgual10: number;
    totalCanciones: number;
    interludios: number;
    generos?: string[];
    yearRange: string;
    avgDuration: string;
  }

  /**
   * Describes a single column in the albums table.
   *
   * Fields:
   *   header            - Header text shown in the table.
   *   property          - AlbumInfo field key, or '#' for the 1-based rank number.
   *   persistent        - true = the user fills this in Excel manually;
   *                       the script reads it back before clearing so the data survives re-runs.
   *   sortProperty      - AlbumInfo key used for sorting (defaults to `property`).
   *                       Useful when the stored field differs from the sort field,
   *                       e.g. 'duration' is stored as a string but we sort by 'durationMinutes'.
   *   align             - Cell alignment: 'center' | 'left' | 'right'.
   *   bold              - Whether the column uses bold text.
   *   colorFn           - Returns a background color hex string given the cell's numeric value.
   *   artistHeader      - Override the header text in the artist summary table.
   *   artistValue       - How to compute the cell value in the artist summary table.
   *                       Defaults to stats[property] if not provided.
   *   artistSortProperty - ArtistaStats key used to sort the artist table (defaults to `property`).
   *   parseValue        - How to parse the raw Excel cell value when reading back persistent data.
   *                       Return null to skip storing the value.
   *   postAssign        - Called on the album after its persistent value is set.
   *                       Use this to update derived fields (e.g., compute durationMinutes from duration).
   */
  interface ColumnDef {
    header: string;
    numberFormat?: string;
    property: keyof AlbumInfo | '#';
    persistent: boolean;
    sortProperty?: keyof AlbumInfo;
    align?: 'center' | 'left' | 'right';
    bold?: boolean;
    colorFn?: (value: number) => string;
    artistHeader?: string;
    artistValue?: (stats: ArtistaStats, index: number) => string | number;
    artistSortProperty?: keyof ArtistaStats;
    parseValue?: (raw: string | number | boolean) => string | number | null;
    postAssign?: (album: AlbumInfo) => void;
  }

  // =================== HELPER FUNCTIONS ===================

  function getColorForScore(score: number): string {
    const lerp = (start: number, end: number, factor: number) =>
      Math.round(start + (end - start) * factor);

    const colorRojo     = [255, 73,  77 ];
    const colorAmarillo = [255, 245, 67 ];
    const colorAzul     = [0,   176, 240];

    let r: number, g: number, b: number;

    if (score <= 5) {
      [r, g, b] = colorRojo;
    } else if (score <= 7) {
      const f = (score - 5) / 2;
      r = lerp(colorRojo[0], colorAmarillo[0], f);
      g = lerp(colorRojo[1], colorAmarillo[1], f);
      b = lerp(colorRojo[2], colorAmarillo[2], f);
    } else if (score <= 10) {
      const f = (score - 7) / 3;
      r = lerp(colorAmarillo[0], colorAzul[0], f);
      g = lerp(colorAmarillo[1], colorAzul[1], f);
      b = lerp(colorAmarillo[2], colorAzul[2], f);
    } else {
      [r, g, b] = colorAzul;
    }

    const toHex = (n: number) => n.toString(16).padStart(2, '0');
    return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
  }

  function getRankingColor(rank: number): string {
    if (rank === 1) return '#FFD700';
    if (rank === 2) return '#C0C0C0';
    if (rank === 3) return '#CD7F32';
    return '#E8E8E8';
  }

  function parseDurationToMinutes(dur: string): number {
    const hMatch = dur.match(/(\d+)h/);
    const mMatch = dur.match(/(\d+)m/);
    return (hMatch ? parseInt(hMatch[1]) : 0) * 60 + (mMatch ? parseInt(mMatch[1]) : 0);
  }

  function getAlignmentEnum(align: string) {
    if (align === 'center') return ExcelScript.HorizontalAlignment.center;
    if (align === 'right')  return ExcelScript.HorizontalAlignment.right;
    return ExcelScript.HorizontalAlignment.left;
  }

  // =================== COLUMN DEFINITIONS ===================
  //
  // This is the ONLY place you need to touch to add a new column.
  // See the file header comment for step-by-step instructions.

  const COLUMNS: ColumnDef[] = [
    {
      header: '#',
      property: '#',
      persistent: false,
      align: 'center',
      bold: true,
    },
    {
      header: 'Artista',
      property: 'artista',
      persistent: false,
    },
    {
      header: 'Álbum',
      property: 'album',
      persistent: false,
      artistHeader: 'Álbumes',
      artistValue: (stats) => stats.numAlbumes,
      artistSortProperty: 'numAlbumes',
    },
    {
      header: 'Media',
      property: 'media',
      persistent: false,
      align: 'center',
    },
    {
      header: '3rd EYE SCORE',
      property: 'thirdEyeScore',
      persistent: false,
      align: 'center',
      bold: true,
      colorFn: (v) => getColorForScore(v),
    },
    {
      header: 'Mediana',
      property: 'mediana',
      persistent: false,
      align: 'center',
    },
    {
      header: 'Desv. Típica',
      property: 'desviacionTipica',
      persistent: false,
      align: 'center',
    },
    {
      header: 'Notas ≥10',
      property: 'notasMayoresIgual10',
      persistent: false,
      align: 'center',
    },
    {
      header: 'Total Canciones',
      property: 'totalCanciones',
      persistent: false,
      align: 'center',
    },
    {
      header: 'Interludios',
      property: 'interludios',
      persistent: false,
      align: 'center',
    },
    {
      header: 'Género',
      property: 'genero',
      persistent: true,
      artistHeader: 'Géneros',
      artistValue: (stats) => stats.generos ? stats.generos.join(', ') : '',
    },
    {
      header: 'Año',
      property: 'year',
      persistent: true,
      align: 'center',
      parseValue: (raw) => {
        const n = typeof raw === 'number' ? raw : parseInt((raw?.toString() || ''));
        return (!isNaN(n) && n > 1900 && n < 2100) ? n : null;
      },
      artistValue: (stats) => stats.yearRange,
      artistSortProperty: 'yearRange',
    },
    {
      header: 'Duración',
      property: 'duration',
      persistent: true,
      align: 'center',
      sortProperty: 'durationMinutes',
      parseValue: (raw) => {
        const s: string = raw.toString().trim() || '';
        return (s && s !== '0') ? s : null;
      },
      postAssign: (album) => {
        album.durationMinutes = album.duration ? parseDurationToMinutes(album.duration) : 0;
      },
      artistValue: (stats) => stats.avgDuration,
      artistSortProperty: 'avgDuration',
    },
    {
      header: 'Fecha de review',
      numberFormat: '@',
      property: 'dateOfReview',
      persistent: true,
      align: 'center',
      sortProperty: 'dateOfReviewTimestamp',
      parseValue: (raw) => {
        if (typeof raw === 'number' && raw > 0) {
          const serial = Math.floor(raw);
          const date = new Date((serial - 25569) * 86400000);
          const d = date.getUTCDate().toString().padStart(2, '0');
          const mo = (date.getUTCMonth() + 1).toString().padStart(2, '0');
          const y = date.getUTCFullYear().toString();
          return `${d}/${mo}/${y}`;
        }

        const s = raw.toString().trim();
        if (!s) return null;

        // Normalizar string a DD/MM/YYYY
        // Excel suele mandar M/D/YYYY o D/M/YYYY sin ceros — reconstruimos con padding
        const parts = s.split('/');
        if (parts.length === 3) {
          const [p0, p1, p2] = parts.map(p => p.trim());
          const d = p0.padStart(2, '0');   // día
          const mo = p1.padStart(2, '0');  // mes
          const y = p2.length === 2 ? `20${p2}` : p2; // "26" → "2026"
          return `${d}/${mo}/${y}`;
        }

        return s;
      },
      postAssign: (album) => {
        if (album.dateOfReview) {
          const parts = album.dateOfReview.split('/');
          if (parts.length === 3) {
            album.dateOfReviewTimestamp = Date.UTC(
              parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0])
            );
          }
        }
      },
      artistValue: (stats) => '-', // No tiene sentido mostrar esta columna en el resumen de artistas
    }
  ];

  // =================== DERIVED CONFIGURATION ===================

  // Base headers array (without sort asterisk) — derived from COLUMNS
  const headersBase: string[] = COLUMNS.map(col => col.header);

  // Maps header text → AlbumInfo sort property, for detecting which column has the * sort marker
  const headerToProperty: { [key: string]: keyof AlbumInfo } = {};
  for (const col of COLUMNS) {
    if (col.property === '#') continue;
    const sortProp = (col.sortProperty ?? col.property) as keyof AlbumInfo;
    headerToProperty[col.header] = sortProp;
    if (col.artistHeader) headerToProperty[col.artistHeader] = sortProp;
  }

  // =================== SCAN SPREADSHEET FOR ALBUMS ===================

  const artistasMap: { [artista: string]: AlbumInfo[] } = {};
  const albums: AlbumInfo[] = [];
  const todasLasNotas: number[] = [];
  const values = usedRange.getValues();
  const numRows = values.length;
  const numCols = values[0].length;

  for (let row = 0; row < numRows; row++) {
    for (let col = 0; col < numCols; col++) {
      const cellValue = values[row][col];

      if (typeof cellValue === 'string' && cellValue.startsWith('*')) {
        const tituloCompleto = cellValue.substring(1).trim();

        const primerGuionIndex = tituloCompleto.indexOf('-');
        let artista = tituloCompleto;
        let album = '';

        if (primerGuionIndex !== -1) {
          artista = tituloCompleto.substring(0, primerGuionIndex).trim();
          album = tituloCompleto.substring(primerGuionIndex + 1).trim();
        }

        const notas: number[] = [];
        let totalCanciones = 0;
        let interludios = 0;
        let currentRow = row + 1;

        while (currentRow < numRows) {
          const cancionNombre = values[currentRow][col];
          const notaValue = values[currentRow][col + 1];

          if (!cancionNombre ||
            (typeof cancionNombre === 'string' && cancionNombre.startsWith('*'))) {
            break;
          }

          totalCanciones++;

          if (typeof notaValue === 'number' && notaValue >= 0 && notaValue <= 10.5) {
            notas.push(notaValue);
            todasLasNotas.push(notaValue);
          } else {
            interludios++;
          }

          currentRow++;
        }

        if (notas.length > 0) {
          const media = notas.reduce((sum, nota) => sum + nota, 0) / notas.length;

          const notasOrdenadas = [...notas].sort((a, b) => a - b);
          const mitad = Math.floor(notasOrdenadas.length / 2);
          const mediana = notasOrdenadas.length % 2 === 0
            ? (notasOrdenadas[mitad - 1] + notasOrdenadas[mitad]) / 2
            : notasOrdenadas[mitad];

          const varianza = notas.reduce((sum, nota) => sum + Math.pow(nota - media, 2), 0) / notas.length;
          const desviacionTipica = Math.sqrt(varianza);

          const notasMayoresIgual10 = notas.filter(nota => nota >= 10).length;
          const num105 = notas.filter(nota => nota === 10.5).length;

          const thirdEyeScoreRaw =
            media -
            (desviacionTipica * 0.125) +
            ((mediana - media) * 0.2) +
            (num105 * 0.15) +
            ((notasMayoresIgual10 / totalCanciones) * 0.2);

          const thirdEyeScore = Math.round(thirdEyeScoreRaw * 100) / 100;

          albums.push({
            titulo: tituloCompleto,
            artista,
            album,
            media: Math.round(media * 100) / 100,
            mediana: Math.round(mediana * 100) / 100,
            desviacionTipica: Math.round(desviacionTipica * 100) / 100,
            notasMayoresIgual10,
            totalCanciones,
            interludios,
            fila: row + 1,
            genero: '',
            num105,
            thirdEyeScore,
            year: 0,
            duration: '',
            durationMinutes: 0,
            dateOfReview: '',
            dateOfReviewTimestamp: 0,
          });
        }
      }
    }
  }

  // =================== TABLE LAYOUT CONSTANTS ===================

  const columnaRanking = 17; // Column R (A=0 … R=17) — the '#' column
  const columnaInicio  = 18; // Column S — data columns start here
  const startRow = 0;

  // =================== PRE-CLEAR: READ HEADERS & PERSISTENT DATA ===================
  // We read headers and user-entered data BEFORE clearing the area, so the values survive re-runs.

  const headerRowRange = sheet.getRangeByIndexes(startRow + 1, columnaRanking, 1, headersBase.length);
  const currentHeaders = headerRowRange.getValues()[0];

  const artistasStartRowPreClear = startRow + albums.length + 4;
  const artistasHeaderRowRangePreClear = sheet.getRangeByIndexes(artistasStartRowPreClear, columnaRanking, 1, headersBase.length);
  const currentArtistasHeaders = artistasHeaderRowRangePreClear.getValues()[0];

  // Build a lookup: clean header text → column index in the current table
  const currentHeaderIndex: { [header: string]: number } = {};
  for (let i = 0; i < currentHeaders.length; i++) {
    const h = currentHeaders[i]?.toString().replace(/\*/g, '').trim();
    if (h) currentHeaderIndex[h] = i;
  }

  const artistaColIdx = currentHeaderIndex['Artista'] ?? -1;
  const albumColIdx   = currentHeaderIndex['Álbum']   ?? -1;

  // One store per persistent column: { albumKey → value }
  const persistentData: { [colHeader: string]: { [albumKey: string]: string | number } } = {};
  for (const col of COLUMNS) {
    if (col.persistent) persistentData[col.header] = {};
  }

  if (artistaColIdx !== -1 && albumColIdx !== -1) {
    const dataReadRange = sheet.getRangeByIndexes(startRow + 2, columnaRanking, 1000, currentHeaders.length);
    const dataReadValues = dataReadRange.getValues();

    for (let i = 0; i < dataReadValues.length; i++) {
      const art = dataReadValues[i][artistaColIdx]?.toString().trim();
      const alb = dataReadValues[i][albumColIdx]?.toString().trim();

      if (!art && !alb) break;
      if (!art || !alb) continue;

      const key = `${art}|${alb}`;

      for (const col of COLUMNS) {
        if (!col.persistent) continue;
        const colIdx = currentHeaderIndex[col.header];
        if (colIdx === undefined) continue;

        const raw = dataReadValues[i][colIdx];
        const parsed = col.parseValue ? col.parseValue(raw) : (raw?.toString().trim() || null);

        if (parsed !== null && parsed !== undefined && parsed !== '') {
          persistentData[col.header][key] = parsed;
        }
      }
    }

    const summary = COLUMNS
      .filter(c => c.persistent)
      .map(c => `${c.header}: ${Object.keys(persistentData[c.header]).length}`)
      .join(', ');
    console.log(`Datos persistentes leídos → ${summary}`);
  } else {
    console.log(`AVISO: No se pudieron leer datos persistentes (artistaColIdx=${artistaColIdx}, albumColIdx=${albumColIdx}). ¿Es la primera ejecución?`);
  }

  // Assign persistent values to album objects, then run postAssign hooks
  for (const album of albums) {
    const key = `${album.artista}|${album.album}`;
    for (const col of COLUMNS) {
      if (!col.persistent) continue;
      const val = persistentData[col.header][key];
      if (val !== undefined) {
        Object.assign(album, { [col.property as string]: val });
        if (col.postAssign) col.postAssign(album);
      }
    }
  }

  // =================== CLEAR TABLE AREA ===================

  sheet.getRangeByIndexes(startRow, columnaRanking, albums.length * 2 + 100, headersBase.length + 2 + 2)
    .clear(ExcelScript.ClearApplyTo.all);

  // =================== SORT DETECTION — MAIN ALBUMS TABLE ===================
  // Whichever header has a '*' suffix in the existing table becomes the sort column.
  // If none (or more than one), defaults to '3rd EYE SCORE'.

  let sortBy: keyof AlbumInfo = 'thirdEyeScore';
  let headerWithAsterisk: string | null = null;
  let asteriskCount = 0;

  for (let i = 0; i < currentHeaders.length; i++) {
    const headerValue = currentHeaders[i]?.toString().trim();
    if (headerValue && headerValue.includes('*')) {
      asteriskCount++;
      const cleanHeader = headerValue.replace(/\*/g, '').trim();
      if (headerToProperty[cleanHeader]) {
        headerWithAsterisk = cleanHeader;
        sortBy = headerToProperty[cleanHeader];
      }
    }
  }

  if (asteriskCount !== 1) {
    sortBy = 'thirdEyeScore';
    headerWithAsterisk = '3rd EYE SCORE';
  }

  console.log(`Ordenando tabla principal por: ${sortBy} (header con asterisco: "${headerWithAsterisk}")`);

  const headers = headersBase.map(h => h === headerWithAsterisk ? `${h} *` : h);

  albums.sort((a, b) => {
    const valueA = a[sortBy];
    const valueB = b[sortBy];
    if (typeof valueA === 'number' && typeof valueB === 'number') return valueB - valueA;
    if (typeof valueA === 'string' && typeof valueB === 'string') {
      const cmp = valueA.localeCompare(valueB);
      return cmp !== 0 ? cmp : b.media - a.media;
    }
    return 0;
  });

  // =================== WRITE MAIN ALBUMS TABLE ===================

  // Title
  const mainTituloRange = sheet.getRangeByIndexes(startRow, columnaRanking, 1, headers.length);
  mainTituloRange.merge();
  sheet.getCell(startRow, columnaRanking).setValue('ÁLBUMES');
  mainTituloRange.getFormat().getFont().setBold(true);
  mainTituloRange.getFormat().getFont().setSize(13);
  mainTituloRange.getFormat().getFill().setColor('#1A252F');
  mainTituloRange.getFormat().getFont().setColor('#FFFFFF');
  mainTituloRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // Headers
  const headerRange = sheet.getRangeByIndexes(startRow + 1, columnaRanking, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.getFormat().getFont().setBold(true);
  headerRange.getFormat().getFill().setColor('#2C3E50');
  headerRange.getFormat().getFont().setColor('#FFFFFF');
  headerRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // Data rows — built from COLUMNS, no hardcoded field list
  const dataRows: (string | number)[][] = albums.map((album, index) =>
    COLUMNS.map(col => {
      if (col.property === '#') return index + 1;
      const val = album[col.property as keyof AlbumInfo];
      // Treat 0 as empty for persistent columns (e.g. year=0 means "not set")
      if (val === null || val === undefined || val === '' || (col.persistent && val === 0)) return '';
      return val as string | number;
    })
  );

  if (dataRows.length > 0) {
    for (let colIdx = 0; colIdx < COLUMNS.length; colIdx++) {
      const col = COLUMNS[colIdx];
      const colRange = sheet.getRangeByIndexes(startRow + 2, columnaRanking + colIdx, dataRows.length, 1);
      if (col.numberFormat) colRange.setNumberFormatLocal(col.numberFormat);
    }

    sheet.getRangeByIndexes(startRow + 2, columnaRanking, dataRows.length, headers.length)
      .setValues(dataRows);

    // Column formats derived from COLUMNS
    for (let colIdx = 0; colIdx < COLUMNS.length; colIdx++) {
      const col = COLUMNS[colIdx];
      const colRange = sheet.getRangeByIndexes(startRow + 2, columnaRanking + colIdx, dataRows.length, 1);
      if (col.bold)  colRange.getFormat().getFont().setBold(true);
      if (col.colorFn) colRange.getFormat().getFont().setColor('#000000');
      if (col.align) colRange.getFormat().setHorizontalAlignment(getAlignmentEnum(col.align));
      if (col.numberFormat) colRange.setNumberFormatLocal(col.numberFormat);
    }

    // Alternating row background (skip the # column)
    for (let i = 0; i < albums.length; i++) {
      sheet.getRangeByIndexes(startRow + 2 + i, columnaInicio, 1, headers.length - 1)
        .getFormat().getFill().setColor(i % 2 === 0 ? '#F5F5F5' : '#FFFFFF');
    }

    // Per-cell colors: ranking medals and colorFn columns
    for (let i = 0; i < albums.length; i++) {
      const row = startRow + 2 + i;

      // # column: gold / silver / bronze / grey
      sheet.getCell(row, columnaRanking).getFormat().getFill().setColor(getRankingColor(i + 1));

      // Gradient columns (e.g. 3rd EYE SCORE)
      for (let colIdx = 0; colIdx < COLUMNS.length; colIdx++) {
        const col = COLUMNS[colIdx];
        if (col.colorFn) {
          const val = albums[i][col.property as keyof AlbumInfo] as number;
          sheet.getCell(row, columnaRanking + colIdx).getFormat().getFill().setColor(col.colorFn(val));
        }
      }
    }
  }

  sheet.getRangeByIndexes(startRow, columnaRanking, albums.length + 2, headers.length + 1)
    .getFormat().autofitColumns();

  console.log(`Procesados ${albums.length} álbumes y ordenados por ${sortBy}.`);

  // =================== BUILD ARTIST STATS ===================

  for (const album of albums) {
    if (!artistasMap[album.artista]) artistasMap[album.artista] = [];
    artistasMap[album.artista].push(album);
  }

  const artistasRepetidos = Object.keys(artistasMap).filter(a => artistasMap[a].length > 1);
  let artistasStats: ArtistaStats[] = [];

  if (artistasRepetidos.length > 0) {
    artistasStats = artistasRepetidos.map(artista => {
      const albumesArtista = artistasMap[artista];
      const numAlbumes = albumesArtista.length;
      const avgThirdEye = albumesArtista.reduce((sum, a) => sum + a.thirdEyeScore, 0) / numAlbumes;

      const anosConDato = albumesArtista.filter(a => a.year > 0).map(a => a.year);
      let yearRange = '';
      if (anosConDato.length > 0) {
        const anoMin = Math.min(...anosConDato);
        const anoMax = Math.max(...anosConDato);
        yearRange = anoMin === anoMax ? `${anoMin}` : `${anoMin} - ${anoMax}`;
      }

      const albumsConDurArtista = albumesArtista.filter(a => a.durationMinutes > 0);
      let avgDuration = '';
      if (albumsConDurArtista.length > 0) {
        const avgMin = Math.round(albumsConDurArtista.reduce((s, a) => s + a.durationMinutes, 0) / albumsConDurArtista.length);
        const h = Math.floor(avgMin / 60);
        const m = avgMin % 60;
        avgDuration = h > 0 ? `${h}h ${m}m` : `${m}m`;
      }

      return {
        artista,
        numAlbumes,
        media: Math.round((albumesArtista.reduce((sum, a) => sum + a.media, 0) / numAlbumes) * 100) / 100,
        mediana: Math.round((albumesArtista.reduce((sum, a) => sum + a.mediana, 0) / numAlbumes) * 100) / 100,
        desviacionTipica: Math.round((albumesArtista.reduce((sum, a) => sum + a.desviacionTipica, 0) / numAlbumes) * 100) / 100,
        notasMayoresIgual10: albumesArtista.reduce((sum, a) => sum + a.notasMayoresIgual10, 0),
        totalCanciones: albumesArtista.reduce((sum, a) => sum + a.totalCanciones, 0),
        interludios: albumesArtista.reduce((sum, a) => sum + a.interludios, 0),
        thirdEyeScore: Math.round(avgThirdEye * 100) / 100,
        yearRange,
        avgDuration,
      };
    });

    // Compute unique genres per artist
    for (const artistaStat of artistasStats) {
      const allGenres: string[] = artistasMap[artistaStat.artista]
        .flatMap(a => a.genero.split(',').map(g => g.trim()));
      artistaStat.generos = Array.from(new Set(allGenres.filter(g => g)));
    }

    // =================== SORT DETECTION — ARTIST TABLE ===================

    const artistasStartRow = startRow + albums.length + 3;

    let artistasSortBy: keyof ArtistaStats = 'thirdEyeScore';
    let artistasHeaderWithAsterisk: string | null = null;
    let artistasAsteriskCount = 0;

    for (let i = 0; i < currentArtistasHeaders.length; i++) {
      const headerValue = currentArtistasHeaders[i]?.toString().trim();
      if (headerValue && headerValue.includes('*')) {
        artistasAsteriskCount++;
        const cleanHeader = headerValue.replace(/\*/g, '').trim();

        // Find the matching column by header or artistHeader
        const matchingCol = COLUMNS.find(c =>
          (c.artistHeader ?? c.header) === cleanHeader || c.header === cleanHeader
        );
        if (matchingCol) {
          artistasHeaderWithAsterisk = matchingCol.artistHeader ?? matchingCol.header;
          artistasSortBy = (
            matchingCol.artistSortProperty ??
            matchingCol.sortProperty ??
            matchingCol.property
          ) as keyof ArtistaStats;
        }
      }
    }

    if (artistasAsteriskCount !== 1) {
      artistasSortBy = 'thirdEyeScore';
      artistasHeaderWithAsterisk = '3rd EYE SCORE';
    }

    console.log(`Ordenando tabla artistas por: ${artistasSortBy} (header con asterisco: "${artistasHeaderWithAsterisk}")`);

    artistasStats.sort((a, b) => {
      const valueA = a[artistasSortBy];
      const valueB = b[artistasSortBy];
      if (typeof valueA === 'number' && typeof valueB === 'number') return valueB - valueA;
      if (typeof valueA === 'string' && typeof valueB === 'string') {
        const cmp = valueA.localeCompare(valueB);
        return cmp !== 0 ? cmp : b.media - a.media;
      }
      return 0;
    });

    // Artist table headers — use artistHeader overrides, apply asterisk
    const artistasHeadersBase = COLUMNS.map(col => col.artistHeader ?? col.header);
    const artistasHeaders = artistasHeadersBase.map(h =>
      h === artistasHeaderWithAsterisk ? `${h} *` : h
    );

    // Title
    const artistasTituloRange = sheet.getRangeByIndexes(artistasStartRow, columnaRanking, 1, artistasHeaders.length);
    artistasTituloRange.merge();
    sheet.getCell(artistasStartRow, columnaRanking).setValue('RESUMEN ARTISTAS REPETIDOS');
    artistasTituloRange.getFormat().getFont().setBold(true);
    artistasTituloRange.getFormat().getFont().setSize(13);
    artistasTituloRange.getFormat().getFill().setColor('#1A252F');
    artistasTituloRange.getFormat().getFont().setColor('#FFFFFF');
    artistasTituloRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    // Headers
    const artistasHeaderRange = sheet.getRangeByIndexes(artistasStartRow + 1, columnaRanking, 1, artistasHeaders.length);
    artistasHeaderRange.setValues([artistasHeaders]);
    artistasHeaderRange.getFormat().getFont().setBold(true);
    artistasHeaderRange.getFormat().getFill().setColor('#2C3E50');
    artistasHeaderRange.getFormat().getFont().setColor('#FFFFFF');
    artistasHeaderRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    // Data rows — built from COLUMNS, no hardcoded field list
    const artistasDataRows: (string | number)[][] = artistasStats.map((stats, index) =>
      COLUMNS.map(col => {
        if (col.property === '#') return index + 1;
        if (col.artistValue) return col.artistValue(stats, index);
        const val: string | number = (stats as Record<string, string | number>)[col.property as string];
        if (val === null || val === undefined || val === '') return '';
        return val;
      })
    );

    sheet.getRangeByIndexes(artistasStartRow + 2, columnaRanking, artistasDataRows.length, artistasHeaders.length)
      .setValues(artistasDataRows);

    // Column formats derived from COLUMNS
    for (let colIdx = 0; colIdx < COLUMNS.length; colIdx++) {
      const col = COLUMNS[colIdx];
      const colRange = sheet.getRangeByIndexes(artistasStartRow + 2, columnaRanking + colIdx, artistasDataRows.length, 1);
      if (col.bold)  colRange.getFormat().getFont().setBold(true);
      if (col.colorFn) colRange.getFormat().getFont().setColor('#000000');
      if (col.align) colRange.getFormat().setHorizontalAlignment(getAlignmentEnum(col.align));
      if (col.numberFormat) colRange.setNumberFormatLocal(col.numberFormat);
    }

    // Alternating row background (skip the # column)
    for (let i = 0; i < artistasStats.length; i++) {
      sheet.getRangeByIndexes(artistasStartRow + 2 + i, columnaInicio, 1, artistasHeaders.length - 1)
        .getFormat().getFill().setColor(i % 2 === 0 ? '#F5F5F5' : '#FFFFFF');
    }

    // Per-cell colors: ranking medals and colorFn columns
    for (let i = 0; i < artistasStats.length; i++) {
      const row = artistasStartRow + 2 + i;

      sheet.getCell(row, columnaRanking).getFormat().getFill().setColor(getRankingColor(i + 1));

      for (let colIdx = 0; colIdx < COLUMNS.length; colIdx++) {
        const col = COLUMNS[colIdx];
        if (col.colorFn) {
          const val: number = (artistasStats[i] as Record<string, number>)[col.property as string];
          sheet.getCell(row, columnaRanking + colIdx).getFormat().getFill().setColor(col.colorFn(val));
        }
      }
    }

    sheet.getRangeByIndexes(artistasStartRow, columnaRanking, artistasStats.length + 2, artistasHeaders.length + 1)
      .getFormat().autofitColumns();

    console.log(`Procesados ${artistasStats.length} artistas con múltiples álbumes.`);
  }

  // =================== TOP 20 TABLE (sin repetir artistas, sin OSTs) ===================

  const top20SinRepetir = albums
    .filter((album, index, self) =>
      !album.album.includes("[OST]") &&
      index === self.findIndex(a => a.artista === album.artista)
    )
    .sort((a, b) => b.thirdEyeScore - a.thirdEyeScore)
    .slice(0, 20);

  const headersTop20 = ['#', 'Artista', 'Álbum', 'Media', '3rd EYE SCORE'];

  const top20StartRow: number = startRow + albums.length + 3 +
    (artistasRepetidos.length > 0 ? artistasStats.length + 3 : 0);

  const top20TituloRange = sheet.getRangeByIndexes(top20StartRow, columnaRanking, 1, headersTop20.length);
  top20TituloRange.merge();
  sheet.getCell(top20StartRow, columnaRanking).setValue('TOP 20 (sin repetir artistas y sin incluir OSTs)');
  top20TituloRange.getFormat().getFont().setBold(true);
  top20TituloRange.getFormat().getFont().setSize(13);
  top20TituloRange.getFormat().getFill().setColor('#1A252F');
  top20TituloRange.getFormat().getFont().setColor('#FFFFFF');
  top20TituloRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  const top20HeaderRange = sheet.getRangeByIndexes(top20StartRow + 1, columnaRanking, 1, headersTop20.length);
  top20HeaderRange.setValues([headersTop20]);
  top20HeaderRange.getFormat().getFont().setBold(true);
  top20HeaderRange.getFormat().getFill().setColor('#2C3E50');
  top20HeaderRange.getFormat().getFont().setColor('#FFFFFF');
  top20HeaderRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  const top20DataRows: (string | number)[][] = top20SinRepetir.map((album, index) => [
    index + 1,
    album.artista,
    album.album,
    album.media,
    album.thirdEyeScore,
  ]);

  if (top20DataRows.length > 0) {
    sheet.getRangeByIndexes(top20StartRow + 2, columnaRanking, top20DataRows.length, headersTop20.length)
      .setValues(top20DataRows);

    sheet.getRangeByIndexes(top20StartRow + 2, columnaRanking, top20DataRows.length, 1)
      .getFormat().getFont().setBold(true);
    sheet.getRangeByIndexes(top20StartRow + 2, columnaRanking, top20DataRows.length, 1)
      .getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    // 3rd EYE SCORE column (index 4 in headersTop20)
    const top20ThirdEyeCol = sheet.getRangeByIndexes(top20StartRow + 2, columnaRanking + 4, top20DataRows.length, 1);
    top20ThirdEyeCol.getFormat().getFont().setBold(true);
    top20ThirdEyeCol.getFormat().getFont().setColor('#000000');
    top20ThirdEyeCol.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    // Media column (index 3)
    sheet.getRangeByIndexes(top20StartRow + 2, columnaRanking + 3, top20DataRows.length, 1)
      .getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    for (let i = 0; i < top20SinRepetir.length; i++) {
      sheet.getRangeByIndexes(top20StartRow + 2 + i, columnaInicio, 1, headersTop20.length - 1)
        .getFormat().getFill().setColor(i % 2 === 0 ? '#F5F5F5' : '#FFFFFF');
    }

    for (let i = 0; i < top20SinRepetir.length; i++) {
      const row = top20StartRow + 2 + i;
      sheet.getCell(row, columnaRanking).getFormat().getFill().setColor(getRankingColor(i + 1));
      sheet.getCell(row, columnaRanking + 4).getFormat().getFill()
        .setColor(getColorForScore(top20SinRepetir[i].thirdEyeScore));
    }

    sheet.getRangeByIndexes(top20StartRow, columnaRanking, top20SinRepetir.length + 2, headersTop20.length + 1)
      .getFormat().autofitColumns();
  }

  // =================== TABLA RESUMEN: ESTADÍSTICAS GLOBALES ===================

  if (todasLasNotas.length > 0 && albums.length > 0) {
    const resumenCol = columnaRanking + headersBase.length + 2;
    const resumenStartRow = startRow;

    const notasOrd = [...todasLasNotas].sort((a, b) => a - b);
    const n = notasOrd.length;
    const totalCancionesGlobal = albums.reduce((s, a) => s + a.totalCanciones, 0);
    const totalInterludiosGlobal = albums.reduce((s, a) => s + a.interludios, 0);
    const artistasUnicos = new Set(albums.map(a => a.artista)).size;

    const mediaGlobal = todasLasNotas.reduce((s, v) => s + v, 0) / n;

    const medianaGlobal = n % 2 === 0
      ? (notasOrd[n / 2 - 1] + notasOrd[n / 2]) / 2
      : notasOrd[Math.floor(n / 2)];

    const varianzaGlobal = todasLasNotas.reduce((s, v) => s + Math.pow(v - mediaGlobal, 2), 0) / n;
    const desvGlobal = Math.sqrt(varianzaGlobal);

    const coefVariacion = (desvGlobal / mediaGlobal) * 100;

    const q1Index = Math.floor(n * 0.25);
    const q3Index = Math.floor(n * 0.75);
    const q1 = notasOrd[q1Index];
    const q3 = notasOrd[q3Index];
    const iqr = q3 - q1;

    const skewness = todasLasNotas.reduce((s, v) => s + Math.pow((v - mediaGlobal) / desvGlobal, 3), 0) / n;
    const kurtosis = (todasLasNotas.reduce((s, v) => s + Math.pow((v - mediaGlobal) / desvGlobal, 4), 0) / n) - 3;

    const notaMax = notasOrd[n - 1];
    const notaMin = notasOrd[0];
    const rango = notaMax - notaMin;

    const notasGe10 = todasLasNotas.filter(v => v >= 10).length;
    const pctGe10 = (notasGe10 / n) * 100;
    const pctInterludios = (totalInterludiosGlobal / totalCancionesGlobal) * 100;
    const cancionesPorAlbum = totalCancionesGlobal / albums.length;

    const albumsValidos = albums.filter(a => a.totalCanciones >= 2);
    const albumMasConsistente  = [...albumsValidos].sort((a, b) => a.desviacionTipica - b.desviacionTipica)[0];
    const albumMenosConsistente = [...albumsValidos].sort((a, b) => b.desviacionTipica - a.desviacionTipica)[0];
    const albumMejor = [...albums].sort((a, b) => b.media - a.media)[0];
    const albumPeor  = [...albums].sort((a, b) => a.media - b.media)[0];

    const frecuencias: { [nota: string]: number } = {};
    for (const nota of todasLasNotas) {
      const key = nota.toString();
      frecuencias[key] = (frecuencias[key] || 0) + 1;
    }
    let moda = todasLasNotas[0];
    let maxFreq = 0;
    for (const [nota, freq] of Object.entries(frecuencias)) {
      if (freq > maxFreq) { maxFreq = freq; moda = parseFloat(nota); }
    }

    const rango0a5  = todasLasNotas.filter(v => v < 5).length;
    const rango5a7  = todasLasNotas.filter(v => v >= 5 && v < 7).length;
    const rango7a8  = todasLasNotas.filter(v => v >= 7 && v < 8).length;
    const rango8a9  = todasLasNotas.filter(v => v >= 8 && v < 9).length;
    const rango9a10 = todasLasNotas.filter(v => v >= 9 && v < 10).length;
    const rango10plus = todasLasNotas.filter(v => v >= 10).length;

    const rd = (v: number) => Math.round(v * 100) / 100;

    const resumenData: (string | number)[][] = [
      ['GENERAL', ''],
      ['Total álbumes', albums.length],
      ['Artistas únicos', artistasUnicos],
      ['Total canciones', totalCancionesGlobal],
      ['Canciones con nota', n],
      ['Interludios', totalInterludiosGlobal],
      ['Canciones/álbum (media)', rd(cancionesPorAlbum)],
      ['% Interludios', `${rd(pctInterludios)}%`],
      ['', ''],
      ['NOTAS - TENDENCIA CENTRAL', ''],
      ['Media global', rd(mediaGlobal)],
      ['Mediana global', rd(medianaGlobal)],
      ['Moda (nota más frecuente)', `${moda} (×${maxFreq})`],
      ['', ''],
      ['NOTAS - DISPERSIÓN', ''],
      ['Desviación típica', rd(desvGlobal)],
      ['Coef. de variación', `${rd(coefVariacion)}%`],
      ['Rango (máx - mín)', `${rd(rango)} (${notaMin} - ${notaMax})`],
      ['Rango intercuartílico (Q3-Q1)', `${rd(iqr)} (${rd(q1)} - ${rd(q3)})`],
      ['', ''],
      ['NOTAS - FORMA DE LA DISTRIBUCIÓN', ''],
      ['Asimetría (skewness)', `${rd(skewness)} ${skewness < -0.2 ? '← sesgo alto' : skewness > 0.2 ? '← sesgo bajo' : '← simétrica'}`],
      ['Curtosis (excess)', `${rd(kurtosis)} ${kurtosis > 0.5 ? '← colas pesadas' : kurtosis < -0.5 ? '← muy agrupadas' : '← normal'}`],
      ['', ''],
      ['DISTRIBUCIÓN POR RANGOS', ''],
      ['[0, 5)',    `${rango0a5}   (${rd(rango0a5   / n * 100)}%)`],
      ['[5, 7)',    `${rango5a7}   (${rd(rango5a7   / n * 100)}%)`],
      ['[7, 8)',    `${rango7a8}   (${rd(rango7a8   / n * 100)}%)`],
      ['[8, 9)',    `${rango8a9}   (${rd(rango8a9   / n * 100)}%)`],
      ['[9, 10)',   `${rango9a10}  (${rd(rango9a10  / n * 100)}%)`],
      ['[10, 10.5]',`${rango10plus}(${rd(pctGe10)}%)`],
      ['', ''],
      ['DESTACADOS', ''],
      ['Mejor álbum',                    `${albumMejor.artista} - ${albumMejor.album} (${albumMejor.media})`],
      ['Peor álbum',                     `${albumPeor.artista} - ${albumPeor.album} (${albumPeor.media})`],
      ['Más consistente (menor desv.)',  `${albumMasConsistente.artista} - ${albumMasConsistente.album} (σ=${albumMasConsistente.desviacionTipica})`],
      ['Más irregular (mayor desv.)',    `${albumMenosConsistente.artista} - ${albumMenosConsistente.album} (σ=${albumMenosConsistente.desviacionTipica})`],
    ];

    // --- Genre stats ---
    const generoStatsMap: { [g: string]: { count: number; totalMedia: number } } = {};
    let albumsConGenero = 0;

    for (const album of albums) {
      if (album.genero) {
        albumsConGenero++;
        const generos = album.genero.split(',').map(g => g.trim()).filter(g => g);
        for (const g of generos) {
          if (!generoStatsMap[g]) generoStatsMap[g] = { count: 0, totalMedia: 0 };
          generoStatsMap[g].count++;
          generoStatsMap[g].totalMedia += album.media;
        }
      }
    }

    const generosUnicos = Object.keys(generoStatsMap);
    if (generosUnicos.length > 0) {
      const generosSorted = generosUnicos
        .map(g => ({ nombre: g, count: generoStatsMap[g].count, media: rd(generoStatsMap[g].totalMedia / generoStatsMap[g].count) }))
        .sort((a, b) => b.count - a.count);

      const generoMasFrecuente = generosSorted[0];
      const generosPorMedia = [...generosSorted].sort((a, b) => b.media - a.media);
      const generoMejor = generosPorMedia[0];
      const generoPeor  = generosPorMedia[generosPorMedia.length - 1];

      resumenData.push(
        ['', ''],
        ['GÉNEROS', ''],
        ['Álbumes con género',  `${albumsConGenero} de ${albums.length}`],
        ['Géneros únicos',      generosUnicos.length],
        ['Más frecuente',       `${generoMasFrecuente.nombre} (${generoMasFrecuente.count})`],
        ['Mejor media',         `${generoMejor.nombre} (${generoMejor.media})`],
        ['Peor media',          `${generoPeor.nombre} (${generoPeor.media})`],
        ['', ''],
        ['DESGLOSE POR GÉNERO', '']
      );
      for (const g of generosPorMedia) {
        resumenData.push([g.nombre, `${g.count} (${g.media})`]);
      }
    }

    // --- Year stats ---
    const albumsConAno = albums.filter(a => a.year > 0);
    if (albumsConAno.length > 0) {
      const anos = albumsConAno.map(a => a.year);
      const anoMin = Math.min(...anos);
      const anoMax = Math.max(...anos);

      const decadaMap: { [d: string]: { count: number; totalScore: number } } = {};
      for (const a of albumsConAno) {
        const decada = `${Math.floor(a.year / 10) * 10}s`;
        if (!decadaMap[decada]) decadaMap[decada] = { count: 0, totalScore: 0 };
        decadaMap[decada].count++;
        decadaMap[decada].totalScore += a.thirdEyeScore;
      }
      const decadasSorted = Object.keys(decadaMap).sort();

      const decadaTop = decadasSorted.reduce((best, d) =>
        decadaMap[d].count > decadaMap[best].count ? d : best, decadasSorted[0]);
      const decadaTopScore = decadasSorted.reduce((best, d) =>
        (decadaMap[d].totalScore / decadaMap[d].count) > (decadaMap[best].totalScore / decadaMap[best].count) ? d : best, decadasSorted[0]);

      resumenData.push(
        ['', ''],
        ['AÑO', ''],
        ['Álbumes con año',          `${albumsConAno.length} de ${albums.length}`],
        ['Año más antiguo',          anoMin],
        ['Año más reciente',         anoMax],
        ['Décadas con más álbumes',  `${decadaTop} (${decadaMap[decadaTop].count})`],
        ['Década mejor puntuada',    `${decadaTopScore} (${rd(decadaMap[decadaTopScore].totalScore / decadaMap[decadaTopScore].count)})`],
        ['', ''],
        ['ÁLBUMES POR DÉCADA', '']
      );
      for (const d of decadasSorted) {
        resumenData.push([d, `${decadaMap[d].count} álbumes · media ${rd(decadaMap[d].totalScore / decadaMap[d].count)}`]);
      }
    }

    // --- Duration stats ---
    const albumsConDur = albums.filter(a => a.durationMinutes > 0);
    if (albumsConDur.length > 0) {
      const durs = albumsConDur.map(a => a.durationMinutes);
      const durMin = Math.min(...durs);
      const durMax = Math.max(...durs);
      const durMedia = durs.reduce((s, v) => s + v, 0) / durs.length;

      const albumMasLargo = albumsConDur.reduce((best, a) => a.durationMinutes > best.durationMinutes ? a : best);
      const albumMasCorto = albumsConDur.reduce((best, a) => a.durationMinutes < best.durationMinutes ? a : best);

      function minutesToDisplay(m: number): string {
        const h = Math.floor(m / 60);
        const min = m % 60;
        return h > 0 ? `${h}h ${min}m` : `${min}m`;
      }

      resumenData.push(
        ['', ''],
        ['DURACIÓN', ''],
        ['Álbumes con duración', `${albumsConDur.length} de ${albums.length}`],
        ['Duración media',       minutesToDisplay(Math.round(durMedia))],
        ['Más largo',            `${albumMasLargo.artista} - ${albumMasLargo.album} (${minutesToDisplay(durMax)})`],
        ['Más corto',            `${albumMasCorto.artista} - ${albumMasCorto.album} (${minutesToDisplay(durMin)})`],
      );
    }

    // --- Correlation stats ---
    function pearsonCorr(xs: number[], ys: number[]): number {
      const nn = xs.length;
      if (nn < 3) return 0;
      const mx = xs.reduce((s, v) => s + v, 0) / nn;
      const my = ys.reduce((s, v) => s + v, 0) / nn;
      const num = xs.reduce((s, v, i) => s + (v - mx) * (ys[i] - my), 0);
      const dx = Math.sqrt(xs.reduce((s, v) => s + Math.pow(v - mx, 2), 0));
      const dy = Math.sqrt(ys.reduce((s, v) => s + Math.pow(v - my, 2), 0));
      if (dx === 0 || dy === 0) return 0;
      return Math.round((num / (dx * dy)) * 1000) / 1000;
    }

    function corrLabel(r: number): string {
      const abs = Math.abs(r);
      const dir = r >= 0 ? '↑' : '↓';
      if (abs >= 0.7) return `${dir} fuerte`;
      if (abs >= 0.4) return `${dir} moderada`;
      if (abs >= 0.2) return `${dir} débil`;
      return '≈ nula';
    }

    const scores = albums.map(a => a.thirdEyeScore);
    const corrCanciones   = pearsonCorr(albums.map(a => a.totalCanciones),   scores);
    const corrInterludios = pearsonCorr(albums.map(a => a.interludios),       scores);
    const corrDesv        = pearsonCorr(albums.map(a => a.desviacionTipica),  scores);
    const corrNotas10     = pearsonCorr(albums.map(a => a.notasMayoresIgual10), scores);

    const albumsConAnoScores = albums.filter(a => a.year > 0);
    const corrAno = albumsConAnoScores.length >= 3
      ? pearsonCorr(albumsConAnoScores.map(a => a.year), albumsConAnoScores.map(a => a.thirdEyeScore))
      : null;

    const albumsConDurScores = albums.filter(a => a.durationMinutes > 0);
    const corrDur = albumsConDurScores.length >= 3
      ? pearsonCorr(albumsConDurScores.map(a => a.durationMinutes), albumsConDurScores.map(a => a.thirdEyeScore))
      : null;

    const corrPctInterludio = pearsonCorr(
      albums.map(a => a.totalCanciones > 0 ? a.interludios / a.totalCanciones : 0),
      scores
    );

    resumenData.push(
      ['', ''],
      ['CORRELACIONES CON 3RD EYE SCORE', ''],
      ['(r: −1 negativa · 0 nula · +1 positiva)', ''],
      ['Total canciones',    `r=${corrCanciones}   · ${corrLabel(corrCanciones)}`],
      ['Interludios (absoluto)', `r=${corrInterludios} · ${corrLabel(corrInterludios)}`],
      ['% Interludios',      `r=${corrPctInterludio} · ${corrLabel(corrPctInterludio)}`],
      ['Desviación típica',  `r=${corrDesv}         · ${corrLabel(corrDesv)}`],
      ['Notas ≥10',          `r=${corrNotas10}      · ${corrLabel(corrNotas10)}`],
    );

    if (corrAno !== null) resumenData.push(['Año de publicación', `r=${corrAno} · ${corrLabel(corrAno)}`]);
    if (corrDur !== null) resumenData.push(['Duración (minutos)', `r=${corrDur} · ${corrLabel(corrDur)}`]);

    const corrPairs: [string, number][] = [
      ['Total canciones', corrCanciones],
      ['Interludios',     corrInterludios],
      ['% Interludios',   corrPctInterludio],
      ['Desv. típica',    corrDesv],
      ['Notas ≥10',       corrNotas10],
    ];
    if (corrAno !== null) corrPairs.push(['Año',      corrAno]);
    if (corrDur !== null) corrPairs.push(['Duración', corrDur]);

    const strongestCorr = corrPairs.reduce((best, cur) => Math.abs(cur[1]) > Math.abs(best[1]) ? cur : best);
    const weakestCorr   = corrPairs.reduce((best, cur) => Math.abs(cur[1]) < Math.abs(best[1]) ? cur : best);

    resumenData.push(
      ['', ''],
      ['Mayor correlación', `${strongestCorr[0]} (r=${strongestCorr[1]})`],
      ['Menor correlación', `${weakestCorr[0]}   (r=${weakestCorr[1]})`],
    );

    // Write stats table
    const tituloRange = sheet.getRangeByIndexes(resumenStartRow, resumenCol, 1, 2);
    tituloRange.merge();
    sheet.getCell(resumenStartRow, resumenCol).setValue('RESUMEN GLOBAL');
    tituloRange.getFormat().getFont().setBold(true);
    tituloRange.getFormat().getFont().setSize(13);
    tituloRange.getFormat().getFill().setColor('#1A252F');
    tituloRange.getFormat().getFont().setColor('#FFFFFF');
    tituloRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    sheet.getRangeByIndexes(resumenStartRow + 1, resumenCol, resumenData.length, 2)
      .setValues(resumenData);

    for (let i = 0; i < resumenData.length; i++) {
      const row = resumenStartRow + 1 + i;
      const label = resumenData[i][0]?.toString() || '';
      const cellRange = sheet.getRangeByIndexes(row, resumenCol, 1, 2);

      if (label === '' && resumenData[i][1] === '') continue;

      if (label === label.toUpperCase() && label.length > 1 && resumenData[i][1] === '') {
        cellRange.getFormat().getFont().setBold(true);
        cellRange.getFormat().getFill().setColor('#34495E');
        cellRange.getFormat().getFont().setColor('#FFFFFF');
        cellRange.getFormat().getFont().setSize(10);
      } else {
        cellRange.getFormat().getFill().setColor(i % 2 === 0 ? '#F5F5F5' : '#FFFFFF');
        sheet.getCell(row, resumenCol + 1).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
      }
    }

    sheet.getRangeByIndexes(resumenStartRow, resumenCol, resumenData.length + 1, 2)
      .getFormat().autofitColumns();

    console.log('Tabla resumen global generada.');
  }
}
