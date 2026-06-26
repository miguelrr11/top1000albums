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
  const albumsSheet = workbook.getWorksheet('Albums'); // Hoja Albums: fuente de datos
  const tablaAlbumsSheet = workbook.getWorksheet('Tabla Albums');
  const tablasTopSheet = workbook.getWorksheet('Tablas TOP');
  const resumenSheet = workbook.getWorksheet('Resumen');
  const topCancionesSheet = workbook.getWorksheet('TOP Canciones');
  const mapeoGenerosSheet = workbook.getWorksheet('Mapeo Generos');
  const topGenerosSheet = workbook.getWorksheet('TOP Generos');
  const usedRange = albumsSheet.getUsedRange();

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
    notaCancionesFull: number[];  //no rellena inteludios
    notaCancionesFullFull: (number | null)[]; //incluye null para interludios
    interludios: number;
    fila: number;
    subgeneros: string;
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
    subgeneros?: string[];
    yearRange: string;
    avgDuration: string;
  }

  interface CancionInfo {
    titulo: string;
    artista: string;
    albumTitulo: string;
    genero: string;
    albumThirdEyeScore: number;
    nota: number;
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
    artistAlign?: 'center' | 'left' | 'right';
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

    const colorRojo = [255, 73, 77];
    const colorAmarillo = [255, 245, 67];
    const colorAzul = [0, 176, 240];

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
    if (align === 'right') return ExcelScript.HorizontalAlignment.right;
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
      artistAlign: 'center',
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
      header: 'Subgéneros',
      property: 'subgeneros',
      persistent: true,
      artistHeader: 'Subgéneros',
      artistValue: (stats) => stats.subgeneros ? stats.subgeneros.join(', ') : '',
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
  const canciones105: CancionInfo[] = [];
  const canciones10: CancionInfo[] = [];
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
        const notasFull: (number | null)[] = [];
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

          notasFull.push(typeof notaValue === 'number' ? notaValue : null);

          if (typeof notaValue === 'number' && notaValue >= 0 && notaValue <= 10.5) {
            notas.push(notaValue);
            todasLasNotas.push(notaValue);
            if (notaValue === 10.5) {
              canciones105.push({
                titulo: cancionNombre.toString(),
                artista,
                albumTitulo: album,
                genero: '',
                albumThirdEyeScore: 0,
              });
            } else if (notaValue === 10) {
              canciones10.push({
                titulo: cancionNombre.toString(),
                artista,
                albumTitulo: album,
                genero: '',
                albumThirdEyeScore: 0,
              });
            }
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
            notaCancionesFull: notas,
            notaCancionesFullFull: notasFull,
            interludios,
            fila: row + 1,
            subgeneros: '',
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

  const columnaRanking = 0; // Column A — the '#' column
  const columnaInicio = 1;  // Column B — data columns start here
  const startRow = 0;

  // =================== PRE-CLEAR: READ HEADERS & PERSISTENT DATA ===================
  // We read headers and user-entered data BEFORE clearing the area, so the values survive re-runs.

  // Lee cabeceras de "Tabla Albums" (fila 1 = cabeceras, tras el título en fila 0)
  const headerRowRange = tablaAlbumsSheet.getRangeByIndexes(1, 0, 1, headersBase.length);
  const currentHeaders = headerRowRange.getValues()[0];

  // Lee cabeceras de artistas desde "Tablas TOP" (fila 1 = cabeceras de RESUMEN ARTISTAS)
  const artistasHeaderRowRangePreClear = tablasTopSheet.getRangeByIndexes(1, 0, 1, headersBase.length);
  const currentArtistasHeaders = artistasHeaderRowRangePreClear.getValues()[0];

  // Build a lookup: clean header text → column index in the current table
  const currentHeaderIndex: { [header: string]: number } = {};
  for (let i = 0; i < currentHeaders.length; i++) {
    const h = currentHeaders[i]?.toString().replace(/\*/g, '').trim();
    if (h) currentHeaderIndex[h] = i;
  }

  const artistaColIdx = currentHeaderIndex['Artista'] ?? -1;
  const albumColIdx = currentHeaderIndex['Álbum'] ?? -1;

  // One store per persistent column: { albumKey → value }
  const persistentData: { [colHeader: string]: { [albumKey: string]: string | number } } = {};
  for (const col of COLUMNS) {
    if (col.persistent) persistentData[col.header] = {};
  }

  if (artistaColIdx !== -1 && albumColIdx !== -1) {
    const dataReadRange = tablaAlbumsSheet.getRangeByIndexes(2, 0, 1000, currentHeaders.length);
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

  // Enriquecer canciones 10.5 y 10 con género y thirdEyeScore del álbum (disponibles tras asignar persistentes)
  for (const cancion of canciones105) {
    const matchingAlbum = albums.find(a => a.artista === cancion.artista && a.album === cancion.albumTitulo);
    if (matchingAlbum) {
      cancion.genero = matchingAlbum.subgeneros;
      cancion.albumThirdEyeScore = matchingAlbum.thirdEyeScore;
    }
  }
  for (const cancion of canciones10) {
    const matchingAlbum = albums.find(a => a.artista === cancion.artista && a.album === cancion.albumTitulo);
    if (matchingAlbum) {
      cancion.genero = matchingAlbum.subgeneros;
      cancion.albumThirdEyeScore = matchingAlbum.thirdEyeScore;
    }
  }

  // =================== CLEAR TABLE AREA ===================

  tablaAlbumsSheet.getUsedRange()?.clear(ExcelScript.ClearApplyTo.all);
  tablasTopSheet.getUsedRange()?.clear(ExcelScript.ClearApplyTo.all);
  resumenSheet.getUsedRange()?.clear(ExcelScript.ClearApplyTo.all);
  topCancionesSheet.getUsedRange()?.clear(ExcelScript.ClearApplyTo.all);
  topGenerosSheet?.getUsedRange()?.clear(ExcelScript.ClearApplyTo.all);
  // NOTA: "Mapeo Generos" NO se limpia: contiene el mapeo subgénero→padre que rellenas a mano.

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
  const mainTituloRange = tablaAlbumsSheet.getRangeByIndexes(startRow, columnaRanking, 1, headers.length);
  mainTituloRange.merge();
  tablaAlbumsSheet.getCell(startRow, columnaRanking).setValue('ÁLBUMES');
  mainTituloRange.getFormat().getFont().setBold(true);
  mainTituloRange.getFormat().getFont().setSize(13);
  mainTituloRange.getFormat().getFill().setColor('#1A252F');
  mainTituloRange.getFormat().getFont().setColor('#FFFFFF');
  mainTituloRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // Headers
  const headerRange = tablaAlbumsSheet.getRangeByIndexes(startRow + 1, columnaRanking, 1, headers.length);
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
      const colRange = tablaAlbumsSheet.getRangeByIndexes(startRow + 2, columnaRanking + colIdx, dataRows.length, 1);
      if (col.numberFormat) colRange.setNumberFormatLocal(col.numberFormat);
    }

    tablaAlbumsSheet.getRangeByIndexes(startRow + 2, columnaRanking, dataRows.length, headers.length)
      .setValues(dataRows);

    // Column formats derived from COLUMNS
    for (let colIdx = 0; colIdx < COLUMNS.length; colIdx++) {
      const col = COLUMNS[colIdx];
      const colRange = tablaAlbumsSheet.getRangeByIndexes(startRow + 2, columnaRanking + colIdx, dataRows.length, 1);
      if (col.bold) colRange.getFormat().getFont().setBold(true);
      if (col.colorFn) colRange.getFormat().getFont().setColor('#000000');
      if (col.align) colRange.getFormat().setHorizontalAlignment(getAlignmentEnum(col.align));
      if (col.artistAlign) colRange.getFormat().setHorizontalAlignment(getAlignmentEnum(col.artistAlign));
      if (col.numberFormat) colRange.setNumberFormatLocal(col.numberFormat);
    }

    // Alternating row background (skip the # column)
    for (let i = 0; i < albums.length; i++) {
      tablaAlbumsSheet.getRangeByIndexes(startRow + 2 + i, columnaInicio, 1, headers.length - 1)
        .getFormat().getFill().setColor(i % 2 === 0 ? '#F5F5F5' : '#FFFFFF');
    }

    // Per-cell colors: ranking medals and colorFn columns
    for (let i = 0; i < albums.length; i++) {
      const row = startRow + 2 + i;

      // # column: gold / silver / bronze / grey
      tablaAlbumsSheet.getCell(row, columnaRanking).getFormat().getFill().setColor(getRankingColor(i + 1));

      // Gradient columns (e.g. 3rd EYE SCORE)
      for (let colIdx = 0; colIdx < COLUMNS.length; colIdx++) {
        const col = COLUMNS[colIdx];
        if (col.colorFn) {
          const val = albums[i][col.property as keyof AlbumInfo] as number;
          tablaAlbumsSheet.getCell(row, columnaRanking + colIdx).getFormat().getFill().setColor(col.colorFn(val));
        }
      }
    }
  }

  tablaAlbumsSheet.getRangeByIndexes(startRow, columnaRanking, albums.length + 2, headers.length + 1)
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

    // Compute unique subgenres per artist
    for (const artistaStat of artistasStats) {
      const allSubgenres: string[] = artistasMap[artistaStat.artista]
        .flatMap(a => a.subgeneros.split(',').map(g => g.trim()));
      artistaStat.subgeneros = Array.from(new Set(allSubgenres.filter(g => g)));
    }

    // =================== SORT DETECTION — ARTIST TABLE ===================

    const artistasStartRow = 0; // Empieza en la primera fila de "Tablas TOP"

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
    const artistasTituloRange = tablasTopSheet.getRangeByIndexes(artistasStartRow, columnaRanking, 1, artistasHeaders.length);
    artistasTituloRange.merge();
    tablasTopSheet.getCell(artistasStartRow, columnaRanking).setValue('RESUMEN ARTISTAS REPETIDOS');
    artistasTituloRange.getFormat().getFont().setBold(true);
    artistasTituloRange.getFormat().getFont().setSize(13);
    artistasTituloRange.getFormat().getFill().setColor('#1A252F');
    artistasTituloRange.getFormat().getFont().setColor('#FFFFFF');
    artistasTituloRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    // Headers
    const artistasHeaderRange = tablasTopSheet.getRangeByIndexes(artistasStartRow + 1, columnaRanking, 1, artistasHeaders.length);
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

    tablasTopSheet.getRangeByIndexes(artistasStartRow + 2, columnaRanking, artistasDataRows.length, artistasHeaders.length)
      .setValues(artistasDataRows);

    // Column formats derived from COLUMNS
    for (let colIdx = 0; colIdx < COLUMNS.length; colIdx++) {
      const col = COLUMNS[colIdx];
      const colRange = tablasTopSheet.getRangeByIndexes(artistasStartRow + 2, columnaRanking + colIdx, artistasDataRows.length, 1);
      if (col.bold) colRange.getFormat().getFont().setBold(true);
      if (col.colorFn) colRange.getFormat().getFont().setColor('#000000');
      if (col.align) colRange.getFormat().setHorizontalAlignment(getAlignmentEnum(col.align));
      if (col.artistAlign) colRange.getFormat().setHorizontalAlignment(getAlignmentEnum(col.artistAlign));
      if (col.numberFormat) colRange.setNumberFormatLocal(col.numberFormat);
    }

    // Alternating row background (skip the # column)
    for (let i = 0; i < artistasStats.length; i++) {
      tablasTopSheet.getRangeByIndexes(artistasStartRow + 2 + i, columnaInicio, 1, artistasHeaders.length - 1)
        .getFormat().getFill().setColor(i % 2 === 0 ? '#F5F5F5' : '#FFFFFF');
    }

    // Per-cell colors: ranking medals and colorFn columns
    for (let i = 0; i < artistasStats.length; i++) {
      const row = artistasStartRow + 2 + i;

      tablasTopSheet.getCell(row, columnaRanking).getFormat().getFill().setColor(getRankingColor(i + 1));

      for (let colIdx = 0; colIdx < COLUMNS.length; colIdx++) {
        const col = COLUMNS[colIdx];
        if (col.colorFn) {
          const val: number = (artistasStats[i] as Record<string, number>)[col.property as string];
          tablasTopSheet.getCell(row, columnaRanking + colIdx).getFormat().getFill().setColor(col.colorFn(val));
        }
      }
    }

    tablasTopSheet.getRangeByIndexes(artistasStartRow, columnaRanking, artistasStats.length + 2, artistasHeaders.length + 1)
      .getFormat().autofitColumns();

    console.log(`Procesados ${artistasStats.length} artistas con múltiples álbumes.`);
  }

  // funcion que hace tablas top de artistas
  function renderRankingTable(
    sheet: ExcelScript.Worksheet,
    ranking: AlbumInfo[], // ya filtrado, ordenado y recortado
    titulo: string,
    startRow: number,
    startCol: number,
    getRankingColor: (pos: number) => string,
    getColorForScore: (score: number) => string,
  ): void {
    const headers = ['#', 'Artista', 'Álbum', 'Media', '3rd EYE SCORE'];
    const numCols = headers.length;

    // Título
    const tituloRange = sheet.getRangeByIndexes(startRow, startCol, 1, numCols);
    tituloRange.merge();
    sheet.getCell(startRow, startCol).setValue(titulo);
    tituloRange.getFormat().getFont().setBold(true);
    tituloRange.getFormat().getFont().setSize(13);
    tituloRange.getFormat().getFill().setColor('#1A252F');
    tituloRange.getFormat().getFont().setColor('#FFFFFF');
    tituloRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    // Cabeceras
    const headerRange = sheet.getRangeByIndexes(startRow + 1, startCol, 1, numCols);
    headerRange.setValues([headers]);
    headerRange.getFormat().getFont().setBold(true);
    headerRange.getFormat().getFill().setColor('#2C3E50');
    headerRange.getFormat().getFont().setColor('#FFFFFF');
    headerRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    // Datos
    const dataRows: (string | number)[][] = ranking.map((album, index) => [
      index + 1,
      album.artista,
      album.album,
      album.media,
      album.thirdEyeScore,
    ]);

    if (dataRows.length === 0) return;

    const dataStartRow = startRow + 2;

    sheet.getRangeByIndexes(dataStartRow, startCol, dataRows.length, numCols)
      .setValues(dataRows);

    // Columna # (negrita, centrada)
    const colNumero = sheet.getRangeByIndexes(dataStartRow, startCol, dataRows.length, 1);
    colNumero.getFormat().getFont().setBold(true);
    colNumero.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    // Columna 3rd EYE SCORE (índice 4)
    const colScore = sheet.getRangeByIndexes(dataStartRow, startCol + 4, dataRows.length, 1);
    colScore.getFormat().getFont().setBold(true);
    colScore.getFormat().getFont().setColor('#000000');
    colScore.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    // Columna Media (índice 3)
    sheet.getRangeByIndexes(dataStartRow, startCol + 3, dataRows.length, 1)
      .getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    // Filas alternadas (de la columna 1 a la 3, dejando # y score con su color propio)
    for (let i = 0; i < ranking.length; i++) {
      sheet.getRangeByIndexes(dataStartRow + i, startCol + 1, 1, numCols - 2)
        .getFormat().getFill().setColor(i % 2 === 0 ? '#F5F5F5' : '#FFFFFF');
    }

    // Colores por posición y por score
    for (let i = 0; i < ranking.length; i++) {
      const row = dataStartRow + i;
      sheet.getCell(row, startCol).getFormat().getFill().setColor(getRankingColor(i + 1));
      sheet.getCell(row, startCol + 4).getFormat().getFill()
        .setColor(getColorForScore(ranking[i].thirdEyeScore));
    }

    sheet.getRangeByIndexes(startRow, startCol, ranking.length + 2, numCols + 1)
      .getFormat().autofitColumns();
  }

  // En "Tablas TOP": empieza tras RESUMEN ARTISTAS (si existe) o desde el principio
  const top20StartRow: number = artistasRepetidos.length > 0 ? artistasStats.length + 3 : 0;

  const nRows: number = 25;

  // Tabla 1: top 25 sin repetir artistas y sin OST/LIVE
  const rankingSinRepetir: AlbumInfo[] = albums
    .filter(a => !a.album.includes("[OST]") && !a.album.includes("[LIVE]"))
    .filter((album, index, arr) => index === arr.findIndex(a => a.artista === album.artista))
    .sort((a, b) => b.thirdEyeScore - a.thirdEyeScore)
    .slice(0, nRows);

  renderRankingTable(
    tablasTopSheet,
    rankingSinRepetir,
    'TOP 25 (sin repetir artistas y sin incluir OSTs y LIVEs)',
    top20StartRow,
    columnaRanking,
    getRankingColor,
    getColorForScore,
  );

  // Tabla 2: top 25 sin filtros, justo debajo
  const rankingSinReglas: AlbumInfo[] = albums
    .slice() // copia para no mutar el original al ordenar
    .sort((a, b) => b.thirdEyeScore - a.thirdEyeScore)
    .slice(0, nRows);

  renderRankingTable(
    tablasTopSheet,
    rankingSinReglas,
    'TOP 25 (sin reglas)',
    top20StartRow + nRows + 3,
    columnaRanking,
    getRankingColor,
    getColorForScore,
  );

  function renderTablasPorGenero(
    sheet: ExcelScript.Worksheet,
    albums: AlbumInfo[],
    startRow: number,
    startCol: number,
    getRankingColor: (pos: number) => string,
    getColorForScore: (score: number) => string,
  ): number {
    const MIN_APARICIONES = 5;
    const TOP_N = 5;
    const FILAS_POR_TABLA = TOP_N + 3;
    const SEPARACION = 0;
    const TABLAS_POR_FILA = 1;
    const ANCHO_TABLA = 6;

    // 1. Contar apariciones de cada subgénero
    const conteo: { [subgenero: string]: number } = {};
    for (const album of albums) {
      if (!album.subgeneros) continue;
      const subgeneros = album.subgeneros.split(',').map(g => g.trim()).filter(g => g.length > 0);
      for (const g of subgeneros) {
        conteo[g] = (conteo[g] || 0) + 1;
      }
    }

    // 2. Filtrar subgéneros con >= MIN_APARICIONES y ordenar por frecuencia descendente
    const generosValidos: string[] = Object.keys(conteo)
      .filter(g => conteo[g] >= MIN_APARICIONES)
      .sort((a, b) => conteo[b] - conteo[a]);

    // 3. Por cada subgénero, construir el top fusionando artistas repetidos
    let tablasRenderizadas = 0;
    let filaActual = startRow;
    let ultimaFilaUsada = startRow;

    for (const genero of generosValidos) {
      const candidatos: AlbumInfo[] = albums
        .filter(a => {
          if (!a.subgeneros) return false;
          const subgeneros = a.subgeneros.split(',').map(g => g.trim());
          return subgeneros.indexOf(genero) !== -1;
        })
        .slice()
        .sort((a, b) => b.thirdEyeScore - a.thirdEyeScore);

      const topFusionado = construirTopFusionado(candidatos, TOP_N);

      const columnaTabla = startCol + (tablasRenderizadas % TABLAS_POR_FILA) * ANCHO_TABLA;
      const filaTabla = filaActual;
      const generoCapitalizado = genero.split(' ')
        .map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');

      renderRankingTable(
        sheet,
        topFusionado,
        `TOP  ${generoCapitalizado} (${conteo[genero]} álbumes)`,
        filaTabla,
        columnaTabla,
        getRankingColor,
        getColorForScore,
      );

      const filaFinTabla = filaTabla + FILAS_POR_TABLA - 1;
      if (filaFinTabla > ultimaFilaUsada) ultimaFilaUsada = filaFinTabla;
      tablasRenderizadas++;

      if (tablasRenderizadas % TABLAS_POR_FILA === 0) {
        filaActual = ultimaFilaUsada + SEPARACION + 1;
      }
    }

    return ultimaFilaUsada;
  }

  /**
   * Recorre los candidatos ya ordenados por thirdEyeScore desc y construye
   * un top de hasta `topN` entradas. Si un artista ya está en el top, sus
   * álbumes adicionales se fusionan en la entrada existente (álbumes
   * concatenados, media y thirdEyeScore promediados). Como cada fusión
   * libera un hueco, sigue avanzando hasta llenar topN entradas únicas.
   */
  function construirTopFusionado(candidatos: AlbumInfo[], topN: number): AlbumInfo[] {
    const top: AlbumInfo[] = [];
    // Mantenemos arrays paralelos con los valores originales de cada entrada
    // fusionada para poder recalcular promedios correctamente al añadir más.
    const mediasPorEntrada: number[][] = [];
    const scoresPorEntrada: number[][] = [];

    for (const album of candidatos) {
      const idx = top.findIndex(a => a.artista === album.artista);

      if (idx === -1) {
        // Nuevo artista: entra como fila propia
        if (top.length >= topN) break;
        top.push({ ...album }); // copia para no mutar el original
        mediasPorEntrada.push([album.media]);
        scoresPorEntrada.push([album.thirdEyeScore]);
      } else {
        // Artista ya presente: fusionamos en la entrada existente
        top[idx].album = `${top[idx].album}, ${album.album}`;
        mediasPorEntrada[idx].push(album.media);
        scoresPorEntrada[idx].push(album.thirdEyeScore);
        top[idx].media = promedio(mediasPorEntrada[idx]);
        top[idx].thirdEyeScore = promedio(scoresPorEntrada[idx]);
        // Importante: NO incrementamos el "tamaño" del top, así que el siguiente
        // candidato podrá llenar el hueco hasta llegar a topN entradas únicas.
      }
    }

    return top;
  }

  function promedio(nums: number[]): number {
    let suma = 0;
    for (const n of nums) suma += n;
    let media = suma / nums.length;
    let rounded = Math.round(media * 100) / 100;
    return rounded
  }

  const filaInicioGeneros: number = top20StartRow + nRows + 3 + nRows + 3 + 3;
  // (la primera tabla ocupa nRows+3 filas, la segunda otras nRows+3, +3 de margen)

  renderTablasPorGenero(
    tablasTopSheet,
    albums,
    filaInicioGeneros,
    columnaRanking,
    getRankingColor,
    getColorForScore,
  );

  // =================== MAPEO SUBGÉNERO → GÉNERO PADRE ===================
  // La hoja "Mapeo Generos" la rellena el usuario a mano:
  //   Col A "Subgénero" · Col B "Género Padre" · Col C "Géneros Únicos" (la genera el script).
  // El script: (1) lee el mapeo existente, (2) añade los subgéneros usados que falten
  // (con padre vacío para que el usuario los complete), (3) reescribe la lista de géneros únicos.

  // Todos los subgéneros realmente usados en los álbumes
  const subgenerosUsados = new Set<string>();
  for (const album of albums) {
    if (!album.subgeneros) continue;
    for (const s of album.subgeneros.split(',').map(g => g.trim()).filter(g => g)) {
      subgenerosUsados.add(s);
    }
  }

  // Mapa subgénero → género padre, leído de la hoja
  const subgenreToParent: { [sub: string]: string } = {};

  if (mapeoGenerosSheet) {
    const mapeoUsed = mapeoGenerosSheet.getUsedRange();
    const existingSubs: string[] = []; // subgéneros ya presentes en col A
    let mapeoRowCount = 1; // nº de filas usadas (mínimo: la cabecera)

    if (mapeoUsed) {
      const mv = mapeoUsed.getValues();
      mapeoRowCount = mv.length;
      for (let r = 1; r < mv.length; r++) { // fila 0 = cabeceras
        const sub = mv[r][0]?.toString().trim();
        const parent = mv[r][1]?.toString().trim();
        if (!sub) continue;
        existingSubs.push(sub);
        if (parent) subgenreToParent[sub] = parent;
      }
    }

    // Subgéneros usados que faltan en la hoja → se añaden bajo la última fila usada, con padre vacío
    const existingSet = new Set(existingSubs);
    const faltantes = Array.from(subgenerosUsados).filter(s => !existingSet.has(s)).sort();

    if (faltantes.length > 0) {
      const startAppendRow = mapeoRowCount; // primera fila libre tras lo ya escrito
      mapeoGenerosSheet.getRangeByIndexes(startAppendRow, 0, faltantes.length, 2)
        .setValues(faltantes.map(s => [s, '']));
      console.log(`Mapeo Generos: añadidos ${faltantes.length} subgéneros nuevos sin padre → ${faltantes.join(', ')}`);
    }

    // Aviso de subgéneros usados que aún no tienen género padre asignado
    const sinPadre = Array.from(subgenerosUsados).filter(s => !subgenreToParent[s]).sort();
    if (sinPadre.length > 0) {
      console.log(`AVISO: ${sinPadre.length} subgéneros sin género padre (rellénalos en 'Mapeo Generos'): ${sinPadre.join(', ')}`);
    }

    // Col C "Géneros Únicos": géneros padre distintos (orden alfabético)
    const padresUnicos = Array.from(new Set(Object.values(subgenreToParent))).filter(p => p).sort();
    const filasALimpiar = mapeoRowCount + faltantes.length + 10;
    mapeoGenerosSheet.getRangeByIndexes(1, 2, filasALimpiar, 1).clear(ExcelScript.ClearApplyTo.contents);
    if (padresUnicos.length > 0) {
      mapeoGenerosSheet.getRangeByIndexes(1, 2, padresUnicos.length, 1)
        .setValues(padresUnicos.map(p => [p]));
    }
  } else {
    console.log("AVISO: no existe la hoja 'Mapeo Generos'; se omiten géneros padre.");
  }

  // Devuelve los géneros padre de un álbum (sin duplicar si dos subgéneros comparten padre)
  function getGenerosPadre(album: AlbumInfo): string[] {
    if (!album.subgeneros) return [];
    const parents = new Set<string>();
    for (const sub of album.subgeneros.split(',').map(g => g.trim()).filter(g => g)) {
      const parent = subgenreToParent[sub];
      if (parent) parents.add(parent);
    }
    return Array.from(parents);
  }

  // Agrupa álbumes por género padre (cada álbum cuenta una sola vez por género)
  const generosPadreMap: { [parent: string]: AlbumInfo[] } = {};
  for (const album of albums) {
    for (const parent of getGenerosPadre(album)) {
      if (!generosPadreMap[parent]) generosPadreMap[parent] = [];
      generosPadreMap[parent].push(album);
    }
  }

  // =================== TOP GÉNEROS PADRE (hoja "TOP Generos") ===================
  // Una tabla por género padre con TODOS sus álbumes (una fila por álbum), 3rd EYE SCORE desc.

  if (topGenerosSheet) {
    const parentsByCount = Object.keys(generosPadreMap)
      .sort((a, b) => generosPadreMap[b].length - generosPadreMap[a].length);

    let filaGenero = 0;
    for (const parent of parentsByCount) {
      const ranking = generosPadreMap[parent]
        .slice()
        .sort((a, b) => b.thirdEyeScore - a.thirdEyeScore);

      const parentCap = parent.split(' ')
        .map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');

      renderRankingTable(
        topGenerosSheet,
        ranking,
        `TOP ${parentCap} (${ranking.length} álbumes)`,
        filaGenero,
        columnaRanking,
        getRankingColor,
        getColorForScore,
      );

      // título(1) + cabecera(1) + datos(n) + 1 fila de margen
      filaGenero += ranking.length + 3;
    }
    console.log(`TOP Generos: ${parentsByCount.length} géneros padre renderizados.`);
  }

  // =================== TOP CANCIONES 10.5 ===================

  function renderTopCanciones(
    sheet: ExcelScript.Worksheet,
    canciones: CancionInfo[],
    maxRows: number,
    maxPerAlbum: number,
    cancionesRelleno: CancionInfo[] = [],
  ): void {
    const headers = ['#', 'Canción', 'Álbum', 'Artista', 'Subgénero'];
    const numColsTabla = headers.length;

    // Agrupar canciones por álbum, luego recorrer álbumes ordenados por thirdEyeScore desc.
    // Por cada álbum se añaden hasta maxPerAlbum canciones.
    const songsByAlbum: { [key: string]: CancionInfo[] } = {};
    for (const c of canciones) {
      const key = `${c.artista}|${c.albumTitulo}`;
      if (!songsByAlbum[key]) songsByAlbum[key] = [];
      songsByAlbum[key].push(c);
    }

    const albumKeys = Object.keys(songsByAlbum)
      .sort((a, b) => songsByAlbum[b][0].albumThirdEyeScore - songsByAlbum[a][0].albumThirdEyeScore);

    const result: CancionInfo[] = [];

    for (const key of albumKeys) {
      if (result.length >= maxRows) break;
      const songs = songsByAlbum[key];
      const toAdd = songs.slice(0, Math.min(maxPerAlbum, maxRows - result.length));
      for (const s of toAdd) result.push(s);
    }

    // Relleno con canciones 10 si no se ha llegado a maxRows
    if (result.length < maxRows && cancionesRelleno.length > 0) {
      const albumsYaUsados = new Set(result.map(c => `${c.artista}|${c.albumTitulo}`));

      const fillerByAlbum: { [key: string]: CancionInfo[] } = {};
      for (const c of cancionesRelleno) {
        const key = `${c.artista}|${c.albumTitulo}`;
        if (albumsYaUsados.has(key)) continue;
        if (!fillerByAlbum[key]) fillerByAlbum[key] = [];
        fillerByAlbum[key].push(c);
      }

      const fillerAlbumKeys = Object.keys(fillerByAlbum)
        .sort((a, b) => fillerByAlbum[b][0].albumThirdEyeScore - fillerByAlbum[a][0].albumThirdEyeScore);

      for (const key of fillerAlbumKeys) {
        if (result.length >= maxRows) break;
        const songs = fillerByAlbum[key];
        const toAdd = songs.slice(0, Math.min(maxPerAlbum, maxRows - result.length));
        for (const s of toAdd) result.push(s);
      }
    }

    // Título
    const fillerCount = result.filter(c => !canciones.some(x => x.titulo === c.titulo && x.artista === c.artista && x.albumTitulo === c.albumTitulo)).length;
    const titulo = fillerCount > 0
      ? `TOP ${maxRows} CANCIONES 10.5 + relleno 10 · máx. ${maxPerAlbum} por álbum`
      : `TOP ${maxRows} CANCIONES 10.5 (máx. ${maxPerAlbum} por álbum)`;
    const tituloRange = sheet.getRangeByIndexes(0, 0, 1, numColsTabla);
    tituloRange.merge();
    sheet.getCell(0, 0).setValue(titulo);
    tituloRange.getFormat().getFont().setBold(true);
    tituloRange.getFormat().getFont().setSize(13);
    tituloRange.getFormat().getFill().setColor('#1A252F');
    tituloRange.getFormat().getFont().setColor('#FFFFFF');
    tituloRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    // Cabeceras
    const headerRange = sheet.getRangeByIndexes(1, 0, 1, numColsTabla);
    headerRange.setValues([headers]);
    headerRange.getFormat().getFont().setBold(true);
    headerRange.getFormat().getFill().setColor('#2C3E50');
    headerRange.getFormat().getFont().setColor('#FFFFFF');
    headerRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    if (result.length === 0) return;

    // Forzar formato texto en Canción para que nombres numéricos no se conviertan a número
    sheet.getRangeByIndexes(2, 1, result.length, 1).setNumberFormatLocal('@');

    // Forzar tambien a album
    sheet.getRangeByIndexes(2, 2, result.length, 1).setNumberFormatLocal('@');

    // Datos
    const dataRows: (string | number)[][] = result.map((c, i) => [
      i + 1,
      c.titulo,
      c.albumTitulo,
      c.artista,
      c.genero,
    ]);

    sheet.getRangeByIndexes(2, 0, dataRows.length, numColsTabla).setValues(dataRows);

    // Columna # (negrita, centrada)
    const colNum = sheet.getRangeByIndexes(2, 0, dataRows.length, 1);
    colNum.getFormat().getFont().setBold(true);
    colNum.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    // Filas alternadas (columnas 1-4, dejando # con su color propio)
    for (let i = 0; i < result.length; i++) {
      sheet.getRangeByIndexes(2 + i, 1, 1, numColsTabla - 1)
        .getFormat().getFill().setColor(i % 2 === 0 ? '#F5F5F5' : '#FFFFFF');
    }

    // Colores medalla en columna #
    for (let i = 0; i < result.length; i++) {
      sheet.getCell(2 + i, 0).getFormat().getFill().setColor(getRankingColor(i + 1));
    }

    // Columna #: autofit; columnas de texto: anchos fijos generosos; Subgénero: autofit
    sheet.getRangeByIndexes(0, 0, dataRows.length + 2, 1).getFormat().autofitColumns(); // #
    sheet.getRangeByIndexes(0, 1, 1, 1).getFormat().setColumnWidth(220); // Canción
    sheet.getRangeByIndexes(0, 2, 1, 1).getFormat().setColumnWidth(200); // Álbum
    sheet.getRangeByIndexes(0, 3, 1, 1).getFormat().setColumnWidth(150); // Artista
    sheet.getRangeByIndexes(0, 4, dataRows.length + 2, 1).getFormat().autofitColumns(); // Subgénero

    console.log(`TOP Canciones: ${result.length} canciones generadas.`);
  }

  renderTopCanciones(topCancionesSheet, canciones105, 100, 1, canciones10);

  // =================== TABLA RESUMEN: ESTADÍSTICAS GLOBALES ===================

  if (todasLasNotas.length > 0 && albums.length > 0) {
    const resumenCol = 0; // Empieza en la primera columna de "Resumen"
    const resumenStartRow = 0;

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
    const albumMasConsistente = [...albumsValidos].sort((a, b) => a.desviacionTipica - b.desviacionTipica)[0];
    const albumMenosConsistente = [...albumsValidos].sort((a, b) => b.desviacionTipica - a.desviacionTipica)[0];
    const albumMejor = [...albums].sort((a, b) => b.media - a.media)[0];
    const albumPeor = [...albums].sort((a, b) => a.media - b.media)[0];

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

    const rango0a5 = todasLasNotas.filter(v => v < 5).length;
    const rango5a7 = todasLasNotas.filter(v => v >= 5 && v < 7).length;
    const rango7a8 = todasLasNotas.filter(v => v >= 7 && v < 8).length;
    const rango8a9 = todasLasNotas.filter(v => v >= 8 && v < 9).length;
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
      ['[0, 5)', `${rango0a5}   (${rd(rango0a5 / n * 100)}%)`],
      ['[5, 7)', `${rango5a7}   (${rd(rango5a7 / n * 100)}%)`],
      ['[7, 8)', `${rango7a8}   (${rd(rango7a8 / n * 100)}%)`],
      ['[8, 9)', `${rango8a9}   (${rd(rango8a9 / n * 100)}%)`],
      ['[9, 10)', `${rango9a10}  (${rd(rango9a10 / n * 100)}%)`],
      ['[10, 10.5]', `${rango10plus}(${rd(pctGe10)}%)`],
      ['', ''],
      ['DESTACADOS', ''],
      ['Mejor álbum', `${albumMejor.artista} - ${albumMejor.album} (${albumMejor.media})`],
      ['Peor álbum', `${albumPeor.artista} - ${albumPeor.album} (${albumPeor.media})`],
      ['Más consistente (menor desv.)', `${albumMasConsistente.artista} - ${albumMasConsistente.album} (σ=${albumMasConsistente.desviacionTipica})`],
      ['Más irregular (mayor desv.)', `${albumMenosConsistente.artista} - ${albumMenosConsistente.album} (σ=${albumMenosConsistente.desviacionTipica})`],
    ];

    // --- Subgenre stats ---
    const subgeneroStatsMap: { [g: string]: { count: number; totalMedia: number; totalScore: number } } = {};
    let albumsConSubgenero = 0;

    for (const album of albums) {
      if (album.subgeneros) {
        albumsConSubgenero++;
        const subs = album.subgeneros.split(',').map(g => g.trim()).filter(g => g);
        for (const g of subs) {
          if (!subgeneroStatsMap[g]) subgeneroStatsMap[g] = { count: 0, totalMedia: 0, totalScore: 0 };
          subgeneroStatsMap[g].count++;
          subgeneroStatsMap[g].totalMedia += album.media;
          subgeneroStatsMap[g].totalScore += album.thirdEyeScore;
        }
      }
    }

    const subgenerosUnicos = Object.keys(subgeneroStatsMap);
    if (subgenerosUnicos.length > 0) {
      const subSorted = subgenerosUnicos
        .map(g => ({
          nombre: g,
          count: subgeneroStatsMap[g].count,
          media: rd(subgeneroStatsMap[g].totalMedia / subgeneroStatsMap[g].count),
          score: rd(subgeneroStatsMap[g].totalScore / subgeneroStatsMap[g].count),
        }))
        .sort((a, b) => b.count - a.count);

      const subMasFrecuente = subSorted[0];
      const subPorMedia = [...subSorted].sort((a, b) => b.media - a.media);
      const subMejor = subPorMedia[0];
      const subPeor = subPorMedia[subPorMedia.length - 1];

      resumenData.push(
        ['', ''],
        ['SUBGÉNEROS', ''],
        ['Álbumes con subgénero', `${albumsConSubgenero} de ${albums.length}`],
        ['Subgéneros únicos', subgenerosUnicos.length],
        ['Más frecuente', `${subMasFrecuente.nombre} (${subMasFrecuente.count})`],
        ['Mejor media', `${subMejor.nombre} (${subMejor.media})`],
        ['Peor media', `${subPeor.nombre} (${subPeor.media})`],
        ['', ''],
        ['DESGLOSE POR SUBGÉNERO', '']
      );
      for (const g of subPorMedia) {
        resumenData.push([g.nombre, `${g.count} · media ${g.media} · score ${g.score}`]);
      }
    }

    // --- Parent genre stats ---
    const parentNames = Object.keys(generosPadreMap);
    if (parentNames.length > 0) {
      // Álbumes mapeados a algún género padre (cada álbum cuenta una vez)
      const albumsMapeados = albums.filter(a => getGenerosPadre(a).length > 0).length;

      // Subgéneros distintos que cuelgan de cada género padre (según el mapeo)
      const subsPorPadre: { [parent: string]: Set<string> } = {};
      for (const sub of Object.keys(subgenreToParent)) {
        const parent = subgenreToParent[sub];
        if (!subsPorPadre[parent]) subsPorPadre[parent] = new Set<string>();
        subsPorPadre[parent].add(sub);
      }

      const padreSorted = parentNames
        .map(p => {
          const lista = generosPadreMap[p];
          const mejor = lista.reduce((best, a) => a.thirdEyeScore > best.thirdEyeScore ? a : best);
          return {
            nombre: p,
            count: lista.length,
            media: rd(lista.reduce((s, a) => s + a.media, 0) / lista.length),
            score: rd(lista.reduce((s, a) => s + a.thirdEyeScore, 0) / lista.length),
            nSubs: subsPorPadre[p] ? subsPorPadre[p].size : 0,
            mejor,
          };
        })
        .sort((a, b) => b.count - a.count);

      const padreMasFrecuente = padreSorted[0];
      const padrePorMedia = [...padreSorted].sort((a, b) => b.media - a.media);
      const padrePorScore = [...padreSorted].sort((a, b) => b.score - a.score);
      const padreMasVariado = [...padreSorted].sort((a, b) => b.nSubs - a.nSubs)[0];

      resumenData.push(
        ['', ''],
        ['GÉNEROS PADRE', ''],
        ['Álbumes con género padre', `${albumsMapeados} de ${albums.length}`],
        ['Géneros padre únicos', parentNames.length],
        ['Más frecuente', `${padreMasFrecuente.nombre} (${padreMasFrecuente.count})`],
        ['Mejor media', `${padrePorMedia[0].nombre} (${padrePorMedia[0].media})`],
        ['Peor media', `${padrePorMedia[padrePorMedia.length - 1].nombre} (${padrePorMedia[padrePorMedia.length - 1].media})`],
        ['Mejor 3rd EYE SCORE', `${padrePorScore[0].nombre} (${padrePorScore[0].score})`],
        ['Peor 3rd EYE SCORE', `${padrePorScore[padrePorScore.length - 1].nombre} (${padrePorScore[padrePorScore.length - 1].score})`],
        ['Más variado (nº subgéneros)', `${padreMasVariado.nombre} (${padreMasVariado.nSubs})`],
        ['', ''],
        ['DESGLOSE POR GÉNERO PADRE', '']
      );
      for (const p of padrePorScore) {
        resumenData.push([p.nombre, `${p.count} álb · ${p.nSubs} subg · media ${p.media} · score ${p.score}`]);
      }

      resumenData.push(
        ['', ''],
        ['MEJOR ÁLBUM POR GÉNERO PADRE', '']
      );
      for (const p of padrePorScore) {
        resumenData.push([p.nombre, `${p.mejor.artista} - ${p.mejor.album} (${p.mejor.thirdEyeScore})`]);
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
        ['Álbumes con año', `${albumsConAno.length} de ${albums.length}`],
        ['Año más antiguo', anoMin],
        ['Año más reciente', anoMax],
        ['Décadas con más álbumes', `${decadaTop} (${decadaMap[decadaTop].count})`],
        ['Década mejor puntuada', `${decadaTopScore} (${rd(decadaMap[decadaTopScore].totalScore / decadaMap[decadaTopScore].count)})`],
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
        ['Duración media', minutesToDisplay(Math.round(durMedia))],
        ['Más largo', `${albumMasLargo.artista} - ${albumMasLargo.album} (${minutesToDisplay(durMax)})`],
        ['Más corto', `${albumMasCorto.artista} - ${albumMasCorto.album} (${minutesToDisplay(durMin)})`],
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
    const corrCanciones = pearsonCorr(albums.map(a => a.totalCanciones), scores);
    const corrInterludios = pearsonCorr(albums.map(a => a.interludios), scores);
    const corrDesv = pearsonCorr(albums.map(a => a.desviacionTipica), scores);
    const corrNotas10 = pearsonCorr(albums.map(a => a.notasMayoresIgual10), scores);

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
      ['Total canciones', `r=${corrCanciones}   · ${corrLabel(corrCanciones)}`],
      ['Interludios (absoluto)', `r=${corrInterludios} · ${corrLabel(corrInterludios)}`],
      ['% Interludios', `r=${corrPctInterludio} · ${corrLabel(corrPctInterludio)}`],
      ['Desviación típica', `r=${corrDesv}         · ${corrLabel(corrDesv)}`],
      ['Notas ≥10', `r=${corrNotas10}      · ${corrLabel(corrNotas10)}`],
    );

    if (corrAno !== null) resumenData.push(['Año de publicación', `r=${corrAno} · ${corrLabel(corrAno)}`]);
    if (corrDur !== null) resumenData.push(['Duración (minutos)', `r=${corrDur} · ${corrLabel(corrDur)}`]);

    const corrPairs: [string, number][] = [
      ['Total canciones', corrCanciones],
      ['Interludios', corrInterludios],
      ['% Interludios', corrPctInterludio],
      ['Desv. típica', corrDesv],
      ['Notas ≥10', corrNotas10],
    ];
    if (corrAno !== null) corrPairs.push(['Año', corrAno]);
    if (corrDur !== null) corrPairs.push(['Duración', corrDur]);

    const strongestCorr = corrPairs.reduce((best, cur) => Math.abs(cur[1]) > Math.abs(best[1]) ? cur : best);
    const weakestCorr = corrPairs.reduce((best, cur) => Math.abs(cur[1]) < Math.abs(best[1]) ? cur : best);

    resumenData.push(
      ['', ''],
      ['Mayor correlación', `${strongestCorr[0]} (r=${strongestCorr[1]})`],
      ['Menor correlación', `${weakestCorr[0]}   (r=${weakestCorr[1]})`],
    );

    resumenData.push(['', ''], ['MEDIA DE NOTAS POR NÚMERO DE CANCIÓN', '']);
    for (let i = 0; i < 100; i++) {
      // contar cuantos albumes tienen i canciones
      let albumesConCancionN = albums.filter(a => a.notaCancionesFull.length > i).length;
      if (albumesConCancionN < 15) break;
      const notasCancionN = albums.map(a => a.notaCancionesFull[i]).filter(n => n !== undefined);
      if (notasCancionN.length > 0) {
        const mediaN = notasCancionN.reduce((sum, n) => sum + n, 0) / notasCancionN.length;
        //vamos a calcular tambien la media de las desviaciones de cada indice de cancion (media album menos nota cancion n)
        const mediaDesvN = albums.map(a => {
          if (a.notaCancionesFull.length > i && a.notaCancionesFull[i] !== null) {
            const desv = a.notaCancionesFull[i] - a.media;
            return desv;
          }
          return null;
        }).filter(d => d !== null).reduce((sum, d) => sum + d!, 0) / notasCancionN.length;

        resumenData.push([`Álbumes con al menos ${i + 1} cancion${i === 0 ? '' : 'es'}: ${albumesConCancionN}`, `${mediaN.toFixed(2)}, [${mediaDesvN > 0 ? '+' : ''}${mediaDesvN.toFixed(2)}]`]);
      }
    }

    //vamos a hacer exactamente lo mismo pero porcentualmente en saltos de 10 en 10, asi una quinta cancion de un album de 6 canciones contaria como 83% y entraria en el grupo de "álbumes con al menos 80% de canciones"
    resumenData.push(['', ''], ['MEDIA DE NOTAS POR PORCENTAJE DE CANCIÓN', '']);
    for (let p = 10; p <= 100; p += 10) {
      let albumesConCancionP = albums.filter(a => a.notaCancionesFull.length >= Math.ceil(a.totalCanciones * p / 100)).length;
      if (albumesConCancionP < 1) break;
      const notasCancionP = albums.map(a => {
        if (a.notaCancionesFull.length >= Math.ceil(a.totalCanciones * p / 100)) {
          const index = Math.ceil(a.totalCanciones * p / 100) - 1;
          return a.notaCancionesFull[index];
        }
        return null;
      }).filter(n => n !== null) as number[];
      if (notasCancionP.length > 0) {
        const mediaP = notasCancionP.reduce((sum, n) => sum + n, 0) / notasCancionP.length;
        const mediaDesvP = albums.map(a => {
          if (a.notaCancionesFull.length >= Math.ceil(a.totalCanciones * p / 100) && a.notaCancionesFull[Math.ceil(a.totalCanciones * p / 100) - 1] !== null) {
            const index = Math.ceil(a.totalCanciones * p / 100) - 1;
            const desv = a.notaCancionesFull[index] - a.media;
            return desv;
          }
          return null;
        }).filter(d => d !== null).reduce((sum, d) => sum + d!, 0) / notasCancionP.length;

        resumenData.push([`Media canciones al ${p}% del álbum`, `${mediaP.toFixed(2)}, [${mediaDesvP > 0 ? '+' : ''}${mediaDesvP.toFixed(2)}]`]);
      }
    }


    // Write stats table
    const tituloRange = resumenSheet.getRangeByIndexes(resumenStartRow, resumenCol, 1, 2);
    tituloRange.merge();
    resumenSheet.getCell(resumenStartRow, resumenCol).setValue('RESUMEN GLOBAL');
    tituloRange.getFormat().getFont().setBold(true);
    tituloRange.getFormat().getFont().setSize(13);
    tituloRange.getFormat().getFill().setColor('#1A252F');
    tituloRange.getFormat().getFont().setColor('#FFFFFF');
    tituloRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    resumenSheet.getRangeByIndexes(resumenStartRow + 1, resumenCol, resumenData.length, 2)
      .setValues(resumenData);

    for (let i = 0; i < resumenData.length; i++) {
      const row = resumenStartRow + 1 + i;
      const label = resumenData[i][0]?.toString() || '';
      const cellRange = resumenSheet.getRangeByIndexes(row, resumenCol, 1, 2);

      if (label === '' && resumenData[i][1] === '') continue;

      if (label === label.toUpperCase() && label.length > 1 && resumenData[i][1] === '') {
        cellRange.getFormat().getFont().setBold(true);
        cellRange.getFormat().getFill().setColor('#34495E');
        cellRange.getFormat().getFont().setColor('#FFFFFF');
        cellRange.getFormat().getFont().setSize(10);
      } else {
        cellRange.getFormat().getFill().setColor(i % 2 === 0 ? '#F5F5F5' : '#FFFFFF');
        resumenSheet.getCell(row, resumenCol + 1).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
      }
    }

    resumenSheet.getRangeByIndexes(resumenStartRow, resumenCol, resumenData.length + 1, 2)
      .getFormat().autofitColumns();

    console.log('Tabla resumen global generada.');
  }

  // =================== GRÁFICOS (hoja "Graficos") ===================
  // Esta hoja se regenera COMPLETA en cada ejecución: se borran todos los
  // gráficos y el contenido previo, se reescriben las tablas de datos auxiliares
  // (en una zona alejada a la derecha) y se vuelven a crear todos los charts.
  {
    let graficosSheet = workbook.getWorksheet('Graficos');
    if (!graficosSheet) graficosSheet = workbook.addWorksheet('Graficos');

    // --- Limpieza total de la hoja (gráficos + datos) ---
    for (const ch of graficosSheet.getCharts()) ch.delete();
    graficosSheet.getUsedRange()?.clear(ExcelScript.ClearApplyTo.all);

    const rd2 = (v: number) => Math.round(v * 100) / 100;

    // --- Zona de datos auxiliares: muy a la derecha para no estorbar a los charts ---
    const DATA_COL = 45; // columna ~AT
    let dataRow = 0;

    // --- Rejilla de colocación de gráficos (coordenadas absolutas en px) ---
    const CHART_W = 480;
    const CHART_H = 300;
    const CHARTS_PER_ROW = 2;
    const GAP_X = 25;
    const GAP_Y = 25;
    const MARGIN = 10;
    let chartIndex = 0;

    // Escribe una tabla [cabecera + filas] en la zona de datos y devuelve su rango.
    // labelsAsText: fuerza la primera columna (etiquetas) a formato texto ANTES de escribir,
    // para que Excel la trate como eje de categorías y no como una segunda serie numérica.
    // Debe ser false en los scatter, donde la primera columna es la X numérica.
    function writeBlock(
      header: string[],
      rows: (string | number)[][],
      labelsAsText: boolean = true,
    ): ExcelScript.Range | null {
      if (rows.length === 0) return null;
      const width = header.length;
      const startR = dataRow;
      if (labelsAsText) {
        graficosSheet.getRangeByIndexes(startR, DATA_COL, rows.length + 1, 1).setNumberFormatLocal('@');
      }
      graficosSheet.getRangeByIndexes(startR, DATA_COL, 1, width).setValues([header]);
      graficosSheet.getRangeByIndexes(startR + 1, DATA_COL, rows.length, width).setValues(rows);
      const range = graficosSheet.getRangeByIndexes(startR, DATA_COL, rows.length + 1, width);
      dataRow = startR + rows.length + 1 + 2; // hueco entre bloques
      return range;
    }

    // Crea un gráfico a partir de un rango y lo coloca en la rejilla.
    function makeChart(
      type: ExcelScript.ChartType,
      range: ExcelScript.Range | null,
      title: string,
      seriesBy: ExcelScript.ChartSeriesBy = ExcelScript.ChartSeriesBy.columns,
      showLegend: boolean = false,
    ): ExcelScript.Chart | null {
      if (!range) return null;
      const chart = graficosSheet.addChart(type, range, seriesBy);
      chart.getTitle().setText(title);
      chart.getLegend().setVisible(showLegend);

      const colPos = chartIndex % CHARTS_PER_ROW;
      const rowPos = Math.floor(chartIndex / CHARTS_PER_ROW);
      chart.setLeft(MARGIN + colPos * (CHART_W + GAP_X));
      chart.setTop(MARGIN + rowPos * (CHART_H + GAP_Y));
      chart.setWidth(CHART_W);
      chart.setHeight(CHART_H);
      chartIndex++;
      return chart;
    }

    // ---------- 1 & 2. Evolución temporal de reviews ----------
    const reviewed = albums
      .filter(a => a.dateOfReviewTimestamp > 0)
      .sort((a, b) => a.dateOfReviewTimestamp - b.dateOfReviewTimestamp);

    if (reviewed.length > 0) {
      const monthCounts: { [k: string]: number } = {};
      for (const a of reviewed) {
        const d = new Date(a.dateOfReviewTimestamp);
        const key = `${d.getUTCFullYear()}-${(d.getUTCMonth() + 1).toString().padStart(2, '0')}`;
        monthCounts[key] = (monthCounts[key] || 0) + 1;
      }
      const monthKeys = Object.keys(monthCounts).sort();
      const [fy, fm] = monthKeys[0].split('-').map(s => Number(s));
      const [ly, lm] = monthKeys[monthKeys.length - 1].split('-').map(s => Number(s));

      const mesAcum: (string | number)[][] = [];
      const mesMensual: (string | number)[][] = [];
      let y = fy, m = fm, cum = 0;
      while (y < ly || (y === ly && m <= lm)) {
        const key = `${y}-${m.toString().padStart(2, '0')}`;
        const c = monthCounts[key] || 0;
        cum += c;
        mesAcum.push([key, cum]);
        mesMensual.push([key, c]);
        m++; if (m > 12) { m = 1; y++; }
      }

      makeChart(
        ExcelScript.ChartType.line,
        writeBlock(['Mes', 'Reviews acumuladas'], mesAcum),
        `Reviews acumuladas en el tiempo (${reviewed.length} con fecha)`,
      );
      makeChart(
        ExcelScript.ChartType.columnClustered,
        writeBlock(['Mes', 'Reviews del mes'], mesMensual),
        'Reviews por mes (ritmo de reseñas)',
      );
    }

    // ---------- 3. Histograma de notas (todas las canciones) ----------
    if (todasLasNotas.length > 0) {
      const counts: number[] = new Array(22).fill(0);
      for (const v of todasLasNotas) {
        let idx = Math.floor(v / 0.5);
        if (idx < 0) idx = 0; if (idx > 21) idx = 21;
        counts[idx]++;
      }
      const notaBins: (string | number)[][] = [];
      for (let i = 0; i <= 21; i++) notaBins.push([(i * 0.5).toFixed(1), counts[i]]);
      makeChart(
        ExcelScript.ChartType.columnClustered,
        writeBlock(['Nota', 'Frecuencia'], notaBins),
        `Histograma de notas (${todasLasNotas.length} canciones)`,
      );
    }

    // ---------- 4 & 5 & 6. Por década / por año ----------
    const albumsAno = albums.filter(a => a.year > 0);
    if (albumsAno.length > 0) {
      const decadaMap: { [d: string]: { count: number; score: number } } = {};
      for (const a of albumsAno) {
        const dec = `${Math.floor(a.year / 10) * 10}s`;
        if (!decadaMap[dec]) decadaMap[dec] = { count: 0, score: 0 };
        decadaMap[dec].count++;
        decadaMap[dec].score += a.thirdEyeScore;
      }
      const decadas = Object.keys(decadaMap).sort();
      makeChart(
        ExcelScript.ChartType.columnClustered,
        writeBlock(['Década', 'Álbumes'], decadas.map(d => [d, decadaMap[d].count])),
        'Álbumes por década',
      );
      makeChart(
        ExcelScript.ChartType.columnClustered,
        writeBlock(['Década', '3rd EYE SCORE medio'],
          decadas.map(d => [d, rd2(decadaMap[d].score / decadaMap[d].count)])),
        '3rd EYE SCORE medio por década',
      );

      const yearMap: { [y: number]: number } = {};
      for (const a of albumsAno) yearMap[a.year] = (yearMap[a.year] || 0) + 1;
      const years = Object.keys(yearMap).map(s => Number(s)).sort((a, b) => a - b);
      makeChart(
        ExcelScript.ChartType.lineMarkers,
        writeBlock(['Año', 'Álbumes'], years.map(yr => [yr.toString(), yearMap[yr]])),
        'Álbumes por año de publicación',
      );
    }

    // ---------- 7. Top 15 subgéneros por frecuencia ----------
    {
      const subCount: { [s: string]: number } = {};
      for (const a of albums) {
        if (!a.subgeneros) continue;
        for (const s of a.subgeneros.split(',').map(g => g.trim()).filter(g => g)) {
          subCount[s] = (subCount[s] || 0) + 1;
        }
      }
      const topSubs = Object.keys(subCount).sort((a, b) => subCount[b] - subCount[a]).slice(0, 15);
      // En barras horizontales Excel pinta el primero abajo: invertimos para que el mayor quede arriba.
      const subRows = topSubs.map(s => [s, subCount[s]] as (string | number)[]).reverse();
      makeChart(
        ExcelScript.ChartType.barClustered,
        writeBlock(['Subgénero', 'Álbumes'], subRows),
        'Top 15 subgéneros más frecuentes',
      );
    }

    // ---------- 8, 9, 10. Géneros padre ----------
    const parents = Object.keys(generosPadreMap);
    if (parents.length > 0) {
      const parentData = parents.map(p => {
        const lista = generosPadreMap[p];
        return {
          nombre: p,
          count: lista.length,
          media: rd2(lista.reduce((s, a) => s + a.media, 0) / lista.length),
          score: rd2(lista.reduce((s, a) => s + a.thirdEyeScore, 0) / lista.length),
        };
      });

      const porCount = [...parentData].sort((a, b) => b.count - a.count);
      const pie = makeChart(
        ExcelScript.ChartType.pie,
        writeBlock(['Género padre', 'Álbumes'], porCount.map(p => [p.nombre, p.count])),
        'Reparto de álbumes por género padre',
        ExcelScript.ChartSeriesBy.columns,
        true,
      );
      if (pie) {
        const dl = pie.getDataLabels();
        dl.setShowPercentage(true);
        dl.setShowValue(false);
      }

      const porMedia = [...parentData].sort((a, b) => a.media - b.media); // asc → mayor arriba en barras
      makeChart(
        ExcelScript.ChartType.barClustered,
        writeBlock(['Género padre', 'Media'], porMedia.map(p => [p.nombre, p.media])),
        'Media de notas por género padre',
      );

      const porScore = [...parentData].sort((a, b) => a.score - b.score);
      makeChart(
        ExcelScript.ChartType.barClustered,
        writeBlock(['Género padre', '3rd EYE SCORE'], porScore.map(p => [p.nombre, p.score])),
        '3rd EYE SCORE medio por género padre',
      );
    }

    // ---------- 11. Nº de géneros padre por álbum ----------
    {
      const npMap: { [n: number]: number } = {};
      let totalParents = 0;
      for (const a of albums) {
        const n = getGenerosPadre(a).length;
        npMap[n] = (npMap[n] || 0) + 1;
        totalParents += n;
      }
      const mediaPadres = rd2(totalParents / albums.length);
      const ns = Object.keys(npMap).map(s => Number(s)).sort((a, b) => a - b);
      makeChart(
        ExcelScript.ChartType.columnClustered,
        writeBlock(['Géneros padre', 'Álbumes'], ns.map(n => [`${n}`, npMap[n]])),
        `Nº de géneros padre por álbum (media ${mediaPadres})`,
      );
    }

    // ---------- 12. Media de nota según posición de la canción ----------
    {
      const posRows: (string | number)[][] = [];
      for (let i = 0; i < 100; i++) {
        const cnt = albums.filter(a => a.notaCancionesFull.length > i).length;
        if (cnt < 15) break;
        const notasN = albums.map(a => a.notaCancionesFull[i]).filter(n => n !== undefined);
        const mediaN = notasN.reduce((s, n) => s + n, 0) / notasN.length;
        posRows.push([`${i + 1}`, rd2(mediaN)]);
      }
      makeChart(
        ExcelScript.ChartType.lineMarkers,
        writeBlock(['Nº de canción', 'Media'], posRows),
        'Media de nota según la posición de la canción',
      );
    }

    // ---------- 13. Media de nota según % de avance del álbum ----------
    {
      const pctRows: (string | number)[][] = [];
      for (let p = 10; p <= 100; p += 10) {
        const notasP = albums.map(a => {
          const need = Math.ceil(a.totalCanciones * p / 100);
          return a.notaCancionesFull.length >= need ? a.notaCancionesFull[need - 1] : null;
        }).filter(n => n !== null) as number[];
        if (notasP.length > 0) {
          pctRows.push([`${p}%`, rd2(notasP.reduce((s, n) => s + n, 0) / notasP.length)]);
        }
      }
      makeChart(
        ExcelScript.ChartType.lineMarkers,
        writeBlock(['% del álbum', 'Media'], pctRows),
        'Media de nota según el % de avance del álbum',
      );
    }

    // ---------- 14. Histograma de 3rd EYE SCORE (álbumes) ----------
    function histograma(valores: number[], header: string, titulo: string): void {
      if (valores.length === 0) return;
      const sMin = Math.floor(Math.min(...valores) / 0.5) * 0.5;
      const sMax = Math.ceil(Math.max(...valores) / 0.5) * 0.5;
      const nb = Math.round((sMax - sMin) / 0.5) + 1;
      const sCounts: number[] = new Array(nb).fill(0);
      for (const v of valores) {
        let idx = Math.round((Math.floor(v / 0.5) * 0.5 - sMin) / 0.5);
        if (idx < 0) idx = 0; if (idx >= nb) idx = nb - 1;
        sCounts[idx]++;
      }
      const bins: (string | number)[][] = [];
      for (let i = 0; i < nb; i++) bins.push([(sMin + i * 0.5).toFixed(1), sCounts[i]]);
      makeChart(ExcelScript.ChartType.columnClustered, writeBlock([header, 'Álbumes'], bins), titulo);
    }
    histograma(albums.map(a => a.thirdEyeScore), '3rd EYE SCORE', 'Distribución de 3rd EYE SCORE (álbumes)');
    histograma(albums.map(a => a.media), 'Media', 'Distribución de la media por álbum');

    // ---------- 16-19. Dispersiones contra el 3rd EYE SCORE ----------
    function scatter(puntos: [number, number][], headerX: string, titulo: string): void {
      const rows = puntos.map(p => [p[0], p[1]] as (string | number)[]);
      makeChart(ExcelScript.ChartType.xyscatter, writeBlock([headerX, '3rd EYE SCORE'], rows, false), titulo);
    }
    scatter(albumsAno.map(a => [a.year, a.thirdEyeScore]), 'Año', 'Año vs 3rd EYE SCORE');
    const albumsDur = albums.filter(a => a.durationMinutes > 0);
    scatter(albumsDur.map(a => [a.durationMinutes, a.thirdEyeScore]), 'Duración (min)', 'Duración vs 3rd EYE SCORE');
    scatter(albums.map(a => [a.totalCanciones, a.thirdEyeScore]), 'Nº canciones', 'Nº de canciones vs 3rd EYE SCORE');
    scatter(albums.map(a => [a.desviacionTipica, a.thirdEyeScore]), 'Desv. típica', 'Consistencia (desv. típica) vs 3rd EYE SCORE');

    // ---------- 20. Top 15 álbumes por 3rd EYE SCORE ----------
    {
      const top15 = albums.slice().sort((a, b) => b.thirdEyeScore - a.thirdEyeScore).slice(0, 15);
      const rows = top15.map(a => [`${a.artista} - ${a.album}`, a.thirdEyeScore] as (string | number)[]).reverse();
      makeChart(ExcelScript.ChartType.barClustered, writeBlock(['Álbum', '3rd EYE SCORE'], rows),
        'Top 15 álbumes por 3rd EYE SCORE');
    }

    // ---------- 21 & 22. Artistas con varios álbumes ----------
    {
      const repe = Object.keys(artistasMap).filter(a => artistasMap[a].length > 1);
      if (repe.length > 0) {
        const porMedia = repe
          .map(a => ({
            art: a,
            n: artistasMap[a].length,
            media: rd2(artistasMap[a].reduce((s, x) => s + x.media, 0) / artistasMap[a].length),
          }))
          .sort((x, y) => y.media - x.media)
          .slice(0, 15);
        makeChart(
          ExcelScript.ChartType.barClustered,
          writeBlock(['Artista', 'Media'],
            porMedia.map(r => [`${r.art} (${r.n})`, r.media] as (string | number)[]).reverse()),
          'Top 15 artistas (≥2 álbumes) por media',
        );

        const porCount = repe
          .map(a => ({ art: a, n: artistasMap[a].length }))
          .sort((x, y) => y.n - x.n)
          .slice(0, 15);
        makeChart(
          ExcelScript.ChartType.barClustered,
          writeBlock(['Artista', 'Álbumes'],
            porCount.map(r => [r.art, r.n] as (string | number)[]).reverse()),
          'Top 15 artistas por nº de álbumes',
        );
      }
    }

    console.log(`Hoja "Graficos" regenerada con ${chartIndex} gráficos.`);
  }
}
