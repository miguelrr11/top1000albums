// @ts-nocheck
/**
 * Script TypeScript para Excel - Automatización de listado de álbumes
 * 
 * 
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
 *       Excel proporciona automáticamente las definiciones de ExcelScript.
 */

function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();

    if (!usedRange) {
        console.log("No hay datos en la hoja");
        return;
    }



    // Estructura para almacenar información de cada álbum
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
        fila: number; // Para debugging
        genero: string;
        num105: number;
        thirdEyeScore: number;
    }

    const artistasMap: {
        [artista: string]: AlbumInfo[]
    } = {};

    const albums: AlbumInfo[] = [];
    const todasLasNotas: number[] = []; // Todas las notas individuales para estadísticas globales
    const values = usedRange.getValues();
    const numRows = values.length;
    const numCols = values[0].length;

    // Buscar todas las celdas que empiezan con '*'
    for (let row = 0; row < numRows; row++) {
        for (let col = 0; col < numCols; col++) {
            const cellValue = values[row][col];

            // Verificar si la celda comienza con '*'
            if (typeof cellValue === 'string' && cellValue.startsWith('*')) {
                // Remover el '*' del título
                const tituloCompleto = cellValue.substring(1).trim();

                // Separar artista y álbum por el primer '-'
                const primerGuionIndex = tituloCompleto.indexOf('-');
                let artista = tituloCompleto;
                let album = '';

                if (primerGuionIndex !== -1) {
                    artista = tituloCompleto.substring(0, primerGuionIndex).trim();
                    album = tituloCompleto.substring(primerGuionIndex + 1).trim();
                }

                // Extraer notas de las canciones (filas siguientes en la misma columna)
                const notas: number[] = [];
                let totalCanciones = 0;
                let interludios = 0;
                let currentRow = row + 1;

                // Recorrer filas hacia abajo hasta encontrar una celda vacía o otro álbum
                while (currentRow < numRows) {
                    const cancionNombre = values[currentRow][col];
                    const notaValue = values[currentRow][col + 1];

                    // Si encontramos una celda vacía o otro álbum, terminamos
                    if (!cancionNombre ||
                        (typeof cancionNombre === 'string' && cancionNombre.startsWith('*'))) {
                        break;
                    }

                    totalCanciones++;

                    // Verificar si tiene nota
                    if (typeof notaValue === 'number' && notaValue >= 0 && notaValue <= 10.5) {
                        notas.push(notaValue);
                        todasLasNotas.push(notaValue);
                    } else {
                        // Es un interludio (canción sin nota)
                        interludios++;
                    }

                    currentRow++;
                }

                // Calcular estadísticas solo si hay notas
                if (notas.length > 0) {
                    // Media
                    const media = notas.reduce((sum, nota) => sum + nota, 0) / notas.length;

                    // Mediana
                    const notasOrdenadas = [...notas].sort((a, b) => a - b);
                    let mediana: number;
                    const mitad = Math.floor(notasOrdenadas.length / 2);
                    if (notasOrdenadas.length % 2 === 0) {
                        // Si hay cantidad par, promedio de los dos del medio
                        mediana = (notasOrdenadas[mitad - 1] + notasOrdenadas[mitad]) / 2;
                    } else {
                        // Si hay cantidad impar, el del medio
                        mediana = notasOrdenadas[mitad];
                    }

                    // Desviación típica
                    const varianza = notas.reduce((sum, nota) => sum + Math.pow(nota - media, 2), 0) / notas.length;
                    const desviacionTipica = Math.sqrt(varianza);

                    // Contar notas >= 10
                    const notasMayoresIgual10 = notas.filter(nota => nota >= 10).length;
                    const num105 = notas.filter(nota => nota === 10.5).length;

                    const thirdEyeScoreRaw =
                        media -
                        (desviacionTipica * 0.3) +
                        ((mediana - media) * 0.2) +
                        (num105 * 0.15) +
                        ((notasMayoresIgual10 / totalCanciones) * 0.2);

                    const thirdEyeScore = Math.round(thirdEyeScoreRaw * 100) / 100;

                    albums.push({
                        titulo: tituloCompleto,
                        artista,
                        album,
                        media: Math.round(media * 100) / 100, // Redondear a 2 decimales
                        mediana: Math.round(mediana * 100) / 100, // Redondear a 2 decimales
                        desviacionTipica: Math.round(desviacionTipica * 100) / 100,
                        notasMayoresIgual10,
                        totalCanciones,
                        interludios,
                        fila: row + 1, // +1 para número de fila legible
                        genero: '',
                        num105,
                        thirdEyeScore,
                    });
                }
            }
        }
    }

    // Posición fija para los headers
    const columnaRanking = 17; // Columna R (A=0, B=1, ..., R=17) para los números
    const columnaInicio = 18; // Columna S (A=0, B=1, ..., S=18) para los datos
    const startRow = 0; // Fila 1 (índice 0)

    // Headers base (sin asteriscos)
    const headersBase = [
        '#',
        'Artista',
        'Álbum',
        'Media',
        '3rd EYE SCORE',
        'Mediana',
        'Desv. Típica',
        'Notas ≥10',
        'Total Canciones',
        'Interludios',
        'Género'
    ];

    // ========== LEER HEADERS ANTES DE LIMPIAR ==========
    // Leer los headers actuales de ambas tablas para detectar asteriscos ANTES de borrar
    const headerRowRange = sheet.getRangeByIndexes(startRow + 1, columnaRanking, 1, headersBase.length);
    const currentHeaders = headerRowRange.getValues()[0];

    const artistasStartRowPreClear = startRow + albums.length + 4;
    const artistasHeaderRowRangePreClear = sheet.getRangeByIndexes(artistasStartRowPreClear, columnaRanking, 1, headersBase.length);
    const currentArtistasHeaders = artistasHeaderRowRangePreClear.getValues()[0];

    // ========== LEER GÉNEROS ANTES DE LIMPIAR ==========
    const generoMap: {
        [key: string]: string
    } = {};

    let generoColIdx = -1;
    let artistaColIdx = -1;
    let albumColIdx = -1;

    for (let i = 0; i < currentHeaders.length; i++) {
        const h = currentHeaders[i]?.toString().replace(/\*/g, '').trim();
        if (h === 'Género') generoColIdx = i;
        if (h === 'Artista') artistaColIdx = i;
        if (h === 'Álbum') albumColIdx = i;
    }

    if (generoColIdx !== -1 && artistaColIdx !== -1 && albumColIdx !== -1) {
        const maxRows = 500;
        const dataReadRange = sheet.getRangeByIndexes(startRow + 1, columnaRanking, maxRows, currentHeaders.length);
        const dataReadValues = dataReadRange.getValues();

        for (let i = 0; i < dataReadValues.length; i++) {
            const art = dataReadValues[i][artistaColIdx]?.toString().trim();
            const alb = dataReadValues[i][albumColIdx]?.toString().trim();
            const gen = dataReadValues[i][generoColIdx]?.toString().trim();

            if (!art && !alb) break;
            if (art && alb && gen) {
                generoMap[`${art}|${alb}`] = gen;
            }
        }
    }

    // Asignar géneros leídos a los álbumes
    for (const alb of albums) {
        alb.genero = generoMap[`${alb.artista}|${alb.album}`] || alb.genero;
    }

    // ========== LIMPIAR ÁREA DE TABLAS ==========
    // Limpiar tablas principales + tabla resumen a la derecha (con 1 columna de separación + 2 columnas de resumen)
    const cleanRange = sheet.getRangeByIndexes(
        startRow,
        columnaRanking,
        1000,
        headersBase.length + 2 + 2 // tablas principales + 2 separación + resumen (2 columnas)
    );
    cleanRange.clear(ExcelScript.ClearApplyTo.all);

    // ========== DETECTAR CRITERIO DE ORDENACIÓN PARA TABLA PRINCIPAL ==========

    // Mapeo de headers (sin asterisco) a propiedades del álbum
    const headerToProperty: {
        [key: string]: keyof AlbumInfo
    } = {
        'Artista': 'artista',
        'Álbum': 'album',
        'Álbumes': 'album', // Para compatibilidad con tabla de artistas
        'Media': 'media',
        'Mediana': 'mediana',
        'Desv. Típica': 'desviacionTipica',
        'Notas ≥10': 'notasMayoresIgual10',
        'Total Canciones': 'totalCanciones',
        'Interludios': 'interludios',
        'Género': 'genero',
        '3rd EYE SCORE': 'thirdEyeScore',
    };

    // Detectar qué header tiene asterisco en tabla principal
    let sortBy: keyof AlbumInfo = 'thirdEyeScore'; // Default
    let headerWithAsterisk: string | null = null;
    let asteriskCount = 0;

    for (let i = 0; i < currentHeaders.length; i++) {
        const headerValue = currentHeaders[i]?.toString().trim();
        if (headerValue && headerValue.includes('*')) {
            asteriskCount++;
            // Remover el asterisco para buscar la propiedad correspondiente
            const cleanHeader = headerValue.replace(/\*/g, '').trim();
            if (headerToProperty[cleanHeader]) {
                headerWithAsterisk = cleanHeader;
                sortBy = headerToProperty[cleanHeader];
            }
        }
    }

    // Si no hay asterisco o hay más de uno, usar Third Eye Score por defecto
    if (asteriskCount !== 1) {
        sortBy = 'thirdEyeScore';
        headerWithAsterisk = '3rd EYE SCORE';
    }

    console.log(`Ordenando tabla principal por: ${sortBy} (header con asterisco: "${headerWithAsterisk}")`);

    // Crear headers con asterisco en el header correspondiente
    const headers = headersBase.map(header => {
        if (header === headerWithAsterisk) {
            return `${header} *`;
        }
        return header;
    });

    // Ordenar álbumes según el criterio seleccionado
    albums.sort((a, b) => {
        const valueA = a[sortBy];
        const valueB = b[sortBy];

        // Para números: orden descendente (mayor a menor)
        if (typeof valueA === 'number' && typeof valueB === 'number') {
            return valueB - valueA;
        }

        // Para strings: orden alfabético ascendente (A-Z), desempate por media desc
        if (typeof valueA === 'string' && typeof valueB === 'string') {
            const cmp = valueA.localeCompare(valueB);
            if (cmp !== 0) return cmp;
            return b.media - a.media;
        }

        return 0;
    });

    // Título de la tabla principal
    const mainTituloRange = sheet.getRangeByIndexes(startRow, columnaRanking, 1, headers.length);
    mainTituloRange.merge();
    const mainTituloCell = sheet.getCell(startRow, columnaRanking);
    mainTituloCell.setValue('ÁLBUMES');
    mainTituloRange.getFormat().getFont().setBold(true);
    mainTituloRange.getFormat().getFont().setSize(13);
    mainTituloRange.getFormat().getFill().setColor('#1A252F');
    mainTituloRange.getFormat().getFont().setColor('#FFFFFF');
    mainTituloRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    // OPTIMIZACIÓN: Escribir todos los encabezados de una vez
    const headerRange = sheet.getRangeByIndexes(startRow + 1, columnaRanking, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.getFormat().getFont().setBold(true);
    headerRange.getFormat().getFill().setColor('#2C3E50');
    headerRange.getFormat().getFont().setColor('#FFFFFF');
    headerRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    // Función para convertir nota a color (degradado personalizado)
    function getColorForScore(score: number): string {
        // Función de interpolación lineal (lerp)
        const lerp = (start: number, end: number, factor: number): number => {
            return Math.round(start + (end - start) * factor);
        };

        let r: number, g: number, b: number;

        // Colores de referencia
        const colorRojo = [255, 73, 77]; // 0-5
        const colorAmarillo = [255, 245, 67]; // 5-7
        const colorAzul = [0, 176, 240]; // 7-10+

        if (score <= 5) {
            // 0 a 5: Rojo sólido [255, 73, 77]
            [r, g, b] = colorRojo;
        } else if (score <= 7) {
            // 5 a 7: Lerp de rojo a amarillo
            const factor = (score - 5) / 2; // Normalizar 5-7 a 0-1
            r = lerp(colorRojo[0], colorAmarillo[0], factor);
            g = lerp(colorRojo[1], colorAmarillo[1], factor);
            b = lerp(colorRojo[2], colorAmarillo[2], factor);
        } else if (score <= 10) {
            // 7 a 10: Lerp de amarillo a azul
            const factor = (score - 7) / 3; // Normalizar 7-10 a 0-1
            r = lerp(colorAmarillo[0], colorAzul[0], factor);
            g = lerp(colorAmarillo[1], colorAzul[1], factor);
            b = lerp(colorAmarillo[2], colorAzul[2], factor);
        } else {
            // 10+: Azul sólido [0, 176, 240]
            [r, g, b] = colorAzul;
        }

        // Convertir RGB a hexadecimal
        const toHex = (n: number) => {
            const hex = n.toString(16);
            return hex.length === 1 ? '0' + hex : hex;
        };

        return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
    }

    // Función para obtener color de ranking (medallas para top 3)
    function getRankingColor(rank: number): string {
        if (rank === 1) return '#FFD700'; // Oro
        if (rank === 2) return '#C0C0C0'; // Plata
        if (rank === 3) return '#CD7F32'; // Bronce
        return '#E8E8E8'; // Gris claro para el resto
    }

    // OPTIMIZACIÓN: Preparar todos los datos en una matriz
    const dataRows: (string | number)[][] = albums.map((album, index) => [
        index + 1,
        album.artista,
        album.album,
        album.media,
        album.thirdEyeScore,
        album.mediana,
        album.desviacionTipica,
        album.notasMayoresIgual10,
        album.totalCanciones,
        album.interludios,
        album.genero
    ]);

    // OPTIMIZACIÓN: Escribir todos los datos de una vez
    if (dataRows.length > 0) {
        const dataRange = sheet.getRangeByIndexes(
            startRow + 2,
            columnaRanking,
            dataRows.length,
            headers.length
        );
        dataRange.setValues(dataRows);

        // Aplicar formatos en lotes
        // Formato de columna de ranking
        const rankingColumn = sheet.getRangeByIndexes(startRow + 2, columnaRanking, dataRows.length, 1);
        rankingColumn.getFormat().getFont().setBold(true);
        rankingColumn.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

        // Formato de columna de media
        const thirdEyeColumn = sheet.getRangeByIndexes(startRow + 2, columnaInicio + 3, dataRows.length, 1);
        thirdEyeColumn.getFormat().getFont().setBold(true);
        thirdEyeColumn.getFormat().getFont().setColor('#000000');
        thirdEyeColumn.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

        const mediaColumn = sheet.getRangeByIndexes(startRow + 2, columnaInicio + 2, dataRows.length, 1);
        mediaColumn.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

        // Aplicar colores de forma más eficiente
        // Primero aplicar colores de filas alternas a todo el rango
        for (let i = 0; i < albums.length; i += 2) {
            const rowRange = sheet.getRangeByIndexes(
                startRow + 2 + i,
                columnaInicio,
                1,
                headers.length - 1
            );
            rowRange.getFormat().getFill().setColor('#F5F5F5');
        }

        for (let i = 1; i < albums.length; i += 2) {
            const rowRange = sheet.getRangeByIndexes(
                startRow + 2 + i,
                columnaInicio,
                1,
                headers.length - 1
            );
            rowRange.getFormat().getFill().setColor('#FFFFFF');
        }

        // Luego aplicar colores específicos solo donde es necesario
        for (let i = 0; i < albums.length; i++) {
            const row = startRow + 2 + i;
            const rank = i + 1;

            // Color de ranking (solo para top 3, el resto queda con el fondo de fila)
            if (rank <= 3) {
                sheet.getCell(row, columnaRanking).getFormat().getFill().setColor(getRankingColor(rank));
            } else {
                sheet.getCell(row, columnaRanking).getFormat().getFill().setColor('#E8E8E8');
            }

            // Color de media (degradado según puntuación)
            sheet.getCell(row, columnaInicio + 3)
                .getFormat()
                .getFill()
                .setColor(getColorForScore(albums[i].thirdEyeScore));
        }
    }

    // Aplicar formato a la tabla de resultados (incluyendo columna de ranking)
    const resultRange = sheet.getRangeByIndexes(
        startRow,
        columnaRanking, // Empezar desde la columna R (#)
        albums.length + 2, // +1 título + +1 encabezados
        headers.length + 1 // +1 para incluir columna de ranking
    );

    resultRange.getFormat().autofitColumns();

    console.log(`Procesados ${albums.length} álbumes y ordenados por media.`);

    // ========== TABLA DE ARTISTAS REPETIDOS ==========

    // Agrupar álbumes por artista

    for (const album of albums) {
        if (!artistasMap[album.artista]) {
            artistasMap[album.artista] = [];
        }
        artistasMap[album.artista].push(album);
    }

    // Filtrar solo artistas con más de un álbum
    const artistasRepetidos = Object.keys(artistasMap).filter(artista => artistasMap[artista].length > 1);
    interface ArtistaStats {
        artista: string;
        numAlbumes: number;
        media: number;
        thirdEyeScore: number; // NEW
        mediana: number;
        desviacionTipica: number;
        notasMayoresIgual10: number;
        totalCanciones: number;
        interludios: number;
        generos ? : string[];
    }
    let artistasStats: ArtistaStats[] = [];

    if (artistasRepetidos.length > 0) {
        // Calcular estadísticas para cada artista

        artistasStats = artistasRepetidos.map(artista => {
            const albumesArtista = artistasMap[artista];
            const numAlbumes = albumesArtista.length;
            const avgThirdEye = albumesArtista.reduce((sum, a) => sum + a.thirdEyeScore, 0) / numAlbumes;

            return {
                artista: artista,
                numAlbumes: numAlbumes,
                media: Math.round((albumesArtista.reduce((sum, a) => sum + a.media, 0) / numAlbumes) * 100) / 100,
                mediana: Math.round((albumesArtista.reduce((sum, a) => sum + a.mediana, 0) / numAlbumes) * 100) / 100,
                desviacionTipica: Math.round((albumesArtista.reduce((sum, a) => sum + a.desviacionTipica, 0) / numAlbumes) * 100) / 100,
                // Total de notas ≥10 en todos los álbumes del artista
                notasMayoresIgual10: albumesArtista.reduce((sum, a) => sum + a.notasMayoresIgual10, 0),
                // Total de canciones en todos los álbumes del artista
                totalCanciones: albumesArtista.reduce((sum, a) => sum + a.totalCanciones, 0),
                // Total de interludios en todos los álbumes del artista
                interludios: albumesArtista.reduce((sum, a) => sum + a.interludios, 0),
                thirdEyeScore: Math.round(avgThirdEye * 100) / 100, // NEW
            };
        });

        // ========== DETECTAR CRITERIO DE ORDENACIÓN PARA TABLA DE ARTISTAS ==========

        // Posición de la nueva tabla: una fila en blanco después de la tabla principal (título+cabecera+datos+blanco)
        const artistasStartRow = startRow + albums.length + 3; // +1 título principal, +1 cabecera, +1 fila en blanco

        // Detectar qué header tiene asterisco en tabla de artistas (leídos antes del clear)
        let artistasSortBy: keyof ArtistaStats = 'media'; // Default
        let artistasHeaderWithAsterisk: string | null = null;
        let artistasAsteriskCount = 0;

        for (let i = 0; i < currentArtistasHeaders.length; i++) {
            const headerValue = currentArtistasHeaders[i]?.toString().trim();
            if (headerValue && headerValue.includes('*')) {
                artistasAsteriskCount++;
                // Remover el asterisco para buscar la propiedad correspondiente
                const cleanHeader = headerValue.replace(/\*/g, '').trim();

                // Mapeo específico para artistas
                if (cleanHeader === 'Álbumes' || cleanHeader === 'Álbum') {
                    artistasHeaderWithAsterisk = 'Álbumes';
                    artistasSortBy = 'numAlbumes';
                } else if (headerToProperty[cleanHeader]) {
                    artistasHeaderWithAsterisk = cleanHeader;
                    artistasSortBy = headerToProperty[cleanHeader] as keyof ArtistaStats;
                }
            }
        }

        // Si no hay asterisco o hay más de uno, usar Media por defecto
        if (artistasAsteriskCount !== 1) {
            artistasSortBy = 'media';
            artistasHeaderWithAsterisk = 'Media';
        }

        console.log(`Ordenando tabla artistas por: ${artistasSortBy} (header con asterisco: "${artistasHeaderWithAsterisk}")`);

        // Ordenar artistas según su propio criterio
        artistasStats.sort((a, b) => {
            const valueA: string = a[artistasSortBy];
            const valueB: string = b[artistasSortBy];

            if (typeof valueA === 'number' && typeof valueB === 'number') {
                return valueB - valueA;
            }

            if (typeof valueA === 'string' && typeof valueB === 'string') {
                const cmp = valueA.localeCompare(valueB);
                if (cmp !== 0) return cmp;
                return b.media - a.media;
            }

            return 0;
        });

        // Crear headers personalizados para artistas (cambiar "Álbum" por "Álbumes" y aplicar asterisco)
        const artistasHeadersBase = headersBase
        const artistasHeaders = artistasHeadersBase.map(header => {
            // Cambiar "Álbum" por "Álbumes"
            let headerText = header;

            if (headerText === 'Álbum') {
                headerText = 'Álbumes';
            }

            //cambiar genero por generos
            if (headerText === 'Género') {
                headerText = 'Géneros';
            }

            // Aplicar asterisco si corresponde
            if (headerText === artistasHeaderWithAsterisk) {
                return `${headerText} *`;
            }
            return headerText;
        });

        // Título de la tabla de artistas
        const artistasTituloRange = sheet.getRangeByIndexes(artistasStartRow, columnaRanking, 1, artistasHeaders.length);
        artistasTituloRange.merge();
        const artistasTituloCell = sheet.getCell(artistasStartRow, columnaRanking);
        artistasTituloCell.setValue('RESUMEN ARTISTAS REPETIDOS');
        artistasTituloRange.getFormat().getFont().setBold(true);
        artistasTituloRange.getFormat().getFont().setSize(13);
        artistasTituloRange.getFormat().getFill().setColor('#1A252F');
        artistasTituloRange.getFormat().getFont().setColor('#FFFFFF');
        artistasTituloRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

        // Escribir headers
        const artistasHeaderRange = sheet.getRangeByIndexes(artistasStartRow + 1, columnaRanking, 1, artistasHeaders.length);
        artistasHeaderRange.setValues([artistasHeaders]);
        artistasHeaderRange.getFormat().getFont().setBold(true);
        artistasHeaderRange.getFormat().getFill().setColor('#2C3E50');
        artistasHeaderRange.getFormat().getFont().setColor('#FFFFFF');
        artistasHeaderRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

        // Rellenar artistas.generos con los géneros únicos de sus álbumes
        for (const artistaStat of artistasStats) {
            const albumesArtista = artistasMap[artistaStat.artista];
            //separate a.genero into different values by commas and trim spaces
            const allGenres: string[] = albumesArtista.flatMap(a => a.genero.split(',').map(g => g.trim())); //declare type of allGenres as string[]
            const generosUnicos = new Set(allGenres);
            artistaStat.generos = Array.from(generosUnicos);
        }

        // Preparar datos
        const dataRows: (string | number)[][] = albums.map((album, index) => [
            index + 1,
            album.artista,
            album.album,
            album.media,
            album.thirdEyeScore,
            album.mediana,
            album.desviacionTipica,
            album.notasMayoresIgual10,
            album.totalCanciones,
            album.interludios,
            album.genero
        ]);

        // Preparar datos
        const artistasDataRows: (string | number)[][] = artistasStats.map((artistaStats, index) => [
            index + 1,
            artistaStats.artista,
            artistaStats.numAlbumes,
            artistaStats.media,
            artistaStats.thirdEyeScore, // NEW
            artistaStats.mediana,
            artistaStats.desviacionTipica,
            artistaStats.notasMayoresIgual10,
            artistaStats.totalCanciones,
            artistaStats.interludios,
            artistaStats.generos ? artistaStats.generos.join(', ') : ''
        ]);

        // Escribir datos
        const artistasDataRange = sheet.getRangeByIndexes(
            artistasStartRow + 2,
            columnaRanking,
            artistasDataRows.length,
            artistasHeaders.length
        );
        artistasDataRange.setValues(artistasDataRows);

        // Aplicar formatos
        const artistasRankingColumn = sheet.getRangeByIndexes(artistasStartRow + 2, columnaRanking, artistasDataRows.length, 1);
        artistasRankingColumn.getFormat().getFont().setBold(true);
        artistasRankingColumn.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

        // Centrar columna "Álbumes" (segunda columna de datos: columnaInicio + 1)
        const artistasAlbumesColumn = sheet.getRangeByIndexes(artistasStartRow + 2, columnaInicio + 1, artistasDataRows.length, 1);
        artistasAlbumesColumn.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

        const artistasThirdEyeColumn = sheet.getRangeByIndexes(artistasStartRow + 2, columnaInicio + 3, artistasDataRows.length, 1);
        artistasThirdEyeColumn.getFormat().getFont().setBold(true);
        artistasThirdEyeColumn.getFormat().getFont().setColor('#000000');
        artistasThirdEyeColumn.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

        const artistasMediaColumn = sheet.getRangeByIndexes(artistasStartRow + 2, columnaInicio + 2, artistasDataRows.length, 1);
        artistasMediaColumn.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

        // Aplicar colores de filas alternas
        for (let i = 0; i < artistasStats.length; i += 2) {
            const rowRange = sheet.getRangeByIndexes(
                artistasStartRow + 2 + i,
                columnaInicio,
                1,
                artistasHeaders.length - 1
            );
            rowRange.getFormat().getFill().setColor('#F5F5F5');
        }

        for (let i = 1; i < artistasStats.length; i += 2) {
            const rowRange = sheet.getRangeByIndexes(
                artistasStartRow + 2 + i,
                columnaInicio,
                1,
                artistasHeaders.length - 1
            );
            rowRange.getFormat().getFill().setColor('#FFFFFF');
        }

        // Aplicar colores específicos
        for (let i = 0; i < artistasStats.length; i++) {
            const row = artistasStartRow + 2 + i;
            const rank = i + 1;

            if (rank <= 3) {
                sheet.getCell(row, columnaRanking).getFormat().getFill().setColor(getRankingColor(rank));
            } else {
                sheet.getCell(row, columnaRanking).getFormat().getFill().setColor('#E8E8E8');
            }

            sheet.getCell(row, columnaInicio + 3)
                .getFormat()
                .getFill()
                .setColor(getColorForScore(artistasStats[i].thirdEyeScore));
        }

        // Autofit
        const artistasResultRange = sheet.getRangeByIndexes(
            artistasStartRow,
            columnaRanking,
            artistasStats.length + 2,
            artistasHeaders.length + 1
        );
        artistasResultRange.getFormat().autofitColumns();

        console.log(`Procesados ${artistasStats.length} artistas con múltiples álbumes.`);
    }

    // tabla top 20 albumes sin repetir artistas (sort by thirdEyeScore siempre)
    const top20SinRepetir = albums
        .filter((album, index, self) =>
            index === self.findIndex(a => a.artista === album.artista)
        )
        .sort((a, b) => b.thirdEyeScore - a.thirdEyeScore)
        .slice(0, 20);

    const headersTop20 = [
        '#',
        'Artista',
        'Álbum',
        'Media',
        '3rd EYE SCORE',
    ]

    const top20StartRow: number = startRow + albums.length + 3 + (artistasRepetidos.length > 0 ? artistasStats.length + 3 : 0);

    // Título de la tabla top 20 sin repetir artistas
    const top20TituloRange = sheet.getRangeByIndexes(top20StartRow, columnaRanking, 1, headersTop20.length);
    top20TituloRange.merge();
    const top20TituloCell = sheet.getCell(top20StartRow, columnaRanking);
    top20TituloCell.setValue('TOP 20 SIN REPETIR ARTISTAS');
    top20TituloRange.getFormat().getFont().setBold(true);
    top20TituloRange.getFormat().getFont().setSize(13);
    top20TituloRange.getFormat().getFill().setColor('#1A252F');
    top20TituloRange.getFormat().getFont().setColor('#FFFFFF');
    top20TituloRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    // Escribir headers para top 20 sin repetir artistas
    const top20HeaderRange = sheet.getRangeByIndexes(top20StartRow + 1, columnaRanking, 1, headersTop20.length);
    top20HeaderRange.setValues([headersTop20]);
    top20HeaderRange.getFormat().getFont().setBold(true);
    top20HeaderRange.getFormat().getFill().setColor('#2C3E50');
    top20HeaderRange.getFormat().getFont().setColor('#FFFFFF');
    top20HeaderRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    // Escribir datos para top 20 sin repetir artistas
    const top20DataRows: (string | number)[][] = top20SinRepetir.map((album, index) => [
        index + 1,
        album.artista,
        album.album,
        album.media,
        album.thirdEyeScore,
    ]);
    if (top20DataRows.length > 0) {
        const top20DataRange = sheet.getRangeByIndexes(
            top20StartRow + 2,
            columnaRanking,
            top20DataRows.length,
            headersTop20.length
        );
        top20DataRange.setValues(top20DataRows);
        // Aplicar formatos
        const top20RankingColumn = sheet.getRangeByIndexes(top20StartRow + 2, columnaRanking, top20DataRows.length, 1);
        top20RankingColumn.getFormat().getFont().setBold(true);
        top20RankingColumn.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        const top20ThirdEyeColumn = sheet.getRangeByIndexes(top20StartRow + 2, columnaInicio + 3, top20DataRows.length, 1);
        top20ThirdEyeColumn.getFormat().getFont().setBold(true);
        top20ThirdEyeColumn.getFormat().getFont().setColor('#000000');
        top20ThirdEyeColumn.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        const top20MediaColumn = sheet.getRangeByIndexes(top20StartRow + 2, columnaInicio + 2, top20DataRows.length, 1);
        top20MediaColumn.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        // Aplicar colores de filas alternas
        for (let i = 0; i < top20SinRepetir.length; i += 2) {
            const rowRange = sheet.getRangeByIndexes(
                top20StartRow + 2 + i,
                columnaInicio,
                1,
                headersTop20.length - 1
            );
            rowRange.getFormat().getFill().setColor('#F5F5F5');
        }
        for (let i = 1; i < top20SinRepetir.length; i += 2) {
            const rowRange = sheet.getRangeByIndexes(
                top20StartRow + 2 + i,
                columnaInicio,
                1,
                headersTop20.length - 1
            );
            rowRange.getFormat().getFill().setColor('#FFFFFF');
        }
        // Aplicar colores específicos
        for (let i = 0; i < top20SinRepetir.length; i++) {
            const row = top20StartRow + 2 + i;
            const rank = i + 1;
            if (rank <= 3) {
                sheet.getCell(row, columnaRanking).getFormat().getFill().setColor(getRankingColor(rank));
            } else {
                sheet.getCell(row, columnaRanking).getFormat().getFill().setColor('#E8E8E8');
            }
            sheet.getCell(row, columnaInicio + 3)
                .getFormat()
                .getFill()
                .setColor(getColorForScore(top20SinRepetir[i].thirdEyeScore));
        }
        // Autofit
        const top20ResultRange = sheet.getRangeByIndexes(
            top20StartRow,
            columnaRanking,
            top20SinRepetir.length + 2,
            headersTop20.length + 1
        );
        top20ResultRange.getFormat().autofitColumns();
    }




    // ========== TABLA RESUMEN: ESTADÍSTICAS GLOBALES ==========

    if (todasLasNotas.length > 0 && albums.length > 0) {
        // Posición: a la derecha de la tabla principal, arriba del todo
        const resumenCol = columnaRanking + headersBase.length + 2; // 2 columnas de separación
        const resumenStartRow = startRow;

        // --- Estadísticas globales de notas ---
        const notasOrd = [...todasLasNotas].sort((a, b) => a - b);
        const n = notasOrd.length;
        const totalCancionesGlobal = albums.reduce((s, a) => s + a.totalCanciones, 0);
        const totalInterludiosGlobal = albums.reduce((s, a) => s + a.interludios, 0);
        const artistasUnicos = new Set(albums.map(a => a.artista)).size;

        // Media global
        const mediaGlobal = todasLasNotas.reduce((s, v) => s + v, 0) / n;

        // Mediana global
        const medianaGlobal = n % 2 === 0 ?
            (notasOrd[n / 2 - 1] + notasOrd[n / 2]) / 2 :
            notasOrd[Math.floor(n / 2)];

        // Desviación típica global
        const varianzaGlobal = todasLasNotas.reduce((s, v) => s + Math.pow(v - mediaGlobal, 2), 0) / n;
        const desvGlobal = Math.sqrt(varianzaGlobal);

        // Coeficiente de variación (dispersión relativa)
        const coefVariacion = (desvGlobal / mediaGlobal) * 100;

        // Cuartiles (Q1, Q3) e IQR
        const q1Index = Math.floor(n * 0.25);
        const q3Index = Math.floor(n * 0.75);
        const q1 = notasOrd[q1Index];
        const q3 = notasOrd[q3Index];
        const iqr = q3 - q1;

        // Asimetría (skewness) de Fisher - indica si las notas se concentran arriba o abajo
        // > 0: cola derecha (mayoría de notas bajas), < 0: cola izquierda (mayoría de notas altas)
        const skewness = todasLasNotas.reduce((s, v) => s + Math.pow((v - mediaGlobal) / desvGlobal, 3), 0) / n;

        // Curtosis excess (Fisher) - indica si hay muchas notas extremas o están agrupadas
        // > 0: más extremos de lo normal, < 0: más agrupadas que lo normal
        const kurtosis = (todasLasNotas.reduce((s, v) => s + Math.pow((v - mediaGlobal) / desvGlobal, 4), 0) / n) - 3;

        // Nota máxima y mínima
        const notaMax = notasOrd[n - 1];
        const notaMin = notasOrd[0];
        const rango = notaMax - notaMin;

        // Porcentajes
        const notasGe10 = todasLasNotas.filter(v => v >= 10).length;
        const pctGe10 = (notasGe10 / n) * 100;
        const pctInterludios = (totalInterludiosGlobal / totalCancionesGlobal) * 100;

        // Canciones por álbum
        const cancionesPorAlbum = totalCancionesGlobal / albums.length;

        // Álbum más y menos consistente
        const albumMasConsistente = [...albums].sort((a, b) => a.desviacionTipica - b.desviacionTipica)[0];
        const albumMenosConsistente = [...albums].sort((a, b) => b.desviacionTipica - a.desviacionTipica)[0];

        // Álbum mejor y peor
        const albumMejor = [...albums].sort((a, b) => b.media - a.media)[0];
        const albumPeor = [...albums].sort((a, b) => a.media - b.media)[0];

        // Moda (nota más frecuente)
        const frecuencias: {
            [nota: string]: number
        } = {};
        for (const nota of todasLasNotas) {
            const key = nota.toString();
            frecuencias[key] = (frecuencias[key] || 0) + 1;
        }
        let moda = todasLasNotas[0];
        let maxFreq = 0;
        for (const [nota, freq] of Object.entries(frecuencias)) {
            if (freq > maxFreq) {
                maxFreq = freq;
                moda = parseFloat(nota);
            }
        }

        // Distribución por rangos
        const rango0a5 = todasLasNotas.filter(v => v < 5).length;
        const rango5a7 = todasLasNotas.filter(v => v >= 5 && v < 7).length;
        const rango7a8 = todasLasNotas.filter(v => v >= 7 && v < 8).length;
        const rango8a9 = todasLasNotas.filter(v => v >= 8 && v < 9).length;
        const rango9a10 = todasLasNotas.filter(v => v >= 9 && v < 10).length;
        const rango10plus = todasLasNotas.filter(v => v >= 10).length;

        const rd = (v: number) => Math.round(v * 100) / 100;

        // Construir filas: [Estadística, Valor]
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
            ['[0, 5)', `${rango0a5} (${rd(rango0a5 / n * 100)}%)`],
            ['[5, 7)', `${rango5a7} (${rd(rango5a7 / n * 100)}%)`],
            ['[7, 8)', `${rango7a8} (${rd(rango7a8 / n * 100)}%)`],
            ['[8, 9)', `${rango8a9} (${rd(rango8a9 / n * 100)}%)`],
            ['[9, 10)', `${rango9a10} (${rd(rango9a10 / n * 100)}%)`],
            ['[10, 10.5]', `${rango10plus} (${rd(pctGe10)}%)`],
            ['', ''],
            ['DESTACADOS', ''],
            ['Mejor álbum', `${albumMejor.artista} - ${albumMejor.album} (${albumMejor.media})`],
            ['Peor álbum', `${albumPeor.artista} - ${albumPeor.album} (${albumPeor.media})`],
            ['Más consistente (menor desv.)', `${albumMasConsistente.artista} - ${albumMasConsistente.album} (σ=${albumMasConsistente.desviacionTipica})`],
            ['Más irregular (mayor desv.)', `${albumMenosConsistente.artista} - ${albumMenosConsistente.album} (σ=${albumMenosConsistente.desviacionTipica})`],
        ];

        // --- Estadísticas de géneros ---
        const generoStatsMap: {
            [g: string]: {
                count: number,
                totalMedia: number
            }
        } = {};
        let albumsConGenero = 0;

        for (const album of albums) {
            if (album.genero) {
                albumsConGenero++;
                const generos = album.genero.split(',').map(g => g.trim()).filter(g => g);
                for (const g of generos) {
                    if (!generoStatsMap[g]) generoStatsMap[g] = {
                        count: 0,
                        totalMedia: 0
                    };
                    generoStatsMap[g].count++;
                    generoStatsMap[g].totalMedia += album.media;
                }
            }
        }

        const generosUnicos = Object.keys(generoStatsMap);
        if (generosUnicos.length > 0) {
            const generosSorted = generosUnicos
                .map(g => ({
                    nombre: g,
                    count: generoStatsMap[g].count,
                    media: rd(generoStatsMap[g].totalMedia / generoStatsMap[g].count)
                }))
                .sort((a, b) => b.count - a.count);

            const generoMasFrecuente = generosSorted[0];
            const generosPorMedia = [...generosSorted].sort((a, b) => b.media - a.media);
            const generoMejor = generosPorMedia[0];
            const generoPeor = generosPorMedia[generosPorMedia.length - 1];

            resumenData.push(
                ['', ''],
                ['GÉNEROS', ''],
                ['Álbumes con género', `${albumsConGenero} de ${albums.length}`],
                ['Géneros únicos', generosUnicos.length],
                ['Más frecuente', `${generoMasFrecuente.nombre} (${generoMasFrecuente.count})`],
                ['Mejor media', `${generoMejor.nombre} (${generoMejor.media})`],
                ['Peor media', `${generoPeor.nombre} (${generoPeor.media})`],
                ['', ''],
                ['DESGLOSE POR GÉNERO', '']
            );

            for (const g of generosPorMedia) {
                resumenData.push([g.nombre, `${g.count} (${g.media})`]);
            }
        }

        // Escribir título de la tabla
        const tituloRange = sheet.getRangeByIndexes(resumenStartRow, resumenCol, 1, 2);
        tituloRange.merge();
        const tituloCell = sheet.getCell(resumenStartRow, resumenCol);
        tituloCell.setValue('RESUMEN GLOBAL');
        tituloRange.getFormat().getFont().setBold(true);
        tituloRange.getFormat().getFont().setSize(13);
        tituloRange.getFormat().getFill().setColor('#1A252F');
        tituloRange.getFormat().getFont().setColor('#FFFFFF');
        tituloRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

        // Escribir datos
        const resumenRange = sheet.getRangeByIndexes(
            resumenStartRow + 1,
            resumenCol,
            resumenData.length,
            2
        );
        resumenRange.setValues(resumenData);

        // Aplicar formato por filas
        for (let i = 0; i < resumenData.length; i++) {
            const row = resumenStartRow + 1 + i;
            const label = resumenData[i][0]?.toString() || '';
            const cellRange = sheet.getRangeByIndexes(row, resumenCol, 1, 2);

            if (label === '' && resumenData[i][1] === '') {
                // Fila separadora vacía
                continue;
            }

            if (label === label.toUpperCase() && label.length > 1 && resumenData[i][1] === '') {
                // Encabezado de sección (todo mayúsculas, sin valor)
                cellRange.getFormat().getFont().setBold(true);
                cellRange.getFormat().getFill().setColor('#34495E');
                cellRange.getFormat().getFont().setColor('#FFFFFF');
                cellRange.getFormat().getFont().setSize(10);
            } else {
                // Fila de dato normal - colores alternos
                cellRange.getFormat().getFill().setColor(i % 2 === 0 ? '#F5F5F5' : '#FFFFFF');
                // Alinear valor a la derecha
                sheet.getCell(row, resumenCol + 1).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
            }
        }

        // Autofit
        const resumenFullRange = sheet.getRangeByIndexes(
            resumenStartRow,
            resumenCol,
            resumenData.length + 1,
            2
        );
        resumenFullRange.getFormat().autofitColumns();

        console.log('Tabla resumen global generada.');
    }
}