/**
 * Converts a zero-based column index to A1 notation letter(s).
 * e.g., 0 -> 'A', 8 -> 'I', 9 -> 'J', 25 -> 'Z', 26 -> 'AA'
 */
export function colIndexToLetter(colIndex) {
  let temp, letter = '';
  let c = colIndex + 1;
  while (c > 0) {
    temp = (c - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    c = (c - temp - 1) / 26;
  }
  return letter;
}

/**
 * Creates a normalized key map of header names to column metadata.
 */
export function createHeaderMap(headerRow) {
  const map = {};
  if (!headerRow || !Array.isArray(headerRow)) return map;

  headerRow.forEach((header, index) => {
    if (header) {
      const key = header.toString().toLowerCase().replace(/[^a-z0-9]/g, '');
      map[key] = {
        index: index,
        letter: colIndexToLetter(index),
        originalName: header
      };
    }
  });
  return map;
}

/**
 * Finds column info matching several candidate header alias names.
 */
export function getColInfo(headerMap, ...possibleNames) {
  for (const name of possibleNames) {
    const cleanName = name.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (headerMap[cleanName]) {
      return headerMap[cleanName];
    }
  }
  return null;
}