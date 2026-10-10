/**
 * XML and OpenXML utility functions
 */

export function parseXml(xmlString: string): Document {
  let cleaned = xmlString.replace(/^\uFEFF/, "").trimStart();
  // Strip duplicate or misplaced XML declarations from the beginning.
  // In DOMParser, XML declarations are optional, but any leading whitespace or misplaced
  // declaration causes: "error on line 2 at column 6: XML declaration allowed only at the start of the document".
  while (/^<\?xml[^>]*\?>/i.test(cleaned)) {
    cleaned = cleaned.replace(/^<\?xml[^>]*\?>\s*/i, "").trimStart();
  }

  const parser = new DOMParser();
  const doc = parser.parseFromString(cleaned, "application/xml");
  const parserErrors = doc.getElementsByTagName("parsererror");
  if (parserErrors.length > 0) {
    throw new Error("XML parse error: " + parserErrors[0].textContent);
  }
  return doc;
}

export function serializeXml(doc: Document): string {
  const root = doc.documentElement;
  const serializer = new XMLSerializer();
  let xmlStr = serializer.serializeToString(root);
  // Strip any accidental leading XML declaration or whitespace
  xmlStr = xmlStr.replace(/^<\?xml[^>]*\?>\s*/i, "").trimStart();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xmlStr}`;
}

export function getDocNamespace(doc: Document): string {
  return doc.documentElement?.namespaceURI || "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
}

export function findRow(sheetDoc: Document, rowNum: number): Element | null {
  const rows = sheetDoc.getElementsByTagName("row");
  for (let i = 0; i < rows.length; i++) {
    if (parseInt(rows[i].getAttribute("r") || "0", 10) === rowNum) {
      return rows[i];
    }
  }
  return null;
}

export function findCell(rowEl: Element, cellRef: string): Element | null {
  const cells = rowEl.getElementsByTagName("c");
  const target = cellRef.toUpperCase();
  for (let i = 0; i < cells.length; i++) {
    if ((cells[i].getAttribute("r") || "").toUpperCase() === target) {
      return cells[i];
    }
  }
  return null;
}

export function getOrCreateCell(doc: Document, rowEl: Element, cellRef: string): Element {
  let cell = findCell(rowEl, cellRef);
  if (cell) return cell;

  const ns = getDocNamespace(doc);
  cell = doc.createElementNS(ns, "c");
  cell.setAttribute("r", cellRef);

  // Insert in column sorted order
  const targetColLetters = cellRef.replace(/[0-9]/g, "");
  const targetColNum = colLettersToNumber(targetColLetters);

  const existingCells = Array.from(rowEl.getElementsByTagName("c"));
  let inserted = false;
  for (const ec of existingCells) {
    const cRef = ec.getAttribute("r") || "";
    const colLetters = cRef.replace(/[0-9]/g, "");
    if (colLettersToNumber(colLetters) > targetColNum) {
      rowEl.insertBefore(cell, ec);
      inserted = true;
      break;
    }
  }
  if (!inserted) {
    rowEl.appendChild(cell);
  }
  return cell;
}

export function colLettersToNumber(col: string): number {
  let num = 0;
  const upper = col.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    num = num * 26 + (upper.charCodeAt(i) - 64);
  }
  return num;
}

export function colNumberToLetters(num: number): string {
  let letters = "";
  let temp = num;
  while (temp > 0) {
    const rem = (temp - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    temp = Math.floor((temp - 1) / 26);
  }
  return letters;
}

export function parseCellRef(ref: string): { col: string; row: number; colNum: number } | null {
  const match = ref.match(/^([A-Za-z]+)([0-9]+)$/);
  if (!match) return null;
  return {
    col: match[1].toUpperCase(),
    row: parseInt(match[2], 10),
    colNum: colLettersToNumber(match[1]),
  };
}

export function buildCellRef(col: string, row: number): string {
  return `${col.toUpperCase()}${row}`;
}

export function extractSharedStrings(sharedStringsDoc: Document | null): string[] {
  if (!sharedStringsDoc) return [];
  const siNodes = sharedStringsDoc.getElementsByTagName("si");
  const strings: string[] = [];

  for (let i = 0; i < siNodes.length; i++) {
    const si = siNodes[i];
    // Gather all <t> nodes inside <si>
    const tNodes = si.getElementsByTagName("t");
    let fullText = "";
    for (let j = 0; j < tNodes.length; j++) {
      fullText += tNodes[j].textContent || "";
    }
    strings.push(fullText);
  }
  return strings;
}

export function getCellValue(cell: Element | null, sharedStrings: string[]): string {
  if (!cell) return "";
  const t = cell.getAttribute("t");

  if (t === "s") {
    const v = cell.getElementsByTagName("v")[0]?.textContent;
    if (v !== undefined) {
      const idx = parseInt(v, 10);
      return sharedStrings[idx] ?? "";
    }
    return "";
  }

  // Check <is> (inline string)
  const is = cell.getElementsByTagName("is")[0];
  if (is) {
    const tNodes = is.getElementsByTagName("t");
    let text = "";
    for (let j = 0; j < tNodes.length; j++) {
      text += tNodes[j].textContent || "";
    }
    return text;
  }

  // Number or direct value or cached formula value
  const v = cell.getElementsByTagName("v")[0]?.textContent;
  if (v !== undefined) {
    return v;
  }

  return "";
}

export function setCellInlineString(doc: Document, cell: Element, text: string): void {
  const ns = getDocNamespace(doc);
  cell.setAttribute("t", "inlineStr");
  while (cell.firstChild) {
    cell.removeChild(cell.firstChild);
  }

  const is = doc.createElementNS(ns, "is");
  const t = doc.createElementNS(ns, "t");
  if (text.startsWith(" ") || text.endsWith(" ") || text.includes("  ")) {
    t.setAttribute("xml:space", "preserve");
  }
  t.textContent = text;
  is.appendChild(t);
  cell.appendChild(is);
}

export function setCellNumber(doc: Document, cell: Element, num: number): void {
  const ns = getDocNamespace(doc);
  cell.removeAttribute("t");
  while (cell.firstChild) {
    cell.removeChild(cell.firstChild);
  }

  const v = doc.createElementNS(ns, "v");
  v.textContent = String(num);
  cell.appendChild(v);
}

export function setCellFormula(doc: Document, cell: Element, formula: string, cachedValue?: string | number): void {
  const ns = getDocNamespace(doc);
  while (cell.firstChild) {
    cell.removeChild(cell.firstChild);
  }

  const f = doc.createElementNS(ns, "f");
  f.textContent = formula.startsWith("=") ? formula.substring(1) : formula;
  cell.appendChild(f);

  if (cachedValue !== undefined) {
    const v = doc.createElementNS(ns, "v");
    v.textContent = String(cachedValue);
    cell.appendChild(v);
    if (typeof cachedValue === "string") {
      cell.setAttribute("t", "str");
    } else {
      cell.removeAttribute("t");
    }
  } else {
    cell.removeAttribute("t");
  }
}

export function emptyCell(cell: Element): void {
  cell.removeAttribute("t");
  while (cell.firstChild) {
    cell.removeChild(cell.firstChild);
  }
}

/**
 * Shifts row references in formula strings.
 * If row >= insertionRow, shifts row by +k.
 */
export function shiftFormulaRows(formula: string, insertionRow: number, k: number): string {
  // Matches cell references like A1, $A$1, A$1, $A1, etc.
  // Using [A-Za-z]{1,3} and negative lookahead for '(' so function names with numbers aren't matched
  return formula.replace(/(\$?[A-Za-z]{1,3})(\$?)([0-9]+)(?!\s*\()/g, (match, colPart, dollarRow, rowNumStr) => {
    const row = parseInt(rowNumStr, 10);
    if (row >= insertionRow) {
      return `${colPart}${dollarRow}${row + k}`;
    }
    return match;
  });
}

/**
 * When copying a row (from fromRow to toRow), rewrite relative row references to fromRow to toRow.
 */
export function rewriteRelativeRowInFormula(formula: string, fromRow: number, toRow: number): string {
  return formula.replace(/(\$?[A-Za-z]{1,3})(\$?)([0-9]+)(?!\s*\()/g, (match, colPart, dollarRow, rowNumStr) => {
    const row = parseInt(rowNumStr, 10);
    // If it is an absolute row reference like $12, don't rewrite
    if (dollarRow === "$") {
      return match;
    }
    if (row === fromRow) {
      return `${colPart}${toRow}`;
    }
    return match;
  });
}

/**
 * Shifts defined names (like print area) in workbook.xml
 */
export function shiftDefinedNameRange(formula: string, sheetName: string, insertionRow: number, k: number): string {
  // sheetName can be quoted with ' or &apos; or unquoted
  const escapedSheet = sheetName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const regex = new RegExp(
    `('(?:${escapedSheet})'|&(?:apos);(?:${escapedSheet})&(?:apos);|${escapedSheet})!(\\$?[A-Za-z]+\\$?[0-9]+):(\\$?[A-Za-z]+\\$?[0-9]+)`,
    "gi"
  );

  return formula.replace(regex, (match, sheetPart, startCell, endCell) => {
    const shiftCell = (cellStr: string) => {
      return cellStr.replace(/(\$?[A-Za-z]+)(\$?)([0-9]+)/g, (cMatch, col, dollar, rNum) => {
        const r = parseInt(rNum, 10);
        if (r >= insertionRow) {
          return `${col}${dollar}${r + k}`;
        }
        return cMatch;
      });
    };
    return `${sheetPart}!${shiftCell(startCell)}:${shiftCell(endCell)}`;
  });
}

/**
 * Estimate wrapped lines using greedy word wrap:
 * chars per line ~ colWidth * 0.88
 */
export function calculateWrappedLines(text: string, colWidth: number): number {
  if (!text || text.trim() === "") return 1;
  const effectiveWidth = Math.max(8, colWidth);
  const maxCharsPerLine = Math.max(10, Math.floor(effectiveWidth * 0.88));

  // Split by explicit line breaks first
  const paragraphs = text.split(/\r?\n/);
  let totalLines = 0;

  for (const para of paragraphs) {
    if (para.length === 0) {
      totalLines += 1;
      continue;
    }

    const words = para.split(/\s+/);
    let currentLineLength = 0;
    let linesInPara = 1;

    for (const word of words) {
      if (word.length === 0) continue;
      if (currentLineLength === 0) {
        currentLineLength = word.length;
      } else if (currentLineLength + 1 + word.length <= maxCharsPerLine) {
        currentLineLength += 1 + word.length;
      } else {
        // wrap to next line
        linesInPara += 1;
        currentLineLength = word.length;
      }
    }

    totalLines += linesInPara;
  }

  return Math.max(1, totalLines);
}
