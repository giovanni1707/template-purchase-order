import JSZip from "jszip";
import { HyperFormula } from "hyperformula";
import type { QuotationExtraction, PoDetailsOption } from "../types/index.ts";
import {
  parseXml,
  serializeXml,
  extractSharedStrings,
  getCellValue,
  setCellInlineString,
  setCellNumber,
  setCellFormula,
  emptyCell,
  shiftFormulaRows,
  rewriteRelativeRowInFormula,
  shiftDefinedNameRange,
  calculateWrappedLines,
  colLettersToNumber,
  findRow,
  findCell,
  getOrCreateCell,
  getDocNamespace,
} from "./xmlUtils.ts";

export interface WorkbookInspection {
  isValid: boolean;
  errorMessage?: string;
  orderListSheetName: string;
  poMultivacSheetName: string;
  orderListPath: string;
  poMultivacPath: string;
  firstPoDetailsText: string | null;
}

const FORBIDDEN_SHEET_TAGS = [
  "conditionalFormatting",
  "dataValidations",
  "autoFilter",
  "rowBreaks",
  "colBreaks",
  "hyperlinks",
  "tableParts",
  "drawing",
  "extLst",
  "ignoredErrors",
];

function getFormulaText(sheetDoc: Document, cell: Element, fromRow: number, toRow: number): string {
  const fNode = cell.getElementsByTagName("f")[0];
  if (!fNode) return "";
  let formula = (fNode.textContent || "").trim();
  if (!formula) {
    const si = fNode.getAttribute("si");
    if (si) {
      const allF = sheetDoc.getElementsByTagName("f");
      for (let i = 0; i < allF.length; i++) {
        if (allF[i].getAttribute("si") === si && (allF[i].textContent || "").trim()) {
          const masterFormula = allF[i].textContent!.trim();
          const masterCell = allF[i].parentElement;
          const match = (masterCell?.getAttribute("r") || "").match(/[0-9]+/);
          const masterRow = match ? parseInt(match[0], 10) : fromRow;
          return rewriteRelativeRowInFormula(masterFormula, masterRow, toRow);
        }
      }
    }
  }
  return rewriteRelativeRowInFormula(formula, fromRow, toRow);
}

export async function inspectWorkbook(zip: JSZip): Promise<WorkbookInspection> {
  const workbookXmlFile = zip.file("xl/workbook.xml");
  const workbookRelsFile = zip.file("xl/_rels/workbook.xml.rels");

  if (!workbookXmlFile || !workbookRelsFile) {
    return {
      isValid: false,
      errorMessage: "This is not a valid Excel workbook (missing xl/workbook.xml or rels).",
      orderListSheetName: "",
      poMultivacSheetName: "",
      orderListPath: "",
      poMultivacPath: "",
      firstPoDetailsText: null,
    };
  }

  const workbookXml = await workbookXmlFile.async("string");
  const workbookRelsXml = await workbookRelsFile.async("string");

  const wbDoc = parseXml(workbookXml);
  const relsDoc = parseXml(workbookRelsXml);

  const sheets = wbDoc.getElementsByTagName("sheet");
  let orderListSheetName = "";
  let poMultivacSheetName = "";
  let orderListRId = "";
  let poMultivacRId = "";

  for (let i = 0; i < sheets.length; i++) {
    const s = sheets[i];
    const name = s.getAttribute("name") || "";
    const rId = s.getAttribute("r:id") || s.getAttribute("id") || "";
    if (name.trim().toLowerCase() === "order list") {
      orderListSheetName = name;
      orderListRId = rId;
    } else if (name.trim().toLowerCase() === "po multivac") {
      poMultivacSheetName = name;
      poMultivacRId = rId;
    }
  }

  if (!orderListSheetName || !poMultivacSheetName) {
    return {
      isValid: false,
      errorMessage: "This is not the Template_Purchase_Order workbook (sheets 'Order list' and 'PO Multivac' are required).",
      orderListSheetName: "",
      poMultivacSheetName: "",
      orderListPath: "",
      poMultivacPath: "",
      firstPoDetailsText: null,
    };
  }

  // Resolve target paths via rels
  const relationships = relsDoc.getElementsByTagName("Relationship");
  let orderListTarget = "";
  let poMultivacTarget = "";

  for (let i = 0; i < relationships.length; i++) {
    const rel = relationships[i];
    const id = rel.getAttribute("Id");
    const target = rel.getAttribute("Target") || "";
    if (id === orderListRId) {
      orderListTarget = target;
    }
    if (id === poMultivacRId) {
      poMultivacTarget = target;
    }
  }

  const resolvePath = (t: string) => {
    if (t.startsWith("/")) return t.substring(1);
    if (t.startsWith("xl/")) return t;
    return `xl/${t.replace(/^\.\//, "")}`;
  };

  const orderListPath = resolvePath(orderListTarget);
  const poMultivacPath = resolvePath(poMultivacTarget);

  if (!zip.file(orderListPath) || !zip.file(poMultivacPath)) {
    return {
      isValid: false,
      errorMessage: "Could not find worksheet XML files for 'Order list' and 'PO Multivac' in the workbook package.",
      orderListSheetName,
      poMultivacSheetName,
      orderListPath,
      poMultivacPath,
      firstPoDetailsText: null,
    };
  }

  // Read initial details text
  let firstPoDetailsText: string | null = null;
  const sharedStringsFile = zip.file("xl/sharedStrings.xml");
  let sharedStrings: string[] = [];
  if (sharedStringsFile) {
    const ssXml = await sharedStringsFile.async("string");
    sharedStrings = extractSharedStrings(parseXml(ssXml));
  }

  const poXmlStr = await zip.file(poMultivacPath)!.async("string");
  const poDoc = parseXml(poXmlStr);

  const poRows = poDoc.getElementsByTagName("row");
  let poHeaderRowIndex = -1;

  for (let i = 0; i < poRows.length; i++) {
    const rowEl = poRows[i];
    const rNum = parseInt(rowEl.getAttribute("r") || "0", 10);
    if (rNum >= 1 && rNum <= 60) {
      const colACell = findCell(rowEl, `A${rNum}`);
      if (colACell) {
        const val = getCellValue(colACell, sharedStrings).trim().toLowerCase();
        if (val === "pos" || val === "pos.") {
          poHeaderRowIndex = rNum;
          break;
        }
      }
    }
  }

  if (poHeaderRowIndex !== -1) {
    const firstItemRow = poHeaderRowIndex + 1;
    const itemRowEl = findRow(poDoc, firstItemRow);
    if (itemRowEl) {
      const detailsCell = findCell(itemRowEl, `E${firstItemRow}`);
      if (detailsCell) {
        const val = getCellValue(detailsCell, sharedStrings).trim();
        if (val) {
          firstPoDetailsText = val;
        }
      }
    }
  }

  return {
    isValid: true,
    orderListSheetName,
    poMultivacSheetName,
    orderListPath,
    poMultivacPath,
    firstPoDetailsText,
  };
}

export async function processWorkbook(
  originalZipBytes: Uint8Array,
  quotation: QuotationExtraction,
  exchangeRate: number,
  poDetailsOption: PoDetailsOption,
  poDetailsCustomText: string
): Promise<{
  filledZipBytes: Uint8Array;
  orderListSheetName: string;
  poMultivacSheetName: string;
  orderListRowsCount: number;
  poRowsCount: number;
  totalExworksOrderList: number;
  poTotalCalculated: number;
}> {
  const zip = await JSZip.loadAsync(originalZipBytes);
  const inspection = await inspectWorkbook(zip);

  if (!inspection.isValid) {
    throw new Error(inspection.errorMessage || "Invalid workbook template.");
  }

  const { orderListPath, poMultivacPath, orderListSheetName, poMultivacSheetName } = inspection;

  const sharedStringsFile = zip.file("xl/sharedStrings.xml");
  let sharedStrings: string[] = [];
  if (sharedStringsFile) {
    const ssXml = await sharedStringsFile.async("string");
    sharedStrings = extractSharedStrings(parseXml(ssXml));
  }

  const orderListXml = await zip.file(orderListPath)!.async("string");
  const poMultivacXml = await zip.file(poMultivacPath)!.async("string");

  const orderDoc = parseXml(orderListXml);
  const poDoc = parseXml(poMultivacXml);

  // 1. Safety check: forbidden XML tags
  for (const tag of FORBIDDEN_SHEET_TAGS) {
    if (orderDoc.getElementsByTagName(tag).length > 0) {
      throw new Error(`Safety check failed: Sheet '${orderListSheetName}' contains forbidden element <${tag}>.`);
    }
    if (poDoc.getElementsByTagName(tag).length > 0) {
      throw new Error(`Safety check failed: Sheet '${poMultivacSheetName}' contains forbidden element <${tag}>.`);
    }
  }

  // 2. Identify anchors in Order list
  // Header row = row (1-10) whose col A text is "Reference"
  const orderRows = Array.from(orderDoc.getElementsByTagName("row"));
  let orderHeaderRow = -1;
  for (const r of orderRows) {
    const rNum = parseInt(r.getAttribute("r") || "0", 10);
    if (rNum >= 1 && rNum <= 10) {
      const cellA = findCell(r, `A${rNum}`);
      if (cellA) {
        const val = getCellValue(cellA, sharedStrings).trim().toLowerCase();
        if (val === "reference" || val.startsWith("ref")) {
          orderHeaderRow = rNum;
          break;
        }
      }
    }
  }

  if (orderHeaderRow === -1) {
    throw new Error(`Safety check failed: Could not find header row with column A text 'Reference' in '${orderListSheetName}'.`);
  }

  const orderFirstItemRow = orderHeaderRow + 1;

  // Totals row = first later row whose column E formula starts with SUM(E
  let orderTotalsRow = -1;
  for (const r of orderRows) {
    const rNum = parseInt(r.getAttribute("r") || "0", 10);
    if (rNum > orderFirstItemRow) {
      const cellE = findCell(r, `E${rNum}`);
      const fNode = cellE?.getElementsByTagName("f")[0];
      const formula = (fNode?.textContent || "").trim().toUpperCase().replace(/^=/, "").replace(/\s+/g, "");
      if (formula.startsWith("SUM(E") || formula.startsWith("SUM($E")) {
        orderTotalsRow = rNum;
        break;
      }
    }
  }

  if (orderTotalsRow === -1) {
    throw new Error(`Safety check failed: Could not find totals row with SUM(E formula in '${orderListSheetName}'.`);
  }

  const orderTemplateItemRows = orderTotalsRow - orderFirstItemRow;
  const numPositions = quotation.positions.length;

  if (orderTemplateItemRows > numPositions) {
    throw new Error(
      `Safety check failed: The template has more item rows (${orderTemplateItemRows}) than the quotation has items (${numPositions}).`
    );
  }

  // 3. Identify anchors in PO Multivac
  // Header row = row (1-60) whose col A text is "Pos"
  const poRows = Array.from(poDoc.getElementsByTagName("row"));
  let poHeaderRow = -1;
  for (const r of poRows) {
    const rNum = parseInt(r.getAttribute("r") || "0", 10);
    if (rNum >= 1 && rNum <= 60) {
      const cellA = findCell(r, `A${rNum}`);
      if (cellA) {
        const val = getCellValue(cellA, sharedStrings).trim().toLowerCase();
        if (val === "pos" || val === "pos.") {
          poHeaderRow = rNum;
          break;
        }
      }
    }
  }

  if (poHeaderRow === -1) {
    throw new Error(`Safety check failed: Could not find header row with column A text 'Pos' in '${poMultivacSheetName}'.`);
  }

  const poFirstItemRow = poHeaderRow + 1;

  // Total row = first later row whose col D text is "Total" (or any cell in row matching "total")
  let poTotalRow = -1;
  for (const r of poRows) {
    const rNum = parseInt(r.getAttribute("r") || "0", 10);
    if (rNum > poFirstItemRow) {
      const cellD = findCell(r, `D${rNum}`);
      if (cellD && getCellValue(cellD, sharedStrings).trim().toLowerCase() === "total") {
        poTotalRow = rNum;
        break;
      }
      const cells = r.getElementsByTagName("c");
      for (let c = 0; c < cells.length; c++) {
        const val = getCellValue(cells[c], sharedStrings).trim().toLowerCase();
        if (val === "total" || val.startsWith("total ")) {
          poTotalRow = rNum;
          break;
        }
      }
      if (poTotalRow !== -1) break;
    }
  }

  if (poTotalRow === -1) {
    throw new Error(`Safety check failed: Could not find Total row with text 'Total' in '${poMultivacSheetName}'.`);
  }

  const poTemplateItemRows = poTotalRow - poFirstItemRow;
  if (poTemplateItemRows > numPositions) {
    throw new Error(
      `Safety check failed: Sheet '${poMultivacSheetName}' has more template item rows (${poTemplateItemRows}) than quotation items (${numPositions}).`
    );
  }

  // Find row in PO Multivac whose text starts with "Your Quotation"
  let poQuotationRow = -1;
  let poQuotationCol = "A";
  for (const r of poRows) {
    const rNum = parseInt(r.getAttribute("r") || "0", 10);
    const cells = r.getElementsByTagName("c");
    for (let c = 0; c < cells.length; c++) {
      const val = getCellValue(cells[c], sharedStrings).trim().toLowerCase();
      if (val.includes("your quotation") || val.startsWith("your quote")) {
        poQuotationRow = rNum;
        poQuotationCol = (cells[c].getAttribute("r") || "A").replace(/[0-9]/g, "");
        break;
      }
    }
    if (poQuotationRow !== -1) break;
  }

  if (poQuotationRow === -1) {
    throw new Error(`Safety check failed: Could not find row starting with 'Your Quotation' in '${poMultivacSheetName}'.`);
  }

  // Safety check: verify no moved cell has cross-sheet formula
  const kOrder = numPositions - orderTemplateItemRows;
  const kPo = numPositions - poTemplateItemRows;

  // Check moved cells in Order list (rows >= orderTotalsRow)
  for (const r of orderRows) {
    const rNum = parseInt(r.getAttribute("r") || "0", 10);
    if (rNum >= orderTotalsRow) {
      const cells = r.getElementsByTagName("c");
      for (let i = 0; i < cells.length; i++) {
        const fNode = cells[i].getElementsByTagName("f")[0];
        if (fNode && fNode.textContent && fNode.textContent.includes("!")) {
          throw new Error(
            `Safety check failed: Moved cell ${cells[i].getAttribute("r")} in '${orderListSheetName}' has cross-sheet formula: ${fNode.textContent}.`
          );
        }
      }
    }
  }

  // Check moved cells in PO Multivac (rows >= poTotalRow)
  for (const r of poRows) {
    const rNum = parseInt(r.getAttribute("r") || "0", 10);
    if (rNum >= poTotalRow) {
      const cells = r.getElementsByTagName("c");
      for (let i = 0; i < cells.length; i++) {
        const fNode = cells[i].getElementsByTagName("f")[0];
        if (fNode && fNode.textContent && fNode.textContent.includes("!")) {
          throw new Error(
            `Safety check failed: Moved cell ${cells[i].getAttribute("r")} in '${poMultivacSheetName}' has cross-sheet formula: ${fNode.textContent}.`
          );
        }
      }
    }
  }

  // Read Duty and VAT from template's first item row in Order list
  let templateDuty = 0;
  let templateVat = 0.15;
  const firstOrderRowEl = findRow(orderDoc, orderFirstItemRow);
  if (firstOrderRowEl) {
    const cellO = findCell(firstOrderRowEl, `O${orderFirstItemRow}`);
    const cellP = findCell(firstOrderRowEl, `P${orderFirstItemRow}`);
    if (cellO) {
      const parsed = parseFloat(getCellValue(cellO, sharedStrings));
      if (!isNaN(parsed)) templateDuty = parsed;
    }
    if (cellP) {
      const parsed = parseFloat(getCellValue(cellP, sharedStrings));
      if (!isNaN(parsed)) templateVat = parsed;
    }
  }

  // Read base row height from PO template first item row
  let poBaseHeight = 15;
  const firstPoRowEl = findRow(poDoc, poFirstItemRow);
  if (firstPoRowEl && firstPoRowEl.getAttribute("ht")) {
    const ht = parseFloat(firstPoRowEl.getAttribute("ht") || "15");
    if (!isNaN(ht) && ht > 0) poBaseHeight = ht;
  }

  // -------------------------------------------------------------------------
  // 4. Transform "Order list"
  // -------------------------------------------------------------------------
  const orderSheetData = orderDoc.getElementsByTagName("sheetData")[0];
  if (!orderSheetData) throw new Error("Missing sheetData in Order list");

  if (kOrder > 0) {
    // 4a. Shift all formulas referencing row >= orderTotalsRow down by kOrder
    const allCells = orderDoc.getElementsByTagName("c");
    for (let i = 0; i < allCells.length; i++) {
      const fNode = allCells[i].getElementsByTagName("f")[0];
      if (fNode && fNode.textContent) {
        fNode.textContent = shiftFormulaRows(fNode.textContent, orderTotalsRow, kOrder);
        // Shared-formula ref attribute
        const fRef = fNode.getAttribute("ref");
        if (fRef) {
          fNode.setAttribute("ref", shiftFormulaRows(fRef, orderTotalsRow, kOrder));
        }
      }
    }

    // 4b. Renumber all rows and cells at or below orderTotalsRow by +kOrder (reverse order)
    const orderRowsList = Array.from(orderDoc.getElementsByTagName("row"));
    orderRowsList.sort((a, b) => parseInt(b.getAttribute("r") || "0", 10) - parseInt(a.getAttribute("r") || "0", 10));

    for (const r of orderRowsList) {
      const oldRNum = parseInt(r.getAttribute("r") || "0", 10);
      if (oldRNum >= orderTotalsRow) {
        const newRNum = oldRNum + kOrder;
        r.setAttribute("r", String(newRNum));
        const cells = r.getElementsByTagName("c");
        for (let j = 0; j < cells.length; j++) {
          const cRef = cells[j].getAttribute("r") || "";
          const colLetters = cRef.replace(/[0-9]/g, "");
          cells[j].setAttribute("r", `${colLetters}${newRNum}`);
        }
      }
    }

    // 4c. Deep copy last template row kOrder times
    const lastTemplateRowNum = orderTotalsRow - 1;
    const lastTemplateRowEl = findRow(orderDoc, lastTemplateRowNum);
    if (!lastTemplateRowEl) {
      throw new Error(`Could not find last template row ${lastTemplateRowNum} in Order list`);
    }

    // Reference row where we will insert: new orderTotalsRow = orderTotalsRow + kOrder
    const totalsRowEl = findRow(orderDoc, orderTotalsRow + kOrder);

    for (let insertIdx = 0; insertIdx < kOrder; insertIdx++) {
      const newRowNum = orderTotalsRow + insertIdx;
      const newRow = lastTemplateRowEl.cloneNode(true) as Element;
      newRow.setAttribute("r", String(newRowNum));

      const newCells = newRow.getElementsByTagName("c");
      for (let j = 0; j < newCells.length; j++) {
        const cell = newCells[j];
        const colLetters = (cell.getAttribute("r") || "").replace(/[0-9]/g, "");
        cell.setAttribute("r", `${colLetters}${newRowNum}`);

        const fNode = cell.getElementsByTagName("f")[0];
        if (fNode) {
          const formulaText = getFormulaText(orderDoc, cell, lastTemplateRowNum, newRowNum);
          fNode.removeAttribute("t");
          fNode.removeAttribute("ref");
          fNode.removeAttribute("si");
          fNode.textContent = formulaText;
        }
      }

      orderSheetData.insertBefore(newRow, totalsRowEl);
    }

    // 4d. Shift mergeCells and dimension in Order list
    const mergeCells = orderDoc.getElementsByTagName("mergeCell");
    for (let i = 0; i < mergeCells.length; i++) {
      const refAttr = mergeCells[i].getAttribute("ref") || "";
      mergeCells[i].setAttribute("ref", shiftFormulaRows(refAttr, orderTotalsRow, kOrder));
    }

    const dimensionNode = orderDoc.getElementsByTagName("dimension")[0];
    if (dimensionNode) {
      const dRef = dimensionNode.getAttribute("ref") || "";
      dimensionNode.setAttribute("ref", shiftFormulaRows(dRef, orderTotalsRow, kOrder));
    }

    // 4e. Update totals row SUM ranges in Order list
    const newTotalsRowNum = orderTotalsRow + kOrder;
    const newLastItemRowNum = newTotalsRowNum - 1;
    const updatedTotalsRowEl = findRow(orderDoc, newTotalsRowNum);
    if (updatedTotalsRowEl) {
      const sumCells = updatedTotalsRowEl.getElementsByTagName("c");
      for (let i = 0; i < sumCells.length; i++) {
        const fNode = sumCells[i].getElementsByTagName("f")[0];
        if (fNode && fNode.textContent) {
          fNode.textContent = fNode.textContent.replace(
            /SUM\(\$?([A-Za-z]+)\$?[0-9]+:\$?([A-Za-z]+)\$?[0-9]+\)/gi,
            (_m, col1, col2) => `SUM(${col1}${orderFirstItemRow}:${col2}${newLastItemRowNum})`
          );
        }
      }
    }
  }

  // 4f. In every item row in Order list, write values
  for (let i = 0; i < numPositions; i++) {
    const pos = quotation.positions[i];
    const targetRowNum = orderFirstItemRow + i;
    const rowEl = findRow(orderDoc, targetRowNum);
    if (!rowEl) throw new Error(`Missing item row ${targetRowNum} in Order list`);

    // A: reference
    const cellA = getOrCreateCell(orderDoc, rowEl, `A${targetRowNum}`);
    setCellInlineString(orderDoc, cellA, pos.reference);

    // B: partName
    const cellB = getOrCreateCell(orderDoc, rowEl, `B${targetRowNum}`);
    setCellInlineString(orderDoc, cellB, pos.partName);

    // E: netPrice
    const cellE = getOrCreateCell(orderDoc, rowEl, `E${targetRowNum}`);
    setCellNumber(orderDoc, cellE, pos.netPrice);

    // F: quantity
    const cellF = getOrCreateCell(orderDoc, rowEl, `F${targetRowNum}`);
    setCellNumber(orderDoc, cellF, pos.quantity);

    // I: exchange rate
    const cellI = getOrCreateCell(orderDoc, rowEl, `I${targetRowNum}`);
    setCellNumber(orderDoc, cellI, exchangeRate);

    // O: duty (0)
    const cellO = getOrCreateCell(orderDoc, rowEl, `O${targetRowNum}`);
    setCellNumber(orderDoc, cellO, templateDuty);

    // P: VAT (0.15)
    const cellP = getOrCreateCell(orderDoc, rowEl, `P${targetRowNum}`);
    setCellNumber(orderDoc, cellP, templateVat);
  }

  // -------------------------------------------------------------------------
  // 5. Transform "PO Multivac"
  // -------------------------------------------------------------------------
  const poSheetData = poDoc.getElementsByTagName("sheetData")[0];
  if (!poSheetData) throw new Error("Missing sheetData in PO Multivac");

  if (kPo > 0) {
    // 5a. Shift formulas in PO Multivac referencing row >= poTotalRow down by kPo
    const allCells = poDoc.getElementsByTagName("c");
    for (let i = 0; i < allCells.length; i++) {
      const fNode = allCells[i].getElementsByTagName("f")[0];
      if (fNode && fNode.textContent) {
        fNode.textContent = shiftFormulaRows(fNode.textContent, poTotalRow, kPo);
        const fRef = fNode.getAttribute("ref");
        if (fRef) {
          fNode.setAttribute("ref", shiftFormulaRows(fRef, poTotalRow, kPo));
        }
      }
    }

    // 5b. Renumber rows and cells at or below poTotalRow by +kPo
    const poRowsList = Array.from(poDoc.getElementsByTagName("row"));
    poRowsList.sort((a, b) => parseInt(b.getAttribute("r") || "0", 10) - parseInt(a.getAttribute("r") || "0", 10));

    for (const r of poRowsList) {
      const oldRNum = parseInt(r.getAttribute("r") || "0", 10);
      if (oldRNum >= poTotalRow) {
        const newRNum = oldRNum + kPo;
        r.setAttribute("r", String(newRNum));
        const cells = r.getElementsByTagName("c");
        for (let j = 0; j < cells.length; j++) {
          const cRef = cells[j].getAttribute("r") || "";
          const colLetters = cRef.replace(/[0-9]/g, "");
          cells[j].setAttribute("r", `${colLetters}${newRNum}`);
        }
      }
    }

    // 5c. Deep copy last template row kPo times
    const lastTemplatePoRowNum = poTotalRow - 1;
    const lastTemplatePoRowEl = findRow(poDoc, lastTemplatePoRowNum);
    if (!lastTemplatePoRowEl) {
      throw new Error(`Could not find last template row ${lastTemplatePoRowNum} in PO Multivac`);
    }

    const totalRowEl = findRow(poDoc, poTotalRow + kPo);

    for (let insertIdx = 0; insertIdx < kPo; insertIdx++) {
      const newRowNum = poTotalRow + insertIdx;
      const newRow = lastTemplatePoRowEl.cloneNode(true) as Element;
      newRow.setAttribute("r", String(newRowNum));

      const newCells = newRow.getElementsByTagName("c");
      for (let j = 0; j < newCells.length; j++) {
        const cell = newCells[j];
        const colLetters = (cell.getAttribute("r") || "").replace(/[0-9]/g, "");
        cell.setAttribute("r", `${colLetters}${newRowNum}`);

        const fNode = cell.getElementsByTagName("f")[0];
        if (fNode) {
          const formulaText = getFormulaText(poDoc, cell, lastTemplatePoRowNum, newRowNum);
          fNode.removeAttribute("t");
          fNode.removeAttribute("ref");
          fNode.removeAttribute("si");
          fNode.textContent = formulaText;
        }
      }

      poSheetData.insertBefore(newRow, totalRowEl);
    }

    // 5d. Shift mergeCells and dimension in PO Multivac
    const mergeCells = poDoc.getElementsByTagName("mergeCell");
    for (let i = 0; i < mergeCells.length; i++) {
      const refAttr = mergeCells[i].getAttribute("ref") || "";
      mergeCells[i].setAttribute("ref", shiftFormulaRows(refAttr, poTotalRow, kPo));
    }

    const dimensionNode = poDoc.getElementsByTagName("dimension")[0];
    if (dimensionNode) {
      const dRef = dimensionNode.getAttribute("ref") || "";
      dimensionNode.setAttribute("ref", shiftFormulaRows(dRef, poTotalRow, kPo));
    }

    // 5e. Update total row SUM range in PO Multivac
    const newPoTotalRowNum = poTotalRow + kPo;
    const newLastPoItemRowNum = newPoTotalRowNum - 1;
    const updatedPoTotalRowEl = findRow(poDoc, newPoTotalRowNum);
    if (updatedPoTotalRowEl) {
      const cellG = findCell(updatedPoTotalRowEl, `G${newPoTotalRowNum}`);
      const fNode = cellG?.getElementsByTagName("f")[0];
      if (fNode) {
        fNode.textContent = `SUM(G${poFirstItemRow}:G${newLastPoItemRowNum})`;
      }
    }
  }

  // Read column widths from PO Multivac <cols>
  let colBWidth = 14;
  let colCWidth = 35;
  let colEWidth = 20;

  const colsList = poDoc.getElementsByTagName("col");
  for (let i = 0; i < colsList.length; i++) {
    const colEl = colsList[i];
    const min = parseInt(colEl.getAttribute("min") || "0", 10);
    const max = parseInt(colEl.getAttribute("max") || "0", 10);
    const w = parseFloat(colEl.getAttribute("width") || "0");
    if (!isNaN(w) && w > 0) {
      if (min <= 2 && max >= 2) colBWidth = w;
      if (min <= 3 && max >= 3) colCWidth = w;
      if (min <= 5 && max >= 5) colEWidth = w;
    }
  }

  // 5f. Populate every item row in PO Multivac
  for (let i = 0; i < numPositions; i++) {
    const pos = quotation.positions[i];
    const srcRow = orderFirstItemRow + i;
    const targetRow = poFirstItemRow + i;
    const rowEl = findRow(poDoc, targetRow);
    if (!rowEl) throw new Error(`Missing item row ${targetRow} in PO Multivac`);

    // A: i + 1 (number)
    const cellA = getOrCreateCell(poDoc, rowEl, `A${targetRow}`);
    setCellNumber(poDoc, cellA, i + 1);

    // B: 'Order list'!A<src>
    const cellB = getOrCreateCell(poDoc, rowEl, `B${targetRow}`);
    setCellFormula(poDoc, cellB, `'${orderListSheetName}'!A${srcRow}`, pos.reference);

    // C: 'Order list'!B<src>
    const cellC = getOrCreateCell(poDoc, rowEl, `C${targetRow}`);
    setCellFormula(poDoc, cellC, `'${orderListSheetName}'!B${srcRow}`, pos.partName);

    // D: 'Order list'!F<src>
    const cellD = getOrCreateCell(poDoc, rowEl, `D${targetRow}`);
    setCellFormula(poDoc, cellD, `'${orderListSheetName}'!F${srcRow}`, pos.quantity);

    // F: 'Order list'!E<src>
    const cellF = getOrCreateCell(poDoc, rowEl, `F${targetRow}`);
    setCellFormula(poDoc, cellF, `'${orderListSheetName}'!E${srcRow}`, pos.netPrice);

    // G: F<r>*D<r>
    const cellG = getOrCreateCell(poDoc, rowEl, `G${targetRow}`);
    const lineTotal = pos.netPrice * pos.quantity;
    setCellFormula(poDoc, cellG, `F${targetRow}*D${targetRow}`, lineTotal);

    // E: Details column
    const cellE = getOrCreateCell(poDoc, rowEl, `E${targetRow}`);
    let detailsTextForWrap = "";
    if (i === 0) {
      if (poDetailsOption === "keep") {
        detailsTextForWrap = inspection.firstPoDetailsText || "";
      } else if (poDetailsOption === "replace") {
        setCellInlineString(poDoc, cellE, poDetailsCustomText);
        detailsTextForWrap = poDetailsCustomText;
      } else {
        emptyCell(cellE);
      }
    } else {
      emptyCell(cellE);
    }

    // Compute row height
    const linesB = calculateWrappedLines(pos.reference, colBWidth);
    const linesC = calculateWrappedLines(pos.partName, colCWidth);
    const linesE = calculateWrappedLines(detailsTextForWrap, colEWidth);
    const maxLines = Math.max(1, linesB, linesC, linesE);

    if (maxLines > 1) {
      rowEl.setAttribute("ht", (poBaseHeight * maxLines).toFixed(2));
      rowEl.setAttribute("customHeight", "1");
    }
  }

  // Update Quotation number in PO Multivac
  const actualPoQuotationRow = poQuotationRow >= poTotalRow ? poQuotationRow + kPo : poQuotationRow;
  const quotRowEl = findRow(poDoc, actualPoQuotationRow);
  if (quotRowEl) {
    const cellQuot = getOrCreateCell(poDoc, quotRowEl, `${poQuotationCol}${actualPoQuotationRow}`);
    setCellInlineString(poDoc, cellQuot, `Your Quotation: ${quotation.quotationNumber}`);
  }

  // -------------------------------------------------------------------------
  // 6. HyperFormula Cached Values Calculation
  // -------------------------------------------------------------------------
  const hf = HyperFormula.buildEmpty({
    licenseKey: "gpl-v3",
  });

  hf.addSheet(orderListSheetName);
  hf.addSheet(poMultivacSheetName);
  const orderSheetHfId = hf.getSheetId(orderListSheetName)!;
  const poSheetHfId = hf.getSheetId(poMultivacSheetName)!;

  // Load Order list formulas and values into HyperFormula
  const finalOrderRows = Array.from(orderDoc.getElementsByTagName("row"));
  for (const r of finalOrderRows) {
    const rNum = parseInt(r.getAttribute("r") || "0", 10);
    const cells = r.getElementsByTagName("c");
    for (let c = 0; c < cells.length; c++) {
      const cell = cells[c];
      const cRef = cell.getAttribute("r") || "";
      const colLetters = cRef.replace(/[0-9]/g, "");
      const colNum = colLettersToNumber(colLetters) - 1; // 0-based
      const rowIdx = rNum - 1; // 0-based

      const fNode = cell.getElementsByTagName("f")[0];
      if (fNode && fNode.textContent) {
        const formulaText = fNode.textContent.trim().replace(/^=/, "");
        hf.setCellContents({ col: colNum, row: rowIdx, sheet: orderSheetHfId }, [[`=${formulaText}`]]);
      } else {
        const t = cell.getAttribute("t");
        if (t === "inlineStr") {
          const isNode = cell.getElementsByTagName("is")[0];
          const text = isNode?.textContent || "";
          hf.setCellContents({ col: colNum, row: rowIdx, sheet: orderSheetHfId }, [[text]]);
        } else if (t === "s") {
          const v = cell.getElementsByTagName("v")[0]?.textContent;
          const sText = v !== undefined ? sharedStrings[parseInt(v, 10)] ?? "" : "";
          hf.setCellContents({ col: colNum, row: rowIdx, sheet: orderSheetHfId }, [[sText]]);
        } else {
          const v = cell.getElementsByTagName("v")[0]?.textContent;
          if (v !== undefined) {
            const num = parseFloat(v);
            hf.setCellContents({ col: colNum, row: rowIdx, sheet: orderSheetHfId }, [[isNaN(num) ? v : num]]);
          }
        }
      }
    }
  }

  // Load PO Multivac formulas and values into HyperFormula
  const finalPoRows = Array.from(poDoc.getElementsByTagName("row"));
  for (const r of finalPoRows) {
    const rNum = parseInt(r.getAttribute("r") || "0", 10);
    // Ignore rows 1-5 which might refer to Template sheet (A4, F4)
    if (rNum <= 5) continue;

    const cells = r.getElementsByTagName("c");
    for (let c = 0; c < cells.length; c++) {
      const cell = cells[c];
      const cRef = cell.getAttribute("r") || "";
      const colLetters = cRef.replace(/[^A-Za-z]/g, "");
      const colNum = colLettersToNumber(colLetters) - 1;
      const rowIdx = rNum - 1;

      const fNode = cell.getElementsByTagName("f")[0];
      if (fNode && fNode.textContent) {
        if (fNode.textContent.includes("!") && !fNode.textContent.includes(orderListSheetName)) {
          continue;
        }
        const formulaText = fNode.textContent.trim().replace(/^=/, "");
        hf.setCellContents({ col: colNum, row: rowIdx, sheet: poSheetHfId }, [[`=${formulaText}`]]);
      } else {
        const t = cell.getAttribute("t");
        if (t === "inlineStr") {
          const text = cell.getElementsByTagName("is")[0]?.textContent || "";
          hf.setCellContents({ col: colNum, row: rowIdx, sheet: poSheetHfId }, [[text]]);
        } else if (t === "s") {
          const v = cell.getElementsByTagName("v")[0]?.textContent;
          const sText = v !== undefined ? sharedStrings[parseInt(v, 10)] ?? "" : "";
          hf.setCellContents({ col: colNum, row: rowIdx, sheet: poSheetHfId }, [[sText]]);
        } else {
          const v = cell.getElementsByTagName("v")[0]?.textContent;
          if (v !== undefined) {
            const num = parseFloat(v);
            hf.setCellContents({ col: colNum, row: rowIdx, sheet: poSheetHfId }, [[isNaN(num) ? v : num]]);
          }
        }
      }
    }
  }

  // Write HyperFormula computed values back to <v> of formula cells in both sheets
  const orderDocNs = getDocNamespace(orderDoc);
  for (const r of finalOrderRows) {
    const rNum = parseInt(r.getAttribute("r") || "0", 10);
    const cells = r.getElementsByTagName("c");
    for (let c = 0; c < cells.length; c++) {
      const cell = cells[c];
      const fNode = cell.getElementsByTagName("f")[0];
      if (fNode && fNode.textContent) {
        const cRef = cell.getAttribute("r") || "";
        const colLetters = cRef.replace(/[^A-Za-z]/g, "");
        const colNum = colLettersToNumber(colLetters) - 1;
        const val = hf.getCellValue({ col: colNum, row: rNum - 1, sheet: orderSheetHfId });

        if (typeof val === "number" && !isNaN(val) && isFinite(val)) {
          let vNode = cell.getElementsByTagName("v")[0];
          if (!vNode) {
            vNode = orderDoc.createElementNS(orderDocNs, "v");
            cell.appendChild(vNode);
          }
          vNode.textContent = String(val);
          cell.removeAttribute("t");
        } else if (typeof val === "string" && !val.startsWith("#")) {
          let vNode = cell.getElementsByTagName("v")[0];
          if (!vNode) {
            vNode = orderDoc.createElementNS(orderDocNs, "v");
            cell.appendChild(vNode);
          }
          vNode.textContent = val;
          cell.setAttribute("t", "str");
        }
      }
    }
  }

  const poDocNs = getDocNamespace(poDoc);
  for (const r of finalPoRows) {
    const rNum = parseInt(r.getAttribute("r") || "0", 10);
    if (rNum <= 5) continue; // Keep original cached values for external refs A4/F4
    const cells = r.getElementsByTagName("c");
    for (let c = 0; c < cells.length; c++) {
      const cell = cells[c];
      const fNode = cell.getElementsByTagName("f")[0];
      if (fNode && fNode.textContent) {
        if (fNode.textContent.includes("!") && !fNode.textContent.includes(orderListSheetName)) {
          continue;
        }
        const cRef = cell.getAttribute("r") || "";
        const colLetters = cRef.replace(/[^A-Za-z]/g, "");
        const colNum = colLettersToNumber(colLetters) - 1;
        const val = hf.getCellValue({ col: colNum, row: rNum - 1, sheet: poSheetHfId });

        if (typeof val === "number" && !isNaN(val) && isFinite(val)) {
          let vNode = cell.getElementsByTagName("v")[0];
          if (!vNode) {
            vNode = poDoc.createElementNS(poDocNs, "v");
            cell.appendChild(vNode);
          }
          vNode.textContent = String(val);
          cell.removeAttribute("t");
        } else if (typeof val === "string" && !val.startsWith("#")) {
          let vNode = cell.getElementsByTagName("v")[0];
          if (!vNode) {
            vNode = poDoc.createElementNS(poDocNs, "v");
            cell.appendChild(vNode);
          }
          vNode.textContent = val;
          cell.setAttribute("t", "str");
        }
      }
    }
  }

  // Calculate totals for verification and summary
  const finalOrderTotalsRowNum = orderTotalsRow + kOrder;
  const totalExworksVal = hf.getCellValue({
    col: colLettersToNumber("G") - 1,
    row: finalOrderTotalsRowNum - 1,
    sheet: orderSheetHfId,
  });
  const totalExworksOrderList = typeof totalExworksVal === "number" ? totalExworksVal : 0;

  const finalPoTotalRowNum = poTotalRow + kPo;
  const poTotalVal = hf.getCellValue({
    col: colLettersToNumber("G") - 1,
    row: finalPoTotalRowNum - 1,
    sheet: poSheetHfId,
  });
  const poTotalCalculated = typeof poTotalVal === "number" ? poTotalVal : 0;

  // -------------------------------------------------------------------------
  // 7. Package-Level Edits
  // -------------------------------------------------------------------------
  // 7a. Update workbook.xml defined names and calcPr
  const wbFile = zip.file("xl/workbook.xml");
  if (wbFile) {
    let wbXml = await wbFile.async("string");
    const wbDoc = parseXml(wbXml);

    // Shift defined names
    const definedNames = wbDoc.getElementsByTagName("definedName");
    for (let i = 0; i < definedNames.length; i++) {
      const dn = definedNames[i];
      let content = dn.textContent || "";
      if (kOrder > 0) {
        content = shiftDefinedNameRange(content, orderListSheetName, orderTotalsRow, kOrder);
      }
      if (kPo > 0) {
        content = shiftDefinedNameRange(content, poMultivacSheetName, poTotalRow, kPo);
      }
      dn.textContent = content;
    }

    // calcPr fullCalcOnLoad="1"
    let calcPr = wbDoc.getElementsByTagName("calcPr")[0];
    if (!calcPr) {
      const wbDocNs = getDocNamespace(wbDoc);
      calcPr = wbDoc.createElementNS(wbDocNs, "calcPr");
      const extLst = wbDoc.getElementsByTagName("extLst")[0];
      if (extLst) {
        wbDoc.documentElement.insertBefore(calcPr, extLst);
      } else {
        wbDoc.documentElement.appendChild(calcPr);
      }
    }
    calcPr.setAttribute("fullCalcOnLoad", "1");

    zip.file("xl/workbook.xml", serializeXml(wbDoc));
  }

  // 7b. Delete xl/calcChain.xml and clean up rels & [Content_Types].xml
  if (zip.file("xl/calcChain.xml")) {
    zip.remove("xl/calcChain.xml");
  }

  const wbRelsFile = zip.file("xl/_rels/workbook.xml.rels");
  if (wbRelsFile) {
    const wbRelsXml = await wbRelsFile.async("string");
    if (wbRelsXml.includes("calcChain.xml")) {
      const wbRelsDoc = parseXml(wbRelsXml);
      const rels = wbRelsDoc.getElementsByTagName("Relationship");
      for (let i = rels.length - 1; i >= 0; i--) {
        const target = rels[i].getAttribute("Target") || "";
        if (target.includes("calcChain.xml")) {
          wbRelsDoc.documentElement.removeChild(rels[i]);
        }
      }
      zip.file("xl/_rels/workbook.xml.rels", serializeXml(wbRelsDoc));
    }
  }

  const contentTypesFile = zip.file("[Content_Types].xml");
  if (contentTypesFile) {
    const ctXml = await contentTypesFile.async("string");
    if (ctXml.includes("calcChain.xml")) {
      const ctDoc = parseXml(ctXml);
      const overrides = ctDoc.getElementsByTagName("Override");
      for (let i = overrides.length - 1; i >= 0; i--) {
        const partName = overrides[i].getAttribute("PartName") || "";
        if (partName.includes("calcChain.xml")) {
          ctDoc.documentElement.removeChild(overrides[i]);
        }
      }
      zip.file("[Content_Types].xml", serializeXml(ctDoc));
    }
  }

  // 7c. Order list comments + VML parts (if kOrder > 0)
  if (kOrder > 0) {
    const orderRelsPath = orderListPath.replace("worksheets/", "worksheets/_rels/").replace(".xml", ".xml.rels");
    const orderRelsFile = zip.file(orderRelsPath);
    if (orderRelsFile) {
      const orderRelsXml = await orderRelsFile.async("string");
      const orderRelsDoc = parseXml(orderRelsXml);
      const rels = orderRelsDoc.getElementsByTagName("Relationship");

      for (let i = 0; i < rels.length; i++) {
        const type = rels[i].getAttribute("Type") || "";
        const target = rels[i].getAttribute("Target") || "";
        const resolvedPath = target.startsWith("/")
          ? target.substring(1)
          : `xl/${target.replace(/^\.\.\//, "")}`;

        if (type.includes("comments") && zip.file(resolvedPath)) {
          let commentsXml = await zip.file(resolvedPath)!.async("string");
          const commentsDoc = parseXml(commentsXml);
          const commentNodes = commentsDoc.getElementsByTagName("comment");

          for (let c = 0; c < commentNodes.length; c++) {
            const refAttr = commentNodes[c].getAttribute("ref") || "";
            const match = refAttr.match(/^([A-Za-z]+)([0-9]+)$/);
            if (match) {
              const r = parseInt(match[2], 10);
              if (r >= orderTotalsRow) {
                commentNodes[c].setAttribute("ref", `${match[1]}${r + kOrder}`);
              }
            }
          }
          zip.file(resolvedPath, serializeXml(commentsDoc));
        }

        if (type.includes("vmlDrawing") && zip.file(resolvedPath)) {
          let vmlText = await zip.file(resolvedPath)!.async("string");
          // Update <x:Row> and <x:Anchor>
          vmlText = vmlText.replace(/<x:Row>([0-9]+)<\/x:Row>/g, (m, rStr) => {
            const r = parseInt(rStr, 10);
            if (r >= orderTotalsRow - 1) {
              return `<x:Row>${r + kOrder}</x:Row>`;
            }
            return m;
          });

          vmlText = vmlText.replace(/<x:Anchor>\s*([^<]+)\s*<\/x:Anchor>/g, (m, coordsStr) => {
            const parts = coordsStr.split(",").map((p: string) => p.trim());
            if (parts.length >= 8) {
              const row1 = parseInt(parts[2], 10);
              const row2 = parseInt(parts[6], 10);
              if (row1 >= orderTotalsRow - 1) {
                parts[2] = String(row1 + kOrder);
              }
              if (row2 >= orderTotalsRow - 1) {
                parts[6] = String(row2 + kOrder);
              }
              return `<x:Anchor>${parts.join(", ")}</x:Anchor>`;
            }
            return m;
          });

          zip.file(resolvedPath, vmlText);
        }
      }
    }
  }

  // Save modified sheet XMLs
  zip.file(orderListPath, serializeXml(orderDoc));
  zip.file(poMultivacPath, serializeXml(poDoc));

  // Rezip with DEFLATE
  const filledZipBytes = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  return {
    filledZipBytes,
    orderListSheetName,
    poMultivacSheetName,
    orderListRowsCount: numPositions,
    poRowsCount: numPositions,
    totalExworksOrderList,
    poTotalCalculated,
  };
}
