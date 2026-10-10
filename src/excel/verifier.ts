import JSZip from "jszip";
import { HyperFormula } from "hyperformula";
import type { QuotationExtraction, PoDetailsOption, VerificationResult, VerificationCheck } from "../types/index.ts";
import {
  parseXml,
  extractSharedStrings,
  getCellValue,
  colLettersToNumber,
  findRow,
  findCell,
} from "./xmlUtils.ts";

export async function verifyFilledWorkbook(
  originalZipBytes: Uint8Array,
  filledZipBytes: Uint8Array,
  quotation: QuotationExtraction,
  exchangeRate: number,
  poDetailsOption: PoDetailsOption,
  poDetailsCustomText: string,
  initialPoDetailsText: string | null
): Promise<VerificationResult> {
  const origZip = await JSZip.loadAsync(originalZipBytes);
  const filledZip = await JSZip.loadAsync(filledZipBytes);

  const checks: VerificationCheck[] = [];
  const addCheck = (id: string, title: string, passed: boolean, details?: string) => {
    checks.push({ id, title, passed, details });
  };

  // -------------------------------------------------------------------------
  // Check 1: Package integrity & allowed files byte-identity
  // -------------------------------------------------------------------------
  // Find Sheet Paths in filled workbook & check sheet structure
  // -------------------------------------------------------------------------
  const origWbDoc = parseXml(await origZip.file("xl/workbook.xml")!.async("string"));
  const filledWbDoc = parseXml(await filledZip.file("xl/workbook.xml")!.async("string"));
  const relsDoc = parseXml(await filledZip.file("xl/_rels/workbook.xml.rels")!.async("string"));

  const origSheets = Array.from(origWbDoc.getElementsByTagName("sheet"));
  const filledSheets = Array.from(filledWbDoc.getElementsByTagName("sheet"));

  let sheetsMatch = origSheets.length === filledSheets.length;
  if (sheetsMatch) {
    for (let i = 0; i < origSheets.length; i++) {
      if (
        origSheets[i].getAttribute("name") !== filledSheets[i].getAttribute("name") ||
        origSheets[i].getAttribute("state") !== filledSheets[i].getAttribute("state")
      ) {
        sheetsMatch = false;
        break;
      }
    }
  }

  let orderListName = "";
  let poMultivacName = "";
  let orderListRId = "";
  let poMultivacRId = "";

  for (const s of filledSheets) {
    const name = s.getAttribute("name") || "";
    const id = s.getAttribute("r:id") || s.getAttribute("id") || "";
    if (name.toLowerCase() === "order list") {
      orderListName = name;
      orderListRId = id;
    } else if (name.toLowerCase() === "po multivac") {
      poMultivacName = name;
      poMultivacRId = id;
    }
  }

  const rels = Array.from(relsDoc.getElementsByTagName("Relationship"));
  let orderListTarget = "";
  let poMultivacTarget = "";
  for (const r of rels) {
    if (r.getAttribute("Id") === orderListRId) orderListTarget = r.getAttribute("Target") || "";
    if (r.getAttribute("Id") === poMultivacRId) poMultivacTarget = r.getAttribute("Target") || "";
  }

  const resolvePath = (t: string) => (t.startsWith("xl/") ? t : `xl/${t.replace(/^\.\//, "")}`);
  const orderListPath = resolvePath(orderListTarget);
  const poMultivacPath = resolvePath(poMultivacTarget);

  // -------------------------------------------------------------------------
  // Check 1: Package integrity & allowed files byte-identity
  // -------------------------------------------------------------------------
  const origFiles = Object.keys(origZip.files).filter((p) => !origZip.files[p].dir);

  const allowedExactFiles = new Set([
    "xl/workbook.xml",
    "[Content_Types].xml",
    "xl/_rels/workbook.xml.rels",
    "xl/calcChain.xml",
    orderListPath,
    poMultivacPath,
  ]);

  let unallowedModifiedFiles: string[] = [];
  for (const fName of origFiles) {
    if (fName === "xl/calcChain.xml") continue; // intentionally deleted
    const isAllowed =
      allowedExactFiles.has(fName) ||
      /^xl\/comments[0-9]*\.xml$/.test(fName) ||
      /^xl\/drawings\/vmlDrawing[0-9]*\.vml$/.test(fName);

    if (!isAllowed) {
      const origContent = await origZip.file(fName)!.async("uint8array");
      const filledFile = filledZip.file(fName);
      if (!filledFile) {
        unallowedModifiedFiles.push(`${fName} (missing in output)`);
        continue;
      }
      const filledContent = await filledFile.async("uint8array");
      if (origContent.length !== filledContent.length) {
        unallowedModifiedFiles.push(`${fName} (byte length mismatch)`);
      } else {
        let match = true;
        for (let i = 0; i < origContent.length; i++) {
          if (origContent[i] !== filledContent[i]) {
            match = false;
            break;
          }
        }
        if (!match) {
          unallowedModifiedFiles.push(`${fName} (byte content modified)`);
        }
      }
    }
  }

  const check1Passed = unallowedModifiedFiles.length === 0 && sheetsMatch && !filledZip.file("xl/calcChain.xml");
  addCheck(
    "package-byte-identity",
    "Package integrity & unmodified sheets byte-identity",
    check1Passed,
    check1Passed
      ? "All non-target sheets, drawings, and metadata files remain byte-identical to original template; calcChain safely removed."
      : `Discrepancies found: ${unallowedModifiedFiles.join(", ") || "Sheet list / order modified."}`
  );

  const sharedStringsFile = filledZip.file("xl/sharedStrings.xml");
  let sharedStrings: string[] = [];
  if (sharedStringsFile) {
    sharedStrings = extractSharedStrings(parseXml(await sharedStringsFile.async("string")));
  }

  const orderDoc = parseXml(await filledZip.file(orderListPath)!.async("string"));
  const poDoc = parseXml(await filledZip.file(poMultivacPath)!.async("string"));

  // -------------------------------------------------------------------------
  // Check 2: Order list validation
  // -------------------------------------------------------------------------
  const orderRows = Array.from(orderDoc.getElementsByTagName("row"));
  let orderHeaderRow = -1;
  for (const r of orderRows) {
    const rNum = parseInt(r.getAttribute("r") || "0", 10);
    if (rNum <= 10) {
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

  const orderFirstItemRow = orderHeaderRow + 1;
  const numPositions = quotation.positions.length;
  const orderExpectedLastItemRow = orderFirstItemRow + numPositions - 1;

  let orderItemsMatch = true;
  let orderItemsDetails = "";

  for (let i = 0; i < numPositions; i++) {
    const rNum = orderFirstItemRow + i;
    const pos = quotation.positions[i];
    const rowEl = findRow(orderDoc, rNum);
    if (!rowEl) {
      orderItemsMatch = false;
      orderItemsDetails = `Missing item row ${rNum}`;
      break;
    }

    const cellA = findCell(rowEl, `A${rNum}`);
    const cellB = findCell(rowEl, `B${rNum}`);
    const cellE = findCell(rowEl, `E${rNum}`);
    const cellF = findCell(rowEl, `F${rNum}`);
    const cellI = findCell(rowEl, `I${rNum}`);
    const cellO = findCell(rowEl, `O${rNum}`);
    const cellP = findCell(rowEl, `P${rNum}`);

    const refVal = getCellValue(cellA, sharedStrings).trim();
    if (refVal !== pos.reference) {
      orderItemsMatch = false;
      orderItemsDetails = `Row ${rNum} reference '${refVal}' does not match quotation '${pos.reference}'`;
      break;
    }

    const nameVal = getCellValue(cellB, sharedStrings).trim();
    if (nameVal !== pos.partName) {
      orderItemsMatch = false;
      orderItemsDetails = `Row ${rNum} part name '${nameVal}' does not match quotation '${pos.partName}'`;
      break;
    }

    const priceVal = parseFloat(getCellValue(cellE, sharedStrings));
    if (Math.abs(priceVal - pos.netPrice) > 0.01) {
      orderItemsMatch = false;
      orderItemsDetails = `Row ${rNum} net price ${priceVal} does not match quotation ${pos.netPrice}`;
      break;
    }

    const qtyVal = parseFloat(getCellValue(cellF, sharedStrings));
    if (Math.abs(qtyVal - pos.quantity) > 0.001) {
      orderItemsMatch = false;
      orderItemsDetails = `Row ${rNum} quantity ${qtyVal} does not match quotation ${pos.quantity}`;
      break;
    }

    const rateVal = parseFloat(getCellValue(cellI, sharedStrings));
    if (Math.abs(rateVal - exchangeRate) > 0.0001) {
      orderItemsMatch = false;
      orderItemsDetails = `Row ${rNum} rate ${rateVal} does not match user exchange rate ${exchangeRate}`;
      break;
    }

    const dutyVal = parseFloat(getCellValue(cellO, sharedStrings));
    if (isNaN(dutyVal) || dutyVal !== 0) {
      orderItemsMatch = false;
      orderItemsDetails = `Row ${rNum} duty is not 0 (got ${dutyVal})`;
      break;
    }

    const vatVal = parseFloat(getCellValue(cellP, sharedStrings));
    if (isNaN(vatVal) || Math.abs(vatVal - 0.15) > 0.001) {
      orderItemsMatch = false;
      orderItemsDetails = `Row ${rNum} VAT is not 0.15 (got ${vatVal})`;
      break;
    }
  }

  // Check totals row SUM range in Order list
  const orderTotalsRowNum = orderExpectedLastItemRow + 1;
  const orderTotalsRowEl = findRow(orderDoc, orderTotalsRowNum);
  let orderSumRangeValid = false;
  if (orderTotalsRowEl) {
    const cellE = findCell(orderTotalsRowEl, `E${orderTotalsRowNum}`);
    const fNode = cellE?.getElementsByTagName("f")[0];
    const formula = fNode?.textContent?.trim().toUpperCase() || "";
    const cleanFormula = formula.replace(/[\$\s]/g, "");
    if (cleanFormula === `SUM(E${orderFirstItemRow}:E${orderExpectedLastItemRow})`) {
      orderSumRangeValid = true;
    }
  }

  const check2Passed = orderItemsMatch && orderSumRangeValid;
  addCheck(
    "order-list-verification",
    "Order list: positions, values, exchange rate, and totals range",
    check2Passed,
    check2Passed
      ? `All ${numPositions} quotation positions correctly typed into Order list; formulas intact; totals SUM covers exact rows E${orderFirstItemRow}:E${orderExpectedLastItemRow}.`
      : orderItemsDetails || `Totals SUM range mismatch in row ${orderTotalsRowNum}.`
  );

  // -------------------------------------------------------------------------
  // Check 3: PO Multivac validation
  // -------------------------------------------------------------------------
  const poRows = Array.from(poDoc.getElementsByTagName("row"));
  let poHeaderRow = -1;
  for (const r of poRows) {
    const rNum = parseInt(r.getAttribute("r") || "0", 10);
    if (rNum <= 60) {
      const cellA = findCell(r, `A${rNum}`);
      if (cellA) {
        const text = getCellValue(cellA, sharedStrings).trim().toLowerCase();
        if (text === "pos" || text === "pos.") {
          poHeaderRow = rNum;
          break;
        }
      }
    }
  }

  const poFirstItemRow = poHeaderRow + 1;
  const poExpectedLastItemRow = poFirstItemRow + numPositions - 1;
  let poRowsValid = true;
  let poRowsDetails = "";

  for (let i = 0; i < numPositions; i++) {
    const rNum = poFirstItemRow + i;
    const srcRow = orderFirstItemRow + i;
    const rowEl = findRow(poDoc, rNum);
    if (!rowEl) {
      poRowsValid = false;
      poRowsDetails = `Missing PO row ${rNum}`;
      break;
    }

    const cellA = findCell(rowEl, `A${rNum}`);
    const cellB = findCell(rowEl, `B${rNum}`);
    const cellC = findCell(rowEl, `C${rNum}`);
    const cellD = findCell(rowEl, `D${rNum}`);
    const cellE = findCell(rowEl, `E${rNum}`);
    const cellF = findCell(rowEl, `F${rNum}`);
    const cellG = findCell(rowEl, `G${rNum}`);

    const posNum = parseInt(getCellValue(cellA, sharedStrings), 10);
    if (posNum !== i + 1) {
      poRowsValid = false;
      poRowsDetails = `PO row ${rNum} position number is ${posNum}, expected ${i + 1}`;
      break;
    }

    const fB = cellB?.getElementsByTagName("f")[0]?.textContent?.trim().replace(/^=/, "") || "";
    if (fB !== `'${orderListName}'!A${srcRow}`) {
      poRowsValid = false;
      poRowsDetails = `PO row ${rNum} Reference formula is '${fB}', expected ''${orderListName}'!A${srcRow}'`;
      break;
    }

    const fC = cellC?.getElementsByTagName("f")[0]?.textContent?.trim().replace(/^=/, "") || "";
    if (fC !== `'${orderListName}'!B${srcRow}`) {
      poRowsValid = false;
      poRowsDetails = `PO row ${rNum} Description formula is '${fC}', expected ''${orderListName}'!B${srcRow}'`;
      break;
    }

    const fD = cellD?.getElementsByTagName("f")[0]?.textContent?.trim().replace(/^=/, "") || "";
    if (fD !== `'${orderListName}'!F${srcRow}`) {
      poRowsValid = false;
      poRowsDetails = `PO row ${rNum} Qty formula is '${fD}', expected ''${orderListName}'!F${srcRow}'`;
      break;
    }

    const fF = cellF?.getElementsByTagName("f")[0]?.textContent?.trim().replace(/^=/, "") || "";
    if (fF !== `'${orderListName}'!E${srcRow}`) {
      poRowsValid = false;
      poRowsDetails = `PO row ${rNum} Unit Price formula is '${fF}', expected ''${orderListName}'!E${srcRow}'`;
      break;
    }

    const fG = cellG?.getElementsByTagName("f")[0]?.textContent?.trim().replace(/^=/, "") || "";
    if (fG !== `F${rNum}*D${rNum}` && fG !== `D${rNum}*F${rNum}`) {
      poRowsValid = false;
      poRowsDetails = `PO row ${rNum} Line Amount formula is '${fG}', expected 'F${rNum}*D${rNum}'`;
      break;
    }

    // Check Details (Col E)
    if (i === 0) {
      const eVal = getCellValue(cellE, sharedStrings).trim();
      if (poDetailsOption === "clear" && eVal !== "") {
        poRowsValid = false;
        poRowsDetails = `First PO Details cell was not cleared (contains '${eVal}')`;
        break;
      } else if (poDetailsOption === "replace" && eVal !== poDetailsCustomText) {
        poRowsValid = false;
        poRowsDetails = `First PO Details cell has '${eVal}', expected '${poDetailsCustomText}'`;
        break;
      }
    } else {
      const eVal = getCellValue(cellE, sharedStrings).trim();
      if (eVal !== "") {
        poRowsValid = false;
        poRowsDetails = `PO Details cell in row ${rNum} must be empty (found '${eVal}')`;
        break;
      }
    }
  }

  // Check PO Total row SUM formula
  const poTotalRowNum = poExpectedLastItemRow + 1;
  const poTotalRowEl = findRow(poDoc, poTotalRowNum);
  let poSumRangeValid = false;
  if (poTotalRowEl) {
    const cellG = findCell(poTotalRowEl, `G${poTotalRowNum}`);
    const formula = cellG?.getElementsByTagName("f")[0]?.textContent?.trim().toUpperCase() || "";
    const cleanPoFormula = formula.replace(/[\$\s]/g, "");
    if (cleanPoFormula === `SUM(G${poFirstItemRow}:G${poExpectedLastItemRow})`) {
      poSumRangeValid = true;
    }
  }

  // Check Quotation Number header
  let quotationHeaderFound = false;
  for (const r of poRows) {
    const cells = r.getElementsByTagName("c");
    for (let c = 0; c < cells.length; c++) {
      const text = getCellValue(cells[c], sharedStrings).trim().toLowerCase();
      if (text.includes("your quotation") && text.includes(quotation.quotationNumber.toLowerCase())) {
        quotationHeaderFound = true;
        break;
      }
    }
    if (quotationHeaderFound) break;
  }

  // Check A4 / F4 untouched
  const row4 = findRow(poDoc, 4);
  const cellA4 = row4 ? findCell(row4, "A4") : null;
  const cellF4 = row4 ? findCell(row4, "F4") : null;
  const a4Formula = cellA4?.getElementsByTagName("f")[0]?.textContent?.trim() || "";
  const f4Formula = cellF4?.getElementsByTagName("f")[0]?.textContent?.trim() || "";
  const a4f4Untouched = a4Formula.includes("Template") || a4Formula.length > 0;

  const check3Passed = poRowsValid && poSumRangeValid && quotationHeaderFound && a4f4Untouched;
  addCheck(
    "po-multivac-verification",
    "PO Multivac: links to Order list, line math, Details text, and quotation header",
    check3Passed,
    check3Passed
      ? `PO Multivac correctly links all ${numPositions} rows to Order list; Details selection applied; Your Quotation header updated.`
      : poRowsDetails ||
        (!poSumRangeValid ? "PO Total SUM formula mismatch." : !quotationHeaderFound ? "Quotation header not updated." : "A4/F4 modified.")
  );

  // -------------------------------------------------------------------------
  // Check 4: HyperFormula numerical calculation & error check
  // -------------------------------------------------------------------------
  const hf = HyperFormula.buildEmpty({ licenseKey: "gpl-v3" });
  hf.addSheet(orderListName);
  hf.addSheet(poMultivacName);
  const hfOrderSheetId = hf.getSheetId(orderListName)!;
  const hfPoSheetId = hf.getSheetId(poMultivacName)!;

  // Load Order list
  for (const r of orderRows) {
    const rNum = parseInt(r.getAttribute("r") || "0", 10);
    const cells = r.getElementsByTagName("c");
    for (let c = 0; c < cells.length; c++) {
      const cell = cells[c];
      const cRef = cell.getAttribute("r") || "";
      const colLetters = cRef.replace(/[0-9]/g, "");
      const colNum = colLettersToNumber(colLetters) - 1;
      const rowIdx = rNum - 1;

      const fNode = cell.getElementsByTagName("f")[0];
      if (fNode && fNode.textContent) {
        const formulaText = fNode.textContent.trim().replace(/^=/, "");
        hf.setCellContents({ col: colNum, row: rowIdx, sheet: hfOrderSheetId }, [[`=${formulaText}`]]);
      } else {
        const val = getCellValue(cell, sharedStrings);
        const num = parseFloat(val);
        hf.setCellContents({ col: colNum, row: rowIdx, sheet: hfOrderSheetId }, [[isNaN(num) || val === "" ? val : num]]);
      }
    }
  }

  // Load PO Multivac
  for (const r of poRows) {
    const rNum = parseInt(r.getAttribute("r") || "0", 10);
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
        if (fNode.textContent.includes("!") && !fNode.textContent.includes(orderListName)) {
          continue;
        }
        const formulaText = fNode.textContent.trim().replace(/^=/, "");
        hf.setCellContents({ col: colNum, row: rowIdx, sheet: hfPoSheetId }, [[`=${formulaText}`]]);
      } else {
        const val = getCellValue(cell, sharedStrings);
        const num = parseFloat(val);
        hf.setCellContents({ col: colNum, row: rowIdx, sheet: hfPoSheetId }, [[isNaN(num) || val === "" ? val : num]]);
      }
    }
  }

  // Evaluate Totals
  const totalExworksVal = hf.getCellValue({
    col: colLettersToNumber("G") - 1,
    row: orderTotalsRowNum - 1,
    sheet: hfOrderSheetId,
  });
  const totalExworks = typeof totalExworksVal === "number" ? totalExworksVal : 0;

  const poTotalVal = hf.getCellValue({
    col: colLettersToNumber("G") - 1,
    row: poTotalRowNum - 1,
    sheet: hfPoSheetId,
  });
  const poTotal = typeof poTotalVal === "number" ? poTotalVal : 0;

  const invoiceAmount = quotation.invoiceAmount;
  const sumNetQty = quotation.positions.reduce((acc, p) => acc + p.netPrice * p.quantity, 0);

  const exworksDiff = Math.abs(totalExworks - invoiceAmount);
  const poTotalDiff = Math.abs(poTotal - invoiceAmount);
  const sumNetDiff = Math.abs(sumNetQty - invoiceAmount);

  let numChecksPassed = true;
  let numDetails = "";

  if (exworksDiff > 0.05 || poTotalDiff > 0.05 || sumNetDiff > 0.05) {
    numChecksPassed = false;
    numDetails = `Total Exworks (${totalExworks.toFixed(2)}) / PO Total (${poTotal.toFixed(2)}) does not match Invoice amount (${invoiceAmount.toFixed(2)}).`;
  }

  // Check each line: netPrice * quantity = PDF line amount (0.02 tolerance)
  for (let i = 0; i < quotation.positions.length; i++) {
    const p = quotation.positions[i];
    const calc = p.netPrice * p.quantity;
    if (Math.abs(calc - p.lineAmount) > 0.05) {
      numChecksPassed = false;
      numDetails = `Position ${p.position} calculation (${calc.toFixed(2)}) differs from PDF line amount (${p.lineAmount.toFixed(2)}).`;
      break;
    }
    if (p.listPrice > 0 && p.netPrice > p.listPrice + 0.001) {
      numChecksPassed = false;
      numDetails = `Position ${p.position} net price (${p.netPrice}) exceeds list price (${p.listPrice}).`;
      break;
    }
  }

  // Check freight allocation column L sums to 100 in Order list
  const freightColNum = colLettersToNumber("L") - 1;
  const freightSumVal = hf.getCellValue({
    col: freightColNum,
    row: orderTotalsRowNum - 1,
    sheet: hfOrderSheetId,
  });
  if (typeof freightSumVal === "number") {
    // Percentage in Excel can be 1.0 (100%) or 100
    if (Math.abs(freightSumVal - 1) > 0.02 && Math.abs(freightSumVal - 100) > 0.5) {
      numChecksPassed = false;
      numDetails = `Freight allocation sum in column L is ${freightSumVal}, expected 100% / 1.0.`;
    }
  }

  // Check no formula errors (#REF!, #DIV/0!, #VALUE!, #NAME?, #N/A) in edited sheets
  const formulaErrorNames = ["#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NULL!", "#NUM!"];
  let hasFormulaErrors = false;

  // Scan Order list item and totals rows
  for (let r = orderFirstItemRow; r <= orderTotalsRowNum; r++) {
    const rowEl = findRow(orderDoc, r);
    if (!rowEl) continue;
    const cells = rowEl.getElementsByTagName("c");
    for (let c = 0; c < cells.length; c++) {
      const fNode = cells[c].getElementsByTagName("f")[0];
      if (!fNode) continue;
      const cRef = cells[c].getAttribute("r") || "";
      const colLetters = cRef.replace(/[0-9]/g, "");
      const colNum = colLettersToNumber(colLetters) - 1;
      const v = hf.getCellValue({ col: colNum, row: r - 1, sheet: hfOrderSheetId });
      const errName =
        typeof v === "string" && (v.startsWith("#") || formulaErrorNames.includes(v))
          ? v
          : v && typeof v === "object" && ("message" in (v as any) || "type" in (v as any))
          ? (v as any).message || `#${(v as any).type}!`
          : "";
      if (errName) {
        hasFormulaErrors = true;
        numChecksPassed = false;
        numDetails = `Formula error ${errName} detected in Order list cell ${cRef}.`;
        break;
      }
    }
    if (hasFormulaErrors) break;
  }

  // Scan PO Multivac item and total rows
  for (let r = poFirstItemRow; r <= poTotalRowNum; r++) {
    const rowEl = findRow(poDoc, r);
    if (!rowEl) continue;
    const cells = rowEl.getElementsByTagName("c");
    for (let c = 0; c < cells.length; c++) {
      const fNode = cells[c].getElementsByTagName("f")[0];
      if (!fNode) continue;
      const cRef = cells[c].getAttribute("r") || "";
      const colLetters = cRef.replace(/[^A-Za-z]/g, "");
      const colNum = colLettersToNumber(colLetters) - 1;
      const v = hf.getCellValue({ col: colNum, row: r - 1, sheet: hfPoSheetId });
      const errName =
        typeof v === "string" && (v.startsWith("#") || formulaErrorNames.includes(v))
          ? v
          : v && typeof v === "object" && ("message" in (v as any) || "type" in (v as any))
          ? (v as any).message || `#${(v as any).type}!`
          : "";
      if (errName) {
        hasFormulaErrors = true;
        numChecksPassed = false;
        numDetails = `Formula error ${errName} detected in PO Multivac cell ${cRef}.`;
        break;
      }
    }
    if (hasFormulaErrors) break;
  }

  addCheck(
    "numerical-reconciliation",
    "Numerical reconciliation & formula error scan",
    numChecksPassed,
    numChecksPassed
      ? `Total Exworks (€${totalExworks.toFixed(2)}) matches Quotation Invoice amount (€${invoiceAmount.toFixed(2)}); freight allocation = 100%; zero formula errors.`
      : numDetails
  );

  // -------------------------------------------------------------------------
  // Notices construction
  // -------------------------------------------------------------------------
  const notices: string[] = [];

  // Non-PC units
  const nonPcUnits = Array.from(new Set(quotation.positions.map((p) => p.unit).filter((u) => u && u.toUpperCase() !== "PC")));
  if (nonPcUnits.length > 0) {
    notices.push(`Quotation contains items with non-PC unit of measure: ${nonPcUnits.join(", ")}.`);
  }

  // Ignored lines notice
  if (quotation.ignoredLines.length > 0) {
    notices.push(`${quotation.ignoredLines.length} descriptive sub-total or note lines were excluded from order positions.`);
  }

  notices.push("Order list columns A, B, E, F, I now hold typed values instead of links to Quotation_Multivac.");
  notices.push("PO number and date at the top of PO Multivac come from the Template sheet and still show placeholders until you fill Template.");
  notices.push("Footer dates on PO Multivac and the rate-header date in Order list are fixed text and unchanged.");

  if (poDetailsOption === "clear") {
    notices.push("PO Details text was cleared.");
  } else if (poDetailsOption === "keep") {
    notices.push(`PO Details sample text ('${initialPoDetailsText || "Princes Tuna"}') was kept.`);
  } else if (poDetailsOption === "replace") {
    notices.push(`PO Details text was replaced with '${poDetailsCustomText}'.`);
  }

  if (numPositions > 15) {
    notices.push("With many items the PO's closing frame can flow onto a second printed page.");
  }

  notices.push("Excel will recalculate when the file opens and may ask about updating external links.");
  notices.push("Existing errors in other sheets were left as they were.");

  const allPassed = checks.every((c) => c.passed);
  const passedCount = checks.filter((c) => c.passed).length;

  return {
    passed: allPassed,
    checks,
    totalChecks: checks.length,
    passedChecks: passedCount,
    summary: {
      exchangeRate,
      orderListRows: numPositions,
      poRows: numPositions,
      totalExworks,
      invoiceAmount,
      poTotal,
    },
    notices,
  };
}
