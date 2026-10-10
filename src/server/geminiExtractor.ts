import { GoogleGenAI, Type } from "@google/genai";
import type { QuotationExtraction, QuotationPosition, UnclearItem } from "../types/index.ts";

const quotationResponseSchema = {
  type: Type.OBJECT,
  properties: {
    quotationNumber: {
      type: Type.STRING,
      description: "The Document number printed at the top of the quotation (e.g. 21400572). NOT customer number, NOT 'Your quote request' reference. If ambiguous, record in unclear.",
    },
    invoiceAmount: {
      type: Type.NUMBER,
      description: "The final 'Invoice amount' in EUR. Plain decimal number (European formatting 1.234,56 converted to 1234.56).",
    },
    currency: {
      type: Type.STRING,
      description: "Quotation currency, e.g. EUR.",
    },
    positions: {
      type: Type.ARRAY,
      description: "All quotation line items in order.",
      items: {
        type: Type.OBJECT,
        properties: {
          position: {
            type: Type.STRING,
            description: "Position number, e.g. '10', '20'.",
          },
          reference: {
            type: Type.STRING,
            description: "The bold material or part number. Leading zeros must be strictly preserved.",
          },
          partName: {
            type: Type.STRING,
            description: "Full part description with its dimension line joined (e.g. 'Tension spring D1=24'). If split across page break, join pieces.",
          },
          netPrice: {
            type: Type.NUMBER,
            description: "Per-unit Net price AFTER customer discount. NOT list price, NOT line amount.",
          },
          listPrice: {
            type: Type.NUMBER,
            description: "List price per unit before discount.",
          },
          quantity: {
            type: Type.NUMBER,
            description: "Quantity as a number.",
          },
          unit: {
            type: Type.STRING,
            description: "Unit of measure, e.g. 'PC', 'm', 'SET'.",
          },
          lineAmount: {
            type: Type.NUMBER,
            description: "The line total amount printed in the PDF.",
          },
        },
        required: ["position", "reference", "partName", "netPrice", "listPrice", "quantity", "unit", "lineAmount"],
      },
    },
    ignoredLines: {
      type: Type.ARRAY,
      description: "Sub-total/text/note lines without price and quantity.",
      items: {
        type: Type.OBJECT,
        properties: {
          lineText: { type: Type.STRING },
          reason: { type: Type.STRING },
        },
        required: ["lineText", "reason"],
      },
    },
    unclear: {
      type: Type.ARRAY,
      description: "Anything ambiguous, unreadable, or uncertain. Never guess.",
      items: {
        type: Type.OBJECT,
        properties: {
          item: { type: Type.STRING },
          reason: { type: Type.STRING },
        },
        required: ["item", "reason"],
      },
    },
  },
  required: ["quotationNumber", "invoiceAmount", "currency", "positions", "ignoredLines", "unclear"],
};

function getGeminiClient(): GoogleGenAI {
  const apiKey = process.env.GEMINI_API_KEY || process.env.API_KEY;
  if (!apiKey) {
    throw new Error("GEMINI_API_KEY_MISSING");
  }
  return new GoogleGenAI({
    apiKey,
    httpOptions: {
      headers: {
        "User-Agent": "aistudio-build",
      },
    },
  });
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function isTransientError(err: any): boolean {
  const message = String(err?.message || "").toLowerCase();
  const status = err?.status || err?.statusCode;
  return (
    status === 429 ||
    status === 503 ||
    message.includes("429") ||
    message.includes("503") ||
    message.includes("resource_exhausted") ||
    message.includes("unavailable") ||
    message.includes("quota") ||
    message.includes("rate limit")
  );
}

async function callGeminiWithRetry(
  ai: GoogleGenAI,
  model: string,
  pdfBase64: string,
  promptText: string
): Promise<QuotationExtraction> {
  const maxRetries = 3;
  let backoffMs = 1500;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = await ai.models.generateContent({
        model,
        contents: [
          {
            inlineData: {
              mimeType: "application/pdf",
              data: pdfBase64,
            },
          },
          {
            text: promptText,
          },
        ],
        config: {
          temperature: 0,
          responseMimeType: "application/json",
          responseSchema: quotationResponseSchema,
        },
      });

      const text = response.text?.trim();
      if (!text) {
        throw new Error("Empty response returned by Gemini model");
      }

      const parsed = JSON.parse(text) as QuotationExtraction;
      return sanitizeExtraction(parsed);
    } catch (err: any) {
      if (err?.message === "GEMINI_API_KEY_MISSING") throw err;

      if (isTransientError(err) && attempt < maxRetries) {
        await delay(backoffMs);
        backoffMs *= 2;
        continue;
      }

      if (attempt === maxRetries) {
        throw err;
      }
      throw err;
    }
  }

  throw new Error("Failed to extract quotation after retries");
}

async function runExtractionPass(
  ai: GoogleGenAI,
  pdfBase64: string,
  promptText: string
): Promise<QuotationExtraction> {
  try {
    return await callGeminiWithRetry(ai, "gemini-3.8-flash", pdfBase64, promptText);
  } catch (err: any) {
    if (isTransientError(err) || String(err?.message || "").includes("model")) {
      // Fallback once to gemini-3.1-flash-lite
      return await callGeminiWithRetry(ai, "gemini-3.1-flash-lite", pdfBase64, promptText);
    }
    throw err;
  }
}

function sanitizeExtraction(data: any): QuotationExtraction {
  return {
    quotationNumber: String(data?.quotationNumber || "").trim(),
    invoiceAmount: Number(data?.invoiceAmount) || 0,
    currency: String(data?.currency || "EUR").trim(),
    positions: Array.isArray(data?.positions)
      ? data.positions.map((p: any) => ({
          position: String(p?.position || "").trim(),
          reference: String(p?.reference || "").trim(),
          partName: String(p?.partName || "").trim(),
          netPrice: Number(p?.netPrice) || 0,
          listPrice: Number(p?.listPrice) || 0,
          quantity: Number(p?.quantity) || 0,
          unit: String(p?.unit || "PC").trim(),
          lineAmount: Number(p?.lineAmount) || 0,
        }))
      : [],
    ignoredLines: Array.isArray(data?.ignoredLines)
      ? data.ignoredLines.map((l: any) => ({
          lineText: String(l?.lineText || "").trim(),
          reason: String(l?.reason || "").trim(),
        }))
      : [],
    unclear: Array.isArray(data?.unclear)
      ? data.unclear.map((u: any) => ({
          item: String(u?.item || "").trim(),
          reason: String(u?.reason || "").trim(),
        }))
      : [],
  };
}

interface DiffResult {
  hasDifferences: boolean;
  differences: string[];
}

function compareExtractions(a: QuotationExtraction, b: QuotationExtraction): DiffResult {
  const diffs: string[] = [];

  if (a.quotationNumber !== b.quotationNumber) {
    diffs.push(`Quotation number mismatch: Pass 1='${a.quotationNumber}', Pass 2='${b.quotationNumber}'`);
  }

  if (Math.abs(a.invoiceAmount - b.invoiceAmount) > 0.01) {
    diffs.push(`Invoice amount mismatch: Pass 1=${a.invoiceAmount}, Pass 2=${b.invoiceAmount}`);
  }

  if (a.positions.length !== b.positions.length) {
    diffs.push(`Position count mismatch: Pass 1 has ${a.positions.length} items, Pass 2 has ${b.positions.length} items`);
    return { hasDifferences: true, differences: diffs };
  }

  for (let i = 0; i < a.positions.length; i++) {
    const p1 = a.positions[i];
    const p2 = b.positions[i];
    const idx = i + 1;

    if (p1.position !== p2.position) {
      diffs.push(`Item #${idx} label mismatch: '${p1.position}' vs '${p2.position}'`);
    }
    if (p1.reference !== p2.reference) {
      diffs.push(`Item #${idx} reference mismatch: '${p1.reference}' vs '${p2.reference}'`);
    }
    if (Math.abs(p1.netPrice - p2.netPrice) > 0.01) {
      diffs.push(`Item #${idx} (${p1.reference}) netPrice mismatch: ${p1.netPrice} vs ${p2.netPrice}`);
    }
    if (Math.abs(p1.quantity - p2.quantity) > 0.001) {
      diffs.push(`Item #${idx} (${p1.reference}) quantity mismatch: ${p1.quantity} vs ${p2.quantity}`);
    }
    if (p1.unit.toLowerCase() !== p2.unit.toLowerCase()) {
      diffs.push(`Item #${idx} (${p1.reference}) unit mismatch: '${p1.unit}' vs '${p2.unit}'`);
    }
    if (Math.abs(p1.lineAmount - p2.lineAmount) > 0.01) {
      diffs.push(`Item #${idx} (${p1.reference}) lineAmount mismatch: ${p1.lineAmount} vs ${p2.lineAmount}`);
    }
  }

  return {
    hasDifferences: diffs.length > 0,
    differences: diffs,
  };
}

export async function extractAndReconcileQuotation(pdfBase64: string): Promise<{
  success: boolean;
  data?: QuotationExtraction;
  unclear?: UnclearItem[];
  error?: string;
}> {
  const ai = getGeminiClient();

  const prompt1 = `You are a high-precision MULTIVAC quotation extractor.
Extract all details from this MULTIVAC quotation PDF into JSON according to the schema.
Instructions:
1. quotationNumber: The Document number printed at the top of the quotation (e.g. 21400572). NOT customer number, NOT 'Your quote request' reference.
2. invoiceAmount: The final 'Invoice amount' in EUR.
3. European numbers: 3,52 = 3.52 and 1.234,56 = 1234.56.
4. positions: Extract every quotation item position.
   - position: Item position number (e.g. '10').
   - reference: Bold material/article number (preserve leading zeros strictly).
   - partName: Full description + its dimension line joined together (e.g. 'Tension spring D1=24'). Join across page breaks if needed.
   - netPrice: Per-unit Net price AFTER customer discount (NOT list price, NOT line amount).
   - listPrice: List price per unit.
   - quantity: Item quantity.
   - unit: Unit string (e.g. PC, m, SET).
   - lineAmount: Line total printed on the quotation.
5. ignoredLines: Any non-position rows, subtotals, or text headers.
6. unclear: If anything is ambiguous, unreadable, or multiple conflicting numbers could be the quotation number, DO NOT GUESS. Add to unclear.`;

  const prompt2 = `You are an independent auditor extracting a MULTIVAC quotation PDF.
Carefully parse all fields with absolute mathematical rigor:
- quotationNumber: Document number header at the top (not customer number, not quote request).
- invoiceAmount: The final Invoice amount at the bottom in EUR.
- positions:
  - position: Row number (e.g. '10')
  - reference: Bold material number, retaining any leading zeros intact.
  - partName: Part name including any secondary dimension line.
  - netPrice: Unit net price after discount (not line total).
  - listPrice: Gross list price.
  - quantity: Quantity number.
  - unit: Unit designation (PC, m, etc.).
  - lineAmount: Exact line total amount.
- Convert European comma decimals to standard numbers.
- Any uncertainties or ambiguous references must be placed in 'unclear'. Never guess.`;

  // Run Call 1 and Call 2 concurrently
  const [pass1, pass2] = await Promise.all([
    runExtractionPass(ai, pdfBase64, prompt1),
    runExtractionPass(ai, pdfBase64, prompt2),
  ]);

  let finalData: QuotationExtraction;
  const comparison = compareExtractions(pass1, pass2);

  if (!comparison.hasDifferences) {
    finalData = pass1;
  } else {
    // Call 3: Senior auditor with differences
    const prompt3 = `You are a Senior Auditor resolving differences between two independent extraction passes of this MULTIVAC quotation PDF.
The two previous extraction passes had the following specific differences:
${comparison.differences.map((d, i) => `${i + 1}. ${d}`).join("\n")}

Inspect the PDF document with utmost precision and return the final, definitive JSON matching the schema.
Ensure all European numbers are converted correctly, leading zeros in reference numbers are preserved, netPrice is unit price after discount, and any genuine ambiguities are placed in unclear.`;

    finalData = await runExtractionPass(ai, pdfBase64, prompt3);
  }

  // Check mathematical and logical integrity
  const uncertainties: UnclearItem[] = [...(finalData.unclear || [])];

  if (!finalData.quotationNumber) {
    uncertainties.push({
      item: "Quotation Number",
      reason: "Could not find a clear quotation document number in the PDF header.",
    });
  }

  if (finalData.positions.length === 0) {
    uncertainties.push({
      item: "Quotation Positions",
      reason: "No line items could be detected in the quotation PDF.",
    });
  }

  // Check each position line calculation
  for (let i = 0; i < finalData.positions.length; i++) {
    const pos = finalData.positions[i];
    const expectedLineAmount = pos.netPrice * pos.quantity;
    if (Math.abs(expectedLineAmount - pos.lineAmount) > 0.03) {
      uncertainties.push({
        item: `Position ${pos.position} (${pos.reference || "item " + (i + 1)})`,
        reason: `Net price (${pos.netPrice}) × Quantity (${pos.quantity}) = ${expectedLineAmount.toFixed(2)}, which does not match printed line amount ${pos.lineAmount}.`,
      });
    }

    if (pos.netPrice > pos.listPrice && pos.listPrice > 0) {
      uncertainties.push({
        item: `Position ${pos.position} (${pos.reference})`,
        reason: `Net price (${pos.netPrice}) exceeds list price (${pos.listPrice}).`,
      });
    }

    if (!pos.reference) {
      uncertainties.push({
        item: `Position ${pos.position}`,
        reason: "Missing material reference number.",
      });
    }
  }

  // Check total sum vs invoice amount
  const sumLineAmounts = finalData.positions.reduce((acc, p) => acc + p.lineAmount, 0);
  if (Math.abs(sumLineAmounts - finalData.invoiceAmount) > 0.05) {
    // Check if difference matches freight or other line in ignoredLines
    uncertainties.push({
      item: "Total Invoice Amount",
      reason: `Sum of line amounts (${sumLineAmounts.toFixed(2)} EUR) does not match total invoice amount (${finalData.invoiceAmount.toFixed(2)} EUR).`,
    });
  }

  if (uncertainties.length > 0) {
    return {
      success: false,
      unclear: uncertainties,
      data: finalData,
    };
  }

  return {
    success: true,
    data: finalData,
  };
}
