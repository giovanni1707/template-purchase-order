import { extractAndReconcileQuotation } from "./geminiExtractor.ts";

export interface ExtractRequestBody {
  pdfBase64?: string;
}

export async function handleExtractRequest(body: ExtractRequestBody): Promise<{
  status: number;
  data: any;
}> {
  const { pdfBase64 } = body;

  if (!pdfBase64) {
    return {
      status: 400,
      data: { error: "Missing required 'pdfBase64' payload in request body." },
    };
  }

  // Check approximate base64 size (50MB is ~67MB in base64)
  if (pdfBase64.length > 70 * 1024 * 1024) {
    return {
      status: 413,
      data: {
        error: "The uploaded quotation PDF is too large for the processing limit (maximum 50 MB).",
      },
    };
  }

  try {
    const result = await extractAndReconcileQuotation(pdfBase64);
    if (!result.success) {
      return {
        status: 200,
        data: {
          success: false,
          unclear: result.unclear || [],
          data: result.data,
        },
      };
    }

    return {
      status: 200,
      data: {
        success: true,
        data: result.data,
      },
    };
  } catch (err: any) {
    console.error("Extraction error:", err);
    const msg = String(err?.message || "");

    if (msg === "GEMINI_API_KEY_MISSING") {
      return {
        status: 500,
        data: {
          error: "Gemini API key is not configured on the server. Please verify your GEMINI_API_KEY secret.",
        },
      };
    }

    if (msg.includes("API key not valid") || msg.includes("INVALID_ARGUMENT") && msg.includes("API key")) {
      return {
        status: 401,
        data: {
          error: "The provided Gemini API key is invalid.",
        },
      };
    }

    if (
      msg.includes("429") ||
      msg.includes("RESOURCE_EXHAUSTED") ||
      msg.includes("quota") ||
      msg.includes("rate limit")
    ) {
      return {
        status: 429,
        data: {
          error: "Gemini quota or rate limit exceeded. Please wait a moment and try again.",
        },
      };
    }

    if (msg.includes("timeout") || msg.includes("timed out") || msg.includes("DEADLINE_EXCEEDED")) {
      return {
        status: 504,
        data: {
          error: "The extraction request timed out. Please retry with a smaller PDF or try again.",
        },
      };
    }

    return {
      status: 500,
      data: {
        error: "An unexpected error occurred during quotation extraction: " + (err?.message || "Unknown error"),
      },
    };
  }
}
