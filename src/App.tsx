import React, { useState } from "react";
import JSZip from "jszip";
import { FileText, Loader2, Sparkles, AlertCircle } from "lucide-react";
import { FileUpload } from "./components/FileUpload.tsx";
import { PoDetailsControl } from "./components/PoDetailsControl.tsx";
import { UncertaintiesScreen } from "./components/UncertaintiesScreen.tsx";
import { ResultScreen } from "./components/ResultScreen.tsx";
import { inspectWorkbook, processWorkbook } from "./excel/workbookProcessor.ts";
import { verifyFilledWorkbook } from "./excel/verifier.ts";
import type {
  QuotationExtraction,
  PoDetailsOption,
  UnclearItem,
  VerificationResult,
} from "./types/index.ts";

type AppStep = "input" | "processing" | "unclear" | "result";

export default function App() {
  // Input states
  const [quotationFile, setQuotationFile] = useState<File | null>(null);
  const [quotationError, setQuotationError] = useState<string>("");

  const [workbookFile, setWorkbookFile] = useState<File | null>(null);
  const [workbookBytes, setWorkbookBytes] = useState<Uint8Array | null>(null);
  const [workbookError, setWorkbookError] = useState<string>("");

  const [firstPoDetailsText, setFirstPoDetailsText] = useState<string | null>(null);
  const [poDetailsOption, setPoDetailsOption] = useState<PoDetailsOption>("clear");
  const [poDetailsCustomText, setPoDetailsCustomText] = useState<string>("");

  // Exchange rate: required, ONE plain positive number, never pre-filled, never remembered
  const [exchangeRateInput, setExchangeRateInput] = useState<string>("");
  const [exchangeRateError, setExchangeRateError] = useState<string>("");

  // Processing & result states
  const [step, setStep] = useState<AppStep>("input");
  const [progressText, setProgressText] = useState<string>("Reading quotation…");
  const [generalError, setGeneralError] = useState<string>("");

  const [uncertainties, setUncertainties] = useState<UnclearItem[]>([]);
  const [extractedQuotation, setExtractedQuotation] = useState<QuotationExtraction | null>(null);
  const [verificationResult, setVerificationResult] = useState<VerificationResult | null>(null);
  const [filledZipBytes, setFilledZipBytes] = useState<Uint8Array | null>(null);

  // File selection handlers
  const handleSelectQuotation = (file: File) => {
    setGeneralError("");
    if (!file.name.toLowerCase().endsWith(".pdf")) {
      setQuotationError("Please upload a PDF quotation file.");
      setQuotationFile(null);
      return;
    }
    setQuotationError("");
    setQuotationFile(file);
  };

  const handleClearQuotation = () => {
    setQuotationFile(null);
    setQuotationError("");
  };

  const handleSelectWorkbook = async (file: File) => {
    setGeneralError("");
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      setWorkbookError("Please upload an .xlsx Excel workbook.");
      setWorkbookFile(null);
      setWorkbookBytes(null);
      setFirstPoDetailsText(null);
      return;
    }

    try {
      const buffer = await file.arrayBuffer();
      const bytes = new Uint8Array(buffer);
      const zip = await JSZip.loadAsync(bytes);

      const inspection = await inspectWorkbook(zip);
      if (!inspection.isValid) {
        setWorkbookError(
          inspection.errorMessage ||
            "This is not the Template_Purchase_Order workbook (sheets 'Order list' and 'PO Multivac' are required)."
        );
        setWorkbookFile(null);
        setWorkbookBytes(null);
        setFirstPoDetailsText(null);
        return;
      }

      setWorkbookError("");
      setWorkbookFile(file);
      setWorkbookBytes(bytes);
      setFirstPoDetailsText(inspection.firstPoDetailsText);
      setPoDetailsOption("clear");
      setPoDetailsCustomText("");
    } catch (err: any) {
      console.error("Workbook read error:", err);
      setWorkbookError(
        "This is not the Template_Purchase_Order workbook (sheets 'Order list' and 'PO Multivac' are required)."
      );
      setWorkbookFile(null);
      setWorkbookBytes(null);
      setFirstPoDetailsText(null);
    }
  };

  const handleClearWorkbook = () => {
    setWorkbookFile(null);
    setWorkbookBytes(null);
    setWorkbookError("");
    setFirstPoDetailsText(null);
    setPoDetailsOption("clear");
    setPoDetailsCustomText("");
  };

  const handleExchangeRateChange = (val: string) => {
    setExchangeRateInput(val);
    setGeneralError("");
    if (!val.trim()) {
      setExchangeRateError("Exchange rate is required.");
      return;
    }
    // Required regex ^\d+(\.\d+)?$
    const isValid = /^\d+(\.\d+)?$/.test(val.trim());
    if (!isValid || parseFloat(val.trim()) <= 0) {
      setExchangeRateError("Must be ONE plain positive number (e.g. 49.50).");
    } else {
      setExchangeRateError("");
    }
  };

  const handleStartOver = () => {
    setQuotationFile(null);
    setQuotationError("");
    setWorkbookFile(null);
    setWorkbookBytes(null);
    setWorkbookError("");
    setFirstPoDetailsText(null);
    setPoDetailsOption("clear");
    setPoDetailsCustomText("");
    setExchangeRateInput("");
    setExchangeRateError("");
    setGeneralError("");
    setUncertainties([]);
    setExtractedQuotation(null);
    setVerificationResult(null);
    setFilledZipBytes(null);
    setStep("input");
  };

  const handleGenerate = async () => {
    if (!quotationFile) {
      setQuotationError("Please upload a quotation PDF.");
      return;
    }
    if (!workbookFile || !workbookBytes) {
      setWorkbookError("Please upload the workbook template.");
      return;
    }
    const rateTrimmed = exchangeRateInput.trim();
    if (!rateTrimmed || !/^\d+(\.\d+)?$/.test(rateTrimmed) || parseFloat(rateTrimmed) <= 0) {
      setExchangeRateError("Must be ONE plain positive number (e.g. 49.50).");
      return;
    }

    const rateNum = parseFloat(rateTrimmed);
    setGeneralError("");
    setStep("processing");

    try {
      // Step 1: Read PDF and convert to base64
      setProgressText("Reading quotation…");
      const pdfBuffer = await quotationFile.arrayBuffer();
      let binary = "";
      const bytes = new Uint8Array(pdfBuffer);
      const len = bytes.byteLength;
      for (let i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
      }
      const pdfBase64 = btoa(binary);

      // Step 2: Server-side Gemini extraction + verification passes
      setProgressText("Running verification passes…");
      const response = await fetch("/api/extract", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ pdfBase64 }),
      });

      const result = await response.json();

      if (!response.ok) {
        throw new Error(result.error || `Server extraction failed (${response.status})`);
      }

      if (!result.success) {
        // Amber stop screen: uncertain or inconsistent values
        setUncertainties(result.unclear || [{ item: "Quotation Values", reason: "Uncertain or inconsistent values detected." }]);
        setStep("unclear");
        return;
      }

      const quotationData: QuotationExtraction = result.data;
      setExtractedQuotation(quotationData);

      // Step 3: Browser Excel XML direct editing
      setProgressText("Filling workbook…");
      const processed = await processWorkbook(
        workbookBytes,
        quotationData,
        rateNum,
        poDetailsOption,
        poDetailsCustomText.trim()
      );

      // Step 4: Verification pass on filled zip bytes
      setProgressText("Verifying…");
      const verification = await verifyFilledWorkbook(
        workbookBytes,
        processed.filledZipBytes,
        quotationData,
        rateNum,
        poDetailsOption,
        poDetailsCustomText.trim(),
        firstPoDetailsText
      );

      setFilledZipBytes(processed.filledZipBytes);
      setVerificationResult(verification);
      setStep("result");
    } catch (err: any) {
      console.error("Processing error:", err);
      setGeneralError(err?.message || "An unexpected error occurred while processing the order.");
      setStep("input");
    }
  };

  const isFormValid =
    quotationFile !== null &&
    workbookFile !== null &&
    exchangeRateInput.trim().length > 0 &&
    /^\d+(\.\d+)?$/.test(exchangeRateInput.trim()) &&
    parseFloat(exchangeRateInput.trim()) > 0 &&
    !quotationError &&
    !workbookError;

  return (
    <div className="min-h-screen bg-slate-100 flex flex-col justify-center items-center p-4 antialiased text-slate-900">
      {/* Centered Card */}
      <main className="w-full max-w-lg bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
        {/* Header */}
        <header className="px-6 py-5 border-b border-slate-100 bg-white">
          <div className="flex items-center gap-2.5">
            <div className="p-2 bg-slate-900 text-white rounded-lg">
              <FileText className="w-4 h-4" />
            </div>
            <div>
              <h1 className="text-base font-semibold text-slate-900 tracking-tight">
                MULTIVAC Purchase Order
              </h1>
              <p className="text-xs text-slate-500">
                Fill Template_Purchase_Order_2026.xlsx from quotation PDF
              </p>
            </div>
          </div>
        </header>

        {/* Content Area */}
        <div className="p-6">
          {generalError && (
            <div className="mb-4 p-3 bg-rose-50 border border-rose-200 rounded-lg flex items-start gap-2.5 text-xs text-rose-800">
              <AlertCircle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
              <span>{generalError}</span>
            </div>
          )}

          {/* STEP 1: Main Input Screen */}
          {step === "input" && (
            <form
              onSubmit={(e) => {
                e.preventDefault();
                handleGenerate();
              }}
              className="space-y-4"
            >
              {/* Input 1: Quotation PDF */}
              <FileUpload
                label="1. MULTIVAC Quotation PDF"
                sublabel="Quotation PDF file"
                accept=".pdf"
                selectedFile={quotationFile}
                onFileSelect={handleSelectQuotation}
                onClear={handleClearQuotation}
                errorMessage={quotationError}
              />

              {/* Input 2: Template Workbook */}
              <FileUpload
                label="2. Workbook Template"
                sublabel="Template_Purchase_Order_2026.xlsx"
                accept=".xlsx"
                selectedFile={workbookFile}
                onFileSelect={handleSelectWorkbook}
                onClear={handleClearWorkbook}
                errorMessage={workbookError}
              />

              {/* Input 3: Exchange Rate */}
              <div className="space-y-1.5">
                <div className="flex items-center justify-between text-xs">
                  <label htmlFor="exchange-rate" className="font-semibold text-slate-800">
                    3. EUR → MUR Exchange Rate
                  </label>
                  <span className="text-slate-400">Required</span>
                </div>
                <input
                  id="exchange-rate"
                  type="text"
                  inputMode="decimal"
                  autoComplete="off"
                  value={exchangeRateInput}
                  onChange={(e) => handleExchangeRateChange(e.target.value)}
                  placeholder="e.g. 50.25"
                  className={`w-full px-3 py-2 text-xs font-mono bg-white border rounded-lg text-slate-900 placeholder:text-slate-400 focus:outline-none focus:ring-1 ${
                    exchangeRateError
                      ? "border-rose-300 focus:ring-rose-500 bg-rose-50/20"
                      : "border-slate-300 focus:ring-slate-900"
                  }`}
                />
                {exchangeRateError && (
                  <p className="text-[11px] text-rose-600 font-medium">
                    {exchangeRateError}
                  </p>
                )}
              </div>

              {/* Conditional Control: PO Details text */}
              {firstPoDetailsText !== null && (
                <PoDetailsControl
                  initialText={firstPoDetailsText}
                  selectedOption={poDetailsOption}
                  customText={poDetailsCustomText}
                  onChangeOption={setPoDetailsOption}
                  onChangeCustomText={setPoDetailsCustomText}
                />
              )}

              {/* Primary Action Button */}
              <div className="pt-2">
                <button
                  type="submit"
                  disabled={!isFormValid}
                  className={`w-full flex items-center justify-center gap-2 py-2.5 px-4 rounded-lg text-xs font-semibold text-white transition-colors ${
                    isFormValid
                      ? "bg-slate-900 hover:bg-slate-800 active:bg-slate-950 shadow-sm cursor-pointer"
                      : "bg-slate-300 cursor-not-allowed"
                  }`}
                >
                  <Sparkles className="w-3.5 h-3.5" />
                  <span>Generate Purchase Order</span>
                </button>
              </div>
            </form>
          )}

          {/* STEP 2: Processing State */}
          {step === "processing" && (
            <div className="py-10 flex flex-col items-center justify-center text-center space-y-3">
              <Loader2 className="w-6 h-6 text-slate-700 animate-spin" />
              <div className="space-y-1">
                <p className="text-xs font-semibold text-slate-800">
                  {progressText}
                </p>
                <p className="text-[11px] text-slate-500">
                  Auditing quotation positions and updating workbook XML
                </p>
              </div>
            </div>
          )}

          {/* STEP 3: Amber Stop Screen */}
          {step === "unclear" && (
            <UncertaintiesScreen
              uncertainties={uncertainties}
              onStartOver={handleStartOver}
            />
          )}

          {/* STEP 4: Result Screen */}
          {step === "result" && extractedQuotation && verificationResult && (
            <ResultScreen
              originalWorkbookName={workbookFile?.name || "Template_Purchase_Order_2026.xlsx"}
              quotation={extractedQuotation}
              verification={verificationResult}
              filledZipBytes={filledZipBytes}
              onStartOver={handleStartOver}
            />
          )}
        </div>
      </main>

      {/* Footer */}
      <footer className="mt-6 text-[11px] text-slate-400">
        Direct XML OpenXML processing · Zero load-and-save alterations
      </footer>
    </div>
  );
}
