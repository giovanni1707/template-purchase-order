import React from "react";
import { CheckCircle2, AlertCircle, Download, ArrowLeft, Info } from "lucide-react";
import type { QuotationExtraction, VerificationResult } from "../types/index.ts";

interface ResultScreenProps {
  originalWorkbookName: string;
  quotation: QuotationExtraction;
  verification: VerificationResult;
  filledZipBytes: Uint8Array | null;
  onStartOver: () => void;
}

export const ResultScreen: React.FC<ResultScreenProps> = ({
  originalWorkbookName,
  quotation,
  verification,
  filledZipBytes,
  onStartOver,
}) => {
  const handleDownload = () => {
    if (!filledZipBytes) return;
    const baseName = originalWorkbookName.replace(/\.xlsx$/i, "");
    const downloadFileName = `${baseName}_filled.xlsx`;

    const blob = new Blob([filledZipBytes as unknown as BlobPart], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = downloadFileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const { passed, checks, totalChecks, passedChecks, summary, notices } = verification;

  const failedChecks = checks.filter((c) => !c.passed);

  return (
    <div className="space-y-4">
      {/* 1. Status Header / Badge */}
      {passed ? (
        <div className="flex items-center gap-2 p-3 bg-emerald-50 border border-emerald-200 rounded-lg text-emerald-900 text-xs">
          <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
          <span className="font-semibold">
            All checks passed ({passedChecks} of {totalChecks})
          </span>
        </div>
      ) : (
        <div className="p-3 bg-rose-50 border border-rose-200 rounded-lg text-rose-900 text-xs space-y-1.5">
          <div className="flex items-center gap-2">
            <AlertCircle className="w-4 h-4 text-rose-600 shrink-0" />
            <span className="font-semibold">
              Verification failed ({failedChecks.length} checks failed)
            </span>
          </div>
          <ul className="list-disc pl-5 space-y-1 text-[11px] text-rose-800">
            {failedChecks.map((f) => (
              <li key={f.id}>
                <strong>{f.title}:</strong> {f.details}
              </li>
            ))}
          </ul>
        </div>
      )}

      {/* 2. One summary line */}
      <div className="p-2.5 bg-slate-50 border border-slate-200 rounded-lg text-xs text-slate-700 font-medium overflow-x-auto whitespace-nowrap">
        <span>Rate: {summary.exchangeRate}</span>
        <span className="mx-2 text-slate-300">·</span>
        <span>Order list: {summary.orderListRows} rows</span>
        <span className="mx-2 text-slate-300">·</span>
        <span>PO: {summary.poRows} rows</span>
        <span className="mx-2 text-slate-300">·</span>
        <span className="tabular-nums">
          Total Exworks: €{summary.totalExworks.toFixed(2)} (Invoice: €{summary.invoiceAmount.toFixed(2)})
        </span>
        <span className="mx-2 text-slate-300">·</span>
        <span className="tabular-nums">PO Total: €{summary.poTotal.toFixed(2)}</span>
      </div>

      {/* 3. Compact scrollable table of extracted items */}
      <div className="space-y-1.5">
        <div className="flex items-center justify-between text-xs">
          <span className="font-semibold text-slate-800">
            Extracted Line Items ({quotation.positions.length})
          </span>
          <span className="text-slate-400">
            Quotation #{quotation.quotationNumber}
          </span>
        </div>

        <div className="border border-slate-200 rounded-lg overflow-hidden">
          <div className="max-h-48 overflow-y-auto">
            <table className="w-full text-[11px] text-left border-collapse">
              <thead className="sticky top-0 bg-slate-100 text-slate-600 border-b border-slate-200">
                <tr>
                  <th className="py-1.5 px-2.5 font-medium">Reference</th>
                  <th className="py-1.5 px-2.5 font-medium">Part Name</th>
                  <th className="py-1.5 px-2.5 font-medium text-right">Net Price</th>
                  <th className="py-1.5 px-2.5 font-medium text-right">Qty</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 bg-white">
                {quotation.positions.map((pos, idx) => (
                  <tr key={idx} className="hover:bg-slate-50/70">
                    <td className="py-1.5 px-2.5 font-mono text-slate-900 font-medium">
                      {pos.reference}
                    </td>
                    <td className="py-1.5 px-2.5 text-slate-700 truncate max-w-[200px]" title={pos.partName}>
                      {pos.partName}
                    </td>
                    <td className="py-1.5 px-2.5 font-mono text-right tabular-nums text-slate-900">
                      €{pos.netPrice.toFixed(2)}
                    </td>
                    <td className="py-1.5 px-2.5 font-mono text-right tabular-nums text-slate-700 whitespace-nowrap">
                      {pos.quantity} {pos.unit}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* 4. Notices box */}
      {notices.length > 0 && (
        <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg space-y-1.5">
          <div className="flex items-center gap-1.5 text-xs font-semibold text-slate-700">
            <Info className="w-3.5 h-3.5 text-slate-500" />
            <span>Notices</span>
          </div>
          <ul className="space-y-1 text-[11px] text-slate-600 leading-relaxed list-disc pl-4">
            {notices.map((n, i) => (
              <li key={i}>{n}</li>
            ))}
          </ul>
        </div>
      )}

      {/* 5. Download Button + Start Over link */}
      <div className="pt-2 flex flex-col items-center gap-2.5">
        {passed && filledZipBytes ? (
          <button
            type="button"
            onClick={handleDownload}
            className="w-full flex items-center justify-center gap-2 py-2.5 px-4 bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white text-xs font-semibold rounded-lg shadow-sm transition-colors"
          >
            <Download className="w-4 h-4" />
            <span>Download filled workbook</span>
          </button>
        ) : (
          <div className="text-xs text-rose-600 font-medium text-center">
            Download disabled due to verification failures.
          </div>
        )}

        <button
          type="button"
          onClick={onStartOver}
          className="inline-flex items-center gap-1.5 text-xs text-slate-500 hover:text-slate-800 transition-colors"
        >
          <ArrowLeft className="w-3.5 h-3.5" />
          <span>Start over</span>
        </button>
      </div>
    </div>
  );
};
