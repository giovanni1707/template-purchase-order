import React from "react";
import { AlertTriangle, ArrowLeft } from "lucide-react";
import type { UnclearItem } from "../types/index.ts";

interface UncertaintiesScreenProps {
  uncertainties: UnclearItem[];
  onStartOver: () => void;
}

export const UncertaintiesScreen: React.FC<UncertaintiesScreenProps> = ({
  uncertainties,
  onStartOver,
}) => {
  return (
    <div className="space-y-5">
      <div className="p-4 border border-amber-300 bg-amber-50/80 rounded-lg">
        <div className="flex items-start gap-3">
          <AlertTriangle className="w-5 h-5 text-amber-700 shrink-0 mt-0.5" />
          <div className="space-y-1">
            <h2 className="text-sm font-semibold text-amber-900">
              Uncertain values detected
            </h2>
            <p className="text-xs text-amber-800">
              Extraction was halted because unambiguous values could not be mathematically or textually verified. No file was generated.
            </p>
          </div>
        </div>
      </div>

      <div className="space-y-2">
        <h3 className="text-xs font-semibold text-slate-800">
          Detected Inconsistencies ({uncertainties.length})
        </h3>
        <div className="border border-slate-200 rounded-lg divide-y divide-slate-100 max-h-64 overflow-y-auto bg-white">
          {uncertainties.map((item, idx) => (
            <div key={idx} className="p-3 text-xs space-y-1">
              <span className="font-semibold text-slate-900">{item.item}</span>
              <p className="text-slate-600 leading-relaxed">{item.reason}</p>
            </div>
          ))}
        </div>
      </div>

      <div className="pt-2 text-center">
        <button
          type="button"
          onClick={onStartOver}
          className="inline-flex items-center gap-1.5 text-xs font-medium text-slate-600 hover:text-slate-900 transition-colors"
        >
          <ArrowLeft className="w-3.5 h-3.5" />
          <span>Start over</span>
        </button>
      </div>
    </div>
  );
};
