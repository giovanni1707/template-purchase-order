import React from "react";
import type { PoDetailsOption } from "../types/index.ts";

interface PoDetailsControlProps {
  initialText: string;
  selectedOption: PoDetailsOption;
  customText: string;
  onChangeOption: (opt: PoDetailsOption) => void;
  onChangeCustomText: (text: string) => void;
  disabled?: boolean;
}

export const PoDetailsControl: React.FC<PoDetailsControlProps> = ({
  initialText,
  selectedOption,
  customText,
  onChangeOption,
  onChangeCustomText,
  disabled = false,
}) => {
  return (
    <div className="space-y-2 p-3 bg-slate-50 border border-slate-200 rounded-lg">
      <div className="flex items-baseline justify-between">
        <label className="text-xs font-semibold text-slate-800">
          PO Details text
        </label>
        <span className="text-[11px] text-slate-500 truncate max-w-[220px]" title={initialText}>
          Template value: &quot;{initialText}&quot;
        </span>
      </div>

      <div className="flex items-center gap-4 text-xs">
        <label className="inline-flex items-center gap-1.5 cursor-pointer">
          <input
            type="radio"
            name="poDetails"
            value="clear"
            checked={selectedOption === "clear"}
            onChange={() => onChangeOption("clear")}
            disabled={disabled}
            className="text-slate-900 focus:ring-slate-900"
          />
          <span className="text-slate-700 font-medium">Clear</span>
        </label>

        <label className="inline-flex items-center gap-1.5 cursor-pointer">
          <input
            type="radio"
            name="poDetails"
            value="keep"
            checked={selectedOption === "keep"}
            onChange={() => onChangeOption("keep")}
            disabled={disabled}
            className="text-slate-900 focus:ring-slate-900"
          />
          <span className="text-slate-700 font-medium">Keep</span>
        </label>

        <label className="inline-flex items-center gap-1.5 cursor-pointer">
          <input
            type="radio"
            name="poDetails"
            value="replace"
            checked={selectedOption === "replace"}
            onChange={() => onChangeOption("replace")}
            disabled={disabled}
            className="text-slate-900 focus:ring-slate-900"
          />
          <span className="text-slate-700 font-medium">Replace with my text</span>
        </label>
      </div>

      {selectedOption === "replace" && (
        <div className="pt-1">
          <input
            type="text"
            value={customText}
            onChange={(e) => onChangeCustomText(e.target.value)}
            disabled={disabled}
            placeholder="Enter custom PO details text..."
            className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-md text-slate-900 placeholder:text-slate-400 focus:outline-none focus:ring-1 focus:ring-slate-900"
          />
        </div>
      )}
    </div>
  );
};
