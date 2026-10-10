import React, { useRef, useState } from "react";
import { FileUp, FileCheck2, X } from "lucide-react";

interface FileUploadProps {
  label: string;
  sublabel: string;
  accept: string;
  selectedFile: File | null;
  onFileSelect: (file: File) => void;
  onClear: () => void;
  errorMessage?: string;
  disabled?: boolean;
}

export const FileUpload: React.FC<FileUploadProps> = ({
  label,
  sublabel,
  accept,
  selectedFile,
  onFileSelect,
  onClear,
  errorMessage,
  disabled = false,
}) => {
  const [isDragging, setIsDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    if (!disabled) setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (disabled) return;
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      onFileSelect(file);
    }
  };

  const formatSize = (bytes: number) => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  return (
    <div className="space-y-1.5">
      <div className="flex items-center justify-between text-xs">
        <label className="font-semibold text-slate-800">{label}</label>
        <span className="text-slate-400">{sublabel}</span>
      </div>

      {!selectedFile ? (
        <div
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          onClick={() => !disabled && inputRef.current?.click()}
          className={`flex flex-col items-center justify-center p-4 border rounded-lg cursor-pointer transition-colors text-center ${
            isDragging
              ? "border-slate-900 bg-slate-50"
              : errorMessage
              ? "border-rose-300 bg-rose-50/40 hover:bg-rose-50/70"
              : "border-slate-200 bg-slate-50/60 hover:bg-slate-100/60"
          } ${disabled ? "opacity-50 cursor-not-allowed" : ""}`}
        >
          <input
            ref={inputRef}
            type="file"
            accept={accept}
            onChange={(e) => {
              if (e.target.files && e.target.files[0]) {
                onFileSelect(e.target.files[0]);
              }
            }}
            disabled={disabled}
            className="hidden"
          />
          <FileUp className="w-5 h-5 text-slate-400 mb-1.5" />
          <p className="text-xs font-medium text-slate-700">
            Click to upload or drag & drop
          </p>
          <p className="text-[11px] text-slate-400 mt-0.5">{accept.toUpperCase()}</p>
        </div>
      ) : (
        <div className="flex items-center justify-between p-3 border border-slate-200 bg-white rounded-lg">
          <div className="flex items-center space-x-2.5 overflow-hidden">
            <FileCheck2 className="w-4 h-4 text-slate-700 shrink-0" />
            <div className="min-w-0">
              <p className="text-xs font-medium text-slate-900 truncate">
                {selectedFile.name}
              </p>
              <p className="text-[11px] text-slate-400 tabular-nums">
                {formatSize(selectedFile.size)}
              </p>
            </div>
          </div>
          {!disabled && (
            <button
              type="button"
              onClick={onClear}
              className="text-slate-400 hover:text-slate-700 p-1 rounded transition-colors"
              title="Remove file"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          )}
        </div>
      )}

      {errorMessage && (
        <p className="text-[11px] text-rose-600 font-medium">{errorMessage}</p>
      )}
    </div>
  );
};
