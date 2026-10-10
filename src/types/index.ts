export interface QuotationPosition {
  position: string;
  reference: string;
  partName: string;
  netPrice: number;
  listPrice: number;
  quantity: number;
  unit: string;
  lineAmount: number;
}

export interface IgnoredLine {
  lineText: string;
  reason: string;
}

export interface UnclearItem {
  item: string;
  reason: string;
}

export interface QuotationExtraction {
  quotationNumber: string;
  invoiceAmount: number;
  currency: string;
  positions: QuotationPosition[];
  ignoredLines: IgnoredLine[];
  unclear: UnclearItem[];
}

export type PoDetailsOption = 'keep' | 'clear' | 'replace';

export interface VerificationCheck {
  id: string;
  title: string;
  passed: boolean;
  details?: string;
}

export interface VerificationResult {
  passed: boolean;
  checks: VerificationCheck[];
  totalChecks: number;
  passedChecks: number;
  summary: {
    exchangeRate: number;
    orderListRows: number;
    poRows: number;
    totalExworks: number;
    invoiceAmount: number;
    poTotal: number;
  };
  notices: string[];
}
