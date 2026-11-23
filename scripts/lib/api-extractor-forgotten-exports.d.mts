export interface ForgottenExportItem {
  entry: string;
  symbol: string;
}
export interface ForgottenExportPolicy {
  staleAllowlist: ForgottenExportItem[];
  staleBaseline: ForgottenExportItem[];
  unallowlisted: ForgottenExportItem[];
  forbiddenAllowlist: ForgottenExportItem[];
  crossEntryNotPublic: ForgottenExportItem[];
}
export function parseForgottenExportMessage(message: string): ForgottenExportItem | null;
export function evaluateForgottenExportPolicy(options: {
  packageName: string;
  isLocal: boolean;
  allowlist: Record<string, string[]>;
  messages: string[];
  protectedSymbols?: string[];
  baseline?: Record<string, string[]>;
  crossEntry?: Record<string, string[]>;
}): ForgottenExportPolicy;
export function nextForgottenExportBaseline(options: {
  baseline?: Record<string, string[]>;
  policy: ForgottenExportPolicy;
  record?: boolean;
}): Record<string, string[]>;
