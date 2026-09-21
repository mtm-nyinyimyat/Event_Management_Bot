/**
 * Event Excel store — public API.
 * Implementation: ./excel/{types,workbookCache,sheetParse,workbookLoad,workbookSearch}
 */
import { clearRagIndex } from "../rag";
import { clearQueryCaches } from "../utils/queryCache";
import {
  clearUploadedWorkbooks as clearUploadedMaps,
  clearWorkbookMemoryCache,
} from "./excel/workbookCache";

export type {
  CachedWorkbook,
  EventRecord,
  SheetData,
  SheetSummary,
  WorkbookLookupOptions,
  WorkbookSearchResult,
} from "./excel/types";

export {
  clearConversationWorkbook,
  getConversationWorkbook,
  getEventsSource,
  getSharedUploadedWorkbook,
  isEventSessionMode,
  resolveExcelPath,
  resolveUploadedWorkbook,
  setConversationWorkbook,
} from "./excel/workbookCache";

export { loadWorkbook, loadWorkbookFromBuffer, sheetsToXlsxBuffer } from "./excel/workbookLoad";

export { searchWorkbook } from "./excel/workbookSearch";

export function clearWorkbookCache(): void {
  clearWorkbookMemoryCache();
  clearRagIndex();
  clearQueryCaches();
}

export function clearUploadedWorkbooks(): void {
  clearUploadedMaps();
}
