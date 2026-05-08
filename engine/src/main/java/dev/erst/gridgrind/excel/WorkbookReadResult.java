package dev.erst.gridgrind.excel;

/** Immutable workbook-core result produced by one read command. */
public sealed interface WorkbookReadResult
    permits WorkbookReadIntrospectionResult, WorkbookReadSurfaceResult, WorkbookReadAnalysisResult {

  /** Stable caller-provided identifier copied from the originating read command. */
  String stepId();
}
