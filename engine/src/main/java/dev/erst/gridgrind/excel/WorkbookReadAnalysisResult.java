package dev.erst.gridgrind.excel;

/** Marker for derived workbook analysis results. */
public sealed interface WorkbookReadAnalysisResult extends WorkbookReadResult
    permits WorkbookAnalysisResult {}
