package dev.erst.gridgrind.contract.query;

/** Marker for derived workbook analysis results. */
public sealed interface InspectionAnalysisResult extends InspectionResult
    permits WorkbookAnalysisResult {}
