package dev.erst.gridgrind.core.plan;

/** Marker for derived workbook analysis results. */
public sealed interface InspectionAnalysisResult extends InspectionResult
    permits WorkbookAnalysisResult {}
