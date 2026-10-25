package dev.erst.gridgrind.core.plan;

/** Marker for fact-only inspection results. */
public sealed interface InspectionIntrospectionResult extends InspectionResult
    permits WorkbookAssetInspectionResult, WorkbookInspectionResult, SheetInspectionResult {}
