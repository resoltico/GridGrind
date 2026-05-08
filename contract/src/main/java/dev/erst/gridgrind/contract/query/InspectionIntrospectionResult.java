package dev.erst.gridgrind.contract.query;

/** Marker for fact-only inspection results. */
public sealed interface InspectionIntrospectionResult extends InspectionResult
    permits WorkbookAssetInspectionResult, WorkbookInspectionResult, SheetInspectionResult {}
