package dev.erst.gridgrind.contract.query;

/** Marker for derived factual surface-summary inspection results. */
public sealed interface InspectionSurfaceResult extends InspectionResult
    permits WorkbookSurfaceInspectionResult {}
