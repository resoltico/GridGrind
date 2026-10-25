package dev.erst.gridgrind.core.plan;

/** Marker for derived factual surface-summary inspection results. */
public sealed interface InspectionSurfaceResult extends InspectionResult
    permits WorkbookSurfaceInspectionResult {}
