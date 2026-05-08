package dev.erst.gridgrind.excel;

/** Marker for derived factual surface-summary workbook reads. */
public sealed interface WorkbookReadSurfaceResult extends WorkbookReadResult
    permits WorkbookSurfaceResult {}
