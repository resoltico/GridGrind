package dev.erst.gridgrind.excel;

/** Marker for fact-only workbook reads. */
public sealed interface WorkbookReadIntrospectionResult extends WorkbookReadResult
    permits WorkbookCoreResult,
        WorkbookSheetResult,
        WorkbookDrawingResult,
        WorkbookRuleResult,
        WorkbookSurfaceResult {}
