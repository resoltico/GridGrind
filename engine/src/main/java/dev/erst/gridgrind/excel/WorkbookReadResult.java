package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.vocabulary.StepId;

/** Immutable workbook-core result produced by one read command. */
public sealed interface WorkbookReadResult
    permits WorkbookReadIntrospectionResult, WorkbookReadSurfaceResult, WorkbookReadAnalysisResult {

  /** Stable caller-provided identifier copied from the originating read command. */
  StepId stepId();
}
