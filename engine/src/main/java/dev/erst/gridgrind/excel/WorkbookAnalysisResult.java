package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.WorkbookResultSupport.requireNonBlank;

import java.util.Objects;

/** Conclusion-bearing workbook analysis results. */
public sealed interface WorkbookAnalysisResult extends WorkbookReadAnalysisResult
    permits WorkbookAnalysisResult.FormulaHealthResult,
        WorkbookAnalysisResult.DataValidationHealthResult,
        WorkbookAnalysisResult.ConditionalFormattingHealthResult,
        WorkbookAnalysisResult.AutofilterHealthResult,
        WorkbookAnalysisResult.TableHealthResult,
        WorkbookAnalysisResult.PivotTableHealthResult,
        WorkbookAnalysisResult.HyperlinkHealthResult,
        WorkbookAnalysisResult.NamedRangeHealthResult,
        WorkbookAnalysisResult.WorkbookFindingsResult {

  /** Returns formula-health findings for one analysis read. */
  record FormulaHealthResult(String stepId, WorkbookAnalysis.FormulaHealth analysis)
      implements WorkbookAnalysisResult {
    public FormulaHealthResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns data-validation-health findings for one analysis read. */
  record DataValidationHealthResult(String stepId, WorkbookAnalysis.DataValidationHealth analysis)
      implements WorkbookAnalysisResult {
    public DataValidationHealthResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns conditional-formatting-health findings for one analysis read. */
  record ConditionalFormattingHealthResult(
      String stepId, WorkbookAnalysis.ConditionalFormattingHealth analysis)
      implements WorkbookAnalysisResult {
    public ConditionalFormattingHealthResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns autofilter-health findings for one analysis read. */
  record AutofilterHealthResult(String stepId, WorkbookAnalysis.AutofilterHealth analysis)
      implements WorkbookAnalysisResult {
    public AutofilterHealthResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns table-health findings for one analysis read. */
  record TableHealthResult(String stepId, WorkbookAnalysis.TableHealth analysis)
      implements WorkbookAnalysisResult {
    public TableHealthResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns pivot-table-health findings for one analysis read. */
  record PivotTableHealthResult(String stepId, WorkbookAnalysis.PivotTableHealth analysis)
      implements WorkbookAnalysisResult {
    public PivotTableHealthResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns hyperlink-health findings for one analysis read. */
  record HyperlinkHealthResult(String stepId, WorkbookAnalysis.HyperlinkHealth analysis)
      implements WorkbookAnalysisResult {
    public HyperlinkHealthResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns named-range-health findings for one analysis read. */
  record NamedRangeHealthResult(String stepId, WorkbookAnalysis.NamedRangeHealth analysis)
      implements WorkbookAnalysisResult {
    public NamedRangeHealthResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns the aggregated workbook findings from the first analysis family. */
  record WorkbookFindingsResult(String stepId, WorkbookAnalysis.WorkbookFindings analysis)
      implements WorkbookAnalysisResult {
    public WorkbookFindingsResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }
}
