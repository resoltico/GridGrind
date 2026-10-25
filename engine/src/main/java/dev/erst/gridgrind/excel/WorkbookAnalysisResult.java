package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.vocabulary.StepId;
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
  record FormulaHealthResult(StepId stepId, WorkbookAnalysis.FormulaHealth analysis)
      implements WorkbookAnalysisResult {
    public FormulaHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns data-validation-health findings for one analysis read. */
  record DataValidationHealthResult(StepId stepId, WorkbookAnalysis.DataValidationHealth analysis)
      implements WorkbookAnalysisResult {
    public DataValidationHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns conditional-formatting-health findings for one analysis read. */
  record ConditionalFormattingHealthResult(
      StepId stepId, WorkbookAnalysis.ConditionalFormattingHealth analysis)
      implements WorkbookAnalysisResult {
    public ConditionalFormattingHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns autofilter-health findings for one analysis read. */
  record AutofilterHealthResult(StepId stepId, WorkbookAnalysis.AutofilterHealth analysis)
      implements WorkbookAnalysisResult {
    public AutofilterHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns table-health findings for one analysis read. */
  record TableHealthResult(StepId stepId, WorkbookAnalysis.TableHealth analysis)
      implements WorkbookAnalysisResult {
    public TableHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns pivot-table-health findings for one analysis read. */
  record PivotTableHealthResult(StepId stepId, WorkbookAnalysis.PivotTableHealth analysis)
      implements WorkbookAnalysisResult {
    public PivotTableHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns hyperlink-health findings for one analysis read. */
  record HyperlinkHealthResult(StepId stepId, WorkbookAnalysis.HyperlinkHealth analysis)
      implements WorkbookAnalysisResult {
    public HyperlinkHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns named-range-health findings for one analysis read. */
  record NamedRangeHealthResult(StepId stepId, WorkbookAnalysis.NamedRangeHealth analysis)
      implements WorkbookAnalysisResult {
    public NamedRangeHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns the aggregated workbook findings from the first analysis family. */
  record WorkbookFindingsResult(StepId stepId, WorkbookAnalysis.WorkbookFindings analysis)
      implements WorkbookAnalysisResult {
    public WorkbookFindingsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }
}
