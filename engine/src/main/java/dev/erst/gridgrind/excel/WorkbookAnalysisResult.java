package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.formula.FormulaHealthReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeHealthReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableHealthReport;
import dev.erst.gridgrind.core.model.table.AutofilterHealthReport;
import dev.erst.gridgrind.core.model.table.TableHealthReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingHealthReport;
import dev.erst.gridgrind.core.model.validation.DataValidationHealthReport;
import dev.erst.gridgrind.core.model.value.HyperlinkHealthReport;
import dev.erst.gridgrind.core.result.WorkbookFindingsReport;
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
  record FormulaHealthResult(StepId stepId, FormulaHealthReport analysis)
      implements WorkbookAnalysisResult {
    public FormulaHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns data-validation-health findings for one analysis read. */
  record DataValidationHealthResult(StepId stepId, DataValidationHealthReport analysis)
      implements WorkbookAnalysisResult {
    public DataValidationHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns conditional-formatting-health findings for one analysis read. */
  record ConditionalFormattingHealthResult(
      StepId stepId, ConditionalFormattingHealthReport analysis) implements WorkbookAnalysisResult {
    public ConditionalFormattingHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns autofilter-health findings for one analysis read. */
  record AutofilterHealthResult(StepId stepId, AutofilterHealthReport analysis)
      implements WorkbookAnalysisResult {
    public AutofilterHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns table-health findings for one analysis read. */
  record TableHealthResult(StepId stepId, TableHealthReport analysis)
      implements WorkbookAnalysisResult {
    public TableHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns pivot-table-health findings for one analysis read. */
  record PivotTableHealthResult(StepId stepId, PivotTableHealthReport analysis)
      implements WorkbookAnalysisResult {
    public PivotTableHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns hyperlink-health findings for one analysis read. */
  record HyperlinkHealthResult(StepId stepId, HyperlinkHealthReport analysis)
      implements WorkbookAnalysisResult {
    public HyperlinkHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns named-range-health findings for one analysis read. */
  record NamedRangeHealthResult(StepId stepId, NamedRangeHealthReport analysis)
      implements WorkbookAnalysisResult {
    public NamedRangeHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns the aggregated workbook findings from the first analysis family. */
  record WorkbookFindingsResult(StepId stepId, WorkbookFindingsReport analysis)
      implements WorkbookAnalysisResult {
    public WorkbookFindingsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }
}
