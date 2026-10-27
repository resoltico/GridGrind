package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Variant;
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

/** Workbook health-analysis results. */
public sealed interface WorkbookAnalysisResult extends InspectionAnalysisResult
    permits WorkbookAnalysisResult.FormulaHealthResult,
        WorkbookAnalysisResult.DataValidationHealthResult,
        WorkbookAnalysisResult.ConditionalFormattingHealthResult,
        WorkbookAnalysisResult.AutofilterHealthResult,
        WorkbookAnalysisResult.TableHealthResult,
        WorkbookAnalysisResult.PivotTableHealthResult,
        WorkbookAnalysisResult.HyperlinkHealthResult,
        WorkbookAnalysisResult.NamedRangeHealthResult,
        WorkbookAnalysisResult.WorkbookFindingsResult {

  /** Returns formula-health findings. */
  @Variant(id = "ANALYZE_FORMULA_HEALTH", summary = "Returns formula-health findings.")
  record FormulaHealthResult(StepId stepId, FormulaHealthReport analysis)
      implements WorkbookAnalysisResult {
    public FormulaHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns data-validation-health findings. */
  @Variant(
      id = "ANALYZE_DATA_VALIDATION_HEALTH",
      summary = "Returns data-validation-health findings.")
  record DataValidationHealthResult(StepId stepId, DataValidationHealthReport analysis)
      implements WorkbookAnalysisResult {
    public DataValidationHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns conditional-formatting-health findings. */
  @Variant(
      id = "ANALYZE_CONDITIONAL_FORMATTING_HEALTH",
      summary = "Returns conditional-formatting-health findings.")
  record ConditionalFormattingHealthResult(
      StepId stepId, ConditionalFormattingHealthReport analysis) implements WorkbookAnalysisResult {
    public ConditionalFormattingHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns autofilter-health findings. */
  @Variant(id = "ANALYZE_AUTOFILTER_HEALTH", summary = "Returns autofilter-health findings.")
  record AutofilterHealthResult(StepId stepId, AutofilterHealthReport analysis)
      implements WorkbookAnalysisResult {
    public AutofilterHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns table-health findings. */
  @Variant(id = "ANALYZE_TABLE_HEALTH", summary = "Returns table-health findings.")
  record TableHealthResult(StepId stepId, TableHealthReport analysis)
      implements WorkbookAnalysisResult {
    public TableHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns pivot-table-health findings. */
  @Variant(id = "ANALYZE_PIVOT_TABLE_HEALTH", summary = "Returns pivot-table-health findings.")
  record PivotTableHealthResult(StepId stepId, PivotTableHealthReport analysis)
      implements WorkbookAnalysisResult {
    public PivotTableHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns hyperlink-health findings. */
  @Variant(id = "ANALYZE_HYPERLINK_HEALTH", summary = "Returns hyperlink-health findings.")
  record HyperlinkHealthResult(StepId stepId, HyperlinkHealthReport analysis)
      implements WorkbookAnalysisResult {
    public HyperlinkHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns named-range-health findings. */
  @Variant(id = "ANALYZE_NAMED_RANGE_HEALTH", summary = "Returns named-range-health findings.")
  record NamedRangeHealthResult(StepId stepId, NamedRangeHealthReport analysis)
      implements WorkbookAnalysisResult {
    public NamedRangeHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns aggregated workbook findings. */
  @Variant(id = "ANALYZE_WORKBOOK_FINDINGS", summary = "Returns aggregated workbook findings.")
  record WorkbookFindingsResult(StepId stepId, WorkbookFindingsReport analysis)
      implements WorkbookAnalysisResult {
    public WorkbookFindingsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }
}
