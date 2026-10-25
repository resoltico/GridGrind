package dev.erst.gridgrind.core.plan;

import com.fasterxml.jackson.annotation.JsonTypeName;
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
  @JsonTypeName("ANALYZE_FORMULA_HEALTH")
  record FormulaHealthResult(StepId stepId, FormulaHealthReport analysis)
      implements WorkbookAnalysisResult {
    public FormulaHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns data-validation-health findings. */
  @JsonTypeName("ANALYZE_DATA_VALIDATION_HEALTH")
  record DataValidationHealthResult(StepId stepId, DataValidationHealthReport analysis)
      implements WorkbookAnalysisResult {
    public DataValidationHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns conditional-formatting-health findings. */
  @JsonTypeName("ANALYZE_CONDITIONAL_FORMATTING_HEALTH")
  record ConditionalFormattingHealthResult(
      StepId stepId, ConditionalFormattingHealthReport analysis) implements WorkbookAnalysisResult {
    public ConditionalFormattingHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns autofilter-health findings. */
  @JsonTypeName("ANALYZE_AUTOFILTER_HEALTH")
  record AutofilterHealthResult(StepId stepId, AutofilterHealthReport analysis)
      implements WorkbookAnalysisResult {
    public AutofilterHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns table-health findings. */
  @JsonTypeName("ANALYZE_TABLE_HEALTH")
  record TableHealthResult(StepId stepId, TableHealthReport analysis)
      implements WorkbookAnalysisResult {
    public TableHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns pivot-table-health findings. */
  @JsonTypeName("ANALYZE_PIVOT_TABLE_HEALTH")
  record PivotTableHealthResult(StepId stepId, PivotTableHealthReport analysis)
      implements WorkbookAnalysisResult {
    public PivotTableHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns hyperlink-health findings. */
  @JsonTypeName("ANALYZE_HYPERLINK_HEALTH")
  record HyperlinkHealthResult(StepId stepId, HyperlinkHealthReport analysis)
      implements WorkbookAnalysisResult {
    public HyperlinkHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns named-range-health findings. */
  @JsonTypeName("ANALYZE_NAMED_RANGE_HEALTH")
  record NamedRangeHealthResult(StepId stepId, NamedRangeHealthReport analysis)
      implements WorkbookAnalysisResult {
    public NamedRangeHealthResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns aggregated workbook findings. */
  @JsonTypeName("ANALYZE_WORKBOOK_FINDINGS")
  record WorkbookFindingsResult(StepId stepId, WorkbookFindingsReport analysis)
      implements WorkbookAnalysisResult {
    public WorkbookFindingsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }
}
