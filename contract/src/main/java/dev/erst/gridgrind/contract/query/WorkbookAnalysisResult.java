package dev.erst.gridgrind.contract.query;

import com.fasterxml.jackson.annotation.JsonTypeName;
import dev.erst.gridgrind.contract.dto.AutofilterHealthReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingHealthReport;
import dev.erst.gridgrind.contract.dto.DataValidationHealthReport;
import dev.erst.gridgrind.contract.dto.FormulaHealthReport;
import dev.erst.gridgrind.contract.dto.HyperlinkHealthReport;
import dev.erst.gridgrind.contract.dto.NamedRangeHealthReport;
import dev.erst.gridgrind.contract.dto.PivotTableHealthReport;
import dev.erst.gridgrind.contract.dto.TableHealthReport;
import dev.erst.gridgrind.contract.dto.WorkbookFindingsReport;
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
  record FormulaHealthResult(String stepId, FormulaHealthReport analysis)
      implements WorkbookAnalysisResult {
    public FormulaHealthResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns data-validation-health findings. */
  @JsonTypeName("ANALYZE_DATA_VALIDATION_HEALTH")
  record DataValidationHealthResult(String stepId, DataValidationHealthReport analysis)
      implements WorkbookAnalysisResult {
    public DataValidationHealthResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns conditional-formatting-health findings. */
  @JsonTypeName("ANALYZE_CONDITIONAL_FORMATTING_HEALTH")
  record ConditionalFormattingHealthResult(
      String stepId, ConditionalFormattingHealthReport analysis) implements WorkbookAnalysisResult {
    public ConditionalFormattingHealthResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns autofilter-health findings. */
  @JsonTypeName("ANALYZE_AUTOFILTER_HEALTH")
  record AutofilterHealthResult(String stepId, AutofilterHealthReport analysis)
      implements WorkbookAnalysisResult {
    public AutofilterHealthResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns table-health findings. */
  @JsonTypeName("ANALYZE_TABLE_HEALTH")
  record TableHealthResult(String stepId, TableHealthReport analysis)
      implements WorkbookAnalysisResult {
    public TableHealthResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns pivot-table-health findings. */
  @JsonTypeName("ANALYZE_PIVOT_TABLE_HEALTH")
  record PivotTableHealthResult(String stepId, PivotTableHealthReport analysis)
      implements WorkbookAnalysisResult {
    public PivotTableHealthResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns hyperlink-health findings. */
  @JsonTypeName("ANALYZE_HYPERLINK_HEALTH")
  record HyperlinkHealthResult(String stepId, HyperlinkHealthReport analysis)
      implements WorkbookAnalysisResult {
    public HyperlinkHealthResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns named-range-health findings. */
  @JsonTypeName("ANALYZE_NAMED_RANGE_HEALTH")
  record NamedRangeHealthResult(String stepId, NamedRangeHealthReport analysis)
      implements WorkbookAnalysisResult {
    public NamedRangeHealthResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns aggregated workbook findings. */
  @JsonTypeName("ANALYZE_WORKBOOK_FINDINGS")
  record WorkbookFindingsResult(String stepId, WorkbookFindingsReport analysis)
      implements WorkbookAnalysisResult {
    public WorkbookFindingsResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }
}
