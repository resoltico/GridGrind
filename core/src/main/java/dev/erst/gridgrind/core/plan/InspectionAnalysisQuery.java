package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.catalog.GridGrindInspectionContractText;
import dev.erst.gridgrind.core.metadata.Variant;

/** Workbook health-analysis queries. */
public sealed interface InspectionAnalysisQuery extends InspectionQuery.Analysis
    permits InspectionAnalysisQuery.AnalyzeFormulaHealth,
        InspectionAnalysisQuery.AnalyzeDataValidationHealth,
        InspectionAnalysisQuery.AnalyzeConditionalFormattingHealth,
        InspectionAnalysisQuery.AnalyzeAutofilterHealth,
        InspectionAnalysisQuery.AnalyzeTableHealth,
        InspectionAnalysisQuery.AnalyzePivotTableHealth,
        InspectionAnalysisQuery.AnalyzeHyperlinkHealth,
        InspectionAnalysisQuery.AnalyzeNamedRangeHealth,
        InspectionAnalysisQuery.AnalyzeWorkbookFindings {

  @Variant(
      id = "ANALYZE_FORMULA_HEALTH",
      summary = GridGrindInspectionContractText.FORMULA_HEALTH_READ_SUMMARY,
      targetSelectors = {SheetSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record AnalyzeFormulaHealth() implements InspectionAnalysisQuery {}

  @Variant(
      id = "ANALYZE_DATA_VALIDATION_HEALTH",
      summary = "Return data-validation-health findings for the selected sheets.",
      targetSelectors = {SheetSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record AnalyzeDataValidationHealth() implements InspectionAnalysisQuery {}

  @Variant(
      id = "ANALYZE_CONDITIONAL_FORMATTING_HEALTH",
      summary = "Return conditional-formatting-health findings for the selected sheets.",
      targetSelectors = {SheetSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record AnalyzeConditionalFormattingHealth() implements InspectionAnalysisQuery {}

  @Variant(
      id = "ANALYZE_AUTOFILTER_HEALTH",
      summary = "Return autofilter-health findings for the selected sheets.",
      targetSelectors = {SheetSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record AnalyzeAutofilterHealth() implements InspectionAnalysisQuery {}

  @Variant(
      id = "ANALYZE_TABLE_HEALTH",
      summary = "Return table-health findings for the selected tables.",
      targetSelectors = {TableSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record AnalyzeTableHealth() implements InspectionAnalysisQuery {}

  @Variant(
      id = "ANALYZE_PIVOT_TABLE_HEALTH",
      summary = "Return pivot-table-health findings for the selected pivot tables.",
      targetSelectors = {PivotTableSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record AnalyzePivotTableHealth() implements InspectionAnalysisQuery {}

  @Variant(
      id = "ANALYZE_HYPERLINK_HEALTH",
      summary = "Return hyperlink-health findings for the selected sheets.",
      targetSelectors = {SheetSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record AnalyzeHyperlinkHealth() implements InspectionAnalysisQuery {}

  @Variant(
      id = "ANALYZE_NAMED_RANGE_HEALTH",
      summary = GridGrindInspectionContractText.NAMED_RANGE_HEALTH_READ_SUMMARY,
      targetSelectors = {NamedRangeSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record AnalyzeNamedRangeHealth() implements InspectionAnalysisQuery {}

  @Variant(
      id = "ANALYZE_WORKBOOK_FINDINGS",
      summary = GridGrindInspectionContractText.WORKBOOK_FINDINGS_READ_SUMMARY,
      targetSelectors = {WorkbookSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record AnalyzeWorkbookFindings() implements InspectionAnalysisQuery {}
}
