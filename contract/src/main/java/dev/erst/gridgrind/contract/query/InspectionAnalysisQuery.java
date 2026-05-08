package dev.erst.gridgrind.contract.query;

import dev.erst.gridgrind.contract.catalog.GridGrindContractText;
import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadata;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;

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

  @ProtocolTypeMetadata(
      id = "ANALYZE_FORMULA_HEALTH",
      summary = GridGrindContractText.FORMULA_HEALTH_READ_SUMMARY,
      targetSelectors = {SheetSelector.class})
  record AnalyzeFormulaHealth() implements InspectionAnalysisQuery {}

  @ProtocolTypeMetadata(
      id = "ANALYZE_DATA_VALIDATION_HEALTH",
      summary = "Return data-validation-health findings for the selected sheets.",
      targetSelectors = {SheetSelector.class})
  record AnalyzeDataValidationHealth() implements InspectionAnalysisQuery {}

  @ProtocolTypeMetadata(
      id = "ANALYZE_CONDITIONAL_FORMATTING_HEALTH",
      summary = "Return conditional-formatting-health findings for the selected sheets.",
      targetSelectors = {SheetSelector.class})
  record AnalyzeConditionalFormattingHealth() implements InspectionAnalysisQuery {}

  @ProtocolTypeMetadata(
      id = "ANALYZE_AUTOFILTER_HEALTH",
      summary = "Return autofilter-health findings for the selected sheets.",
      targetSelectors = {SheetSelector.class})
  record AnalyzeAutofilterHealth() implements InspectionAnalysisQuery {}

  @ProtocolTypeMetadata(
      id = "ANALYZE_TABLE_HEALTH",
      summary = "Return table-health findings for the selected tables.",
      targetSelectors = {TableSelector.class})
  record AnalyzeTableHealth() implements InspectionAnalysisQuery {}

  @ProtocolTypeMetadata(
      id = "ANALYZE_PIVOT_TABLE_HEALTH",
      summary = "Return pivot-table-health findings for the selected pivot tables.",
      targetSelectors = {PivotTableSelector.class})
  record AnalyzePivotTableHealth() implements InspectionAnalysisQuery {}

  @ProtocolTypeMetadata(
      id = "ANALYZE_HYPERLINK_HEALTH",
      summary = "Return hyperlink-health findings for the selected sheets.",
      targetSelectors = {SheetSelector.class})
  record AnalyzeHyperlinkHealth() implements InspectionAnalysisQuery {}

  @ProtocolTypeMetadata(
      id = "ANALYZE_NAMED_RANGE_HEALTH",
      summary = GridGrindContractText.NAMED_RANGE_HEALTH_READ_SUMMARY,
      targetSelectors = {NamedRangeSelector.class})
  record AnalyzeNamedRangeHealth() implements InspectionAnalysisQuery {}

  @ProtocolTypeMetadata(
      id = "ANALYZE_WORKBOOK_FINDINGS",
      summary = GridGrindContractText.WORKBOOK_FINDINGS_READ_SUMMARY,
      targetSelectors = {WorkbookSelector.class})
  record AnalyzeWorkbookFindings() implements InspectionAnalysisQuery {}
}
