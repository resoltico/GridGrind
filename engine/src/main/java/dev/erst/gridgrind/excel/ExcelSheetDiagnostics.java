package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import java.util.List;
import java.util.Objects;

/** Package-local analysis and health checks for one sheet. */
final class ExcelSheetDiagnostics {
  private final ExcelSheet sheet;
  private final ExcelSheetAnalysisSupport analysisSupport;
  private final ExcelSheetMetadataSupport metadataSupport;

  ExcelSheetDiagnostics(
      ExcelSheet sheet,
      ExcelSheetAnalysisSupport analysisSupport,
      ExcelSheetMetadataSupport metadataSupport) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.analysisSupport =
        Objects.requireNonNull(analysisSupport, "analysisSupport must not be null");
    this.metadataSupport =
        Objects.requireNonNull(metadataSupport, "metadataSupport must not be null");
  }

  int formulaCellCount() {
    return analysisSupport.formulaCellCount();
  }

  int conditionalFormattingBlockCount() {
    return metadataSupport.conditionalFormattingBlockCount();
  }

  List<AnalysisFindingReport> formulaHealthFindings() {
    return analysisSupport.formulaHealthFindings(sheet.name());
  }

  List<AnalysisFindingReport> conditionalFormattingHealthFindings() {
    return metadataSupport.conditionalFormattingHealthFindings(sheet.name());
  }

  int hyperlinkCount() {
    return analysisSupport.hyperlinkCount();
  }

  List<AnalysisFindingReport> hyperlinkHealthFindings(WorkbookLocation workbookLocation) {
    return analysisSupport.hyperlinkHealthFindings(workbookLocation);
  }

  List<AnalysisFindingReport> hyperlinkHealthFindings() {
    return analysisSupport.hyperlinkHealthFindings();
  }
}
