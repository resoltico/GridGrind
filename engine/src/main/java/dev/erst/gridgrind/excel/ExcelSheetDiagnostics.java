package dev.erst.gridgrind.excel;

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

  List<WorkbookAnalysis.AnalysisFinding> formulaHealthFindings() {
    return analysisSupport.formulaHealthFindings(sheet.name());
  }

  List<WorkbookAnalysis.AnalysisFinding> conditionalFormattingHealthFindings() {
    return metadataSupport.conditionalFormattingHealthFindings(sheet.name());
  }

  int hyperlinkCount() {
    return analysisSupport.hyperlinkCount();
  }

  List<WorkbookAnalysis.AnalysisFinding> hyperlinkHealthFindings(
      WorkbookLocation workbookLocation) {
    return analysisSupport.hyperlinkHealthFindings(workbookLocation);
  }

  List<WorkbookAnalysis.AnalysisFinding> hyperlinkHealthFindings() {
    return analysisSupport.hyperlinkHealthFindings();
  }
}
