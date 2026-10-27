package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.engine.formulas.FormulaAdapter;
import dev.erst.gridgrind.engine.formulas.FormulaRuntime;
import java.util.List;
import java.util.Objects;

/** Package-local analysis and health checks for one sheet. */
final class SheetDiagnostics {
  private final XssfSheet sheet;
  private final FormulaRuntime formulaRuntime;
  private final SheetMetadata metadataSupport;

  SheetDiagnostics(XssfSheet sheet, FormulaRuntime formulaRuntime, SheetMetadata metadataSupport) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.formulaRuntime = Objects.requireNonNull(formulaRuntime, "formulaRuntime must not be null");
    this.metadataSupport =
        Objects.requireNonNull(metadataSupport, "metadataSupport must not be null");
  }

  int formulaCellCount() {
    return FormulaAdapter.formulaCellCount(sheet.xssfSheet());
  }

  int conditionalFormattingBlockCount() {
    return metadataSupport.conditionalFormattingBlockCount();
  }

  List<AnalysisFindingReport> formulaHealthFindings() {
    return FormulaAdapter.formulaHealthFindings(sheet.name(), sheet.xssfSheet(), formulaRuntime);
  }

  List<AnalysisFindingReport> conditionalFormattingHealthFindings() {
    return metadataSupport.conditionalFormattingHealthFindings(sheet.name());
  }

  int hyperlinkCount() {
    return sheet.links().hyperlinkCount();
  }

  List<AnalysisFindingReport> hyperlinkHealthFindings(WorkbookLocation workbookLocation) {
    return sheet.links().hyperlinkHealthFindings(workbookLocation);
  }

  List<AnalysisFindingReport> hyperlinkHealthFindings() {
    return sheet.links().hyperlinkHealthFindings();
  }
}
