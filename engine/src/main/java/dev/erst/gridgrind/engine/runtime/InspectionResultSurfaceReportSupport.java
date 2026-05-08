package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.FormulaPatternReport;
import dev.erst.gridgrind.contract.dto.FormulaSurfaceReport;
import dev.erst.gridgrind.contract.dto.NamedRangeBackingKind;
import dev.erst.gridgrind.contract.dto.NamedRangeSurfaceEntryReport;
import dev.erst.gridgrind.contract.dto.NamedRangeSurfaceReport;
import dev.erst.gridgrind.contract.dto.SchemaColumnReport;
import dev.erst.gridgrind.contract.dto.SheetFormulaSurfaceReport;
import dev.erst.gridgrind.contract.dto.SheetSchemaReport;
import dev.erst.gridgrind.contract.dto.TypeCountReport;

/** Converts derived factual surface snapshots into protocol surface reports. */
final class InspectionResultSurfaceReportSupport {
  private InspectionResultSurfaceReportSupport() {}

  static FormulaSurfaceReport toFormulaSurfaceReport(
      dev.erst.gridgrind.excel.WorkbookSurfaceResult.FormulaSurface surface) {
    return new FormulaSurfaceReport(
        surface.totalFormulaCellCount(),
        surface.sheets().stream()
            .map(
                sheet ->
                    new SheetFormulaSurfaceReport(
                        sheet.sheetName(),
                        sheet.formulaCellCount(),
                        sheet.distinctFormulaCount(),
                        sheet.formulas().stream()
                            .map(
                                formula ->
                                    new FormulaPatternReport(
                                        formula.formula(),
                                        formula.occurrenceCount(),
                                        formula.addresses()))
                            .toList()))
            .toList());
  }

  static SheetSchemaReport toSheetSchemaReport(
      dev.erst.gridgrind.excel.WorkbookSurfaceResult.SheetSchema surface) {
    return new SheetSchemaReport(
        surface.sheetName(),
        surface.topLeftAddress(),
        surface.rowCount(),
        surface.columnCount(),
        surface.dataRowCount(),
        surface.columns().stream()
            .map(
                column ->
                    new SchemaColumnReport(
                        column.columnIndex(),
                        column.columnAddress(),
                        column.headerDisplayValue(),
                        column.populatedCellCount(),
                        column.blankCellCount(),
                        column.observedTypes().stream()
                            .map(
                                typeCount ->
                                    new TypeCountReport(typeCount.type(), typeCount.count()))
                            .toList(),
                        column.dominantType()))
            .toList());
  }

  static NamedRangeSurfaceReport toNamedRangeSurfaceReport(
      dev.erst.gridgrind.excel.WorkbookSurfaceResult.NamedRangeSurface surface) {
    return new NamedRangeSurfaceReport(
        surface.workbookScopedCount(),
        surface.sheetScopedCount(),
        surface.rangeBackedCount(),
        surface.formulaBackedCount(),
        surface.namedRanges().stream()
            .map(
                entry ->
                    new NamedRangeSurfaceEntryReport(
                        entry.name(),
                        InspectionResultWorkbookCoreReportSupport.toNamedRangeScope(entry.scope()),
                        entry.refersToFormula(),
                        switch (entry.kind()) {
                          case RANGE -> NamedRangeBackingKind.RANGE;
                          case FORMULA -> NamedRangeBackingKind.FORMULA;
                        }))
            .toList());
  }
}
