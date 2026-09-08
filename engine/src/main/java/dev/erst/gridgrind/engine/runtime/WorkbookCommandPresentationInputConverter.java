package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.SheetPresentationInput;
import dev.erst.gridgrind.excel.ExcelIgnoredError;
import dev.erst.gridgrind.excel.ExcelSheetDefaults;
import dev.erst.gridgrind.excel.ExcelSheetDisplay;
import dev.erst.gridgrind.excel.ExcelSheetOutlineSummary;
import dev.erst.gridgrind.excel.ExcelSheetPresentation;

/** Converts the sheet-display and error-presentation portion of a workbook command. */
final class WorkbookCommandPresentationInputConverter {
  private WorkbookCommandPresentationInputConverter() {}

  static ExcelSheetPresentation toExcelSheetPresentation(SheetPresentationInput presentation) {
    return new ExcelSheetPresentation(
        new ExcelSheetDisplay(
            presentation.display().displayGridlines(),
            presentation.display().displayZeros(),
            presentation.display().displayRowColHeadings(),
            presentation.display().displayFormulas(),
            presentation.display().rightToLeft()),
        presentation.tabColor().flatMap(WorkbookCommandCellInputConverter::toExcelColor),
        new ExcelSheetOutlineSummary(
            presentation.outlineSummary().rowSumsBelow(),
            presentation.outlineSummary().rowSumsRight()),
        new ExcelSheetDefaults(
            presentation.sheetDefaults().defaultColumnWidth(),
            presentation.sheetDefaults().defaultRowHeightPoints()),
        presentation.ignoredErrors().stream()
            .map(
                ignoredError ->
                    new ExcelIgnoredError(ignoredError.range(), ignoredError.errorTypes()))
            .toList());
  }
}
