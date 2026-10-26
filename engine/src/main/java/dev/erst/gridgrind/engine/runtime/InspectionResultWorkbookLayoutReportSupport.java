package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.drawing.CellCommentReport;
import dev.erst.gridgrind.core.model.layout.ColumnLayoutReport;
import dev.erst.gridgrind.core.model.layout.PaneReport;
import dev.erst.gridgrind.core.model.layout.RowLayoutReport;
import dev.erst.gridgrind.core.model.layout.SheetDefaultsReport;
import dev.erst.gridgrind.core.model.layout.SheetDisplayReport;
import dev.erst.gridgrind.core.model.layout.SheetLayoutReport;
import dev.erst.gridgrind.core.model.layout.SheetOutlineSummaryReport;
import dev.erst.gridgrind.core.model.layout.SheetPresentationReport;
import dev.erst.gridgrind.core.model.print.PrintAreaReport;
import dev.erst.gridgrind.core.model.print.PrintAreaSpec;
import dev.erst.gridgrind.core.model.print.PrintLayoutReport;
import dev.erst.gridgrind.core.model.print.PrintScalingReport;
import dev.erst.gridgrind.core.model.print.PrintScalingSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsReport;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsReport;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec;
import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorReport;
import dev.erst.gridgrind.core.model.value.CellHyperlinkReport;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.excel.ExcelCellReadProjection;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelSheetPresentationSnapshot;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import java.util.ArrayList;
import java.util.List;

/** Converts sheet/window/layout/readback snapshots into protocol report records. */
@SuppressWarnings(
    "PMD.ExcessiveImports") // Corresponding backend and report types define this adapter.
final class InspectionResultWorkbookLayoutReportSupport {
  private InspectionResultWorkbookLayoutReportSupport() {}

  static SheetProtectionReport toSheetProtectionReport(
      dev.erst.gridgrind.excel.WorkbookSheetResult.SheetProtection protection) {
    return switch (protection) {
      case dev.erst.gridgrind.excel.WorkbookSheetResult.SheetProtection.Unprotected _ ->
          new SheetProtectionReport.Unprotected();
      case dev.erst.gridgrind.excel.WorkbookSheetResult.SheetProtection.Protected protectedState ->
          new SheetProtectionReport.Protected(toSheetProtectionSettings(protectedState.settings()));
    };
  }

  static dev.erst.gridgrind.core.model.layout.WindowReport toWindowReport(
      dev.erst.gridgrind.excel.WorkbookSheetResult.Window window,
      ExcelCellReadProjection projection,
      boolean includeBlanks,
      boolean date1904) {
    dev.erst.gridgrind.core.model.layout.WindowDimensionsReport dimensions =
        new dev.erst.gridgrind.core.model.layout.WindowDimensionsReport(
            window.rowCount(), window.columnCount());
    if (includeBlanks) {
      return new dev.erst.gridgrind.core.model.layout.WindowReport.Dense(
          new SheetLabel(window.sheetName()),
          new CellAddress(window.topLeftAddress()),
          dimensions,
          window.rows().stream()
              .map(
                  row ->
                      new dev.erst.gridgrind.core.model.layout.WindowRowReport(
                          row.rowIndex(),
                          row.cells().stream()
                              .map(
                                  cell ->
                                      InspectionResultCellReportSupport.toCellReport(
                                          cell, projection, date1904))
                              .toList()))
              .toList());
    }
    List<dev.erst.gridgrind.core.model.value.CellReport> populatedCells = new ArrayList<>();
    for (dev.erst.gridgrind.excel.WorkbookSheetResult.WindowRow row : window.rows()) {
      for (dev.erst.gridgrind.excel.ExcelCellSnapshot cell : row.cells()) {
        if (cell instanceof dev.erst.gridgrind.excel.ExcelCellSnapshot.BlankSnapshot) {
          continue;
        }
        populatedCells.add(
            InspectionResultCellReportSupport.toCellReport(cell, projection, date1904));
      }
    }
    return new dev.erst.gridgrind.core.model.layout.WindowReport.Sparse(
        new SheetLabel(window.sheetName()),
        new CellAddress(window.topLeftAddress()),
        dimensions,
        List.copyOf(populatedCells));
  }

  static CellHyperlinkReport toCellHyperlinkReport(
      dev.erst.gridgrind.excel.WorkbookSheetResult.CellHyperlink hyperlink) {
    return new CellHyperlinkReport(
        new CellAddress(hyperlink.address()),
        InspectionResultCellReportSupport.toHyperlinkTarget(hyperlink.hyperlink()).orElseThrow());
  }

  static CellCommentReport toCellCommentReport(
      dev.erst.gridgrind.excel.WorkbookSheetResult.CellComment comment) {
    return new CellCommentReport(
        new CellAddress(comment.address()),
        InspectionResultCellReportSupport.toCommentReport(comment.comment()).orElseThrow());
  }

  static SheetLayoutReport toSheetLayoutReport(
      dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayout layout) {
    return new SheetLayoutReport(
        new SheetLabel(layout.sheetName()),
        toPaneReport(layout.pane()),
        layout.zoomPercent(),
        toSheetPresentationReport(layout.presentation()),
        layout.columns().stream()
            .map(
                column ->
                    new ColumnLayoutReport(
                        column.columnIndex(),
                        column.widthCharacters(),
                        column.hidden(),
                        column.outlineLevel(),
                        column.collapsed()))
            .toList(),
        layout.rows().stream()
            .map(
                row ->
                    new RowLayoutReport(
                        row.rowIndex(),
                        row.heightPoints(),
                        row.hidden(),
                        row.outlineLevel(),
                        row.collapsed()))
            .toList());
  }

  static PrintLayoutReport toPrintLayoutReport(
      dev.erst.gridgrind.excel.WorkbookSheetResult.PrintLayoutResult printLayout) {
    return new PrintLayoutReport(
        new SheetLabel(printLayout.sheetName()),
        toPrintAreaReport(printLayout.printLayout().layout().printArea()),
        printLayout.printLayout().layout().orientation(),
        toPrintScalingReport(printLayout.printLayout().layout().scaling()),
        toPrintTitleRowsReport(printLayout.printLayout().layout().repeatingRows()),
        toPrintTitleColumnsReport(printLayout.printLayout().layout().repeatingColumns()),
        printLayout.printLayout().layout().header(),
        printLayout.printLayout().layout().footer(),
        printLayout.printLayout().setup());
  }

  static SheetPresentationReport toSheetPresentationReport(
      ExcelSheetPresentationSnapshot presentation) {
    return new SheetPresentationReport(
        new SheetDisplayReport(
            presentation.display().displayGridlines(),
            presentation.display().displayZeros(),
            presentation.display().displayRowColHeadings(),
            presentation.display().displayFormulas(),
            presentation.display().rightToLeft()),
        presentation.tabColor(),
        new SheetOutlineSummaryReport(
            presentation.outlineSummary().rowSumsBelow(),
            presentation.outlineSummary().rowSumsRight()),
        new SheetDefaultsReport(
            presentation.sheetDefaults().defaultColumnWidth(),
            presentation.sheetDefaults().defaultRowHeightPoints()),
        presentation.ignoredErrors().stream()
            .map(
                ignoredError ->
                    new IgnoredErrorReport(
                        new RangeLabel(ignoredError.range()), ignoredError.errorTypes()))
            .toList());
  }

  static PaneReport toPaneReport(ExcelSheetPane pane) {
    return switch (pane) {
      case ExcelSheetPane.None _ -> new PaneReport.None();
      case ExcelSheetPane.Frozen frozen ->
          new PaneReport.Frozen(
              frozen.splitColumn(), frozen.splitRow(), frozen.leftmostColumn(), frozen.topRow());
      case ExcelSheetPane.Split split ->
          new PaneReport.Split(
              split.xSplitPosition(),
              split.ySplitPosition(),
              split.leftmostColumn(),
              split.topRow(),
              split.activePane());
    };
  }

  static PrintAreaReport toPrintAreaReport(PrintAreaSpec printArea) {
    return switch (printArea) {
      case PrintAreaSpec.None _ -> new PrintAreaReport.None();
      case PrintAreaSpec.Range range ->
          new PrintAreaReport.Range(new RangeLabel(range.range().value()));
    };
  }

  static PrintScalingReport toPrintScalingReport(PrintScalingSpec scaling) {
    return switch (scaling) {
      case PrintScalingSpec.Automatic _ -> new PrintScalingReport.Automatic();
      case PrintScalingSpec.Fit fit ->
          new PrintScalingReport.Fit(fit.widthPages(), fit.heightPages());
    };
  }

  static PrintTitleRowsReport toPrintTitleRowsReport(PrintTitleRowsSpec repeatingRows) {
    return switch (repeatingRows) {
      case PrintTitleRowsSpec.None _ -> new PrintTitleRowsReport.None();
      case PrintTitleRowsSpec.Band band ->
          new PrintTitleRowsReport.Band(band.firstRowIndex(), band.lastRowIndex());
    };
  }

  static PrintTitleColumnsReport toPrintTitleColumnsReport(PrintTitleColumnsSpec repeatingColumns) {
    return switch (repeatingColumns) {
      case PrintTitleColumnsSpec.None _ -> new PrintTitleColumnsReport.None();
      case PrintTitleColumnsSpec.Band band ->
          new PrintTitleColumnsReport.Band(band.firstColumnIndex(), band.lastColumnIndex());
    };
  }

  private static SheetProtectionSettings toSheetProtectionSettings(
      ExcelSheetProtectionSettings settings) {
    return new SheetProtectionSettings(
        settings.autoFilterLocked(),
        settings.deleteColumnsLocked(),
        settings.deleteRowsLocked(),
        settings.formatCellsLocked(),
        settings.formatColumnsLocked(),
        settings.formatRowsLocked(),
        settings.insertColumnsLocked(),
        settings.insertHyperlinksLocked(),
        settings.insertRowsLocked(),
        settings.objectsLocked(),
        settings.pivotTablesLocked(),
        settings.scenariosLocked(),
        settings.selectLockedCellsLocked(),
        settings.selectUnlockedCellsLocked(),
        settings.sortLocked());
  }
}
