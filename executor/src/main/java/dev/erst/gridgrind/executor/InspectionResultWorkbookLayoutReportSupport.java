package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.GridGrindLayoutSurfaceReports;
import dev.erst.gridgrind.contract.dto.GridGrindWorkbookSurfaceReports;
import dev.erst.gridgrind.contract.dto.HeaderFooterTextReport;
import dev.erst.gridgrind.contract.dto.IgnoredErrorReport;
import dev.erst.gridgrind.contract.dto.PaneReport;
import dev.erst.gridgrind.contract.dto.PrintAreaReport;
import dev.erst.gridgrind.contract.dto.PrintLayoutReport;
import dev.erst.gridgrind.contract.dto.PrintMarginsReport;
import dev.erst.gridgrind.contract.dto.PrintScalingReport;
import dev.erst.gridgrind.contract.dto.PrintSetupReport;
import dev.erst.gridgrind.contract.dto.PrintTitleColumnsReport;
import dev.erst.gridgrind.contract.dto.PrintTitleRowsReport;
import dev.erst.gridgrind.contract.dto.SheetDefaultsReport;
import dev.erst.gridgrind.contract.dto.SheetDisplayReport;
import dev.erst.gridgrind.contract.dto.SheetOutlineSummaryReport;
import dev.erst.gridgrind.contract.dto.SheetPresentationReport;
import dev.erst.gridgrind.contract.dto.SheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelHeaderFooterText;
import dev.erst.gridgrind.excel.ExcelPrintLayout;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelSheetPresentationSnapshot;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;

/** Converts sheet/window/layout/readback snapshots into protocol report records. */
final class InspectionResultWorkbookLayoutReportSupport {
  private InspectionResultWorkbookLayoutReportSupport() {}

  static GridGrindWorkbookSurfaceReports.SheetProtectionReport toSheetProtectionReport(
      dev.erst.gridgrind.excel.WorkbookSheetResult.SheetProtection protection) {
    return switch (protection) {
      case dev.erst.gridgrind.excel.WorkbookSheetResult.SheetProtection.Unprotected _ ->
          new GridGrindWorkbookSurfaceReports.SheetProtectionReport.Unprotected();
      case dev.erst.gridgrind.excel.WorkbookSheetResult.SheetProtection.Protected protectedState ->
          new GridGrindWorkbookSurfaceReports.SheetProtectionReport.Protected(
              toSheetProtectionSettings(protectedState.settings()));
    };
  }

  static GridGrindLayoutSurfaceReports.WindowReport toWindowReport(
      dev.erst.gridgrind.excel.WorkbookSheetResult.Window window) {
    return new GridGrindLayoutSurfaceReports.WindowReport(
        window.sheetName(),
        window.topLeftAddress(),
        window.rowCount(),
        window.columnCount(),
        window.rows().stream()
            .map(
                row ->
                    new GridGrindLayoutSurfaceReports.WindowRowReport(
                        row.rowIndex(),
                        row.cells().stream()
                            .map(InspectionResultCellReportSupport::toCellReport)
                            .toList()))
            .toList());
  }

  static GridGrindLayoutSurfaceReports.CellHyperlinkReport toCellHyperlinkReport(
      dev.erst.gridgrind.excel.WorkbookSheetResult.CellHyperlink hyperlink) {
    return new GridGrindLayoutSurfaceReports.CellHyperlinkReport(
        hyperlink.address(),
        InspectionResultCellReportSupport.toHyperlinkTarget(hyperlink.hyperlink()).orElse(null));
  }

  static GridGrindLayoutSurfaceReports.CellCommentReport toCellCommentReport(
      dev.erst.gridgrind.excel.WorkbookSheetResult.CellComment comment) {
    return new GridGrindLayoutSurfaceReports.CellCommentReport(
        comment.address(),
        InspectionResultCellReportSupport.toCommentReport(comment.comment()).orElse(null));
  }

  static GridGrindLayoutSurfaceReports.SheetLayoutReport toSheetLayoutReport(
      dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayout layout) {
    return new GridGrindLayoutSurfaceReports.SheetLayoutReport(
        layout.sheetName(),
        toPaneReport(layout.pane()),
        layout.zoomPercent(),
        toSheetPresentationReport(layout.presentation()),
        layout.columns().stream()
            .map(
                column ->
                    new GridGrindLayoutSurfaceReports.ColumnLayoutReport(
                        column.columnIndex(),
                        column.widthCharacters(),
                        column.hidden(),
                        column.outlineLevel(),
                        column.collapsed()))
            .toList(),
        layout.rows().stream()
            .map(
                row ->
                    new GridGrindLayoutSurfaceReports.RowLayoutReport(
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
        printLayout.sheetName(),
        toPrintAreaReport(printLayout.printLayout().layout().printArea()),
        printLayout.printLayout().layout().orientation(),
        toPrintScalingReport(printLayout.printLayout().layout().scaling()),
        toPrintTitleRowsReport(printLayout.printLayout().layout().repeatingRows()),
        toPrintTitleColumnsReport(printLayout.printLayout().layout().repeatingColumns()),
        toHeaderFooterTextReport(printLayout.printLayout().layout().header()),
        toHeaderFooterTextReport(printLayout.printLayout().layout().footer()),
        new PrintSetupReport(
            new PrintMarginsReport(
                printLayout.printLayout().setup().margins().left(),
                printLayout.printLayout().setup().margins().right(),
                printLayout.printLayout().setup().margins().top(),
                printLayout.printLayout().setup().margins().bottom(),
                printLayout.printLayout().setup().margins().header(),
                printLayout.printLayout().setup().margins().footer()),
            printLayout.printLayout().setup().printGridlines(),
            printLayout.printLayout().setup().horizontallyCentered(),
            printLayout.printLayout().setup().verticallyCentered(),
            printLayout.printLayout().setup().paperSize(),
            printLayout.printLayout().setup().draft(),
            printLayout.printLayout().setup().blackAndWhite(),
            printLayout.printLayout().setup().copies(),
            printLayout.printLayout().setup().useFirstPageNumber(),
            printLayout.printLayout().setup().firstPageNumber(),
            printLayout.printLayout().setup().rowBreaks(),
            printLayout.printLayout().setup().columnBreaks()));
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
        InspectionResultCellReportSupport.toCellColorReport(presentation.tabColor()),
        new SheetOutlineSummaryReport(
            presentation.outlineSummary().rowSumsBelow(),
            presentation.outlineSummary().rowSumsRight()),
        new SheetDefaultsReport(
            presentation.sheetDefaults().defaultColumnWidth(),
            presentation.sheetDefaults().defaultRowHeightPoints()),
        presentation.ignoredErrors().stream()
            .map(
                ignoredError ->
                    new IgnoredErrorReport(ignoredError.range(), ignoredError.errorTypes()))
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

  static PrintAreaReport toPrintAreaReport(ExcelPrintLayout.Area printArea) {
    return switch (printArea) {
      case ExcelPrintLayout.Area.None _ -> new PrintAreaReport.None();
      case ExcelPrintLayout.Area.Range range -> new PrintAreaReport.Range(range.range());
    };
  }

  static PrintScalingReport toPrintScalingReport(ExcelPrintLayout.Scaling scaling) {
    return switch (scaling) {
      case ExcelPrintLayout.Scaling.Automatic _ -> new PrintScalingReport.Automatic();
      case ExcelPrintLayout.Scaling.Fit fit ->
          new PrintScalingReport.Fit(fit.widthPages(), fit.heightPages());
    };
  }

  static PrintTitleRowsReport toPrintTitleRowsReport(ExcelPrintLayout.TitleRows repeatingRows) {
    return switch (repeatingRows) {
      case ExcelPrintLayout.TitleRows.None _ -> new PrintTitleRowsReport.None();
      case ExcelPrintLayout.TitleRows.Band band ->
          new PrintTitleRowsReport.Band(band.firstRowIndex(), band.lastRowIndex());
    };
  }

  static PrintTitleColumnsReport toPrintTitleColumnsReport(
      ExcelPrintLayout.TitleColumns repeatingColumns) {
    return switch (repeatingColumns) {
      case ExcelPrintLayout.TitleColumns.None _ -> new PrintTitleColumnsReport.None();
      case ExcelPrintLayout.TitleColumns.Band band ->
          new PrintTitleColumnsReport.Band(band.firstColumnIndex(), band.lastColumnIndex());
    };
  }

  static HeaderFooterTextReport toHeaderFooterTextReport(ExcelHeaderFooterText text) {
    return new HeaderFooterTextReport(text.left(), text.center(), text.right());
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
