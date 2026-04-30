package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.PaneInput;
import dev.erst.gridgrind.contract.dto.PrintAreaInput;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.PrintScalingInput;
import dev.erst.gridgrind.contract.dto.PrintSetupInput;
import dev.erst.gridgrind.contract.dto.PrintTitleColumnsInput;
import dev.erst.gridgrind.contract.dto.PrintTitleRowsInput;
import dev.erst.gridgrind.contract.dto.SheetCopyPosition;
import dev.erst.gridgrind.contract.dto.SheetPresentationInput;
import dev.erst.gridgrind.contract.dto.SheetProtectionSettings;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.excel.ExcelHeaderFooterText;
import dev.erst.gridgrind.excel.ExcelIgnoredError;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelPrintLayout;
import dev.erst.gridgrind.excel.ExcelPrintMargins;
import dev.erst.gridgrind.excel.ExcelPrintSetup;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetDefaults;
import dev.erst.gridgrind.excel.ExcelSheetDisplay;
import dev.erst.gridgrind.excel.ExcelSheetOutlineSummary;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelSheetPresentation;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelWorkbookProtectionSettings;

/** Converts named-range, protection, pane, and print/presentation contract inputs. */
final class WorkbookCommandLayoutInputConverter {
  private WorkbookCommandLayoutInputConverter() {}

  static ExcelNamedRangeScope toExcelNamedRangeScope(NamedRangeScope scope) {
    return switch (scope) {
      case NamedRangeScope.Workbook _ -> new ExcelNamedRangeScope.WorkbookScope();
      case NamedRangeScope.Sheet sheet -> new ExcelNamedRangeScope.SheetScope(sheet.sheetName());
    };
  }

  static ExcelNamedRangeScope toExcelNamedRangeScope(NamedRangeSelector.ScopedExact selector) {
    return switch (selector) {
      case NamedRangeSelector.WorkbookScope _ -> new ExcelNamedRangeScope.WorkbookScope();
      case NamedRangeSelector.SheetScope sheet ->
          new ExcelNamedRangeScope.SheetScope(sheet.sheetName());
    };
  }

  static String toExcelNamedRangeName(NamedRangeSelector.ScopedExact selector) {
    return switch (selector) {
      case NamedRangeSelector.WorkbookScope workbook -> workbook.name();
      case NamedRangeSelector.SheetScope sheet -> sheet.name();
    };
  }

  static ExcelNamedRangeTarget toExcelNamedRangeTarget(NamedRangeTarget target) {
    return target.formula() != null
        ? new ExcelNamedRangeTarget(target.formula())
        : new ExcelNamedRangeTarget(target.sheetName(), target.range());
  }

  static ExcelSheetCopyPosition toExcelSheetCopyPosition(SheetCopyPosition position) {
    return switch (position) {
      case SheetCopyPosition.AppendAtEnd _ -> new ExcelSheetCopyPosition.AppendAtEnd();
      case SheetCopyPosition.AtIndex atIndex ->
          new ExcelSheetCopyPosition.AtIndex(atIndex.targetIndex());
    };
  }

  static ExcelSheetProtectionSettings toExcelSheetProtectionSettings(
      SheetProtectionSettings settings) {
    return new ExcelSheetProtectionSettings(
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

  static ExcelWorkbookProtectionSettings toExcelWorkbookProtectionSettings(
      WorkbookProtectionInput protection) {
    return new ExcelWorkbookProtectionSettings(
        protection.structureLocked(),
        protection.windowsLocked(),
        protection.revisionsLocked(),
        protection.workbookPassword(),
        protection.revisionsPassword());
  }

  static ExcelSheetPane toExcelSheetPane(PaneInput pane) {
    return switch (pane) {
      case PaneInput.None _ -> new ExcelSheetPane.None();
      case PaneInput.Frozen frozen ->
          new ExcelSheetPane.Frozen(
              frozen.splitColumn(), frozen.splitRow(), frozen.leftmostColumn(), frozen.topRow());
      case PaneInput.Split split ->
          new ExcelSheetPane.Split(
              split.xSplitPosition(),
              split.ySplitPosition(),
              split.leftmostColumn(),
              split.topRow(),
              split.activePane());
    };
  }

  static ExcelPrintLayout toExcelPrintLayout(PrintLayoutInput printLayout) {
    return new ExcelPrintLayout(
        toExcelPrintArea(printLayout.printArea()),
        printLayout.orientation(),
        toExcelPrintScaling(printLayout.scaling()),
        toExcelPrintTitleRows(printLayout.repeatingRows()),
        toExcelPrintTitleColumns(printLayout.repeatingColumns()),
        new ExcelHeaderFooterText(
            WorkbookCommandSourceSupport.inlineText(printLayout.header().left(), "header left"),
            WorkbookCommandSourceSupport.inlineText(printLayout.header().center(), "header center"),
            WorkbookCommandSourceSupport.inlineText(printLayout.header().right(), "header right")),
        new ExcelHeaderFooterText(
            WorkbookCommandSourceSupport.inlineText(printLayout.footer().left(), "footer left"),
            WorkbookCommandSourceSupport.inlineText(printLayout.footer().center(), "footer center"),
            WorkbookCommandSourceSupport.inlineText(printLayout.footer().right(), "footer right")),
        toExcelPrintSetup(printLayout.setup()));
  }

  static ExcelSheetPresentation toExcelSheetPresentation(SheetPresentationInput presentation) {
    return new ExcelSheetPresentation(
        new ExcelSheetDisplay(
            presentation.display().displayGridlines(),
            presentation.display().displayZeros(),
            presentation.display().displayRowColHeadings(),
            presentation.display().displayFormulas(),
            presentation.display().rightToLeft()),
        presentation
            .tabColor()
            .flatMap(WorkbookCommandCellInputConverter::toExcelColor)
            .orElse(null),
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

  private static ExcelPrintSetup toExcelPrintSetup(PrintSetupInput setup) {
    return new ExcelPrintSetup(
        new ExcelPrintMargins(
            setup.margins().left(),
            setup.margins().right(),
            setup.margins().top(),
            setup.margins().bottom(),
            setup.margins().header(),
            setup.margins().footer()),
        setup.printGridlines(),
        setup.horizontallyCentered(),
        setup.verticallyCentered(),
        setup.paperSize(),
        setup.draft(),
        setup.blackAndWhite(),
        setup.copies(),
        setup.useFirstPageNumber(),
        setup.firstPageNumber(),
        setup.rowBreaks(),
        setup.columnBreaks());
  }

  private static ExcelPrintLayout.Area toExcelPrintArea(PrintAreaInput printArea) {
    return switch (printArea) {
      case PrintAreaInput.None _ -> new ExcelPrintLayout.Area.None();
      case PrintAreaInput.Range range -> new ExcelPrintLayout.Area.Range(range.range());
    };
  }

  private static ExcelPrintLayout.Scaling toExcelPrintScaling(PrintScalingInput scaling) {
    return switch (scaling) {
      case PrintScalingInput.Automatic _ -> new ExcelPrintLayout.Scaling.Automatic();
      case PrintScalingInput.Fit fit ->
          new ExcelPrintLayout.Scaling.Fit(fit.widthPages(), fit.heightPages());
    };
  }

  private static ExcelPrintLayout.TitleRows toExcelPrintTitleRows(
      PrintTitleRowsInput repeatingRows) {
    return switch (repeatingRows) {
      case PrintTitleRowsInput.None _ -> new ExcelPrintLayout.TitleRows.None();
      case PrintTitleRowsInput.Band band ->
          new ExcelPrintLayout.TitleRows.Band(band.firstRowIndex(), band.lastRowIndex());
    };
  }

  private static ExcelPrintLayout.TitleColumns toExcelPrintTitleColumns(
      PrintTitleColumnsInput repeatingColumns) {
    return switch (repeatingColumns) {
      case PrintTitleColumnsInput.None _ -> new ExcelPrintLayout.TitleColumns.None();
      case PrintTitleColumnsInput.Band band ->
          new ExcelPrintLayout.TitleColumns.Band(band.firstColumnIndex(), band.lastColumnIndex());
    };
  }
}
