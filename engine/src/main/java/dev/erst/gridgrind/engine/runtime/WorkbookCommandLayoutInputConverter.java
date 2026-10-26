package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.layout.PaneSpec;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.PrintLayoutSpec;
import dev.erst.gridgrind.core.model.print.ResolvedPrintLayout;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionSpec;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.SheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelWorkbookProtectionSettings;
import java.util.Optional;

/** Converts named-range, protection, pane, and print/presentation contract inputs. */
final class WorkbookCommandLayoutInputConverter {
  private WorkbookCommandLayoutInputConverter() {}

  static ExcelNamedRangeScope toExcelNamedRangeScope(NamedRangeScope scope) {
    return switch (scope) {
      case NamedRangeScope.Workbook _ -> new ExcelNamedRangeScope.WorkbookScope();
      case NamedRangeScope.Sheet sheet ->
          new ExcelNamedRangeScope.SheetScope(sheet.sheetName().value());
    };
  }

  static ExcelNamedRangeScope toExcelNamedRangeScope(NamedRangeSelector.ScopedExact selector) {
    return switch (selector) {
      case NamedRangeSelector.WorkbookScope _ -> new ExcelNamedRangeScope.WorkbookScope();
      case NamedRangeSelector.SheetScope sheet ->
          new ExcelNamedRangeScope.SheetScope(sheet.sheetName().value());
    };
  }

  static String toExcelNamedRangeName(NamedRangeSelector.ScopedExact selector) {
    return switch (selector) {
      case NamedRangeSelector.WorkbookScope workbook -> workbook.name().value();
      case NamedRangeSelector.SheetScope sheet -> sheet.name().value();
    };
  }

  static ExcelNamedRangeTarget toExcelNamedRangeTarget(NamedRangeTarget target) {
    return switch (target) {
      case NamedRangeTarget.Formula formula -> ExcelNamedRangeTarget.formula(formula.formula());
      case NamedRangeTarget.Range range ->
          ExcelNamedRangeTarget.range(range.sheetName().value(), range.range().value());
    };
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
      WorkbookProtectionSpec protection, ExecutionInputBindings bindings) {
    return new ExcelWorkbookProtectionSettings(
        protection.structureLocked(),
        protection.windowsLocked(),
        protection.revisionsLocked(),
        resolveOptionalSecret(protection.workbookPasswordRef(), bindings),
        resolveOptionalSecret(protection.revisionsPasswordRef(), bindings));
  }

  private static Optional<String> resolveOptionalSecret(
      Optional<dev.erst.gridgrind.core.source.SecretReference> reference,
      ExecutionInputBindings bindings) {
    if (reference.isEmpty()) {
      return Optional.empty();
    }
    if (bindings == null) {
      throw new IllegalStateException(
          "secret-bearing workbook protection requires execution bindings");
    }
    return Optional.of(bindings.resolveSecret(reference.orElseThrow()));
  }

  static ExcelSheetPane toExcelSheetPane(PaneSpec pane) {
    return switch (pane) {
      case PaneSpec.None _ -> new ExcelSheetPane.None();
      case PaneSpec.Frozen frozen ->
          new ExcelSheetPane.Frozen(
              frozen.splitColumn(), frozen.splitRow(), frozen.leftmostColumn(), frozen.topRow());
      case PaneSpec.Split split ->
          new ExcelSheetPane.Split(
              split.xSplitPosition(),
              split.ySplitPosition(),
              split.leftmostColumn(),
              split.topRow(),
              split.activePane());
    };
  }

  static ResolvedPrintLayout toResolvedPrintLayout(PrintLayoutSpec printLayout) {
    return new ResolvedPrintLayout(
        printLayout.printArea(),
        printLayout.orientation(),
        printLayout.scaling(),
        printLayout.repeatingRows(),
        printLayout.repeatingColumns(),
        new HeaderFooterTextReport(
            WorkbookCommandSourceSupport.inlineText(printLayout.header().left(), "header left"),
            WorkbookCommandSourceSupport.inlineText(printLayout.header().center(), "header center"),
            WorkbookCommandSourceSupport.inlineText(printLayout.header().right(), "header right")),
        new HeaderFooterTextReport(
            WorkbookCommandSourceSupport.inlineText(printLayout.footer().left(), "footer left"),
            WorkbookCommandSourceSupport.inlineText(printLayout.footer().center(), "footer center"),
            WorkbookCommandSourceSupport.inlineText(printLayout.footer().right(), "footer right")),
        printLayout.setup());
  }
}
