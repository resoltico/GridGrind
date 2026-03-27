package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Reads workbook facts and sheet introspection data from one workbook wrapper. */
final class ExcelWorkbookIntrospector {
  private final ExcelSheetIntrospector sheetIntrospector;

  ExcelWorkbookIntrospector() {
    this(new ExcelSheetIntrospector());
  }

  ExcelWorkbookIntrospector(ExcelSheetIntrospector sheetIntrospector) {
    this.sheetIntrospector =
        Objects.requireNonNull(sheetIntrospector, "sheetIntrospector must not be null");
  }

  /** Executes one fact-only read command against the workbook. */
  WorkbookReadResult.Introspection execute(
      ExcelWorkbook workbook, WorkbookReadCommand.Introspection command) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(command, "command must not be null");

    return switch (command) {
      case WorkbookReadCommand.GetWorkbookSummary getWorkbookSummary ->
          new WorkbookReadResult.WorkbookSummaryResult(
              getWorkbookSummary.requestId(),
              new WorkbookReadResult.WorkbookSummary(
                  workbook.sheetCount(),
                  workbook.sheetNames(),
                  workbook.namedRangeCount(),
                  workbook.forceFormulaRecalculationOnOpenEnabled()));
      case WorkbookReadCommand.GetNamedRanges getNamedRanges ->
          new WorkbookReadResult.NamedRangesResult(
              getNamedRanges.requestId(), selectNamedRanges(workbook, getNamedRanges.selection()));
      case WorkbookReadCommand.GetSheetSummary getSheetSummary ->
          new WorkbookReadResult.SheetSummaryResult(
              getSheetSummary.requestId(),
              sheetIntrospector.summarize(workbook.sheet(getSheetSummary.sheetName())));
      case WorkbookReadCommand.GetCells getCells ->
          new WorkbookReadResult.CellsResult(
              getCells.requestId(),
              getCells.sheetName(),
              sheetIntrospector.cells(workbook.sheet(getCells.sheetName()), getCells.addresses()));
      case WorkbookReadCommand.GetWindow getWindow ->
          new WorkbookReadResult.WindowResult(
              getWindow.requestId(),
              sheetIntrospector.window(
                  workbook.sheet(getWindow.sheetName()),
                  getWindow.topLeftAddress(),
                  getWindow.rowCount(),
                  getWindow.columnCount()));
      case WorkbookReadCommand.GetMergedRegions getMergedRegions ->
          new WorkbookReadResult.MergedRegionsResult(
              getMergedRegions.requestId(),
              getMergedRegions.sheetName(),
              sheetIntrospector.mergedRegions(workbook.sheet(getMergedRegions.sheetName())));
      case WorkbookReadCommand.GetHyperlinks getHyperlinks ->
          new WorkbookReadResult.HyperlinksResult(
              getHyperlinks.requestId(),
              getHyperlinks.sheetName(),
              sheetIntrospector.hyperlinks(
                  workbook.sheet(getHyperlinks.sheetName()), getHyperlinks.selection()));
      case WorkbookReadCommand.GetComments getComments ->
          new WorkbookReadResult.CommentsResult(
              getComments.requestId(),
              getComments.sheetName(),
              sheetIntrospector.comments(
                  workbook.sheet(getComments.sheetName()), getComments.selection()));
      case WorkbookReadCommand.GetSheetLayout getSheetLayout ->
          new WorkbookReadResult.SheetLayoutResult(
              getSheetLayout.requestId(),
              sheetIntrospector.layout(workbook.sheet(getSheetLayout.sheetName())));
    };
  }

  /** Returns the named ranges selected by the provided workbook-core named-range selection. */
  List<ExcelNamedRangeSnapshot> selectNamedRanges(
      ExcelWorkbook workbook, ExcelNamedRangeSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<ExcelNamedRangeSnapshot> namedRanges = workbook.namedRanges();
    return switch (selection) {
      case ExcelNamedRangeSelection.All _ -> namedRanges;
      case ExcelNamedRangeSelection.Selected selected ->
          selectNamedRanges(namedRanges, selected.selectors());
    };
  }

  private List<ExcelNamedRangeSnapshot> selectNamedRanges(
      List<ExcelNamedRangeSnapshot> namedRanges, List<ExcelNamedRangeSelector> selectors) {
    Set<ExcelNamedRangeSnapshot> matched = new LinkedHashSet<>();
    for (ExcelNamedRangeSelector selector : selectors) {
      matched.addAll(matchSelector(namedRanges, selector));
    }
    return List.copyOf(matched);
  }

  /** Returns the named ranges matched by one exact selector or throws when none match. */
  static List<ExcelNamedRangeSnapshot> matchSelector(
      List<ExcelNamedRangeSnapshot> namedRanges, ExcelNamedRangeSelector selector) {
    List<ExcelNamedRangeSnapshot> matches = new ArrayList<>();
    switch (selector) {
      case ExcelNamedRangeSelector.ByName byName -> {
        for (ExcelNamedRangeSnapshot namedRange : namedRanges) {
          if (namedRange.name().equalsIgnoreCase(byName.name())) {
            matches.add(namedRange);
          }
        }
      }
      case ExcelNamedRangeSelector.WorkbookScope workbookScope -> {
        for (ExcelNamedRangeSnapshot namedRange : namedRanges) {
          if (namedRange.name().equalsIgnoreCase(workbookScope.name())
              && namedRange.scope() instanceof ExcelNamedRangeScope.WorkbookScope) {
            matches.add(namedRange);
          }
        }
      }
      case ExcelNamedRangeSelector.SheetScope sheetScope -> {
        for (ExcelNamedRangeSnapshot namedRange : namedRanges) {
          if (namedRange.name().equalsIgnoreCase(sheetScope.name())
              && namedRange.scope() instanceof ExcelNamedRangeScope.SheetScope namedRangeScope
              && namedRangeScope.sheetName().equals(sheetScope.sheetName())) {
            matches.add(namedRange);
          }
        }
      }
    }
    if (matches.isEmpty()) {
      throw notFound(selector);
    }
    return List.copyOf(matches);
  }

  private static NamedRangeNotFoundException notFound(ExcelNamedRangeSelector selector) {
    return switch (selector) {
      case ExcelNamedRangeSelector.ByName byName ->
          new NamedRangeNotFoundException(byName.name(), new ExcelNamedRangeScope.WorkbookScope());
      case ExcelNamedRangeSelector.WorkbookScope workbookScope ->
          new NamedRangeNotFoundException(
              workbookScope.name(), new ExcelNamedRangeScope.WorkbookScope());
      case ExcelNamedRangeSelector.SheetScope sheetScope ->
          new NamedRangeNotFoundException(
              sheetScope.name(), new ExcelNamedRangeScope.SheetScope(sheetScope.sheetName()));
    };
  }
}
