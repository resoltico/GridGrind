package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import org.apache.poi.ss.util.CellReference;

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
      case WorkbookReadCommand.GetFormulaSurface getFormulaSurface ->
          new WorkbookReadResult.FormulaSurfaceResult(
              getFormulaSurface.requestId(),
              formulaSurface(workbook, getFormulaSurface.selection()));
      case WorkbookReadCommand.GetSheetSchema getSheetSchema ->
          new WorkbookReadResult.SheetSchemaResult(
              getSheetSchema.requestId(),
              sheetSchema(
                  workbook,
                  getSheetSchema.sheetName(),
                  getSheetSchema.topLeftAddress(),
                  getSheetSchema.rowCount(),
                  getSheetSchema.columnCount()));
      case WorkbookReadCommand.GetNamedRangeSurface getNamedRangeSurface ->
          new WorkbookReadResult.NamedRangeSurfaceResult(
              getNamedRangeSurface.requestId(),
              namedRangeSurface(workbook, getNamedRangeSurface.selection()));
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

  /** Groups formulas by exact expression across the selected sheets. */
  @SuppressWarnings({"PMD.AvoidInstantiatingObjectsInLoops", "PMD.UseConcurrentHashMap"})
  WorkbookReadResult.FormulaSurface formulaSurface(
      ExcelWorkbook workbook, ExcelSheetSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<String> sheetNames = selectSheets(workbook, selection);
    List<WorkbookReadResult.SheetFormulaSurface> sheets = new ArrayList<>(sheetNames.size());
    int totalFormulaCellCount = 0;
    for (String sheetName : sheetNames) {
      ExcelSheet sheet = workbook.sheet(sheetName);
      List<ExcelCellSnapshot.FormulaSnapshot> formulaCells = sheet.formulaCells();
      Map<String, List<String>> formulas = new LinkedHashMap<>();
      for (ExcelCellSnapshot.FormulaSnapshot formulaCell : formulaCells) {
        formulas
            .computeIfAbsent(formulaCell.formula(), _ -> new ArrayList<>())
            .add(formulaCell.address());
      }
      List<WorkbookReadResult.FormulaPattern> patterns = new ArrayList<>(formulas.size());
      for (Map.Entry<String, List<String>> entry : formulas.entrySet()) {
        patterns.add(
            new WorkbookReadResult.FormulaPattern(
                entry.getKey(), entry.getValue().size(), List.copyOf(entry.getValue())));
      }
      totalFormulaCellCount += formulaCells.size();
      sheets.add(
          new WorkbookReadResult.SheetFormulaSurface(
              sheetName, formulaCells.size(), patterns.size(), List.copyOf(patterns)));
    }
    return new WorkbookReadResult.FormulaSurface(totalFormulaCellCount, List.copyOf(sheets));
  }

  /** Infers a simple schema from the provided sheet window. */
  @SuppressWarnings({"PMD.AvoidInstantiatingObjectsInLoops", "PMD.UseConcurrentHashMap"})
  WorkbookReadResult.SheetSchema sheetSchema(
      ExcelWorkbook workbook,
      String sheetName,
      String topLeftAddress,
      int rowCount,
      int columnCount) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    ExcelSheet sheet = workbook.sheet(sheetName);
    WorkbookReadResult.Window window =
        sheetIntrospector.window(sheet, topLeftAddress, rowCount, columnCount);
    int topLeftColumn = new CellReference(topLeftAddress).getCol();
    int dataRowCount = Math.max(0, rowCount - 1);

    List<WorkbookReadResult.SchemaColumn> columns = new ArrayList<>(columnCount);
    for (int columnOffset = 0; columnOffset < columnCount; columnOffset++) {
      ExcelCellSnapshot headerCell = window.rows().getFirst().cells().get(columnOffset);
      Map<String, Integer> typeCounts = new LinkedHashMap<>();
      int populatedCellCount = 0;
      int blankCellCount = 0;
      for (int rowIndex = 1; rowIndex < window.rows().size(); rowIndex++) {
        ExcelCellSnapshot cell = window.rows().get(rowIndex).cells().get(columnOffset);
        if ("BLANK".equals(cell.effectiveType())) {
          blankCellCount++;
          continue;
        }
        populatedCellCount++;
        typeCounts.merge(cell.effectiveType(), 1, Integer::sum);
      }

      List<WorkbookReadResult.TypeCount> observedTypes = new ArrayList<>(typeCounts.size());
      for (Map.Entry<String, Integer> entry : typeCounts.entrySet()) {
        observedTypes.add(new WorkbookReadResult.TypeCount(entry.getKey(), entry.getValue()));
      }

      columns.add(
          new WorkbookReadResult.SchemaColumn(
              topLeftColumn + columnOffset,
              new CellReference(
                      new CellReference(topLeftAddress).getRow(), topLeftColumn + columnOffset)
                  .formatAsString(),
              headerCell.displayValue(),
              populatedCellCount,
              blankCellCount,
              List.copyOf(observedTypes),
              dominantType(typeCounts)));
    }

    return new WorkbookReadResult.SheetSchema(
        sheetName, topLeftAddress, rowCount, columnCount, dataRowCount, List.copyOf(columns));
  }

  /** Summarizes the scope and backing kind of selected named ranges. */
  WorkbookReadResult.NamedRangeSurface namedRangeSurface(
      ExcelWorkbook workbook, ExcelNamedRangeSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<ExcelNamedRangeSnapshot> namedRanges = selectNamedRanges(workbook, selection);
    int workbookScopedCount = 0;
    int sheetScopedCount = 0;
    int rangeBackedCount = 0;
    int formulaBackedCount = 0;
    List<WorkbookReadResult.NamedRangeSurfaceEntry> entries = new ArrayList<>(namedRanges.size());
    for (ExcelNamedRangeSnapshot namedRange : namedRanges) {
      WorkbookReadResult.NamedRangeBackingKind kind;
      switch (namedRange) {
        case ExcelNamedRangeSnapshot.RangeSnapshot _ -> {
          kind = WorkbookReadResult.NamedRangeBackingKind.RANGE;
          rangeBackedCount++;
        }
        case ExcelNamedRangeSnapshot.FormulaSnapshot _ -> {
          kind = WorkbookReadResult.NamedRangeBackingKind.FORMULA;
          formulaBackedCount++;
        }
      }
      switch (namedRange.scope()) {
        case ExcelNamedRangeScope.WorkbookScope _ -> workbookScopedCount++;
        case ExcelNamedRangeScope.SheetScope _ -> sheetScopedCount++;
      }
      entries.add(
          new WorkbookReadResult.NamedRangeSurfaceEntry(
              namedRange.name(), namedRange.scope(), namedRange.refersToFormula(), kind));
    }
    return new WorkbookReadResult.NamedRangeSurface(
        workbookScopedCount,
        sheetScopedCount,
        rangeBackedCount,
        formulaBackedCount,
        List.copyOf(entries));
  }

  private List<String> selectSheets(ExcelWorkbook workbook, ExcelSheetSelection selection) {
    return switch (selection) {
      case ExcelSheetSelection.All _ -> workbook.sheetNames();
      case ExcelSheetSelection.Selected selected -> List.copyOf(selected.sheetNames());
    };
  }

  private String dominantType(Map<String, Integer> typeCounts) {
    String dominantType = null;
    int dominantCount = -1;
    boolean tie = false;
    for (Map.Entry<String, Integer> entry : typeCounts.entrySet()) {
      if (entry.getValue() > dominantCount) {
        dominantType = entry.getKey();
        dominantCount = entry.getValue();
        tie = false;
      } else if (entry.getValue() == dominantCount) {
        tie = true;
      }
    }
    return tie ? null : dominantType;
  }
}
