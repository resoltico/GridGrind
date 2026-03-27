package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import org.apache.poi.ss.util.CellReference;

/** Derives higher-level workbook analysis from reusable introspection primitives. */
final class WorkbookInsightAnalyzer {
  private final ExcelWorkbookIntrospector workbookIntrospector;
  private final ExcelSheetIntrospector sheetIntrospector;

  WorkbookInsightAnalyzer() {
    this(new ExcelWorkbookIntrospector(), new ExcelSheetIntrospector());
  }

  WorkbookInsightAnalyzer(
      ExcelWorkbookIntrospector workbookIntrospector, ExcelSheetIntrospector sheetIntrospector) {
    this.workbookIntrospector =
        Objects.requireNonNull(workbookIntrospector, "workbookIntrospector must not be null");
    this.sheetIntrospector =
        Objects.requireNonNull(sheetIntrospector, "sheetIntrospector must not be null");
  }

  /** Executes one derived analysis command against the workbook. */
  WorkbookReadResult.Insight execute(ExcelWorkbook workbook, WorkbookReadCommand.Insight command) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(command, "command must not be null");

    return switch (command) {
      case WorkbookReadCommand.AnalyzeFormulaSurface analyzeFormulaSurface ->
          new WorkbookReadResult.FormulaSurfaceResult(
              analyzeFormulaSurface.requestId(),
              analyzeFormulaSurface(workbook, analyzeFormulaSurface.selection()));
      case WorkbookReadCommand.AnalyzeSheetSchema analyzeSheetSchema ->
          new WorkbookReadResult.SheetSchemaResult(
              analyzeSheetSchema.requestId(),
              analyzeSheetSchema(
                  workbook,
                  analyzeSheetSchema.sheetName(),
                  analyzeSheetSchema.topLeftAddress(),
                  analyzeSheetSchema.rowCount(),
                  analyzeSheetSchema.columnCount()));
      case WorkbookReadCommand.AnalyzeNamedRangeSurface analyzeNamedRangeSurface ->
          new WorkbookReadResult.NamedRangeSurfaceResult(
              analyzeNamedRangeSurface.requestId(),
              analyzeNamedRangeSurface(workbook, analyzeNamedRangeSurface.selection()));
    };
  }

  /** Groups formulas by exact expression across the selected sheets. */
  @SuppressWarnings({"PMD.AvoidInstantiatingObjectsInLoops", "PMD.UseConcurrentHashMap"})
  WorkbookReadResult.FormulaSurface analyzeFormulaSurface(
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
  WorkbookReadResult.SheetSchema analyzeSheetSchema(
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
  WorkbookReadResult.NamedRangeSurface analyzeNamedRangeSurface(
      ExcelWorkbook workbook, ExcelNamedRangeSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<ExcelNamedRangeSnapshot> namedRanges =
        workbookIntrospector.selectNamedRanges(workbook, selection);
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
