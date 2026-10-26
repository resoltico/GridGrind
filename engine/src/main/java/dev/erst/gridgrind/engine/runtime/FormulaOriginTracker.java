package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.StepReference;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.concurrent.ConcurrentHashMap;
import java.util.stream.IntStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

/** Tracks request-authored normal-formula cells without evaluating opaque formula content. */
final class FormulaOriginTracker {
  private final Map<CellKey, StepReference> origins = new ConcurrentHashMap<>();

  FormulaWrites plannedWrites(ExcelWorkbook workbook, ResolvedStep step) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(step, "step must not be null");
    return switch (step.action()) {
      case ResolvedMutationAction.SetCell setCell ->
          new FormulaWrites(
              Map.of(
                  new CellCoordinate(
                      ((CellSelector.ByAddress) step.target()).sheetName().value(),
                      ((CellSelector.ByAddress) step.target()).address().value()),
                  formulaFor(setCell.value())));
      case ResolvedMutationAction.SetRange setRange -> rangeWrites(step, setRange);
      case ResolvedMutationAction.SetArrayFormula setArrayFormula ->
          arrayFormulaWrites(step, setArrayFormula);
      case CellMutationAction.ClearRange _ -> clearRangeWrites(step);
      case CellMutationAction.ClearArrayFormula _ ->
          new FormulaWrites(
              Map.of(
                  new CellCoordinate(
                      ((CellSelector.ByAddress) step.target()).sheetName().value(),
                      ((CellSelector.ByAddress) step.target()).address().value()),
                  Optional.empty()));
      case ResolvedMutationAction.AppendRow appendRow -> appendWrites(workbook, step, appendRow);
      default -> FormulaWrites.none();
    };
  }

  void record(ExcelWorkbook workbook, FormulaWrites writes, StepReference authoringStep) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(writes, "writes must not be null");
    Objects.requireNonNull(authoringStep, "authoringStep must not be null");
    pruneStaleOrigins(workbook);
    for (Map.Entry<CellCoordinate, Optional<String>> entry : writes.cells().entrySet()) {
      entry.getValue().ifPresent(formula -> recordFormula(entry.getKey(), formula, authoringStep));
    }
  }

  Optional<StepReference> originFor(Throwable failure) {
    if (!(failure instanceof Exception exception)) {
      return Optional.empty();
    }
    return ExecutionExceptionDiagnosticFields.sheetNameFor(exception)
        .flatMap(
            sheetName ->
                ExecutionExceptionDiagnosticFields.addressFor(exception)
                    .flatMap(address -> originFor(sheetName, address)));
  }

  Optional<StepReference> originFor(String sheetName, String address) {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(address, "address must not be null");
    return origins.entrySet().stream()
        .filter(
            entry ->
                entry.getKey().sheetName().equals(sheetName)
                    && entry.getKey().address().equals(address))
        .map(Map.Entry::getValue)
        .findFirst();
  }

  private FormulaWrites rangeWrites(ResolvedStep step, ResolvedMutationAction.SetRange setRange) {
    CellRangeAddress range =
        CellRangeAddress.valueOf(((RangeSelector.ByRange) step.target()).range().value());
    Map<CellCoordinate, Optional<String>> writes = new ConcurrentHashMap<>();
    for (int rowOffset = 0; rowOffset < setRange.rows().size(); rowOffset++) {
      for (int columnOffset = 0;
          columnOffset < setRange.rows().get(rowOffset).size();
          columnOffset++) {
        writes.put(
            coordinate(
                ((RangeSelector.ByRange) step.target()).sheetName().value(),
                range.getFirstRow() + rowOffset,
                range.getFirstColumn() + columnOffset),
            formulaFor(setRange.rows().get(rowOffset).get(columnOffset)));
      }
    }
    return new FormulaWrites(writes);
  }

  private FormulaWrites arrayFormulaWrites(
      ResolvedStep step, ResolvedMutationAction.SetArrayFormula setArrayFormula) {
    CellRangeAddress range =
        CellRangeAddress.valueOf(((RangeSelector.ByRange) step.target()).range().value());
    Map<CellCoordinate, Optional<String>> writes = new ConcurrentHashMap<>();
    IntStream.rangeClosed(range.getFirstRow(), range.getLastRow())
        .forEach(
            row ->
                IntStream.rangeClosed(range.getFirstColumn(), range.getLastColumn())
                    .forEach(
                        column ->
                            writes.put(
                                coordinate(
                                    ((RangeSelector.ByRange) step.target()).sheetName().value(),
                                    row,
                                    column),
                                Optional.of(setArrayFormula.formula().formula()))));
    return new FormulaWrites(writes);
  }

  private FormulaWrites clearRangeWrites(ResolvedStep step) {
    CellRangeAddress range =
        CellRangeAddress.valueOf(((RangeSelector.ByRange) step.target()).range().value());
    Map<CellCoordinate, Optional<String>> writes = new ConcurrentHashMap<>();
    IntStream.rangeClosed(range.getFirstRow(), range.getLastRow())
        .forEach(
            row ->
                IntStream.rangeClosed(range.getFirstColumn(), range.getLastColumn())
                    .forEach(
                        column ->
                            writes.put(
                                coordinate(
                                    ((RangeSelector.ByRange) step.target()).sheetName().value(),
                                    row,
                                    column),
                                Optional.empty())));
    return new FormulaWrites(writes);
  }

  private FormulaWrites appendWrites(
      ExcelWorkbook workbook, ResolvedStep step, ResolvedMutationAction.AppendRow appendRow) {
    Sheet sheet =
        workbook.xssfWorkbook().getSheet(((SheetSelector.ByName) step.target()).name().value());
    int rowIndex = nextAppendRowIndex(sheet);
    Map<CellCoordinate, Optional<String>> writes = new ConcurrentHashMap<>();
    for (int column = 0; column < appendRow.values().size(); column++) {
      writes.put(
          coordinate(((SheetSelector.ByName) step.target()).name().value(), rowIndex, column),
          formulaFor(appendRow.values().get(column)));
    }
    return new FormulaWrites(writes);
  }

  private static Optional<String> formulaFor(ResolvedCellWrite value) {
    if (value instanceof ResolvedFormula formula) {
      return Optional.of(formula.value());
    }
    return Optional.empty();
  }

  private void recordFormula(
      CellCoordinate coordinate, String formula, StepReference authoringStep) {
    origins.put(new CellKey(coordinate.sheetName(), coordinate.address(), formula), authoringStep);
  }

  private static CellCoordinate coordinate(String sheetName, int rowIndex, int columnIndex) {
    return new CellCoordinate(sheetName, new CellReference(rowIndex, columnIndex).formatAsString());
  }

  private static int nextAppendRowIndex(Sheet sheet) {
    int lastValueBearingRowIndex = -1;
    for (Row row : sheet) {
      if (rowHasValueBearingCell(row)) {
        lastValueBearingRowIndex = row.getRowNum();
      }
    }
    return lastValueBearingRowIndex + 1;
  }

  private static boolean rowHasValueBearingCell(Row row) {
    for (Cell cell : row) {
      if (cell.getCellType() != CellType.BLANK) {
        return true;
      }
    }
    return false;
  }

  private void pruneStaleOrigins(ExcelWorkbook workbook) {
    origins
        .keySet()
        .removeIf(
            key -> {
              Sheet sheet = workbook.xssfWorkbook().getSheet(key.sheetName());
              if (sheet == null) {
                return true;
              }
              CellReference reference = new CellReference(key.address());
              Row row = sheet.getRow(reference.getRow());
              Cell cell = row == null ? null : row.getCell(reference.getCol());
              return cell == null
                  || cell.getCellType() != CellType.FORMULA
                  || !key.formula().equals(cell.getCellFormula());
            });
  }

  record FormulaWrites(Map<CellCoordinate, Optional<String>> cells) {
    FormulaWrites {
      Objects.requireNonNull(cells, "cells must not be null");
      cells = Map.copyOf(cells);
    }

    static FormulaWrites none() {
      return new FormulaWrites(Map.of());
    }
  }

  record CellCoordinate(String sheetName, String address) {
    CellCoordinate {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
    }
  }

  private record CellKey(String sheetName, String address, String formula) {
    private CellKey {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(formula, "formula must not be null");
    }
  }
}
