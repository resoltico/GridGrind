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
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.cells.CellAccess;
import dev.erst.gridgrind.engine.poi.CellBounds;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.concurrent.ConcurrentHashMap;
import java.util.stream.IntStream;

/** Tracks request-authored normal-formula cells without evaluating opaque formula content. */
final class FormulaOriginTracker {
  private final Map<CellKey, StepReference> origins = new ConcurrentHashMap<>();

  FormulaWrites plannedWrites(XssfWorkbook workbook, ResolvedStep step) {
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

  void record(XssfWorkbook workbook, FormulaWrites writes, StepReference authoringStep) {
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
    CellBounds range =
        CellAccess.rangeBounds(((RangeSelector.ByRange) step.target()).range().value());
    Map<CellCoordinate, Optional<String>> writes = new ConcurrentHashMap<>();
    for (int rowOffset = 0; rowOffset < setRange.rows().size(); rowOffset++) {
      for (int columnOffset = 0;
          columnOffset < setRange.rows().get(rowOffset).size();
          columnOffset++) {
        writes.put(
            coordinate(
                ((RangeSelector.ByRange) step.target()).sheetName().value(),
                range.firstRow() + rowOffset,
                range.firstColumn() + columnOffset),
            formulaFor(setRange.rows().get(rowOffset).get(columnOffset)));
      }
    }
    return new FormulaWrites(writes);
  }

  private FormulaWrites arrayFormulaWrites(
      ResolvedStep step, ResolvedMutationAction.SetArrayFormula setArrayFormula) {
    CellBounds range =
        CellAccess.rangeBounds(((RangeSelector.ByRange) step.target()).range().value());
    Map<CellCoordinate, Optional<String>> writes = new ConcurrentHashMap<>();
    IntStream.rangeClosed(range.firstRow(), range.lastRow())
        .forEach(
            row ->
                IntStream.rangeClosed(range.firstColumn(), range.lastColumn())
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
    CellBounds range =
        CellAccess.rangeBounds(((RangeSelector.ByRange) step.target()).range().value());
    Map<CellCoordinate, Optional<String>> writes = new ConcurrentHashMap<>();
    IntStream.rangeClosed(range.firstRow(), range.lastRow())
        .forEach(
            row ->
                IntStream.rangeClosed(range.firstColumn(), range.lastColumn())
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
      XssfWorkbook workbook, ResolvedStep step, ResolvedMutationAction.AppendRow appendRow) {
    int rowIndex =
        workbook
            .sheet(((SheetSelector.ByName) step.target()).name().value())
            .cells()
            .nextAppendRowIndex();
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
    return new CellCoordinate(sheetName, CellAccess.address(rowIndex, columnIndex));
  }

  private void pruneStaleOrigins(XssfWorkbook workbook) {
    origins
        .keySet()
        .removeIf(
            key ->
                !workbook
                    .formulas()
                    .matchesStoredFormula(key.sheetName(), key.address(), key.formula()));
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
