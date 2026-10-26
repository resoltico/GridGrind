package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.plan.CellReadFacet;
import dev.erst.gridgrind.core.plan.CellReadProjection;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.SheetInspectionResult;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookLocation;
import java.util.List;
import java.util.Objects;

/** Evaluates value- and presence-oriented assertion families against canonical read results. */
final class AssertionValueEvaluator {
  private final AssertionObservationExecutor observations;

  AssertionValueEvaluator(AssertionObservationExecutor observations) {
    this.observations = Objects.requireNonNull(observations, "observations must not be null");
  }

  AssertionEvaluation evaluateEntityPresence(
      StepId stepId,
      Selector target,
      String assertionType,
      boolean shouldExist,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    List<InspectionResult> observationList =
        List.of(observations.presenceObservation(stepId, target, workbook, workbookLocation));
    int count = AssertionObservationExecutor.observedCount(observationList.getFirst());
    boolean matchedExpectation = shouldExist ? count > 0 : count == 0;
    return matchedExpectation
        ? AssertionEvaluation.pass(observationList)
        : AssertionEvaluation.fail(
            observationList,
            shouldExist
                ? assertionType + " observed no matching workbook entities"
                : assertionType
                    + " observed "
                    + count
                    + " matching workbook "
                    + (count == 1 ? "entity" : "entities"));
  }

  AssertionEvaluation evaluateCellValue(
      StepId stepId,
      Selector target,
      CellValue expectedValue,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    SheetInspectionResult.CellsResult cellsResult =
        (SheetInspectionResult.CellsResult)
            observations.executeObservation(
                stepId,
                target,
                new SheetIntrospectionQuery.GetCells(
                    java.util.Optional.of(CellReadProjection.of(CellReadFacet.VALUE))),
                workbook,
                workbookLocation);
    if (cellsResult.cells().isEmpty()) {
      return AssertionEvaluation.fail(
          List.of(cellsResult), "EXPECT_CELL_VALUE resolved no matching cells to compare");
    }
    List<String> mismatches =
        cellsResult.cells().stream()
            .filter(cell -> !matchesCellValue(cell, expectedValue))
            .map(cell -> cell.address().value())
            .toList();
    return mismatches.isEmpty()
        ? AssertionEvaluation.pass(List.of(cellsResult))
        : AssertionEvaluation.fail(
            List.of(cellsResult),
            "EXPECT_CELL_VALUE mismatched effective values at " + String.join(", ", mismatches));
  }

  AssertionEvaluation evaluateDisplayValue(
      StepId stepId,
      Selector target,
      String expectedDisplayValue,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    SheetInspectionResult.CellsResult cellsResult =
        (SheetInspectionResult.CellsResult)
            observations.executeObservation(
                stepId,
                target,
                new SheetIntrospectionQuery.GetCells(
                    java.util.Optional.of(CellReadProjection.of(CellReadFacet.FORMAT))),
                workbook,
                workbookLocation);
    if (cellsResult.cells().isEmpty()) {
      return AssertionEvaluation.fail(
          List.of(cellsResult), "EXPECT_DISPLAY_VALUE resolved no matching cells to compare");
    }
    List<String> mismatches =
        cellsResult.cells().stream()
            .filter(cell -> !cell.displayValue().orElseThrow().equals(expectedDisplayValue))
            .map(cell -> cell.address().value())
            .toList();
    return mismatches.isEmpty()
        ? AssertionEvaluation.pass(List.of(cellsResult))
        : AssertionEvaluation.fail(
            List.of(cellsResult),
            "EXPECT_DISPLAY_VALUE mismatched formatted values at " + String.join(", ", mismatches));
  }

  AssertionEvaluation evaluateFormulaText(
      StepId stepId,
      Selector target,
      String expectedFormula,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    SheetInspectionResult.CellsResult cellsResult =
        (SheetInspectionResult.CellsResult)
            observations.executeObservation(
                stepId,
                target,
                new SheetIntrospectionQuery.GetCells(
                    java.util.Optional.of(CellReadProjection.of(CellReadFacet.FORMULA))),
                workbook,
                workbookLocation);
    if (cellsResult.cells().isEmpty()) {
      return AssertionEvaluation.fail(
          List.of(cellsResult), "EXPECT_FORMULA_TEXT resolved no matching cells to compare");
    }
    List<String> mismatches =
        cellsResult.cells().stream()
            .filter(
                cell ->
                    !(cell
                            instanceof
                            dev.erst.gridgrind.core.model.value.CellReport.FormulaReport
                                formulaReport)
                        || !formulaReport.formula().orElseThrow().equals(expectedFormula))
            .map(cell -> cell.address().value())
            .toList();
    return mismatches.isEmpty()
        ? AssertionEvaluation.pass(List.of(cellsResult))
        : AssertionEvaluation.fail(
            List.of(cellsResult),
            "EXPECT_FORMULA_TEXT mismatched formula cells at " + String.join(", ", mismatches));
  }

  AssertionEvaluation evaluateCellStyle(
      StepId stepId,
      Selector target,
      CellStyleReport expectedStyle,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    SheetInspectionResult.CellsResult cellsResult =
        (SheetInspectionResult.CellsResult)
            observations.executeObservation(
                stepId,
                target,
                new SheetIntrospectionQuery.GetCells(
                    java.util.Optional.of(CellReadProjection.of(CellReadFacet.STYLE))),
                workbook,
                workbookLocation);
    if (cellsResult.cells().isEmpty()) {
      return AssertionEvaluation.fail(
          List.of(cellsResult), "EXPECT_CELL_STYLE resolved no matching cells to compare");
    }
    List<String> mismatches =
        cellsResult.cells().stream()
            .filter(cell -> !cell.style().orElseThrow().equals(expectedStyle))
            .map(cell -> cell.address().value())
            .toList();
    return mismatches.isEmpty()
        ? AssertionEvaluation.pass(List.of(cellsResult))
        : AssertionEvaluation.fail(
            List.of(cellsResult),
            "EXPECT_CELL_STYLE mismatched style snapshots at " + String.join(", ", mismatches));
  }

  static boolean matchesCellValue(
      dev.erst.gridgrind.core.model.value.CellReport cell, CellValue expectedValue) {
    if (cell instanceof dev.erst.gridgrind.core.model.value.CellReport.FormulaReport formula) {
      return formula.value().map(expectedValue::equals).orElse(false);
    }
    if (cell instanceof dev.erst.gridgrind.core.model.value.CellReport.BlankReport) {
      return expectedValue instanceof CellValue.Blank;
    }
    return cell.type().equals(expectedValue.type())
        && cell.value().orElseThrow().equals(expectedValue);
  }
}
