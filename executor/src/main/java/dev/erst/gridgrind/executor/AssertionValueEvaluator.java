package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.GridGrindWorkbookSurfaceReports;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.Selector;
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
      String stepId,
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
                : assertionType + " observed " + count + " matching workbook entities");
  }

  AssertionEvaluation evaluateCellValue(
      String stepId,
      Selector target,
      ExpectedCellValue expectedValue,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.CellsResult cellsResult =
        (InspectionResult.CellsResult)
            observations.executeObservation(
                stepId, target, new InspectionQuery.GetCells(), workbook, workbookLocation);
    if (cellsResult.cells().isEmpty()) {
      return AssertionEvaluation.fail(
          List.of(cellsResult), "EXPECT_CELL_VALUE resolved no matching cells to compare");
    }
    List<String> mismatches =
        cellsResult.cells().stream()
            .filter(cell -> !matchesCellValue(cell, expectedValue))
            .map(dev.erst.gridgrind.contract.dto.CellReport::address)
            .toList();
    return mismatches.isEmpty()
        ? AssertionEvaluation.pass(List.of(cellsResult))
        : AssertionEvaluation.fail(
            List.of(cellsResult),
            "EXPECT_CELL_VALUE mismatched effective values at " + String.join(", ", mismatches));
  }

  AssertionEvaluation evaluateDisplayValue(
      String stepId,
      Selector target,
      String expectedDisplayValue,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.CellsResult cellsResult =
        (InspectionResult.CellsResult)
            observations.executeObservation(
                stepId, target, new InspectionQuery.GetCells(), workbook, workbookLocation);
    if (cellsResult.cells().isEmpty()) {
      return AssertionEvaluation.fail(
          List.of(cellsResult), "EXPECT_DISPLAY_VALUE resolved no matching cells to compare");
    }
    List<String> mismatches =
        cellsResult.cells().stream()
            .filter(cell -> !cell.displayValue().equals(expectedDisplayValue))
            .map(dev.erst.gridgrind.contract.dto.CellReport::address)
            .toList();
    return mismatches.isEmpty()
        ? AssertionEvaluation.pass(List.of(cellsResult))
        : AssertionEvaluation.fail(
            List.of(cellsResult),
            "EXPECT_DISPLAY_VALUE mismatched formatted values at " + String.join(", ", mismatches));
  }

  AssertionEvaluation evaluateFormulaText(
      String stepId,
      Selector target,
      String expectedFormula,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.CellsResult cellsResult =
        (InspectionResult.CellsResult)
            observations.executeObservation(
                stepId, target, new InspectionQuery.GetCells(), workbook, workbookLocation);
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
                            dev.erst.gridgrind.contract.dto.CellReport.FormulaReport formulaReport)
                        || !formulaReport.formula().equals(expectedFormula))
            .map(dev.erst.gridgrind.contract.dto.CellReport::address)
            .toList();
    return mismatches.isEmpty()
        ? AssertionEvaluation.pass(List.of(cellsResult))
        : AssertionEvaluation.fail(
            List.of(cellsResult),
            "EXPECT_FORMULA_TEXT mismatched formula cells at " + String.join(", ", mismatches));
  }

  AssertionEvaluation evaluateCellStyle(
      String stepId,
      Selector target,
      GridGrindWorkbookSurfaceReports.CellStyleReport expectedStyle,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.CellsResult cellsResult =
        (InspectionResult.CellsResult)
            observations.executeObservation(
                stepId, target, new InspectionQuery.GetCells(), workbook, workbookLocation);
    if (cellsResult.cells().isEmpty()) {
      return AssertionEvaluation.fail(
          List.of(cellsResult), "EXPECT_CELL_STYLE resolved no matching cells to compare");
    }
    List<String> mismatches =
        cellsResult.cells().stream()
            .filter(cell -> !cell.style().equals(expectedStyle))
            .map(dev.erst.gridgrind.contract.dto.CellReport::address)
            .toList();
    return mismatches.isEmpty()
        ? AssertionEvaluation.pass(List.of(cellsResult))
        : AssertionEvaluation.fail(
            List.of(cellsResult),
            "EXPECT_CELL_STYLE mismatched style snapshots at " + String.join(", ", mismatches));
  }

  static boolean matchesCellValue(
      dev.erst.gridgrind.contract.dto.CellReport cell, ExpectedCellValue expectedValue) {
    if (cell instanceof dev.erst.gridgrind.contract.dto.CellReport.FormulaReport formulaReport) {
      return matchesCellValue(formulaReport.evaluation(), expectedValue);
    }
    return switch (expectedValue) {
      case ExpectedCellValue.Blank _ ->
          cell instanceof dev.erst.gridgrind.contract.dto.CellReport.BlankReport;
      case ExpectedCellValue.Text expectedText ->
          cell instanceof dev.erst.gridgrind.contract.dto.CellReport.TextReport textReport
              && textReport.stringValue().equals(expectedText.text());
      case ExpectedCellValue.NumericValue expectedNumber ->
          cell instanceof dev.erst.gridgrind.contract.dto.CellReport.NumberReport numberReport
              && Double.compare(numberReport.numberValue(), expectedNumber.number()) == 0;
      case ExpectedCellValue.BooleanValue expectedBoolean ->
          cell instanceof dev.erst.gridgrind.contract.dto.CellReport.BooleanReport booleanReport
              && booleanReport.booleanValue().equals(expectedBoolean.value());
      case ExpectedCellValue.ErrorValue expectedError ->
          cell instanceof dev.erst.gridgrind.contract.dto.CellReport.ErrorReport errorReport
              && errorReport.errorValue().equals(expectedError.error());
    };
  }
}
