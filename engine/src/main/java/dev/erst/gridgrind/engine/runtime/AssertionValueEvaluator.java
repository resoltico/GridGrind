package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.CellScalarValue;
import dev.erst.gridgrind.contract.dto.CellStyleReport;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.query.SheetInspectionResult;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
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
                : assertionType
                    + " observed "
                    + count
                    + " matching workbook "
                    + (count == 1 ? "entity" : "entities"));
  }

  AssertionEvaluation evaluateCellValue(
      String stepId,
      Selector target,
      CellScalarValue expectedValue,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    SheetInspectionResult.CellsResult cellsResult =
        (SheetInspectionResult.CellsResult)
            observations.executeObservation(
                stepId, target, new SheetIntrospectionQuery.GetCells(), workbook, workbookLocation);
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
    SheetInspectionResult.CellsResult cellsResult =
        (SheetInspectionResult.CellsResult)
            observations.executeObservation(
                stepId, target, new SheetIntrospectionQuery.GetCells(), workbook, workbookLocation);
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
    SheetInspectionResult.CellsResult cellsResult =
        (SheetInspectionResult.CellsResult)
            observations.executeObservation(
                stepId, target, new SheetIntrospectionQuery.GetCells(), workbook, workbookLocation);
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
      CellStyleReport expectedStyle,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    SheetInspectionResult.CellsResult cellsResult =
        (SheetInspectionResult.CellsResult)
            observations.executeObservation(
                stepId, target, new SheetIntrospectionQuery.GetCells(), workbook, workbookLocation);
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
      dev.erst.gridgrind.contract.dto.CellReport cell, CellScalarValue expectedValue) {
    if (cell instanceof dev.erst.gridgrind.contract.dto.CellReport.FormulaReport formulaReport) {
      return matchesCellValue(formulaReport.evaluation(), expectedValue);
    }
    return switch (expectedValue) {
      case CellScalarValue.Blank _ ->
          cell instanceof dev.erst.gridgrind.contract.dto.CellReport.BlankReport;
      case CellScalarValue.Text expectedText ->
          cell instanceof dev.erst.gridgrind.contract.dto.CellReport.TextReport textReport
              && textReport.stringValue().equals(expectedText.text());
      case CellScalarValue.NumberValue expectedNumber ->
          cell instanceof dev.erst.gridgrind.contract.dto.CellReport.NumberReport numberReport
              && Double.compare(numberReport.numberValue(), expectedNumber.number()) == 0;
      case CellScalarValue.BooleanValue expectedBoolean ->
          cell instanceof dev.erst.gridgrind.contract.dto.CellReport.BooleanReport booleanReport
              && booleanReport.booleanValue().equals(expectedBoolean.bool());
      case CellScalarValue.ErrorValue expectedError ->
          cell instanceof dev.erst.gridgrind.contract.dto.CellReport.ErrorReport errorReport
              && errorReport.errorValue().equals(expectedError.error());
    };
  }
}
