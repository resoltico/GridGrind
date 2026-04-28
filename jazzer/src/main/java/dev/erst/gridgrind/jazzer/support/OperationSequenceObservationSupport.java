package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.nextNamedRangeName;
import static dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.nextPivotTableSelector;
import static dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.nextTableSelector;
import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.assertThat;

import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import java.util.List;

/** Shared selector and assertion generation support for observation-side Jazzer workflows. */
final class OperationSequenceObservationSupport {
  private OperationSequenceObservationSupport() {}

  static ProtocolStepSupport.PendingAssertion nextAssertion(
      GridGrindFuzzData data,
      int index,
      String primarySheet,
      String secondarySheet,
      String workbookNamedRange,
      String sheetNamedRange,
      String pivotTableName) {
    String stepId = "assert-" + index;
    String targetSheet = data.consumeBoolean() ? primarySheet : secondarySheet;
    return switch (selectorSlot(nextSelectorByte(data)) % 7) {
      case 0 -> assertThat(stepId, new WorkbookSelector.Current(), nextWorkbookAssertion(data));
      case 1 ->
          assertThat(
              stepId,
              new CellSelector.ByAddresses(
                  targetSheet,
                  List.of(
                      FuzzDataDecoders.nextNonBlankCellAddress(data, true),
                      data.consumeBoolean() ? "A1" : "C2")),
              nextCellAssertion(data));
      case 2 ->
          assertThat(
              stepId,
              nextNamedRangeSelection(data, targetSheet, workbookNamedRange, sheetNamedRange, true),
              nextNamedRangeAssertion(data));
      case 3 ->
          assertThat(
              stepId, new SheetSelector.ByName(targetSheet), nextSheetAssertion(data, targetSheet));
      case 4 ->
          assertThat(
              stepId,
              nextTableSelector(data, primarySheet, secondarySheet),
              nextTableAssertion(data));
      case 5 ->
          assertThat(
              stepId,
              nextPivotTableSelector(data, primarySheet, secondarySheet, pivotTableName, true),
              nextPivotTableAssertion(data));
      default ->
          assertThat(stepId, new ChartSelector.AllOnSheet(targetSheet), nextChartAssertion(data));
    };
  }

  static NamedRangeSelector nextNamedRangeSelection(
      GridGrindFuzzData data,
      String sheetName,
      String workbookNamedRange,
      String sheetNamedRange,
      boolean validName) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new NamedRangeSelector.All();
      default ->
          new NamedRangeSelector.AnyOf(
              List.of(
                  data.consumeBoolean()
                      ? new NamedRangeSelector.WorkbookScope(
                          validName ? workbookNamedRange : nextNamedRangeName(data, false))
                      : new NamedRangeSelector.SheetScope(
                          validName ? sheetNamedRange : nextNamedRangeName(data, false),
                          sheetName)));
    };
  }

  static CellSelector nextCellSelector(
      GridGrindFuzzData data, String sheetName, boolean validAddress) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new CellSelector.AllUsedInSheet(sheetName);
      default -> new CellSelector.ByAddresses(sheetName, nextReadAddresses(data, validAddress));
    };
  }

  static SheetSelector nextSheetSelector(
      GridGrindFuzzData data, String primarySheet, String secondarySheet) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new SheetSelector.All();
      default ->
          new SheetSelector.ByNames(
              data.consumeBoolean()
                  ? List.of(primarySheet, secondarySheet)
                  : List.of(secondarySheet, primarySheet));
    };
  }

  static List<String> nextReadAddresses(GridGrindFuzzData data, boolean validAddress) {
    String first = FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress);
    String second = FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress);
    if (first.equals(second)) {
      second = validAddress ? ("A1".equals(first) ? "B2" : "A1") : "ZZZ999999";
    }
    return List.of(first, second);
  }

  private static Assertion nextWorkbookAssertion(GridGrindFuzzData data) {
    Assertion primary =
        switch (selectorSlot(nextSelectorByte(data)) % 3) {
          case 0 ->
              new Assertion.WorkbookProtectionFacts(
                  new dev.erst.gridgrind.contract.dto.WorkbookProtectionReport(
                      false, false, false, false, false));
          case 1 ->
              new Assertion.AnalysisMaxSeverity(
                  new InspectionQuery.AnalyzeWorkbookFindings(), nextMaximumSeverity(data));
          default ->
              new Assertion.AnalysisFindingAbsent(
                  new InspectionQuery.AnalyzeWorkbookFindings(),
                  AnalysisFindingCode.FORMULA_EXTERNAL_REFERENCE,
                  nextOptionalSeverity(data),
                  null);
        };
    Assertion alternate =
        new Assertion.AnalysisFindingPresent(
            new InspectionQuery.AnalyzeWorkbookFindings(),
            AnalysisFindingCode.NAMED_RANGE_BROKEN_REFERENCE,
            null,
            null);
    return maybeComposeAssertion(data, primary, alternate);
  }

  private static Assertion nextCellAssertion(GridGrindFuzzData data) {
    Assertion primary =
        switch (selectorSlot(nextSelectorByte(data)) % 4) {
          case 0 -> new Assertion.CellValue(nextExpectedCellValue(data));
          case 1 -> new Assertion.DisplayValue("display-" + data.consumeInt(0, 9));
          case 2 -> new Assertion.FormulaText(data.consumeBoolean() ? "B2*2" : "SUM(B2:B4)");
          default -> new Assertion.DisplayValue("Report");
        };
    Assertion alternate = new Assertion.CellValue(new ExpectedCellValue.Text("Jan"));
    return maybeComposeAssertion(data, primary, alternate);
  }

  private static Assertion nextNamedRangeAssertion(GridGrindFuzzData data) {
    Assertion primary =
        switch (selectorSlot(nextSelectorByte(data)) % 3) {
          case 0 -> new Assertion.NamedRangePresent();
          case 1 ->
              new Assertion.AnalysisMaxSeverity(
                  new InspectionQuery.AnalyzeNamedRangeHealth(), nextMaximumSeverity(data));
          default ->
              new Assertion.AnalysisFindingAbsent(
                  new InspectionQuery.AnalyzeNamedRangeHealth(),
                  AnalysisFindingCode.NAMED_RANGE_BROKEN_REFERENCE,
                  nextOptionalSeverity(data),
                  null);
        };
    Assertion alternate = new Assertion.NamedRangeAbsent();
    return maybeComposeAssertion(data, primary, alternate);
  }

  private static Assertion nextSheetAssertion(GridGrindFuzzData data, String sheetName) {
    Assertion primary =
        switch (selectorSlot(nextSelectorByte(data)) % 6) {
          case 0 ->
              new Assertion.AnalysisMaxSeverity(
                  new InspectionQuery.AnalyzeFormulaHealth(), nextMaximumSeverity(data));
          case 1 ->
              new Assertion.AnalysisFindingAbsent(
                  new InspectionQuery.AnalyzeDataValidationHealth(),
                  AnalysisFindingCode.DATA_VALIDATION_BROKEN_FORMULA,
                  nextOptionalSeverity(data),
                  null);
          case 2 ->
              new Assertion.AnalysisFindingAbsent(
                  new InspectionQuery.AnalyzeConditionalFormattingHealth(),
                  AnalysisFindingCode.CONDITIONAL_FORMATTING_BROKEN_FORMULA,
                  nextOptionalSeverity(data),
                  null);
          case 3 ->
              new Assertion.AnalysisFindingAbsent(
                  new InspectionQuery.AnalyzeAutofilterHealth(),
                  AnalysisFindingCode.AUTOFILTER_INVALID_RANGE,
                  nextOptionalSeverity(data),
                  null);
          case 4 ->
              new Assertion.AnalysisFindingAbsent(
                  new InspectionQuery.AnalyzeHyperlinkHealth(),
                  AnalysisFindingCode.HYPERLINK_MALFORMED_TARGET,
                  nextOptionalSeverity(data),
                  null);
          default ->
              new Assertion.SheetStructureFacts(
                  new dev.erst.gridgrind.contract.dto.GridGrindResponse.SheetSummaryReport(
                      sheetName,
                      ExcelSheetVisibility.VISIBLE,
                      new dev.erst.gridgrind.contract.dto.GridGrindResponse.SheetProtectionReport
                          .Unprotected(),
                      0,
                      -1,
                      -1));
        };
    Assertion alternate =
        new Assertion.AnalysisMaxSeverity(
            new InspectionQuery.AnalyzeFormulaHealth(), AnalysisSeverity.WARNING);
    return maybeComposeAssertion(data, primary, alternate);
  }

  private static Assertion nextTableAssertion(GridGrindFuzzData data) {
    Assertion primary =
        switch (selectorSlot(nextSelectorByte(data)) % 3) {
          case 0 -> new Assertion.TablePresent();
          case 1 ->
              new Assertion.AnalysisMaxSeverity(
                  new InspectionQuery.AnalyzeTableHealth(), nextMaximumSeverity(data));
          default ->
              new Assertion.AnalysisFindingAbsent(
                  new InspectionQuery.AnalyzeTableHealth(),
                  AnalysisFindingCode.TABLE_BROKEN_REFERENCE,
                  nextOptionalSeverity(data),
                  null);
        };
    Assertion alternate = new Assertion.TableAbsent();
    return maybeComposeAssertion(data, primary, alternate);
  }

  private static Assertion nextPivotTableAssertion(GridGrindFuzzData data) {
    Assertion primary =
        switch (selectorSlot(nextSelectorByte(data)) % 3) {
          case 0 -> new Assertion.PivotTablePresent();
          case 1 ->
              new Assertion.AnalysisMaxSeverity(
                  new InspectionQuery.AnalyzePivotTableHealth(), nextMaximumSeverity(data));
          default ->
              new Assertion.AnalysisFindingAbsent(
                  new InspectionQuery.AnalyzePivotTableHealth(),
                  AnalysisFindingCode.PIVOT_TABLE_BROKEN_SOURCE,
                  nextOptionalSeverity(data),
                  null);
        };
    Assertion alternate = new Assertion.PivotTableAbsent();
    return maybeComposeAssertion(data, primary, alternate);
  }

  private static Assertion nextChartAssertion(GridGrindFuzzData data) {
    Assertion primary =
        data.consumeBoolean() ? new Assertion.ChartPresent() : new Assertion.ChartAbsent();
    Assertion alternate =
        data.consumeBoolean() ? new Assertion.ChartAbsent() : new Assertion.ChartPresent();
    return maybeComposeAssertion(data, primary, alternate);
  }

  private static Assertion maybeComposeAssertion(
      GridGrindFuzzData data, Assertion primary, Assertion alternate) {
    return switch (selectorSlot(nextSelectorByte(data)) % 4) {
      case 0 -> primary;
      case 1 -> new Assertion.Not(primary);
      case 2 -> new Assertion.AllOf(List.of(primary, alternate));
      default -> new Assertion.AnyOf(List.of(primary, alternate));
    };
  }

  private static AnalysisSeverity nextMaximumSeverity(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 -> AnalysisSeverity.INFO;
      case 1 -> AnalysisSeverity.WARNING;
      default -> AnalysisSeverity.ERROR;
    };
  }

  private static AnalysisSeverity nextOptionalSeverity(GridGrindFuzzData data) {
    return data.consumeBoolean() ? nextMaximumSeverity(data) : null;
  }

  private static ExpectedCellValue nextExpectedCellValue(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 5) {
      case 0 -> new ExpectedCellValue.Blank();
      case 1 -> new ExpectedCellValue.Text("seed-" + data.consumeInt(0, 9));
      case 2 -> new ExpectedCellValue.NumericValue(data.consumeRegularDouble(0.0d, 10.0d));
      case 3 -> new ExpectedCellValue.BooleanValue(data.consumeBoolean());
      default -> new ExpectedCellValue.ErrorValue("#REF!");
    };
  }

  private static int nextSelectorByte(GridGrindFuzzData data) {
    return Byte.toUnsignedInt(data.consumeByte());
  }

  private static int selectorSlot(int selector) {
    return selector & 0x0F;
  }
}
