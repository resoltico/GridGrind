package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.protocol.dto.*;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadResult;
import java.io.IOException;
import java.time.DateTimeException;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Focused coverage tests for executor-owned protocol-to-engine translation seams. */
class DefaultGridGrindRequestExecutorCoverageTest {
  @Test
  void mapsProtocolEnumsIntoEngineCommands() {
    for (ExcelHorizontalAlignment alignment : ExcelHorizontalAlignment.values()) {
      assertProtocolHorizontalAlignmentMapping(alignment);
    }

    for (ExcelVerticalAlignment alignment : ExcelVerticalAlignment.values()) {
      assertProtocolVerticalAlignmentMapping(alignment);
    }

    for (ExcelBorderStyle style : ExcelBorderStyle.values()) {
      assertProtocolBorderStyleMapping(style);
    }

    for (ExcelComparisonOperator operator : ExcelComparisonOperator.values()) {
      assertProtocolComparisonOperatorMapping(operator);
    }

    for (ExcelSheetVisibility visibility : ExcelSheetVisibility.values()) {
      assertProtocolSheetVisibilityMapping(visibility);
    }

    for (ExcelPaneRegion region : ExcelPaneRegion.values()) {
      assertProtocolPaneRegionMapping(region);
    }
  }

  @Test
  void mapsEngineFactsBackIntoProtocolEnumsAndReports() {
    for (ExcelSheetVisibility visibility : ExcelSheetVisibility.values()) {
      assertEngineSheetVisibilityMapping(visibility);
    }

    for (ExcelPaneRegion region : ExcelPaneRegion.values()) {
      assertEnginePaneRegionMapping(region);
    }
  }

  @Test
  void mapsDataValidationAndConditionalFormattingEnumsBackIntoProtocolReports() {
    for (ExcelDataValidationErrorStyle style : ExcelDataValidationErrorStyle.values()) {
      assertEngineDataValidationErrorStyleMapping(style);
    }

    for (ExcelComparisonOperator operator : ExcelComparisonOperator.values()) {
      assertEngineComparisonOperatorMapping(operator);
    }

    for (ExcelConditionalFormattingThresholdType type :
        ExcelConditionalFormattingThresholdType.values()) {
      assertEngineThresholdTypeMapping(type);
    }

    for (ExcelConditionalFormattingIconSet iconSet : ExcelConditionalFormattingIconSet.values()) {
      assertEngineConditionalFormattingIconSetMapping(iconSet);
    }

    for (ExcelConditionalFormattingUnsupportedFeature feature :
        ExcelConditionalFormattingUnsupportedFeature.values()) {
      assertEngineConditionalFormattingUnsupportedFeatureMapping(feature);
    }

    ConditionalFormattingRuleReport.CellValueRule cellValueRule =
        assertInstanceOf(
            ConditionalFormattingRuleReport.CellValueRule.class,
            WorkbookReadResultConverter.toConditionalFormattingRuleReport(
                new ExcelConditionalFormattingRuleSnapshot.CellValueRule(
                    1, false, ExcelComparisonOperator.GREATER_OR_EQUAL, "1", null, null)));
    assertEquals(ExcelComparisonOperator.GREATER_OR_EQUAL, cellValueRule.operator());
  }

  @Test
  void classifiesEngineExceptionsAndEnrichesProblemContexts() {
    assertEquals(
        GridGrindProblemCode.WORKBOOK_NOT_FOUND,
        DefaultGridGrindRequestExecutor.problemCodeFor(
            new WorkbookNotFoundException(java.nio.file.Path.of("/tmp/missing.xlsx"))));
    assertEquals(
        GridGrindProblemCode.SHEET_NOT_FOUND,
        DefaultGridGrindRequestExecutor.problemCodeFor(new SheetNotFoundException("Budget")));
    assertEquals(
        GridGrindProblemCode.NAMED_RANGE_NOT_FOUND,
        DefaultGridGrindRequestExecutor.problemCodeFor(
            new NamedRangeNotFoundException(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
    assertEquals(
        GridGrindProblemCode.CELL_NOT_FOUND,
        DefaultGridGrindRequestExecutor.problemCodeFor(new CellNotFoundException("B4")));
    assertEquals(
        GridGrindProblemCode.INVALID_CELL_ADDRESS,
        DefaultGridGrindRequestExecutor.problemCodeFor(
            new InvalidCellAddressException("A0", null)));
    assertEquals(
        GridGrindProblemCode.INVALID_RANGE_ADDRESS,
        DefaultGridGrindRequestExecutor.problemCodeFor(
            new InvalidRangeAddressException("A0:B1", null)));
    assertEquals(
        GridGrindProblemCode.UNSUPPORTED_FORMULA,
        DefaultGridGrindRequestExecutor.problemCodeFor(
            new UnsupportedFormulaException("Budget", "B4", "SEQUENCE(2)", "unsupported", null)));
    assertEquals(
        GridGrindProblemCode.INVALID_FORMULA,
        DefaultGridGrindRequestExecutor.problemCodeFor(
            new InvalidFormulaException("Budget", "B4", "SUM(", "invalid", null)));
    assertEquals(
        GridGrindProblemCode.IO_ERROR,
        DefaultGridGrindRequestExecutor.problemCodeFor(new IOException("disk")));
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST,
        DefaultGridGrindRequestExecutor.problemCodeFor(new DateTimeException("bad date")));

    GridGrindResponse.ProblemContext.ApplyOperation applyContext =
        new GridGrindResponse.ProblemContext.ApplyOperation(
            "NEW", "NONE", 0, "SET_CELL", null, null, null, null, null);
    InvalidFormulaException invalidFormula =
        new InvalidFormulaException("Budget", "B4", "SUM(", "invalid", null);
    GridGrindResponse.ProblemContext.ApplyOperation enrichedApply =
        assertInstanceOf(
            GridGrindResponse.ProblemContext.ApplyOperation.class,
            DefaultGridGrindRequestExecutor.enrichContext(applyContext, invalidFormula));
    assertEquals("Budget", enrichedApply.sheetName());
    assertEquals("B4", enrichedApply.address());
    assertEquals("SUM(", enrichedApply.formula());

    GridGrindResponse.ProblemContext.ExecuteRead readContext =
        new GridGrindResponse.ProblemContext.ExecuteRead(
            "NEW", "NONE", 0, "GET_NAMED_RANGES", "ranges", null, null, null, null);
    GridGrindResponse.ProblemContext.ExecuteRead enrichedRead =
        assertInstanceOf(
            GridGrindResponse.ProblemContext.ExecuteRead.class,
            DefaultGridGrindRequestExecutor.enrichContext(
                readContext,
                new NamedRangeNotFoundException(
                    "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
    assertEquals("BudgetTotal", enrichedRead.namedRangeName());

    GridGrindResponse.ProblemContext.ExecuteRead enrichedReadFormula =
        assertInstanceOf(
            GridGrindResponse.ProblemContext.ExecuteRead.class,
            DefaultGridGrindRequestExecutor.enrichContext(readContext, invalidFormula));
    assertEquals("Budget", enrichedReadFormula.sheetName());
    assertEquals("B4", enrichedReadFormula.address());
    assertEquals("SUM(", enrichedReadFormula.formula());

    GridGrindResponse.ProblemContext.ApplyOperation enrichedApplyRange =
        assertInstanceOf(
            GridGrindResponse.ProblemContext.ApplyOperation.class,
            DefaultGridGrindRequestExecutor.enrichContext(
                applyContext, new InvalidRangeAddressException("A1:B2", null)));
    assertEquals("A1:B2", enrichedApplyRange.range());

    GridGrindResponse.ProblemContext.ApplyOperation enrichedApplyNamedRange =
        assertInstanceOf(
            GridGrindResponse.ProblemContext.ApplyOperation.class,
            DefaultGridGrindRequestExecutor.enrichContext(
                applyContext,
                new NamedRangeNotFoundException(
                    "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
    assertEquals("BudgetTotal", enrichedApplyNamedRange.namedRangeName());

    GridGrindResponse.ProblemContext.ReadRequest readRequest =
        new GridGrindResponse.ProblemContext.ReadRequest("id", "$.reads[0]", 3, 7);
    assertSame(
        readRequest, DefaultGridGrindRequestExecutor.enrichContext(readRequest, invalidFormula));
    GridGrindResponse.ProblemContext.ParseArguments parseArguments =
        new GridGrindResponse.ProblemContext.ParseArguments("--request");
    assertSame(
        parseArguments,
        DefaultGridGrindRequestExecutor.enrichContext(parseArguments, invalidFormula));
    GridGrindResponse.ProblemContext.ValidateRequest validateRequest =
        new GridGrindResponse.ProblemContext.ValidateRequest("NEW", "NONE");
    assertSame(
        validateRequest,
        DefaultGridGrindRequestExecutor.enrichContext(validateRequest, invalidFormula));
    GridGrindResponse.ProblemContext.WriteResponse writeResponse =
        new GridGrindResponse.ProblemContext.WriteResponse("/tmp/out.json");
    assertSame(
        writeResponse,
        DefaultGridGrindRequestExecutor.enrichContext(writeResponse, invalidFormula));

    assertEquals("B2", DefaultGridGrindRequestExecutor.addressFor(new CellNotFoundException("B2")));
    assertEquals(
        "C3",
        DefaultGridGrindRequestExecutor.addressFor(new InvalidCellAddressException("C3", null)));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(invalidFormula));
    assertEquals("SUM(", DefaultGridGrindRequestExecutor.formulaFor(invalidFormula));
    assertEquals(
        "A1:B2",
        DefaultGridGrindRequestExecutor.rangeFor(new InvalidRangeAddressException("A1:B2", null)));
    assertEquals(
        "BudgetTotal",
        DefaultGridGrindRequestExecutor.namedRangeNameFor(
            new NamedRangeNotFoundException(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
  }

  private static CellStyleInput styleInput(
      ExcelHorizontalAlignment horizontalAlignment,
      ExcelVerticalAlignment verticalAlignment,
      ExcelBorderStyle borderStyle) {
    return new CellStyleInput(
        null,
        null,
        null,
        null,
        horizontalAlignment,
        verticalAlignment,
        null,
        null,
        null,
        null,
        null,
        null,
        new CellBorderInput(
            null,
            new CellBorderSideInput(borderStyle),
            new CellBorderSideInput(borderStyle),
            new CellBorderSideInput(borderStyle),
            new CellBorderSideInput(borderStyle)));
  }

  private static String comparisonUpperBound(ExcelComparisonOperator operator) {
    return switch (operator) {
      case BETWEEN, NOT_BETWEEN -> "2";
      default -> null;
    };
  }

  private static String comparisonUpperBound(String operatorName) {
    return switch (operatorName) {
      case "BETWEEN", "NOT_BETWEEN" -> "2";
      default -> null;
    };
  }

  private static void assertProtocolHorizontalAlignmentMapping(ExcelHorizontalAlignment alignment) {
    WorkbookCommand.ApplyStyle command =
        assertInstanceOf(
            WorkbookCommand.ApplyStyle.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.ApplyStyle(
                    "Budget",
                    "A1",
                    styleInput(alignment, ExcelVerticalAlignment.TOP, ExcelBorderStyle.THIN))));
    assertEquals(alignment, command.style().horizontalAlignment());
  }

  private static void assertProtocolVerticalAlignmentMapping(ExcelVerticalAlignment alignment) {
    WorkbookCommand.ApplyStyle command =
        assertInstanceOf(
            WorkbookCommand.ApplyStyle.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.ApplyStyle(
                    "Budget",
                    "A1",
                    styleInput(ExcelHorizontalAlignment.LEFT, alignment, ExcelBorderStyle.THIN))));
    assertEquals(alignment, command.style().verticalAlignment());
  }

  private static void assertProtocolBorderStyleMapping(ExcelBorderStyle style) {
    WorkbookCommand.ApplyStyle command =
        assertInstanceOf(
            WorkbookCommand.ApplyStyle.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.ApplyStyle(
                    "Budget",
                    "A1",
                    styleInput(ExcelHorizontalAlignment.LEFT, ExcelVerticalAlignment.TOP, style))));
    assertEquals(style, command.style().border().top().style());
  }

  private static void assertProtocolComparisonOperatorMapping(ExcelComparisonOperator operator) {
    WorkbookCommand.SetDataValidation command =
        assertInstanceOf(
            WorkbookCommand.SetDataValidation.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetDataValidation(
                    "Budget",
                    "A1",
                    new DataValidationInput(
                        new DataValidationRuleInput.WholeNumber(
                            operator, "1", comparisonUpperBound(operator)),
                        false,
                        false,
                        null,
                        null))));
    ExcelDataValidationRule.WholeNumber rule =
        assertInstanceOf(ExcelDataValidationRule.WholeNumber.class, command.validation().rule());
    assertEquals(operator, rule.operator());
  }

  private static void assertProtocolSheetVisibilityMapping(ExcelSheetVisibility visibility) {
    WorkbookCommand.SetSheetVisibility command =
        assertInstanceOf(
            WorkbookCommand.SetSheetVisibility.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetSheetVisibility("Budget", visibility)));
    assertEquals(visibility, command.visibility());
  }

  private static void assertProtocolPaneRegionMapping(ExcelPaneRegion region) {
    WorkbookCommand.SetSheetPane command =
        assertInstanceOf(
            WorkbookCommand.SetSheetPane.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetSheetPane(
                    "Budget", new PaneInput.Split(120, 240, 0, 0, region))));
    ExcelSheetPane.Split split = assertInstanceOf(ExcelSheetPane.Split.class, command.pane());
    assertEquals(region, split.activePane());
  }

  private static void assertEngineSheetVisibilityMapping(ExcelSheetVisibility visibility) {
    WorkbookReadResult.SheetSummaryResult result =
        assertInstanceOf(
            WorkbookReadResult.SheetSummaryResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummaryResult(
                    "sheet",
                    new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummary(
                        "Budget",
                        visibility,
                        new dev.erst.gridgrind.excel.WorkbookReadResult.SheetProtection
                            .Unprotected(),
                        0,
                        -1,
                        -1))));
    assertEquals(visibility, result.sheet().visibility());
  }

  private static void assertEnginePaneRegionMapping(ExcelPaneRegion region) {
    WorkbookReadResult.SheetLayoutResult result =
        assertInstanceOf(
            WorkbookReadResult.SheetLayoutResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayoutResult(
                    "layout",
                    new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayout(
                        "Budget",
                        new ExcelSheetPane.Split(120, 240, 0, 0, region),
                        100,
                        List.of(
                            new dev.erst.gridgrind.excel.WorkbookReadResult.ColumnLayout(0, 12.0)),
                        List.of(
                            new dev.erst.gridgrind.excel.WorkbookReadResult.RowLayout(0, 15.0))))));
    PaneReport.Split split = assertInstanceOf(PaneReport.Split.class, result.layout().pane());
    assertEquals(region, split.activePane());
  }

  private static void assertEngineDataValidationErrorStyleMapping(
      ExcelDataValidationErrorStyle style) {
    DataValidationEntryReport.Supported entry =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class,
            WorkbookReadResultConverter.toDataValidationEntryReport(
                new ExcelDataValidationSnapshot.Supported(
                    List.of("A1"),
                    new ExcelDataValidationDefinition(
                        new ExcelDataValidationRule.WholeNumber(
                            ExcelComparisonOperator.EQUAL, "1", null),
                        false,
                        false,
                        null,
                        new ExcelDataValidationErrorAlert(style, "Title", "Text", true)))));
    assertEquals(style, entry.validation().errorAlert().style());
    DataValidationRuleInput.WholeNumber rule =
        assertInstanceOf(DataValidationRuleInput.WholeNumber.class, entry.validation().rule());
    assertEquals(ExcelComparisonOperator.EQUAL, rule.operator());
  }

  private static void assertEngineComparisonOperatorMapping(ExcelComparisonOperator operator) {
    DataValidationEntryReport.Supported entry =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class,
            WorkbookReadResultConverter.toDataValidationEntryReport(
                new ExcelDataValidationSnapshot.Supported(
                    List.of("A1"),
                    new ExcelDataValidationDefinition(
                        new ExcelDataValidationRule.WholeNumber(
                            operator, "1", comparisonUpperBound(operator.name())),
                        false,
                        false,
                        null,
                        null))));
    DataValidationRuleInput.WholeNumber rule =
        assertInstanceOf(DataValidationRuleInput.WholeNumber.class, entry.validation().rule());
    assertEquals(operator, rule.operator());
  }

  private static void assertEngineThresholdTypeMapping(
      ExcelConditionalFormattingThresholdType type) {
    ConditionalFormattingThresholdReport threshold =
        WorkbookReadResultConverter.toConditionalFormattingThresholdReport(
            new ExcelConditionalFormattingThresholdSnapshot(type, "A1", 2.0));
    assertEquals(type, threshold.type());
  }

  private static void assertEngineConditionalFormattingIconSetMapping(
      ExcelConditionalFormattingIconSet iconSet) {
    ConditionalFormattingRuleReport.IconSetRule rule =
        assertInstanceOf(
            ConditionalFormattingRuleReport.IconSetRule.class,
            WorkbookReadResultConverter.toConditionalFormattingRuleReport(
                new ExcelConditionalFormattingRuleSnapshot.IconSetRule(
                    1,
                    false,
                    iconSet,
                    false,
                    false,
                    List.of(
                        new ExcelConditionalFormattingThresholdSnapshot(
                            ExcelConditionalFormattingThresholdType.MIN, null, null)))));
    assertEquals(iconSet, rule.iconSet());
  }

  private static void assertEngineConditionalFormattingUnsupportedFeatureMapping(
      ExcelConditionalFormattingUnsupportedFeature feature) {
    DifferentialStyleReport report =
        WorkbookReadResultConverter.toDifferentialStyleReport(
            new ExcelDifferentialStyleSnapshot(
                null, null, null, null, null, null, null, null, null, List.of(feature)));
    assertEquals(List.of(feature), report.unsupportedFeatures());
  }
}
