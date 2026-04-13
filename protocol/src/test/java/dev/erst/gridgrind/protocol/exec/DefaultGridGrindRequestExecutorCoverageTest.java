package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.protocol.dto.*;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadResult;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.nio.file.Files;
import java.nio.file.Path;
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
        GridGrindProblemCode.MISSING_EXTERNAL_WORKBOOK,
        DefaultGridGrindRequestExecutor.problemCodeFor(
            new MissingExternalWorkbookException(
                "Budget", "B4", "[Rates.xlsx]Sheet1!A1", "Rates.xlsx", "missing", null)));
    assertEquals(
        GridGrindProblemCode.WORKBOOK_PASSWORD_REQUIRED,
        DefaultGridGrindRequestExecutor.problemCodeFor(
            new WorkbookPasswordRequiredException(Path.of("/tmp/encrypted.xlsx"))));
    assertEquals(
        GridGrindProblemCode.INVALID_WORKBOOK_PASSWORD,
        DefaultGridGrindRequestExecutor.problemCodeFor(
            new InvalidWorkbookPasswordException(Path.of("/tmp/encrypted.xlsx"))));
    assertEquals(
        GridGrindProblemCode.INVALID_SIGNING_CONFIGURATION,
        DefaultGridGrindRequestExecutor.problemCodeFor(
            new InvalidSigningConfigurationException("bad signing")));
    assertEquals(
        GridGrindProblemCode.WORKBOOK_SECURITY_ERROR,
        DefaultGridGrindRequestExecutor.problemCodeFor(
            new WorkbookSecurityException("crypto failed", null)));
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

  @Test
  void coversFormulaLifecycleOperationContextExtractionAndRejectsPrivateHelperMismatches()
      throws ReflectiveOperationException {
    WorkbookOperation.EvaluateFormulaCells evaluateFormulaCells =
        new WorkbookOperation.EvaluateFormulaCells(
            List.of(new FormulaCellTargetInput("Budget", "B4")));
    WorkbookOperation.ClearFormulaCaches clearFormulaCaches =
        new WorkbookOperation.ClearFormulaCaches();
    RuntimeException failure = new IllegalStateException("boom");

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(evaluateFormulaCells, failure));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(clearFormulaCaches, failure));
    assertNull(DefaultGridGrindRequestExecutor.sheetNameFor(evaluateFormulaCells, failure));
    assertNull(DefaultGridGrindRequestExecutor.sheetNameFor(clearFormulaCaches, failure));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(evaluateFormulaCells, failure));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(clearFormulaCaches, failure));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(evaluateFormulaCells, failure));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(clearFormulaCaches, failure));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(evaluateFormulaCells, failure));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(clearFormulaCaches, failure));

    assertPrivateSheetNameHelperRejects(
        "sheetNameForWorkbookScopeOperation",
        new WorkbookOperation.SetCell("Budget", "A1", new CellInput.Text("x")));
    assertPrivateSheetNameHelperRejects(
        "sheetNameForSheetStructureOperation", new WorkbookOperation.SetActiveSheet("Budget"));
    assertPrivateSheetNameHelperRejects(
        "sheetNameForSheetContentOperation", new WorkbookOperation.EnsureSheet("Budget"));
  }

  @Test
  void coversDrawingReadHelpersDrawingWriteHelpersAndNamedRangeSelectionExtraction() {
    DrawingAnchorInput.TwoCell drawingAnchor =
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(1, 2, 0, 0),
            new DrawingMarkerInput(2, 3, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    PictureDataInput pictureData =
        new PictureDataInput(
            ExcelPictureFormat.PNG,
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");
    WorkbookReadOperation.GetDrawingObjects getDrawingObjects =
        new WorkbookReadOperation.GetDrawingObjects("drawing", "Budget");
    WorkbookReadOperation.GetCharts getCharts =
        new WorkbookReadOperation.GetCharts("charts", "Budget");
    WorkbookReadOperation.GetDrawingObjectPayload getDrawingObjectPayload =
        new WorkbookReadOperation.GetDrawingObjectPayload("payload", "Budget", "OpsPicture");
    WorkbookOperation.SetPicture setPicture =
        new WorkbookOperation.SetPicture(
            "Budget", new PictureInput("OpsPicture", pictureData, drawingAnchor, "Queue preview"));
    WorkbookOperation.SetShape setShape =
        new WorkbookOperation.SetShape(
            "Budget",
            new ShapeInput(
                "OpsShape",
                ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                drawingAnchor,
                "rect",
                "Queue"));
    WorkbookOperation.SetEmbeddedObject setEmbeddedObject =
        new WorkbookOperation.SetEmbeddedObject(
            "Budget",
            new EmbeddedObjectInput(
                "OpsEmbed",
                "Payload",
                "payload.txt",
                "payload.txt",
                "cGF5bG9hZA==",
                pictureData,
                drawingAnchor));
    WorkbookOperation.SetDrawingObjectAnchor setDrawingObjectAnchor =
        new WorkbookOperation.SetDrawingObjectAnchor("Budget", "OpsPicture", drawingAnchor);
    WorkbookOperation.DeleteDrawingObject deleteDrawingObject =
        new WorkbookOperation.DeleteDrawingObject("Budget", "OpsPicture");
    WorkbookOperation.SetChart setChart =
        new WorkbookOperation.SetChart(
            "Budget",
            new ChartInput.Bar(
                "OpsChart",
                drawingAnchor,
                new ChartInput.Title.Text("Roadmap"),
                new ChartInput.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
                ExcelChartDisplayBlanksAs.SPAN,
                false,
                true,
                ExcelChartBarDirection.COLUMN,
                List.of(
                    new ChartInput.Series(
                        new ChartInput.Title.Formula("B1"),
                        new ChartInput.DataSource("A2:A4"),
                        new ChartInput.DataSource("B2:B4")))));
    RuntimeException failure = new IllegalStateException("boom");

    assertEquals(
        "GET_DRAWING_OBJECTS", DefaultGridGrindRequestExecutor.readType(getDrawingObjects));
    assertEquals("GET_CHARTS", DefaultGridGrindRequestExecutor.readType(getCharts));
    assertEquals(
        "GET_DRAWING_OBJECT_PAYLOAD",
        DefaultGridGrindRequestExecutor.readType(getDrawingObjectPayload));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(getDrawingObjects));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(getCharts));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(getDrawingObjectPayload));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(getDrawingObjects, failure));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(getCharts, failure));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(getDrawingObjectPayload, failure));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(getDrawingObjects, failure));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(getCharts, failure));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(getDrawingObjectPayload, failure));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setPicture, failure));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setShape, failure));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setEmbeddedObject, failure));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setChart, failure));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setDrawingObjectAnchor, failure));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(deleteDrawingObject, failure));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(setPicture, failure));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(setShape, failure));
    assertEquals(
        "Budget", DefaultGridGrindRequestExecutor.sheetNameFor(setEmbeddedObject, failure));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(setChart, failure));
    assertEquals(
        "Budget", DefaultGridGrindRequestExecutor.sheetNameFor(setDrawingObjectAnchor, failure));
    assertEquals(
        "Budget", DefaultGridGrindRequestExecutor.sheetNameFor(deleteDrawingObject, failure));
    assertEquals(
        "BudgetTotal",
        DefaultGridGrindRequestExecutor.namedRangeNameFor(
            new WorkbookReadOperation.GetNamedRangeSurface(
                "named-range-surface",
                new NamedRangeSelection.Selected(
                    List.of(new NamedRangeSelector.ByName("BudgetTotal")))),
            failure));
    assertEquals(
        "BudgetScoped",
        DefaultGridGrindRequestExecutor.namedRangeNameFor(
            new WorkbookReadOperation.AnalyzeNamedRangeHealth(
                "named-range-health",
                new NamedRangeSelection.Selected(
                    List.of(new NamedRangeSelector.WorkbookScope("BudgetScoped")))),
            failure));
    assertEquals(
        "SheetScoped",
        DefaultGridGrindRequestExecutor.namedRangeNameFor(
            new WorkbookReadOperation.AnalyzeNamedRangeHealth(
                "named-range-health",
                new NamedRangeSelection.Selected(
                    List.of(new NamedRangeSelector.SheetScope("SheetScoped", "Budget")))),
            failure));
    assertNull(
        DefaultGridGrindRequestExecutor.namedRangeNameFor(
            new WorkbookReadOperation.GetNamedRangeSurface(
                "named-range-surface",
                new NamedRangeSelection.Selected(
                    List.of(
                        new NamedRangeSelector.ByName("BudgetTotal"),
                        new NamedRangeSelector.WorkbookScope("BudgetScoped")))),
            failure));
  }

  @Test
  void coversPivotReadAndWriteContextExtraction() {
    WorkbookOperation.SetPivotTable pivotFromRange =
        new WorkbookOperation.SetPivotTable(
            new PivotTableInput(
                "Sales Pivot 2026",
                "Report",
                new PivotTableInput.Source.Range("Data", "A1:D5"),
                new PivotTableInput.Anchor("C5"),
                List.of("Region"),
                List.of("Stage"),
                List.of("Owner"),
                List.of(
                    new PivotTableInput.DataField(
                        "Amount", ExcelPivotDataConsolidateFunction.SUM, null, "#,##0.00"))));
    WorkbookOperation.SetPivotTable pivotFromNamedRange =
        new WorkbookOperation.SetPivotTable(
            new PivotTableInput(
                "Named Pivot",
                "Report",
                new PivotTableInput.Source.NamedRange("PivotSource"),
                new PivotTableInput.Anchor("A3"),
                List.of("Region"),
                List.of(),
                List.of(),
                List.of(
                    new PivotTableInput.DataField(
                        "Amount", ExcelPivotDataConsolidateFunction.SUM, null, null))));
    WorkbookOperation.DeletePivotTable deletePivotTable =
        new WorkbookOperation.DeletePivotTable("Sales Pivot 2026", "Report");
    WorkbookReadOperation.GetPivotTables getPivotTables =
        new WorkbookReadOperation.GetPivotTables(
            "pivots", new PivotTableSelection.ByNames(List.of("Sales Pivot 2026")));
    WorkbookReadOperation.AnalyzePivotTableHealth analyzePivotTableHealth =
        new WorkbookReadOperation.AnalyzePivotTableHealth(
            "pivot-health", new PivotTableSelection.All());
    RuntimeException failure = new IllegalStateException("boom");

    assertEquals("GET_PIVOT_TABLES", DefaultGridGrindRequestExecutor.readType(getPivotTables));
    assertEquals(
        "ANALYZE_PIVOT_TABLE_HEALTH",
        DefaultGridGrindRequestExecutor.readType(analyzePivotTableHealth));
    assertNull(DefaultGridGrindRequestExecutor.sheetNameFor(getPivotTables));
    assertNull(DefaultGridGrindRequestExecutor.sheetNameFor(analyzePivotTableHealth));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(getPivotTables, failure));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(analyzePivotTableHealth, failure));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(getPivotTables, failure));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(analyzePivotTableHealth, failure));

    assertEquals("Report", DefaultGridGrindRequestExecutor.sheetNameFor(pivotFromRange, failure));
    assertEquals("C5", DefaultGridGrindRequestExecutor.addressFor(pivotFromRange, failure));
    assertEquals("A1:D5", DefaultGridGrindRequestExecutor.rangeFor(pivotFromRange, failure));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(pivotFromNamedRange, failure));
    assertEquals(
        "PivotSource",
        DefaultGridGrindRequestExecutor.namedRangeNameFor(pivotFromNamedRange, failure));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(pivotFromRange, failure));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(pivotFromRange, failure));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(deletePivotTable, failure));
    assertEquals("Report", DefaultGridGrindRequestExecutor.sheetNameFor(deletePivotTable, failure));
  }

  @Test
  void coversAdditionalPivotFormulaAndSelectionContextExtraction() {
    WorkbookReadOperation.GetFormulaSurface getFormulaSurface =
        new WorkbookReadOperation.GetFormulaSurface(
            "formula-surface", new SheetSelection.Selected(List.of("Budget")));
    WorkbookReadOperation.AnalyzeFormulaHealth analyzeFormulaHealth =
        new WorkbookReadOperation.AnalyzeFormulaHealth(
            "formula-health", new SheetSelection.Selected(List.of("Budget")));
    WorkbookReadOperation.GetWindow getWindow =
        new WorkbookReadOperation.GetWindow("window", "Budget", "B2", 3, 4);
    WorkbookReadOperation.GetSheetSchema getSheetSchema =
        new WorkbookReadOperation.GetSheetSchema("sheet-schema", "Budget", "C3", 5, 2);
    WorkbookOperation.SetCell setFormulaCell =
        new WorkbookOperation.SetCell("Budget", "D4", new CellInput.Formula("SUM(A1:A3)"));
    WorkbookOperation.SetPivotTable pivotFromTable =
        new WorkbookOperation.SetPivotTable(
            new PivotTableInput(
                "Table Source Pivot",
                "Report",
                new PivotTableInput.Source.Table("SalesTable2026"),
                new PivotTableInput.Anchor("G4"),
                List.of("Region"),
                List.of(),
                List.of(),
                List.of(
                    new PivotTableInput.DataField(
                        "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total Amount", null))));
    WorkbookOperation.SetNamedRange setNamedRange =
        new WorkbookOperation.SetNamedRange(
            "BudgetTotal", new NamedRangeScope.Workbook(), new NamedRangeTarget("Budget", "B4"));
    WorkbookOperation.DeleteNamedRange deleteNamedRange =
        new WorkbookOperation.DeleteNamedRange("BudgetTotal", new NamedRangeScope.Workbook());
    RuntimeException failure = new IllegalStateException("boom");
    InvalidFormulaException invalidFormula =
        new InvalidFormulaException("Budget", "D4", "SUM(", "bad formula", null);

    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(getFormulaSurface));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(analyzeFormulaHealth));
    assertEquals("B2", DefaultGridGrindRequestExecutor.addressFor(getWindow, failure));
    assertEquals("C3", DefaultGridGrindRequestExecutor.addressFor(getSheetSchema, failure));
    assertEquals("SUM(A1:A3)", DefaultGridGrindRequestExecutor.formulaFor(setFormulaCell, failure));
    assertEquals(
        "SUM(", DefaultGridGrindRequestExecutor.formulaFor(setFormulaCell, invalidFormula));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(pivotFromTable, failure));
    assertEquals(
        "BudgetTotal", DefaultGridGrindRequestExecutor.namedRangeNameFor(setNamedRange, failure));
    assertEquals(
        "BudgetTotal",
        DefaultGridGrindRequestExecutor.namedRangeNameFor(deleteNamedRange, failure));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(pivotFromTable, failure));
  }

  @Test
  @SuppressWarnings("PMD.AvoidAccessibilityAlteration")
  void coversExecutionModeHelperBranchesAndBestEffortCleanup() throws Exception {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.None(),
            List.of(),
            List.of());

    Method guardUnexpectedRuntime =
        accessibleMethod(
            DefaultGridGrindRequestExecutor.class,
            "guardUnexpectedRuntime",
            GridGrindProtocolVersion.class,
            GridGrindRequest.class,
            java.util.function.Supplier.class);
    GridGrindResponse.Failure guardedFailure =
        assertInstanceOf(
            GridGrindResponse.Failure.class,
            guardUnexpectedRuntime.invoke(
                executor,
                GridGrindProtocolVersion.current(),
                request,
                (java.util.function.Supplier<GridGrindResponse>)
                    () -> {
                      throw new IllegalStateException("boom");
                    }));
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, guardedFailure.problem().code());
    assertEquals("EXECUTE_REQUEST", guardedFailure.problem().context().stage());

    Method persistStreamingWorkbook =
        accessibleMethod(
            DefaultGridGrindRequestExecutor.class,
            "persistStreamingWorkbook",
            Path.class,
            GridGrindRequest.WorkbookPersistence.class,
            GridGrindRequest.WorkbookSource.class);
    Path materializedWorkbook = Files.createTempFile("gridgrind-streaming-materialized-", ".xlsx");
    try {
      InvocationTargetException failure =
          assertThrows(
              InvocationTargetException.class,
              () ->
                  persistStreamingWorkbook.invoke(
                      executor,
                      materializedWorkbook,
                      new GridGrindRequest.WorkbookPersistence.OverwriteSource(),
                      new GridGrindRequest.WorkbookSource.New()));
      IllegalStateException cause =
          assertInstanceOf(IllegalStateException.class, failure.getCause());
      assertTrue(cause.getMessage().contains("must be NONE or SAVE_AS"));
    } finally {
      Files.deleteIfExists(materializedWorkbook);
    }

    Method deleteIfExists =
        accessibleMethod(DefaultGridGrindRequestExecutor.class, "deleteIfExists", Path.class);
    deleteIfExists.invoke(null, (Object) null);

    Path nonEmptyDirectory = Files.createTempDirectory("gridgrind-non-empty-dir-");
    Path childFile = Files.writeString(nonEmptyDirectory.resolve("payload.txt"), "payload");
    try {
      deleteIfExists.invoke(null, nonEmptyDirectory);
      assertTrue(Files.exists(nonEmptyDirectory));
      assertTrue(Files.exists(childFile));
    } finally {
      Files.deleteIfExists(childFile);
      Files.deleteIfExists(nonEmptyDirectory);
    }

    Class<?> executionModeSelectionClass =
        Class.forName(
            "dev.erst.gridgrind.protocol.exec.DefaultGridGrindRequestExecutor$ExecutionModeSelection");
    java.lang.reflect.Constructor<?> selectionConstructor =
        executionModeSelectionClass.getDeclaredConstructor(
            ExecutionModeInput.ReadMode.class, ExecutionModeInput.WriteMode.class);
    selectionConstructor.setAccessible(true);
    Object eventReadFullWrite =
        selectionConstructor.newInstance(
            ExecutionModeInput.ReadMode.EVENT_READ, ExecutionModeInput.WriteMode.FULL_XSSF);
    Method directEventReadEligible =
        accessibleMethod(
            DefaultGridGrindRequestExecutor.class,
            "directEventReadEligible",
            GridGrindRequest.class,
            executionModeSelectionClass);

    boolean saveAsIsNotDirectEligible =
        (boolean)
            directEventReadEligible.invoke(
                null,
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.ExistingFile("/tmp/source.xlsx"),
                    new GridGrindRequest.WorkbookPersistence.SaveAs("/tmp/copy.xlsx"),
                    new ExecutionModeInput(
                        ExecutionModeInput.ReadMode.EVENT_READ,
                        ExecutionModeInput.WriteMode.FULL_XSSF),
                    null,
                    List.of(),
                    List.of()),
                eventReadFullWrite);
    boolean newSourceIsNotDirectEligible =
        (boolean)
            directEventReadEligible.invoke(
                null,
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    new ExecutionModeInput(
                        ExecutionModeInput.ReadMode.EVENT_READ,
                        ExecutionModeInput.WriteMode.FULL_XSSF),
                    null,
                    List.of(),
                    List.of()),
                eventReadFullWrite);

    assertFalse(saveAsIsNotDirectEligible);
    assertFalse(newSourceIsNotDirectEligible);

    Method executeFullReadsAgainstMaterializedPath =
        accessibleMethod(
            DefaultGridGrindRequestExecutor.class,
            "executeFullReadsAgainstMaterializedPath",
            GridGrindRequest.class,
            WorkbookLocation.class,
            Path.class);
    Path unreadableWorkbookPath = Files.createTempDirectory("gridgrind-workbook-dir-");
    try {
      InvocationTargetException fullReadFailure =
          assertThrows(
              InvocationTargetException.class,
              () ->
                  executeFullReadsAgainstMaterializedPath.invoke(
                      executor,
                      new GridGrindRequest(
                          new GridGrindRequest.WorkbookSource.New(),
                          new GridGrindRequest.WorkbookPersistence.None(),
                          List.of(),
                          List.of(new WorkbookReadOperation.GetWorkbookSummary("workbook"))),
                      new WorkbookLocation.UnsavedWorkbook(),
                      unreadableWorkbookPath));
      assertEquals("ReadExecutionFailure", fullReadFailure.getCause().getClass().getSimpleName());
      assertInstanceOf(IOException.class, fullReadFailure.getCause().getCause());
    } finally {
      Files.deleteIfExists(unreadableWorkbookPath);
    }
  }

  @Test
  void coversDirectEventReadFailureBranchesAndPackageSecurityReadCoordinates() throws Exception {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    Method executeDirectEventReadWorkflow =
        accessibleMethod(
            DefaultGridGrindRequestExecutor.class,
            "executeDirectEventReadWorkflow",
            GridGrindProtocolVersion.class,
            GridGrindRequest.class);
    Method executeEventReads =
        accessibleMethod(
            DefaultGridGrindRequestExecutor.class,
            "executeEventReads",
            Path.class,
            GridGrindRequest.class);
    WorkbookReadOperation.GetPackageSecurity getPackageSecurity =
        new WorkbookReadOperation.GetPackageSecurity("security");
    WorkbookReadOperation.GetWorkbookProtection getWorkbookProtection =
        new WorkbookReadOperation.GetWorkbookProtection("protection");
    IOException failure = new IOException("ignored");
    Path workbookPath = createWorkbookPath("Ops", "A1", "payload");
    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.ExistingFile(workbookPath.toString()),
            new GridGrindRequest.WorkbookPersistence.None(),
            new ExecutionModeInput(
                ExecutionModeInput.ReadMode.EVENT_READ, ExecutionModeInput.WriteMode.FULL_XSSF),
            null,
            List.of(),
            List.of(getPackageSecurity));

    try {
      InvocationTargetException executeEventReadsFailure =
          assertThrows(
              InvocationTargetException.class,
              () -> executeEventReads.invoke(executor, workbookPath, request));
      assertEquals(
          "ReadExecutionFailure", executeEventReadsFailure.getCause().getClass().getSimpleName());

      GridGrindResponse.Failure directFailure =
          assertInstanceOf(
              GridGrindResponse.Failure.class,
              executeDirectEventReadWorkflow.invoke(
                  executor, GridGrindProtocolVersion.current(), request));
      assertEquals("EXECUTE_READ", directFailure.problem().context().stage());

      assertNull(DefaultGridGrindRequestExecutor.sheetNameFor(getPackageSecurity));
      assertNull(DefaultGridGrindRequestExecutor.sheetNameFor(getWorkbookProtection));
      assertNull(DefaultGridGrindRequestExecutor.addressFor(getPackageSecurity, failure));
      assertNull(DefaultGridGrindRequestExecutor.addressFor(getWorkbookProtection, failure));
      assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(getPackageSecurity, failure));
      assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(getWorkbookProtection, failure));
    } finally {
      Files.deleteIfExists(workbookPath);
    }
  }

  @Test
  void coversStreamingPersistenceSecurityHelpers() throws Exception {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    Method persistenceOptions =
        accessibleMethod(
            DefaultGridGrindRequestExecutor.class,
            "persistenceOptions",
            GridGrindRequest.WorkbookPersistence.class);
    Method sourcePackageSecurity =
        accessibleMethod(
            DefaultGridGrindRequestExecutor.class,
            "sourcePackageSecurity",
            GridGrindRequest.WorkbookSource.class);
    Method sourceEncryptionPassword =
        accessibleMethod(
            DefaultGridGrindRequestExecutor.class,
            "sourceEncryptionPassword",
            GridGrindRequest.WorkbookSource.class);
    Method persistStreamingWorkbook =
        accessibleMethod(
            DefaultGridGrindRequestExecutor.class,
            "persistStreamingWorkbook",
            Path.class,
            GridGrindRequest.WorkbookPersistence.class,
            GridGrindRequest.WorkbookSource.class);
    OoxmlPersistenceSecurityInput security =
        new OoxmlPersistenceSecurityInput(
            new OoxmlEncryptionInput("persist-pass", ExcelOoxmlEncryptionMode.STANDARD), null);
    GridGrindRequest.WorkbookSource.New newSource = new GridGrindRequest.WorkbookSource.New();
    GridGrindRequest.WorkbookSource.ExistingFile existingSource =
        new GridGrindRequest.WorkbookSource.ExistingFile(
            "/tmp/source.xlsx", new OoxmlOpenSecurityInput("source-pass"));
    GridGrindRequest.WorkbookSource.ExistingFile unsecuredExistingSource =
        new GridGrindRequest.WorkbookSource.ExistingFile("/tmp/plain-source.xlsx");
    GridGrindRequest.WorkbookPersistence.OverwriteSource overwrite =
        new GridGrindRequest.WorkbookPersistence.OverwriteSource(security);

    assertTrue(
        ((ExcelOoxmlPersistenceOptions)
                persistenceOptions.invoke(null, new GridGrindRequest.WorkbookPersistence.None()))
            .isEmpty());
    assertEquals(
        "persist-pass",
        ((ExcelOoxmlPersistenceOptions) persistenceOptions.invoke(null, overwrite))
            .encryption()
            .password());
    assertFalse(
        ((ExcelOoxmlPackageSecuritySnapshot) sourcePackageSecurity.invoke(null, newSource))
            .isSecure());
    assertFalse(
        ((ExcelOoxmlPackageSecuritySnapshot) sourcePackageSecurity.invoke(null, existingSource))
            .isSecure());
    assertNull(sourceEncryptionPassword.invoke(null, newSource));
    assertEquals("source-pass", sourceEncryptionPassword.invoke(null, existingSource));
    assertNull(sourceEncryptionPassword.invoke(null, unsecuredExistingSource));

    Path materializedWorkbook = createWorkbookPath("Stream", "A1", "Encrypted stream");
    Path persistedWorkbook = Files.createTempFile("gridgrind-streaming-secured-", ".xlsx");
    Files.deleteIfExists(persistedWorkbook);
    try {
      GridGrindResponse.PersistenceOutcome.SavedAs savedAs =
          assertInstanceOf(
              GridGrindResponse.PersistenceOutcome.SavedAs.class,
              persistStreamingWorkbook.invoke(
                  executor,
                  materializedWorkbook,
                  new GridGrindRequest.WorkbookPersistence.SaveAs(
                      persistedWorkbook.toString(), security),
                  newSource));

      assertEquals(persistedWorkbook.toAbsolutePath().toString(), savedAs.executionPath());
      assertFalse(Files.exists(materializedWorkbook));
      assertEquals(
          "Encrypted stream",
          OoxmlSecurityTestSupport.decryptedStringCell(
              persistedWorkbook, "persist-pass", "Stream", "A1"));
    } finally {
      Files.deleteIfExists(materializedWorkbook);
      Files.deleteIfExists(persistedWorkbook);
    }
  }

  private static Path createWorkbookPath(String sheetName, String address, String value)
      throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-protocol-coverage-", ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet(sheetName).setCell(address, new ExcelCellValue.TextValue(value));
      workbook.save(workbookPath);
    }
    return workbookPath;
  }

  private static void assertPrivateSheetNameHelperRejects(
      String methodName, WorkbookOperation operation) throws ReflectiveOperationException {
    Method method =
        accessibleMethod(
            DefaultGridGrindRequestExecutor.class, methodName, WorkbookOperation.class);
    InvocationTargetException failure =
        assertThrows(InvocationTargetException.class, () -> method.invoke(null, operation));
    assertInstanceOf(IllegalStateException.class, failure.getCause());
  }

  @SuppressWarnings("PMD.AvoidAccessibilityAlteration")
  private static Method accessibleMethod(Class<?> type, String name, Class<?>... parameterTypes)
      throws ReflectiveOperationException {
    Method method = type.getDeclaredMethod(name, parameterTypes);
    method.setAccessible(true);
    return method;
  }

  private static CellStyleInput styleInput(
      ExcelHorizontalAlignment horizontalAlignment,
      ExcelVerticalAlignment verticalAlignment,
      ExcelBorderStyle borderStyle) {
    return new CellStyleInput(
        null,
        new CellAlignmentInput(null, horizontalAlignment, verticalAlignment, null, null),
        null,
        null,
        new CellBorderInput(
            null,
            new CellBorderSideInput(borderStyle),
            new CellBorderSideInput(borderStyle),
            new CellBorderSideInput(borderStyle),
            new CellBorderSideInput(borderStyle)),
        null);
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
    assertEquals(alignment, command.style().alignment().horizontalAlignment());
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
    assertEquals(alignment, command.style().alignment().verticalAlignment());
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
                        new ExcelSheetPresentationSnapshot(
                            ExcelSheetDisplay.defaults(),
                            null,
                            ExcelSheetOutlineSummary.defaults(),
                            ExcelSheetDefaults.defaults(),
                            List.of()),
                        List.of(
                            new dev.erst.gridgrind.excel.WorkbookReadResult.ColumnLayout(
                                0, 12.0, false, 0, false)),
                        List.of(
                            new dev.erst.gridgrind.excel.WorkbookReadResult.RowLayout(
                                0, 15.0, false, 0, false))))));
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
