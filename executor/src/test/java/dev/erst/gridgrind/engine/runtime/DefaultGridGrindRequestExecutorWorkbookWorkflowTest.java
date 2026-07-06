package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.*;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.ExcelAutofilterSnapshot;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelColorSnapshot;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRuleSnapshot;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelIgnoredError;
import dev.erst.gridgrind.excel.ExcelSheetDefaults;
import dev.erst.gridgrind.excel.ExcelSheetDisplay;
import dev.erst.gridgrind.excel.ExcelSheetOutlineSummary;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelTableSelection;
import dev.erst.gridgrind.excel.ExcelTableSnapshot;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.XlsxRoundTrip;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationSnapshot;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Workbook workflow integration tests for DefaultGridGrindRequestExecutor. */
class DefaultGridGrindRequestExecutorWorkbookWorkflowTest
    extends DefaultGridGrindRequestExecutorTestSupport {
  @Test
  void executesAssertionStepsAlongsideMutationsAndInspections() {
    GridGrindResponse.Success success =
        success(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    executionPolicy(calculateAll()),
                    null,
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new CellSelector.ByAddress("Budget", "A1"),
                            new CellMutationAction.SetCell(textCell("Owner"))),
                        mutate(
                            new CellSelector.ByAddress("Budget", "B2"),
                            new CellMutationAction.SetCell(formulaCell("2+3")))),
                    assertions(
                        assertThat(
                            "assert-owner",
                            new CellSelector.ByAddress("Budget", "A1"),
                            new CellAssertion.CellValue(
                                new dev.erst.gridgrind.contract.dto.CellScalarValue.Text("Owner"))),
                        assertThat(
                            "assert-formula",
                            new CellSelector.ByAddress("Budget", "B2"),
                            new CellAssertion.FormulaText("2+3"))),
                    inspections(
                        inspect(
                            "cells",
                            new CellSelector.ByAddresses("Budget", List.of("A1", "B2")),
                            allFacetCellsQuery())))));

    assertEquals(List.of("assert-owner", "assert-formula"), assertionIds(success));
    assertEquals(List.of("cells"), inspectionIds(success));
    SheetInspectionResult.CellsResult cells =
        inspection(success, "cells", SheetInspectionResult.CellsResult.class);
    assertEquals(
        "Owner",
        cast(dev.erst.gridgrind.contract.dto.CellReport.TextReport.class, cells.cells().get(0))
            .textValue()
            .orElseThrow());
  }

  @Test
  void surfacesStructuredAssertionFailuresWithObservedFacts() {
    GridGrindResponse.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new CellSelector.ByAddress("Budget", "A1"),
                            new CellMutationAction.SetCell(textCell("Owner")))),
                    assertions(
                        assertThat(
                            "assert-owner",
                            new CellSelector.ByAddress("Budget", "A1"),
                            new CellAssertion.CellValue(
                                new dev.erst.gridgrind.contract.dto.CellScalarValue.Text(
                                    "Wrong")))),
                    inspections())));

    assertEquals(GridGrindProblemCode.ASSERTION_FAILED, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertNotNull(failure.problem().assertionFailure().orElseThrow());
    assertEquals("assert-owner", failure.problem().assertionFailure().orElseThrow().stepId());
    assertEquals(
        "EXPECT_CELL_VALUE", failure.problem().assertionFailure().orElseThrow().assertionType());
    assertEquals(1, failure.problem().assertionFailure().orElseThrow().observations().size());
    assertInstanceOf(
        SheetInspectionResult.CellsResult.class,
        failure.problem().assertionFailure().orElseThrow().observations().getFirst());
  }

  @Test
  void successResponsesCarryStructuredExecutionJournal() {
    GridGrindResponse.Success success =
        success(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                WorkbookPlan.identified(
                    GridGrindProtocolVersion.current(),
                    "ledger-audit",
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    ExecutionPolicyInput.journal(
                        new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE)),
                    FormulaEnvironmentInput.empty(),
                    steps(
                        List.of(
                            mutate(
                                new SheetSelector.ByName("Ledger"),
                                new WorkbookMutationAction.EnsureSheet())),
                        List.of(
                            inspect(
                                "summary",
                                new WorkbookSelector.Current(),
                                new WorkbookIntrospectionQuery.GetWorkbookSummary()))))));

    assertEquals("ledger-audit", success.journal().planId().orElseThrow());
    assertEquals(ExecutionJournalLevel.VERBOSE, success.journal().level());
    assertEquals(2, success.journal().steps().size());
    assertEquals("ENSURE_SHEET", success.journal().steps().getFirst().stepType());
    assertEquals(
        ExecutionJournal.StepOutcome.SUCCEEDED, success.journal().steps().getFirst().outcome());
    assertEquals(
        "Sheet Ledger", success.journal().steps().getFirst().resolvedTargets().getFirst().label());
    assertEquals(ExecutionJournal.Status.SUCCEEDED, success.journal().outcome().status());
    assertFalse(success.journal().events().isEmpty());
  }

  @Test
  void failedResponsesCarryStepFailureClassificationInExecutionJournal() {
    GridGrindResponse.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                WorkbookPlan.identified(
                    GridGrindProtocolVersion.current(),
                    "bad-open",
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    ExecutionPolicyInput.journal(
                        new ExecutionJournalInput(ExecutionJournalLevel.NORMAL)),
                    FormulaEnvironmentInput.empty(),
                    List.of(
                        new InspectionStep(
                            "missing-sheet",
                            new SheetSelector.ByName("Missing"),
                            new SheetIntrospectionQuery.GetSheetSummary())))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    ExecutionJournal.Outcome.Failed outcome =
        assertInstanceOf(ExecutionJournal.Outcome.Failed.class, failure.journal().outcome());
    assertEquals(ExecutionJournal.Status.FAILED, outcome.status());
    assertEquals(0, outcome.failedStep().orElseThrow().failedStepIndex());
    assertEquals("missing-sheet", outcome.failedStep().orElseThrow().failedStepId());
    assertEquals(
        GridGrindProblemCode.SHEET_NOT_FOUND,
        failure.journal().steps().getFirst().failure().orElseThrow().code());
    assertEquals(
        ExecutionJournal.StepOutcome.FAILED, failure.journal().steps().getFirst().outcome());
  }

  @Test
  void executesWorkbookWorkflowAndReturnsOrderedReadResults() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-agent-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindResponse response =
        ExecutionContextFixtureSupport.execute(
            new DefaultGridGrindRequestExecutor(),
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.SaveAs(
                    workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                executionPolicy(calculateAllAndMarkRecalculateOnOpen()),
                null,
                mutations(
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new WorkbookMutationAction.EnsureSheet()),
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new CellMutationAction.AppendRow(
                            new dev.erst.gridgrind.contract.dto.CellRowInput.Typed(
                                List.of(
                                    textCell("Item"), textCell("Amount"), textCell("Billable"))))),
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new CellMutationAction.AppendRow(
                            new dev.erst.gridgrind.contract.dto.CellRowInput.Typed(
                                List.of(
                                    textCell("Hosting"),
                                    new CellInput.NumberValue(49.0),
                                    new CellInput.BooleanValue(true))))),
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new CellMutationAction.AppendRow(
                            new dev.erst.gridgrind.contract.dto.CellRowInput.Typed(
                                List.of(
                                    textCell("Domain"),
                                    new CellInput.NumberValue(12.0),
                                    new CellInput.BooleanValue(false))))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "A4"),
                        new CellMutationAction.SetCell(textCell("Total"))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "B4"),
                        new CellMutationAction.SetCell(formulaCell("SUM(B2:B3)"))),
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new WorkbookMutationAction.AutoSizeColumns())),
                inspect(
                    "workbook",
                    new WorkbookSelector.Current(),
                    new WorkbookIntrospectionQuery.GetWorkbookSummary()),
                inspect(
                    "cells",
                    new CellSelector.ByAddresses("Budget", List.of("A1", "B4", "C2")),
                    allFacetCellsQuery()),
                inspect(
                    "window",
                    new RangeSelector.RectangularWindow("Budget", "A1", 4, 3),
                    new SheetIntrospectionQuery.GetWindow())));

    GridGrindResponse.Success success = success(response);
    WorkbookSummary workbook =
        read(success, "workbook", WorkbookInspectionResult.WorkbookSummaryResult.class).workbook();
    SheetInspectionResult.CellsResult cells =
        read(success, "cells", SheetInspectionResult.CellsResult.class);
    WindowReport window =
        read(success, "window", SheetInspectionResult.WindowResult.class).window();

    assertEquals(GridGrindProtocolVersion.V1, success.protocolVersion());
    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertTrue(Files.exists(workbookPath));
    assertEquals(List.of(), success.warnings());
    assertEquals(List.of("workbook", "cells", "window"), stepIds(success));
    assertEquals(1, workbook.sheetCount());
    assertEquals(List.of("Budget"), workbook.sheetNames());
    assertEquals(0, workbook.namedRangeCount());
    assertTrue(workbook.forceFormulaRecalculationOnOpen());

    assertEquals("Budget", cells.sheetName());
    assertEquals(
        "Item",
        cast(dev.erst.gridgrind.contract.dto.CellReport.TextReport.class, cells.cells().get(0))
            .textValue()
            .orElseThrow());
    dev.erst.gridgrind.contract.dto.CellReport.FormulaReport formulaCell =
        cast(dev.erst.gridgrind.contract.dto.CellReport.FormulaReport.class, cells.cells().get(1));
    assertEquals("SUM(B2:B3)", formulaText(formulaCell));
    assertEquals(
        61.0,
        assertInstanceOf(CellValueReport.NumberValue.class, evaluation(formulaCell)).numberValue());
    assertTrue(
        cast(dev.erst.gridgrind.contract.dto.CellReport.BooleanReport.class, cells.cells().get(2))
            .booleanValue()
            .orElseThrow());

    assertEquals("Budget", window.sheetName());
    assertEquals("A1", window.topLeftAddress());
    assertEquals(4, window.dimensions().rowCount());
    assertEquals("A1", windowCells(window).getFirst().address());
  }

  @Test
  void executesDrawingWorkflowAndReturnsDrawingReadResults() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-drawing-request-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    DrawingAnchorInput.TwoCell firstAnchor =
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(1, 1, 0, 0),
            new DrawingMarkerInput(4, 6, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    DrawingAnchorInput.TwoCell movedAnchor =
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(6, 2, 0, 0),
            new DrawingMarkerInput(9, 7, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
    PictureDataInput pictureData =
        new PictureDataInput(
            ExcelPictureFormat.PNG,
            binary(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII="));

    GridGrindResponse.Success success =
        success(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(
                        workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Ops"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new SheetSelector.ByName("Ops"),
                            new DrawingMutationAction.SetPicture(
                                picture("OpsPicture", pictureData, firstAnchor, null))),
                        mutate(
                            new SheetSelector.ByName("Ops"),
                            new DrawingMutationAction.SetShape(
                                simpleShape("OpsShape", firstAnchor, "rect", text("Queue")))),
                        mutate(
                            new SheetSelector.ByName("Ops"),
                            new DrawingMutationAction.SetEmbeddedObject(
                                new EmbeddedObjectInput(
                                    "OpsEmbed",
                                    "Payload",
                                    "payload.txt",
                                    "payload.txt",
                                    binary("cGF5bG9hZA=="),
                                    pictureData,
                                    firstAnchor))),
                        mutate(
                            new DrawingObjectSelector.ByName("Ops", "OpsPicture"),
                            new DrawingMutationAction.SetDrawingObjectAnchor(movedAnchor)),
                        mutate(
                            new DrawingObjectSelector.ByName("Ops", "OpsShape"),
                            new DrawingMutationAction.DeleteDrawingObject())),
                    inspect(
                        "drawing",
                        new DrawingObjectSelector.AllOnSheet("Ops"),
                        new WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
                    inspect(
                        "payload",
                        new DrawingObjectSelector.ByName("Ops", "OpsEmbed"),
                        new WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload()))));

    WorkbookAssetInspectionResult.DrawingObjectsResult drawingObjects =
        read(success, "drawing", WorkbookAssetInspectionResult.DrawingObjectsResult.class);
    WorkbookAssetInspectionResult.DrawingObjectPayloadResult payload =
        read(success, "payload", WorkbookAssetInspectionResult.DrawingObjectPayloadResult.class);

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertTrue(Files.exists(workbookPath));
    assertEquals(2, drawingObjects.drawingObjects().size());
    assertEquals(
        List.of("OpsPicture", "OpsEmbed"),
        drawingObjects.drawingObjects().stream().map(DrawingObjectReport::name).toList());
    assertEquals(
        ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE,
        assertInstanceOf(
                DrawingAnchorReport.TwoCell.class,
                drawingObjects.drawingObjects().getFirst().anchor())
            .behavior());
    assertEquals("cGF5bG9hZA==", payload.payload().base64Data());
  }

  @Test
  void executesRichTextCellWorkflowAndReportsStructuredRuns() {
    GridGrindResponse response =
        ExecutionContextFixtureSupport.execute(
            new DefaultGridGrindRequestExecutor(),
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                mutations(
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new WorkbookMutationAction.EnsureSheet()),
                    mutate(
                        new RangeSelector.ByRange("Budget", "A1"),
                        new CellMutationAction.ApplyStyle(
                            styleInput(
                                null,
                                null,
                                fontInput(
                                    null,
                                    Boolean.TRUE,
                                    "Aptos",
                                    new FontHeightInput.Twips(260),
                                    ColorInput.rgb("#112233"),
                                    null,
                                    null),
                                null,
                                null,
                                null))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "A1"),
                        new CellMutationAction.SetCell(budgetRichTextCell()))),
                inspect(
                    "cells",
                    new CellSelector.ByAddresses("Budget", List.of("A1")),
                    allFacetCellsQuery())));

    GridGrindResponse.Success success = success(response);
    SheetInspectionResult.CellsResult cells =
        read(success, "cells", SheetInspectionResult.CellsResult.class);
    dev.erst.gridgrind.contract.dto.CellReport.TextReport cell =
        cast(dev.erst.gridgrind.contract.dto.CellReport.TextReport.class, cells.cells().getFirst());

    assertEquals("Budget FY26", textValue(cell));
    assertEquals(2, runs(cell).size());
    assertEquals("Budget", runs(cell).get(0).text());
    assertEquals("Aptos", runs(cell).get(0).font().fontName());
    assertEquals(rgb("#112233"), runs(cell).get(0).font().fontColor());
    assertTrue(runs(cell).get(0).font().italic());
    assertFalse(runs(cell).get(0).font().bold());
    assertEquals(" FY26", runs(cell).get(1).text());
    assertEquals("Aptos", runs(cell).get(1).font().fontName());
    assertEquals(rgb("#FF0000"), runs(cell).get(1).font().fontColor());
    assertTrue(runs(cell).get(1).font().bold());
    assertTrue(runs(cell).get(1).font().italic());
  }

  @Test
  void surfacesRequestWarningsAlongsideSuccessfulExecution() {
    GridGrindResponse response =
        ExecutionContextFixtureSupport.execute(
            new DefaultGridGrindRequestExecutor(),
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                mutations(
                    mutate(
                        new SheetSelector.ByName("Budget Review"),
                        new WorkbookMutationAction.EnsureSheet()),
                    mutate(
                        new SheetSelector.ByName("Summary"),
                        new WorkbookMutationAction.EnsureSheet()),
                    mutate(
                        new CellSelector.ByAddress("Budget Review", "A1"),
                        new CellMutationAction.SetCell(new CellInput.NumberValue(1200.0))),
                    mutate(
                        new CellSelector.ByAddress("Summary", "A1"),
                        new CellMutationAction.SetCell(formulaCell("Budget Review!A1"))))));

    GridGrindResponse.Success success = success(response);

    assertEquals(
        List.of(
            new RequestWarning(
                3,
                "step-04-set-cell",
                "SET_CELL",
                "Formula references same-request sheet names with spaces without single quotes: Budget Review. Use 'Sheet Name'!A1 syntax.")),
        success.warnings());
    assertEquals(List.of(), success.inspections());
  }

  @Test
  void opensExistingWorkbookAndOverwritesSource() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-existing-", ".xlsx");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", ExcelCellValue.text("Before"));
      ExecutionContextFixtureSupport.saveWorkbook(workbook, workbookPath);
    }

    GridGrindResponse response =
        ExecutionContextFixtureSupport.execute(
            new DefaultGridGrindRequestExecutor(),
            request(
                new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                new WorkbookPlan.WorkbookPersistence.Overwrite(),
                mutations(
                    mutate(
                        new CellSelector.ByAddress("Budget", "A1"),
                        new CellMutationAction.SetCell(textCell("After"))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "B1"),
                        new CellMutationAction.SetCell(new CellInput.NumberValue(12.0)))),
                inspect(
                    "cells",
                    new CellSelector.ByAddresses("Budget", List.of("A1", "B1")),
                    allFacetCellsQuery())));

    GridGrindResponse.Success success = success(response);
    SheetInspectionResult.CellsResult cells =
        read(success, "cells", SheetInspectionResult.CellsResult.class);

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(
        "After",
        cast(dev.erst.gridgrind.contract.dto.CellReport.TextReport.class, cells.cells().get(0))
            .textValue()
            .orElseThrow());
    assertEquals(
        12.0,
        cast(dev.erst.gridgrind.contract.dto.CellReport.NumberReport.class, cells.cells().get(1))
            .numberValue()
            .orElseThrow());
  }

  @Test
  void returnsStructuralReadResultsAndPersistsWorkbookShape() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-layout-ops-");

    GridGrindResponse response =
        ExecutionContextFixtureSupport.execute(
            new DefaultGridGrindRequestExecutor(),
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.SaveAs(
                    workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                mutations(
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new WorkbookMutationAction.EnsureSheet()),
                    mutate(
                        new CellSelector.ByAddress("Budget", "A1"),
                        new CellMutationAction.SetCell(textCell("Quarterly"))),
                    mutate(
                        new RangeSelector.ByRange("Budget", "A1:B1"),
                        new WorkbookMutationAction.MergeCells()),
                    mutate(
                        new ColumnBandSelector.Span("Budget", 0, 1),
                        new WorkbookMutationAction.SetColumnWidth(16.0)),
                    mutate(
                        new RowBandSelector.Span("Budget", 0, 0),
                        new WorkbookMutationAction.SetRowHeight(28.5)),
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new WorkbookMutationAction.SetSheetPane(new PaneInput.Frozen(1, 1, 1, 1))),
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new WorkbookMutationAction.SetSheetZoom(125)),
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new WorkbookMutationAction.SetSheetPresentation(
                            new SheetPresentationInput(
                                new SheetDisplayInput(false, false, false, true, true),
                                Optional.of(ColorInput.rgb("#112233")),
                                new SheetOutlineSummaryInput(false, false),
                                new SheetDefaultsInput(11, 18.5d),
                                List.of(
                                    new IgnoredErrorInput(
                                        "A1:B2",
                                        List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT)))))),
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new WorkbookMutationAction.SetPrintLayout(
                            new PrintLayoutInput(
                                new PrintAreaInput.Range("A1:B12"),
                                ExcelPrintOrientation.LANDSCAPE,
                                new PrintScalingInput.Fit(1, 0),
                                new PrintTitleRowsInput.Band(0, 0),
                                new PrintTitleColumnsInput.Band(0, 0),
                                headerFooter("Budget", "", ""),
                                headerFooter("", "Page &P", ""),
                                new PrintSetupInput(
                                    PrintSetupInput.defaults().margins(),
                                    true,
                                    false,
                                    false,
                                    PrintSetupInput.defaults().paperSize(),
                                    false,
                                    false,
                                    PrintSetupInput.defaults().copies(),
                                    false,
                                    PrintSetupInput.defaults().firstPageNumber(),
                                    List.of(),
                                    List.of()))))),
                inspect(
                    "cells",
                    new CellSelector.ByAddresses("Budget", List.of("A1")),
                    allFacetCellsQuery()),
                inspect(
                    "merged",
                    new SheetSelector.ByName("Budget"),
                    new SheetIntrospectionQuery.GetMergedRegions()),
                inspect(
                    "layout",
                    new SheetSelector.ByName("Budget"),
                    new SheetIntrospectionQuery.GetSheetLayout()),
                inspect(
                    "printLayout",
                    new SheetSelector.ByName("Budget"),
                    new SheetIntrospectionQuery.GetPrintLayout()),
                inspect(
                    "workbook",
                    new WorkbookSelector.Current(),
                    new WorkbookIntrospectionQuery.GetWorkbookSummary())));

    GridGrindResponse.Success success = success(response);
    SheetInspectionResult.CellsResult cells =
        read(success, "cells", SheetInspectionResult.CellsResult.class);
    SheetInspectionResult.MergedRegionsResult merged =
        read(success, "merged", SheetInspectionResult.MergedRegionsResult.class);
    SheetLayoutReport layout =
        read(success, "layout", SheetInspectionResult.SheetLayoutResult.class).layout();
    PrintLayoutReport printLayout =
        read(success, "printLayout", SheetInspectionResult.PrintLayoutResult.class).layout();
    WorkbookSummary workbook =
        read(success, "workbook", WorkbookInspectionResult.WorkbookSummaryResult.class).workbook();

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(
        "Quarterly",
        cast(dev.erst.gridgrind.contract.dto.CellReport.TextReport.class, cells.cells().getFirst())
            .textValue()
            .orElseThrow());
    assertEquals(
        List.of("A1:B1"), merged.mergedRegions().stream().map(MergedRegionReport::range).toList());
    assertInstanceOf(PaneReport.Frozen.class, layout.pane());
    PaneReport.Frozen frozen = cast(PaneReport.Frozen.class, layout.pane());
    assertEquals(1, frozen.splitColumn());
    assertEquals(1, frozen.splitRow());
    assertEquals(125, layout.zoomPercent());
    assertEquals(
        new SheetDisplayReport(false, false, false, true, true), layout.presentation().display());
    assertEquals(Optional.of(rgb("#112233")), layout.presentation().tabColor());
    assertEquals(
        new SheetOutlineSummaryReport(false, false), layout.presentation().outlineSummary());
    assertEquals(new SheetDefaultsReport(11, 18.5d), layout.presentation().sheetDefaults());
    assertEquals(
        List.of(
            new IgnoredErrorReport("A1:B2", List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))),
        layout.presentation().ignoredErrors());
    assertEquals(16.0, layout.columns().getFirst().widthCharacters());
    assertEquals(28.5, layout.rows().getFirst().heightPoints());
    assertEquals(ExcelPrintOrientation.LANDSCAPE, printLayout.orientation());
    assertTrue(printLayout.setup().printGridlines());
    assertEquals(List.of("Budget"), workbook.sheetNames());

    assertEquals(List.of("A1:B1"), XlsxRoundTrip.mergedRegions(workbookPath, "Budget"));
    assertEquals(4096, XlsxRoundTrip.columnWidth(workbookPath, "Budget", 0));
    assertEquals((short) 570, XlsxRoundTrip.rowHeightTwips(workbookPath, "Budget", 0));
    assertEquals(new ExcelSheetPane.Frozen(1, 1, 1, 1), XlsxRoundTrip.pane(workbookPath, "Budget"));
    assertEquals(125, XlsxRoundTrip.zoomPercent(workbookPath, "Budget"));
    dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayout reopenedLayout =
        XlsxRoundTrip.sheetLayout(workbookPath, "Budget");
    assertEquals(
        new ExcelSheetDisplay(false, false, false, true, true),
        reopenedLayout.presentation().display());
    assertEquals(
        Optional.of(ExcelColorSnapshot.rgb("#112233")), reopenedLayout.presentation().tabColor());
    assertEquals(
        new ExcelSheetOutlineSummary(false, false), reopenedLayout.presentation().outlineSummary());
    assertEquals(new ExcelSheetDefaults(11, 18.5d), reopenedLayout.presentation().sheetDefaults());
    assertEquals(
        List.of(
            new ExcelIgnoredError("A1:B2", List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))),
        reopenedLayout.presentation().ignoredErrors());
    assertTrue(XlsxRoundTrip.printLayout(workbookPath, "Budget").setup().printGridlines());
  }

  @Test
  void returnsB3SheetLayoutFactsAndPersistsVisibilityGrouping() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-b3-layout-");

    GridGrindResponse response =
        ExecutionContextFixtureSupport.execute(
            new DefaultGridGrindRequestExecutor(),
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.SaveAs(
                    workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                mutations(
                    mutate(
                        new SheetSelector.ByName("Layout"),
                        new WorkbookMutationAction.EnsureSheet()),
                    mutate(
                        new RangeSelector.ByRange("Layout", "A1:F6"),
                        new CellMutationAction.SetRange(
                            new dev.erst.gridgrind.contract.dto.CellGridInput.Typed(
                                List.of(
                                    List.of(
                                        textCell("Item"),
                                        textCell("Qty"),
                                        textCell("Status"),
                                        textCell("Note"),
                                        textCell("Owner"),
                                        textCell("Flag")),
                                    List.of(
                                        textCell("Hosting"),
                                        new CellInput.NumberValue(42.0),
                                        textCell("Open"),
                                        textCell("Alpha"),
                                        textCell("Ada"),
                                        textCell("Y")),
                                    List.of(
                                        textCell("Support"),
                                        new CellInput.NumberValue(84.0),
                                        textCell("Closed"),
                                        textCell("Beta"),
                                        textCell("Lin"),
                                        textCell("N")),
                                    List.of(
                                        textCell("Ops"),
                                        new CellInput.NumberValue(168.0),
                                        textCell("Open"),
                                        textCell("Gamma"),
                                        textCell("Bea"),
                                        textCell("Y")),
                                    List.of(
                                        textCell("QA"),
                                        new CellInput.NumberValue(21.0),
                                        textCell("Queued"),
                                        textCell("Delta"),
                                        textCell("Kai"),
                                        textCell("N")),
                                    List.of(
                                        textCell("Infra"),
                                        new CellInput.NumberValue(7.0),
                                        textCell("Done"),
                                        textCell("Epsilon"),
                                        textCell("Mia"),
                                        textCell("Y")))))),
                    mutate(
                        new RowBandSelector.Span("Layout", 1, 3),
                        new WorkbookMutationAction.GroupRows(true)),
                    mutate(
                        new RowBandSelector.Span("Layout", 5, 5),
                        new WorkbookMutationAction.SetRowVisibility(true)),
                    mutate(
                        new ColumnBandSelector.Span("Layout", 1, 3),
                        new WorkbookMutationAction.GroupColumns(true)),
                    mutate(
                        new ColumnBandSelector.Span("Layout", 5, 5),
                        new WorkbookMutationAction.SetColumnVisibility(true))),
                inspect(
                    "layout",
                    new SheetSelector.ByName("Layout"),
                    new SheetIntrospectionQuery.GetSheetLayout())));

    GridGrindResponse.Success success = success(response);
    SheetLayoutReport layout =
        read(success, "layout", SheetInspectionResult.SheetLayoutResult.class).layout();

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(6, layout.rows().size());
    assertTrue(layout.rows().get(1).hidden());
    assertEquals(1, layout.rows().get(1).outlineLevel());
    assertTrue(layout.rows().get(4).collapsed());
    assertTrue(layout.rows().get(5).hidden());
    assertEquals(6, layout.columns().size());
    assertTrue(layout.columns().get(1).hidden());
    assertEquals(1, layout.columns().get(1).outlineLevel());
    assertTrue(layout.columns().get(4).collapsed());
    assertTrue(layout.columns().get(5).hidden());

    dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayout reopenedLayout =
        XlsxRoundTrip.sheetLayout(workbookPath, "Layout");
    assertTrue(reopenedLayout.rows().get(1).hidden());
    assertEquals(1, reopenedLayout.rows().get(1).outlineLevel());
    assertTrue(reopenedLayout.rows().get(4).collapsed());
    assertTrue(reopenedLayout.rows().get(5).hidden());
    assertTrue(reopenedLayout.columns().get(1).hidden());
    assertEquals(1, reopenedLayout.columns().get(1).outlineLevel());
    assertTrue(reopenedLayout.columns().get(4).collapsed());
    assertTrue(reopenedLayout.columns().get(5).hidden());
  }

  @Test
  void returnsFactualMalformedPositiveLayoutValuesWithoutClampingReadback() throws IOException {
    Path sourceWorkbook = XlsxRoundTrip.newWorkbookPath("gridgrind-layout-factual-source-");
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Layout");
      workbook.sheet("Layout").cells().setCell("A1", ExcelCellValue.text("Header"));
      workbook.sheet("Layout").columns().setWidth(0, 0, 16.0d);
      ExecutionContextFixtureSupport.saveWorkbook(workbook, sourceWorkbook);
    }

    Path malformedWorkbook =
        dev.erst.gridgrind.excel.OoxmlPartMutator.rewriteEntries(
            sourceWorkbook,
            Map.of(
                "xl/worksheets/sheet1.xml",
                xml ->
                    xml.replace("<sheetFormatPr defaultRowHeight=\"15.0\"/>", "")
                        .replace(
                            "<sheetViews>", "<sheetFormatPr defaultRowHeight=\"999\"/><sheetViews>")
                        .replace("width=\"16.0\"", "width=\"300.0\"")));

    GridGrindResponse.Success success =
        success(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(malformedWorkbook.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    inspect(
                        "layout",
                        new SheetSelector.ByName("Layout"),
                        new SheetIntrospectionQuery.GetSheetLayout()))));

    SheetLayoutReport layout =
        read(success, "layout", SheetInspectionResult.SheetLayoutResult.class).layout();
    assertEquals(999.0d, layout.presentation().sheetDefaults().defaultRowHeightPoints());
    assertEquals(300.0d, layout.columns().getFirst().widthCharacters());
    assertEquals(999.0d, layout.rows().getFirst().heightPoints());

    dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayout reopenedLayout =
        XlsxRoundTrip.sheetLayout(malformedWorkbook, "Layout");
    assertEquals(999.0d, reopenedLayout.presentation().sheetDefaults().defaultRowHeightPoints());
    assertEquals(300.0d, reopenedLayout.columns().getFirst().widthCharacters());
    assertEquals(999.0d, reopenedLayout.rows().getFirst().heightPoints());
  }

  @Test
  void executesB3InsertDeleteAndShiftOperationsAndPersistsMovedCells() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-b3-geometry-");

    GridGrindResponse response =
        ExecutionContextFixtureSupport.execute(
            new DefaultGridGrindRequestExecutor(),
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.SaveAs(
                    workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                mutations(
                    mutate(
                        new SheetSelector.ByName("Moves"),
                        new WorkbookMutationAction.EnsureSheet()),
                    mutate(
                        new RangeSelector.ByRange("Moves", "A1:D3"),
                        new CellMutationAction.SetRange(
                            new dev.erst.gridgrind.contract.dto.CellGridInput.Typed(
                                List.of(
                                    List.of(
                                        textCell("Item"),
                                        textCell("Qty"),
                                        textCell("Status"),
                                        textCell("Note")),
                                    List.of(
                                        textCell("Hosting"),
                                        new CellInput.NumberValue(42.0),
                                        textCell("Open"),
                                        textCell("Alpha")),
                                    List.of(
                                        textCell("Support"),
                                        new CellInput.NumberValue(84.0),
                                        textCell("Closed"),
                                        textCell("Beta")))))),
                    mutate(
                        new RowBandSelector.Insertion("Moves", 1, 1),
                        new WorkbookMutationAction.InsertRows()),
                    mutate(
                        new CellSelector.ByAddress("Moves", "A2"),
                        new CellMutationAction.SetCell(textCell("Spacer"))),
                    mutate(
                        new RowBandSelector.Span("Moves", 2, 3),
                        new WorkbookMutationAction.ShiftRows(1)),
                    mutate(
                        new RowBandSelector.Span("Moves", 2, 2),
                        new WorkbookMutationAction.DeleteRows()),
                    mutate(
                        new ColumnBandSelector.Insertion("Moves", 1, 1),
                        new WorkbookMutationAction.InsertColumns()),
                    mutate(
                        new CellSelector.ByAddress("Moves", "B1"),
                        new CellMutationAction.SetCell(textCell("Pad"))),
                    mutate(
                        new ColumnBandSelector.Span("Moves", 2, 4),
                        new WorkbookMutationAction.ShiftColumns(1)),
                    mutate(
                        new ColumnBandSelector.Span("Moves", 2, 2),
                        new WorkbookMutationAction.DeleteColumns())),
                inspect(
                    "cells",
                    new CellSelector.ByAddresses("Moves", List.of("A2", "B1", "A3", "C3", "E4")),
                    allFacetCellsQuery())));

    GridGrindResponse.Success success = success(response);
    SheetInspectionResult.CellsResult cells =
        read(success, "cells", SheetInspectionResult.CellsResult.class);

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(
        "Spacer",
        cast(dev.erst.gridgrind.contract.dto.CellReport.TextReport.class, cells.cells().get(0))
            .textValue()
            .orElseThrow());
    assertEquals(
        "Pad",
        cast(dev.erst.gridgrind.contract.dto.CellReport.TextReport.class, cells.cells().get(1))
            .textValue()
            .orElseThrow());
    assertEquals(
        "Hosting",
        cast(dev.erst.gridgrind.contract.dto.CellReport.TextReport.class, cells.cells().get(2))
            .textValue()
            .orElseThrow());
    assertEquals(
        42.0,
        cast(dev.erst.gridgrind.contract.dto.CellReport.NumberReport.class, cells.cells().get(3))
            .numberValue()
            .orElseThrow());
    assertEquals(
        "Beta",
        cast(dev.erst.gridgrind.contract.dto.CellReport.TextReport.class, cells.cells().get(4))
            .textValue()
            .orElseThrow());

    try (ExcelWorkbook workbook = ExecutionContextFixtureSupport.openWorkbook(workbookPath)) {
      assertEquals("Spacer", workbook.sheet("Moves").cells().text("A2"));
      assertEquals("Pad", workbook.sheet("Moves").cells().text("B1"));
      assertEquals("Hosting", workbook.sheet("Moves").cells().text("A3"));
      assertEquals(42.0, workbook.sheet("Moves").cells().number("C3"));
      assertEquals("Beta", workbook.sheet("Moves").cells().text("E4"));
    }
  }

  @Test
  void returnsStructuredFailureForRowStructuralEditsThatWouldMoveTables() {
    GridGrindResponse.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new RangeSelector.ByRange("Budget", "A1:B3"),
                            new CellMutationAction.SetRange(
                                new dev.erst.gridgrind.contract.dto.CellGridInput.Typed(
                                    List.of(
                                        List.of(textCell("Item"), textCell("Qty")),
                                        List.of(
                                            textCell("Hosting"), new CellInput.NumberValue(42.0)),
                                        List.of(
                                            textCell("Support"),
                                            new CellInput.NumberValue(84.0)))))),
                        mutate(
                            new StructuredMutationAction.SetTable(
                                TableInput.withDefaultMetadata(
                                    "BudgetTable",
                                    "Budget",
                                    "A1:B3",
                                    false,
                                    new TableStyleInput.None()))),
                        mutate(
                            new RowBandSelector.Insertion("Budget", 1, 1),
                            new WorkbookMutationAction.InsertRows())))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertTrue(failure.problem().message().contains("table 'BudgetTable'"));
    assertTrue(failure.problem().message().contains("row structural edits"));
  }

  @Test
  void returnsStructuredFailureForColumnStructuralEditsWhenWorkbookHasFormulas() {
    GridGrindResponse.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new CellSelector.ByAddress("Budget", "A1"),
                            new CellMutationAction.SetCell(textCell("Item"))),
                        mutate(
                            new CellSelector.ByAddress("Budget", "B2"),
                            new CellMutationAction.SetCell(formulaCell("SUM(1, 1)"))),
                        mutate(
                            new ColumnBandSelector.Insertion("Budget", 1, 1),
                            new WorkbookMutationAction.InsertColumns())))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertTrue(failure.problem().message().contains("workbook formulas are present"));
    assertTrue(failure.problem().message().contains("column structural edits"));
  }

  @Test
  void returnsAuthoringMetadataAndNamedRangeReadResults() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-authoring-ops-");

    GridGrindResponse response =
        ExecutionContextFixtureSupport.execute(
            new DefaultGridGrindRequestExecutor(),
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.SaveAs(
                    workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                mutations(
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new WorkbookMutationAction.EnsureSheet()),
                    mutate(
                        new CellSelector.ByAddress("Budget", "A1"),
                        new CellMutationAction.SetCell(textCell("Report"))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "B4"),
                        new CellMutationAction.SetCell(new CellInput.NumberValue(61.0))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "A1"),
                        new CellMutationAction.SetHyperlink(
                            new HyperlinkTarget.Url("https://example.com/report"))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "A1"),
                        new CellMutationAction.SetComment(
                            CommentInput.plain(text("Review"), "GridGrind", true))),
                    mutate(
                        new StructuredMutationAction.SetNamedRange(
                            "BudgetTotal",
                            new NamedRangeScope.Workbook(),
                            NamedRangeTarget.range("Budget", "B4"))),
                    mutate(
                        new StructuredMutationAction.SetNamedRange(
                            "LocalItem",
                            new NamedRangeScope.Sheet("Budget"),
                            NamedRangeTarget.range("Budget", "A1:B2")))),
                inspect(
                    "cells",
                    new CellSelector.ByAddresses("Budget", List.of("A1", "B4")),
                    allFacetCellsQuery()),
                inspect(
                    "hyperlinks",
                    new CellSelector.ByAddresses("Budget", List.of("A1")),
                    new SheetIntrospectionQuery.GetHyperlinks()),
                inspect(
                    "comments",
                    new CellSelector.ByAddresses("Budget", List.of("A1")),
                    new SheetIntrospectionQuery.GetComments()),
                inspect(
                    "ranges",
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                    new WorkbookIntrospectionQuery.GetNamedRanges())));

    GridGrindResponse.Success success = success(response);
    dev.erst.gridgrind.contract.dto.CellReport.TextReport linkedCell =
        cast(
            dev.erst.gridgrind.contract.dto.CellReport.TextReport.class,
            read(success, "cells", SheetInspectionResult.CellsResult.class).cells().getFirst());
    SheetInspectionResult.HyperlinksResult hyperlinks =
        read(success, "hyperlinks", SheetInspectionResult.HyperlinksResult.class);
    SheetInspectionResult.CommentsResult comments =
        read(success, "comments", SheetInspectionResult.CommentsResult.class);
    WorkbookInspectionResult.NamedRangesResult ranges =
        read(success, "ranges", WorkbookInspectionResult.NamedRangesResult.class);

    assertEquals(
        java.util.Optional.of(new HyperlinkTarget.Url("https://example.com/report")),
        linkedCell.hyperlink());
    assertEquals("Review", linkedCell.comment().orElseThrow().text());
    assertEquals("A1", hyperlinks.hyperlinks().getFirst().address());
    assertEquals(
        new HyperlinkTarget.Url("https://example.com/report"),
        hyperlinks.hyperlinks().getFirst().hyperlink());
    assertEquals("A1", comments.comments().getFirst().address());
    assertEquals("Review", comments.comments().getFirst().comment().text());
    assertEquals(2, ranges.namedRanges().size());
    assertTrue(
        ranges
            .namedRanges()
            .contains(
                new NamedRangeReport.RangeReport(
                    "BudgetTotal",
                    new NamedRangeScope.Workbook(),
                    ranges.namedRanges().stream()
                        .filter(namedRange -> "BudgetTotal".equals(namedRange.name()))
                        .findFirst()
                        .orElseThrow()
                        .refersToFormula(),
                    NamedRangeTarget.range("Budget", "B4"))));
    assertTrue(
        ranges
            .namedRanges()
            .contains(
                new NamedRangeReport.RangeReport(
                    "LocalItem",
                    new NamedRangeScope.Sheet("Budget"),
                    ranges.namedRanges().stream()
                        .filter(namedRange -> "LocalItem".equals(namedRange.name()))
                        .findFirst()
                        .orElseThrow()
                        .refersToFormula(),
                    NamedRangeTarget.range("Budget", "A1:B2"))));

    assertEquals(
        new ExcelHyperlink.Url("https://example.com/report"),
        XlsxRoundTrip.cellMetadata(workbookPath, "Budget", "A1").hyperlink().orElseThrow());
    assertEquals(
        new ExcelComment("Review", "GridGrind", true),
        XlsxRoundTrip.cellMetadata(workbookPath, "Budget", "A1")
            .comment()
            .orElseThrow()
            .toPlainComment());
    assertEquals(2, XlsxRoundTrip.namedRanges(workbookPath).size());
  }

  @Test
  void returnsDataValidationReadResultsAndPersistsNormalizedRules() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-data-validation-ops-");

    GridGrindResponse response =
        ExecutionContextFixtureSupport.execute(
            new DefaultGridGrindRequestExecutor(),
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.SaveAs(
                    workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                mutations(
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new WorkbookMutationAction.EnsureSheet()),
                    mutate(
                        new RangeSelector.ByRange("Budget", "A2:C5"),
                        new StructuredMutationAction.SetDataValidation(
                            dataValidation(
                                new DataValidationRuleInput.ExplicitList(List.of("Queued", "Done")),
                                true,
                                false,
                                prompt("Status", "Pick one workflow state.", true),
                                errorAlert(
                                    ExcelDataValidationErrorStyle.STOP,
                                    "Invalid status",
                                    "Use one of the allowed values.",
                                    true)))),
                    mutate(
                        new RangeSelector.ByRanges("Budget", List.of("B3")),
                        new StructuredMutationAction.ClearDataValidations()),
                    mutate(
                        new RangeSelector.ByRange("Budget", "E2:E5"),
                        new StructuredMutationAction.SetDataValidation(
                            dataValidation(
                                new DataValidationRuleInput.FormulaList("#REF!"),
                                false,
                                false,
                                null,
                                null)))),
                inspect(
                    "validations",
                    new RangeSelector.AllOnSheet("Budget"),
                    new SheetIntrospectionQuery.GetDataValidations()),
                inspect(
                    "health",
                    new SheetSelector.ByNames(List.of("Budget")),
                    new InspectionAnalysisQuery.AnalyzeDataValidationHealth())));

    GridGrindResponse.Success success = success(response);
    SheetInspectionResult.DataValidationsResult validations =
        read(success, "validations", SheetInspectionResult.DataValidationsResult.class);
    dev.erst.gridgrind.contract.query.WorkbookAnalysisResult.DataValidationHealthResult health =
        read(
            success,
            "health",
            dev.erst.gridgrind.contract.query.WorkbookAnalysisResult.DataValidationHealthResult
                .class);
    List<ExcelDataValidationSnapshot> persisted =
        XlsxRoundTrip.dataValidations(workbookPath, "Budget");

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(2, validations.validations().size());
    DataValidationEntryReport.Supported explicitList =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class, validations.validations().getFirst());
    DataValidationEntryReport.Supported brokenFormula =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class, validations.validations().get(1));
    assertEquals(List.of("A2:C2", "A4:C5", "A3", "C3"), explicitList.ranges());
    assertInstanceOf(DataValidationRuleInput.ExplicitList.class, explicitList.validation().rule());
    assertEquals(List.of("E2:E5"), brokenFormula.ranges());
    assertEquals(
        "#REF!",
        cast(DataValidationRuleInput.FormulaList.class, brokenFormula.validation().rule())
            .formula());
    assertEquals(2, health.analysis().checkedValidationCount());
    assertEquals(1, health.analysis().summary().errorCount());
    assertEquals(
        AnalysisFindingCode.DATA_VALIDATION_BROKEN_FORMULA,
        health.analysis().findings().getFirst().code());
    assertEquals(2, persisted.size());
    assertEquals(List.of("A2:C2", "A4:C5", "A3", "C3"), persisted.getFirst().ranges());
  }

  @Test
  void returnsConditionalFormattingReadResultsAndPersistsDefinitions() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-conditional-formatting-");

    GridGrindResponse response =
        ExecutionContextFixtureSupport.execute(
            new DefaultGridGrindRequestExecutor(),
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.SaveAs(
                    workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                mutations(
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new WorkbookMutationAction.EnsureSheet()),
                    mutate(
                        new RangeSelector.ByRange("Budget", "A1:B5"),
                        new CellMutationAction.SetRange(
                            new dev.erst.gridgrind.contract.dto.CellGridInput.Typed(
                                List.of(
                                    List.of(textCell("Status"), textCell("Amount")),
                                    List.of(textCell("Queued"), new CellInput.NumberValue(1.0)),
                                    List.of(textCell("Done"), new CellInput.NumberValue(9.0)),
                                    List.of(textCell("Done"), new CellInput.NumberValue(11.0)),
                                    List.of(textCell("Queued"), new CellInput.NumberValue(4.0)))))),
                    mutate(
                        new RangeSelector.ByRange("Budget", "B2:B5"),
                        new StructuredMutationAction.SetConditionalFormatting(
                            conditionalFormattingReadbackBlock()))),
                inspect(
                    "conditional-formatting",
                    new RangeSelector.AllOnSheet("Budget"),
                    new SheetIntrospectionQuery.GetConditionalFormatting()),
                inspect(
                    "conditional-formatting-health",
                    new SheetSelector.ByNames(List.of("Budget")),
                    new InspectionAnalysisQuery.AnalyzeConditionalFormattingHealth())));

    GridGrindResponse.Success success = success(response);
    SheetInspectionResult.ConditionalFormattingResult conditionalFormatting =
        read(
            success,
            "conditional-formatting",
            SheetInspectionResult.ConditionalFormattingResult.class);
    dev.erst.gridgrind.contract.query.WorkbookAnalysisResult.ConditionalFormattingHealthResult
        health =
            read(
                success,
                "conditional-formatting-health",
                dev.erst.gridgrind.contract.query.WorkbookAnalysisResult
                    .ConditionalFormattingHealthResult.class);

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(1, conditionalFormatting.conditionalFormattingBlocks().size());
    ConditionalFormattingEntryReport block =
        conditionalFormatting.conditionalFormattingBlocks().getFirst();
    assertEquals(List.of("B2:B5"), block.ranges());
    assertEquals(2, block.rules().size());
    assertInstanceOf(ConditionalFormattingRuleReport.FormulaRule.class, block.rules().get(0));
    assertInstanceOf(ConditionalFormattingRuleReport.CellValueRule.class, block.rules().get(1));
    assertEquals(1, health.analysis().checkedConditionalFormattingBlockCount());
    assertEquals(List.of(), health.analysis().findings());

    try (ExcelWorkbook workbook = ExecutionContextFixtureSupport.openWorkbook(workbookPath)) {
      WorkbookExecutionEngine readExecutor = new WorkbookExecutionEngine();
      dev.erst.gridgrind.excel.WorkbookRuleResult.ConditionalFormattingResult reopened =
          (dev.erst.gridgrind.excel.WorkbookRuleResult.ConditionalFormattingResult)
              readExecutor
                  .read(
                      workbook,
                      new WorkbookReadCommand.GetConditionalFormatting(
                          "conditional-formatting",
                          "Budget",
                          new dev.erst.gridgrind.excel.ExcelRangeSelection.All()))
                  .getFirst();

      assertEquals(1, reopened.conditionalFormattingBlocks().size());
      ExcelConditionalFormattingBlockSnapshot reopenedBlock =
          reopened.conditionalFormattingBlocks().getFirst();
      assertEquals(List.of("B2:B5"), reopenedBlock.ranges());
      assertEquals(2, reopenedBlock.rules().size());
      assertEquals(
          "B2>5",
          cast(
                  ExcelConditionalFormattingRuleSnapshot.FormulaRule.class,
                  reopenedBlock.rules().get(0))
              .formula());
      assertEquals(
          ExcelComparisonOperator.BETWEEN,
          cast(
                  ExcelConditionalFormattingRuleSnapshot.CellValueRule.class,
                  reopenedBlock.rules().get(1))
              .operator());
    }
  }

  @Test
  void returnsAutofilterAndTableReadResultsAndPersistsDefinitions() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-table-autofilter-");

    GridGrindResponse response =
        ExecutionContextFixtureSupport.execute(
            new DefaultGridGrindRequestExecutor(),
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.SaveAs(
                    workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                mutations(
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new WorkbookMutationAction.EnsureSheet()),
                    mutate(
                        new RangeSelector.ByRange("Budget", "A1:C4"),
                        new CellMutationAction.SetRange(
                            new dev.erst.gridgrind.contract.dto.CellGridInput.Typed(
                                List.of(
                                    List.of(
                                        textCell("Item"), textCell("Amount"), textCell("Billable")),
                                    List.of(
                                        textCell("Hosting"),
                                        new CellInput.NumberValue(49.0),
                                        new CellInput.BooleanValue(true)),
                                    List.of(
                                        textCell("Domain"),
                                        new CellInput.NumberValue(12.0),
                                        new CellInput.BooleanValue(false)),
                                    List.of(
                                        textCell("Support"),
                                        new CellInput.NumberValue(18.0),
                                        new CellInput.BooleanValue(true)))))),
                    mutate(
                        new RangeSelector.ByRange("Budget", "E1:F3"),
                        new CellMutationAction.SetRange(
                            new dev.erst.gridgrind.contract.dto.CellGridInput.Typed(
                                List.of(
                                    List.of(textCell("Queue"), textCell("Owner")),
                                    List.of(textCell("Late invoices"), textCell("Marta")),
                                    List.of(textCell("Badge orders"), textCell("Rihards")))))),
                    mutate(
                        new RangeSelector.ByRange("Budget", "E1:F3"),
                        new StructuredMutationAction.SetAutofilter()),
                    mutate(
                        new StructuredMutationAction.SetTable(
                            TableInput.withDefaultMetadata(
                                "BudgetTable",
                                "Budget",
                                "A1:C4",
                                false,
                                new TableStyleInput.Named(
                                    "TableStyleMedium2", false, false, true, false))))),
                inspect(
                    "filters",
                    new SheetSelector.ByName("Budget"),
                    new SheetIntrospectionQuery.GetAutofilters()),
                inspect(
                    "tables",
                    new TableSelector.All(),
                    new WorkbookAssetIntrospectionQuery.GetTables()),
                inspect(
                    "autofilter-health",
                    new SheetSelector.ByNames(List.of("Budget")),
                    new InspectionAnalysisQuery.AnalyzeAutofilterHealth()),
                inspect(
                    "table-health",
                    new TableSelector.All(),
                    new InspectionAnalysisQuery.AnalyzeTableHealth())));

    GridGrindResponse.Success success = success(response);
    SheetInspectionResult.AutofiltersResult filters =
        read(success, "filters", SheetInspectionResult.AutofiltersResult.class);
    WorkbookAssetInspectionResult.TablesResult tables =
        read(success, "tables", WorkbookAssetInspectionResult.TablesResult.class);
    dev.erst.gridgrind.contract.query.WorkbookAnalysisResult.AutofilterHealthResult
        autofilterHealth =
            read(
                success,
                "autofilter-health",
                dev.erst.gridgrind.contract.query.WorkbookAnalysisResult.AutofilterHealthResult
                    .class);
    dev.erst.gridgrind.contract.query.WorkbookAnalysisResult.TableHealthResult tableHealth =
        read(
            success,
            "table-health",
            dev.erst.gridgrind.contract.query.WorkbookAnalysisResult.TableHealthResult.class);

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(
        List.of(
            new AutofilterEntryReport.SheetOwned("E1:F3"),
            new AutofilterEntryReport.TableOwned("A1:C4", "BudgetTable")),
        filters.autofilters());
    assertEquals(1, tables.tables().size());
    TableEntryReport table = tables.tables().getFirst();
    assertEquals("BudgetTable", table.name());
    assertEquals(List.of("Item", "Amount", "Billable"), table.structure().columnNames());
    assertEquals(
        new TableStyleReport.Named("TableStyleMedium2", false, false, true, false), table.style());
    assertTrue(table.behavior().hasAutofilter());
    assertEquals(2, autofilterHealth.analysis().checkedAutofilterCount());
    assertEquals(List.of(), autofilterHealth.analysis().findings());
    assertEquals(1, tableHealth.analysis().checkedTableCount());
    assertEquals(List.of(), tableHealth.analysis().findings());

    try (ExcelWorkbook workbook = ExecutionContextFixtureSupport.openWorkbook(workbookPath)) {
      WorkbookExecutionEngine readExecutor = new WorkbookExecutionEngine();
      dev.erst.gridgrind.excel.WorkbookRuleResult.AutofiltersResult reopenedFilters =
          (dev.erst.gridgrind.excel.WorkbookRuleResult.AutofiltersResult)
              readExecutor
                  .read(workbook, new WorkbookReadCommand.GetAutofilters("filters", "Budget"))
                  .getFirst();
      dev.erst.gridgrind.excel.WorkbookRuleResult.TablesResult reopenedTables =
          (dev.erst.gridgrind.excel.WorkbookRuleResult.TablesResult)
              readExecutor
                  .read(
                      workbook,
                      new WorkbookReadCommand.GetTables("tables", new ExcelTableSelection.All()))
                  .getFirst();

      assertEquals(
          List.of(
              new ExcelAutofilterSnapshot.SheetOwned("E1:F3"),
              new ExcelAutofilterSnapshot.TableOwned("A1:C4", "BudgetTable")),
          reopenedFilters.autofilters());
      assertEquals(
          List.of(
              new ExcelTableSnapshot(
                  "BudgetTable",
                  "Budget",
                  "A1:C4",
                  new ExcelTableSnapshot.Structure(
                      1,
                      0,
                      List.of("Item", "Amount", "Billable"),
                      List.of(
                          new dev.erst.gridgrind.excel.ExcelTableColumnSnapshot(
                              1L, "Item", "", "", "", ""),
                          new dev.erst.gridgrind.excel.ExcelTableColumnSnapshot(
                              2L, "Amount", "", "", "", ""),
                          new dev.erst.gridgrind.excel.ExcelTableColumnSnapshot(
                              3L, "Billable", "", "", "", ""))),
                  new dev.erst.gridgrind.excel.ExcelTableStyleSnapshot.Named(
                      "TableStyleMedium2", false, false, true, false),
                  new ExcelTableSnapshot.Behavior(true, false, false, false),
                  new ExcelTableSnapshot.Presentation(
                      Optional.empty(), Optional.empty(), Optional.empty(), Optional.empty()))),
          reopenedTables.tables());
    }
  }

  private static ConditionalFormattingDefinitionInput conditionalFormattingReadbackBlock() {
    return new ConditionalFormattingDefinitionInput(
        List.of(
            new ConditionalFormattingRuleInput.FormulaRule(
                "B2>5",
                true,
                java.util.Optional.of(
                    new DifferentialStyleInput(
                        Optional.of("0.00"),
                        Optional.of(true),
                        Optional.empty(),
                        Optional.empty(),
                        java.util.Optional.of("#102030"),
                        Optional.empty(),
                        Optional.empty(),
                        java.util.Optional.of("#E0F0AA"),
                        java.util.Optional.empty()))),
            new ConditionalFormattingRuleInput.CellValueRule(
                ExcelComparisonOperator.BETWEEN,
                "1",
                java.util.Optional.of("10"),
                false,
                java.util.Optional.of(
                    new DifferentialStyleInput(
                        Optional.empty(),
                        Optional.empty(),
                        Optional.of(true),
                        Optional.empty(),
                        java.util.Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        java.util.Optional.empty(),
                        java.util.Optional.empty())))));
  }

  private static CellInput.RichText budgetRichTextCell() {
    return new CellInput.RichText(
        List.of(
            richTextRun("Budget"),
            new RichTextRunInput(
                text(" FY26"),
                maybe(
                    fontInput(
                        Boolean.TRUE, null, null, null, ColorInput.rgb("#FF0000"), null, null)))));
  }
}
