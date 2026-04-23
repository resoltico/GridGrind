package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.excel.ExcelAutofilterSnapshot;
import dev.erst.gridgrind.excel.ExcelBorderSideSnapshot;
import dev.erst.gridgrind.excel.ExcelBorderSnapshot;
import dev.erst.gridgrind.excel.ExcelCellAlignmentSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFillSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFontSnapshot;
import dev.erst.gridgrind.excel.ExcelCellMetadataSnapshot;
import dev.erst.gridgrind.excel.ExcelCellProtectionSnapshot;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelCellStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelColorSnapshot;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelCommentSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRuleSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdSnapshot;
import dev.erst.gridgrind.excel.ExcelDataValidationSnapshot;
import dev.erst.gridgrind.excel.ExcelDifferentialStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelIgnoredError;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeSnapshot;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelPrintMarginsSnapshot;
import dev.erst.gridgrind.excel.ExcelPrintSetupSnapshot;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetDefaults;
import dev.erst.gridgrind.excel.ExcelSheetDisplay;
import dev.erst.gridgrind.excel.ExcelSheetOutlineSummary;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelSheetPresentation;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelTableSelection;
import dev.erst.gridgrind.excel.ExcelTableSnapshot;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.InvalidFormulaException;
import dev.erst.gridgrind.excel.InvalidRangeAddressException;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import dev.erst.gridgrind.excel.UnsupportedFormulaException;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.excel.WorkbookLocation;
import dev.erst.gridgrind.excel.WorkbookNotFoundException;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import dev.erst.gridgrind.excel.XlsxRoundTrip;
import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import dev.erst.gridgrind.excel.foundation.ExcelPaneRegion;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationType;

/** Integration and helper tests for DefaultGridGrindRequestExecutor. */
class DefaultGridGrindRequestExecutorTest {
  @Test
  void executesAssertionStepsAlongsideMutationsAndInspections() {
    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        executionPolicy(calculateAll()),
                        null,
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new CellSelector.ByAddress("Budget", "A1"),
                                new MutationAction.SetCell(textCell("Owner"))),
                            mutate(
                                new CellSelector.ByAddress("Budget", "B2"),
                                new MutationAction.SetCell(formulaCell("2+3")))),
                        assertions(
                            assertThat(
                                "assert-owner",
                                new CellSelector.ByAddress("Budget", "A1"),
                                new Assertion.CellValue(new ExpectedCellValue.Text("Owner"))),
                            assertThat(
                                "assert-formula",
                                new CellSelector.ByAddress("Budget", "B2"),
                                new Assertion.FormulaText("2+3"))),
                        inspections(
                            inspect(
                                "cells",
                                new CellSelector.ByAddresses("Budget", List.of("A1", "B2")),
                                new InspectionQuery.GetCells())))));

    assertEquals(List.of("assert-owner", "assert-formula"), assertionIds(success));
    assertEquals(List.of("cells"), inspectionIds(success));
    InspectionResult.CellsResult cells =
        inspection(success, "cells", InspectionResult.CellsResult.class);
    assertEquals(
        "Owner",
        cast(GridGrindResponse.CellReport.TextReport.class, cells.cells().get(0)).stringValue());
  }

  @Test
  void surfacesStructuredAssertionFailuresWithObservedFacts() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new CellSelector.ByAddress("Budget", "A1"),
                                new MutationAction.SetCell(textCell("Owner")))),
                        assertions(
                            assertThat(
                                "assert-owner",
                                new CellSelector.ByAddress("Budget", "A1"),
                                new Assertion.CellValue(new ExpectedCellValue.Text("Wrong")))),
                        inspections())));

    assertEquals(GridGrindProblemCode.ASSERTION_FAILED, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertNotNull(failure.problem().assertionFailure());
    assertEquals("assert-owner", failure.problem().assertionFailure().stepId());
    assertEquals("EXPECT_CELL_VALUE", failure.problem().assertionFailure().assertionType());
    assertEquals(1, failure.problem().assertionFailure().observations().size());
    assertInstanceOf(
        InspectionResult.CellsResult.class,
        failure.problem().assertionFailure().observations().getFirst());
  }

  @Test
  void successResponsesCarryStructuredExecutionJournal() {
    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    new WorkbookPlan(
                        GridGrindProtocolVersion.current(),
                        "ledger-audit",
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionPolicyInput(
                            null, new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE)),
                        null,
                        steps(
                            List.of(
                                mutate(
                                    new SheetSelector.ByName("Ledger"),
                                    new MutationAction.EnsureSheet())),
                            List.of(
                                inspect(
                                    "summary",
                                    new WorkbookSelector.Current(),
                                    new InspectionQuery.GetWorkbookSummary()))))));

    assertEquals("ledger-audit", success.journal().planId());
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
            new DefaultGridGrindRequestExecutor()
                .execute(
                    new WorkbookPlan(
                        GridGrindProtocolVersion.current(),
                        "bad-open",
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionPolicyInput(
                            null, new ExecutionJournalInput(ExecutionJournalLevel.NORMAL)),
                        null,
                        List.of(
                            new InspectionStep(
                                "missing-sheet",
                                new SheetSelector.ByName("Missing"),
                                new InspectionQuery.GetSheetSummary())))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals(ExecutionJournal.Status.FAILED, failure.journal().outcome().status());
    assertEquals(0, failure.journal().outcome().failedStepIndex());
    assertEquals("missing-sheet", failure.journal().outcome().failedStepId());
    assertEquals(
        GridGrindProblemCode.SHEET_NOT_FOUND,
        failure.journal().steps().getFirst().failure().code());
    assertEquals(
        ExecutionJournal.StepOutcome.FAILED, failure.journal().steps().getFirst().outcome());
  }

  @Test
  void executesWorkbookWorkflowAndReturnsOrderedReadResults() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-agent-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    executionPolicy(calculateAllAndMarkRecalculateOnOpen()),
                    null,
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new MutationAction.AppendRow(
                                List.of(
                                    textCell("Item"), textCell("Amount"), textCell("Billable")))),
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new MutationAction.AppendRow(
                                List.of(
                                    textCell("Hosting"),
                                    new CellInput.Numeric(49.0),
                                    new CellInput.BooleanValue(true)))),
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new MutationAction.AppendRow(
                                List.of(
                                    textCell("Domain"),
                                    new CellInput.Numeric(12.0),
                                    new CellInput.BooleanValue(false)))),
                        mutate(
                            new CellSelector.ByAddress("Budget", "A4"),
                            new MutationAction.SetCell(textCell("Total"))),
                        mutate(
                            new CellSelector.ByAddress("Budget", "B4"),
                            new MutationAction.SetCell(formulaCell("SUM(B2:B3)"))),
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new MutationAction.AutoSizeColumns())),
                    inspect(
                        "workbook",
                        new WorkbookSelector.Current(),
                        new InspectionQuery.GetWorkbookSummary()),
                    inspect(
                        "cells",
                        new CellSelector.ByAddresses("Budget", List.of("A1", "B4", "C2")),
                        new InspectionQuery.GetCells()),
                    inspect(
                        "window",
                        new RangeSelector.RectangularWindow("Budget", "A1", 4, 3),
                        new InspectionQuery.GetWindow())));

    GridGrindResponse.Success success = success(response);
    GridGrindResponse.WorkbookSummary workbook =
        read(success, "workbook", InspectionResult.WorkbookSummaryResult.class).workbook();
    InspectionResult.CellsResult cells = read(success, "cells", InspectionResult.CellsResult.class);
    GridGrindResponse.WindowReport window =
        read(success, "window", InspectionResult.WindowResult.class).window();

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
        cast(GridGrindResponse.CellReport.TextReport.class, cells.cells().get(0)).stringValue());
    GridGrindResponse.CellReport.FormulaReport formulaCell =
        cast(GridGrindResponse.CellReport.FormulaReport.class, cells.cells().get(1));
    assertEquals("SUM(B2:B3)", formulaCell.formula());
    assertEquals(
        61.0,
        cast(GridGrindResponse.CellReport.NumberReport.class, formulaCell.evaluation())
            .numberValue());
    assertTrue(
        cast(GridGrindResponse.CellReport.BooleanReport.class, cells.cells().get(2))
            .booleanValue());

    assertEquals("Budget", window.sheetName());
    assertEquals("A1", window.topLeftAddress());
    assertEquals(4, window.rows().size());
    assertEquals("A1", window.rows().getFirst().cells().getFirst().address());
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
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Ops"), new MutationAction.EnsureSheet()),
                            mutate(
                                new SheetSelector.ByName("Ops"),
                                new MutationAction.SetPicture(
                                    new PictureInput(
                                        "OpsPicture", pictureData, firstAnchor, null))),
                            mutate(
                                new SheetSelector.ByName("Ops"),
                                new MutationAction.SetShape(
                                    new ShapeInput(
                                        "OpsShape",
                                        ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                                        firstAnchor,
                                        "rect",
                                        text("Queue")))),
                            mutate(
                                new SheetSelector.ByName("Ops"),
                                new MutationAction.SetEmbeddedObject(
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
                                new MutationAction.SetDrawingObjectAnchor(movedAnchor)),
                            mutate(
                                new DrawingObjectSelector.ByName("Ops", "OpsShape"),
                                new MutationAction.DeleteDrawingObject())),
                        inspect(
                            "drawing",
                            new DrawingObjectSelector.AllOnSheet("Ops"),
                            new InspectionQuery.GetDrawingObjects()),
                        inspect(
                            "payload",
                            new DrawingObjectSelector.ByName("Ops", "OpsEmbed"),
                            new InspectionQuery.GetDrawingObjectPayload()))));

    InspectionResult.DrawingObjectsResult drawingObjects =
        read(success, "drawing", InspectionResult.DrawingObjectsResult.class);
    InspectionResult.DrawingObjectPayloadResult payload =
        read(success, "payload", InspectionResult.DrawingObjectPayloadResult.class);

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
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
                        mutate(
                            new RangeSelector.ByRange("Budget", "A1"),
                            new MutationAction.ApplyStyle(
                                new CellStyleInput(
                                    null,
                                    null,
                                    new CellFontInput(
                                        null,
                                        Boolean.TRUE,
                                        "Aptos",
                                        new FontHeightInput.Twips(260),
                                        "#112233",
                                        null,
                                        null),
                                    null,
                                    null,
                                    null))),
                        mutate(
                            new CellSelector.ByAddress("Budget", "A1"),
                            new MutationAction.SetCell(
                                new CellInput.RichText(
                                    List.of(
                                        richTextRun("Budget"),
                                        new RichTextRunInput(
                                            text(" FY26"),
                                            new CellFontInput(
                                                Boolean.TRUE,
                                                null,
                                                null,
                                                null,
                                                "#FF0000",
                                                null,
                                                null))))))),
                    inspect(
                        "cells",
                        new CellSelector.ByAddresses("Budget", List.of("A1")),
                        new InspectionQuery.GetCells())));

    GridGrindResponse.Success success = success(response);
    InspectionResult.CellsResult cells = read(success, "cells", InspectionResult.CellsResult.class);
    GridGrindResponse.CellReport.TextReport cell =
        cast(GridGrindResponse.CellReport.TextReport.class, cells.cells().getFirst());

    assertEquals("Budget FY26", cell.stringValue());
    assertNotNull(cell.richText());
    assertEquals(2, cell.richText().size());
    assertEquals("Budget", cell.richText().get(0).text());
    assertEquals("Aptos", cell.richText().get(0).font().fontName());
    assertEquals(rgb("#112233"), cell.richText().get(0).font().fontColor());
    assertTrue(cell.richText().get(0).font().italic());
    assertFalse(cell.richText().get(0).font().bold());
    assertEquals(" FY26", cell.richText().get(1).text());
    assertEquals("Aptos", cell.richText().get(1).font().fontName());
    assertEquals(rgb("#FF0000"), cell.richText().get(1).font().fontColor());
    assertTrue(cell.richText().get(1).font().bold());
    assertTrue(cell.richText().get(1).font().italic());
  }

  @Test
  void surfacesRequestWarningsAlongsideSuccessfulExecution() {
    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Budget Review"),
                            new MutationAction.EnsureSheet()),
                        mutate(
                            new SheetSelector.ByName("Summary"), new MutationAction.EnsureSheet()),
                        mutate(
                            new CellSelector.ByAddress("Budget Review", "A1"),
                            new MutationAction.SetCell(new CellInput.Numeric(1200.0))),
                        mutate(
                            new CellSelector.ByAddress("Summary", "A1"),
                            new MutationAction.SetCell(formulaCell("Budget Review!A1"))))));

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

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget").setCell("A1", ExcelCellValue.text("Before"));
      workbook.save(workbookPath);
    }

    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.OverwriteSource(),
                    mutations(
                        mutate(
                            new CellSelector.ByAddress("Budget", "A1"),
                            new MutationAction.SetCell(textCell("After"))),
                        mutate(
                            new CellSelector.ByAddress("Budget", "B1"),
                            new MutationAction.SetCell(new CellInput.Numeric(12.0)))),
                    inspect(
                        "cells",
                        new CellSelector.ByAddresses("Budget", List.of("A1", "B1")),
                        new InspectionQuery.GetCells())));

    GridGrindResponse.Success success = success(response);
    InspectionResult.CellsResult cells = read(success, "cells", InspectionResult.CellsResult.class);

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(
        "After",
        cast(GridGrindResponse.CellReport.TextReport.class, cells.cells().get(0)).stringValue());
    assertEquals(
        12.0,
        cast(GridGrindResponse.CellReport.NumberReport.class, cells.cells().get(1)).numberValue());
  }

  @Test
  void returnsStructuralReadResultsAndPersistsWorkbookShape() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-layout-ops-");

    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
                        mutate(
                            new CellSelector.ByAddress("Budget", "A1"),
                            new MutationAction.SetCell(textCell("Quarterly"))),
                        mutate(
                            new RangeSelector.ByRange("Budget", "A1:B1"),
                            new MutationAction.MergeCells()),
                        mutate(
                            new ColumnBandSelector.Span("Budget", 0, 1),
                            new MutationAction.SetColumnWidth(16.0)),
                        mutate(
                            new RowBandSelector.Span("Budget", 0, 0),
                            new MutationAction.SetRowHeight(28.5)),
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new MutationAction.SetSheetPane(new PaneInput.Frozen(1, 1, 1, 1))),
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new MutationAction.SetSheetZoom(125)),
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new MutationAction.SetSheetPresentation(
                                new SheetPresentationInput(
                                    new SheetDisplayInput(false, false, false, true, true),
                                    new ColorInput("#112233"),
                                    new SheetOutlineSummaryInput(false, false),
                                    new SheetDefaultsInput(11, 18.5d),
                                    List.of(
                                        new IgnoredErrorInput(
                                            "A1:B2",
                                            List.of(
                                                ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT)))))),
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new MutationAction.SetPrintLayout(
                                new PrintLayoutInput(
                                    new PrintAreaInput.Range("A1:B12"),
                                    ExcelPrintOrientation.LANDSCAPE,
                                    new PrintScalingInput.Fit(1, 0),
                                    new PrintTitleRowsInput.Band(0, 0),
                                    new PrintTitleColumnsInput.Band(0, 0),
                                    headerFooter("Budget", "", ""),
                                    headerFooter("", "Page &P", ""),
                                    new PrintSetupInput(
                                        null, true, null, null, null, null, null, null, null, null,
                                        null, null))))),
                    inspect(
                        "cells",
                        new CellSelector.ByAddresses("Budget", List.of("A1")),
                        new InspectionQuery.GetCells()),
                    inspect(
                        "merged",
                        new SheetSelector.ByName("Budget"),
                        new InspectionQuery.GetMergedRegions()),
                    inspect(
                        "layout",
                        new SheetSelector.ByName("Budget"),
                        new InspectionQuery.GetSheetLayout()),
                    inspect(
                        "printLayout",
                        new SheetSelector.ByName("Budget"),
                        new InspectionQuery.GetPrintLayout()),
                    inspect(
                        "workbook",
                        new WorkbookSelector.Current(),
                        new InspectionQuery.GetWorkbookSummary())));

    GridGrindResponse.Success success = success(response);
    InspectionResult.CellsResult cells = read(success, "cells", InspectionResult.CellsResult.class);
    InspectionResult.MergedRegionsResult merged =
        read(success, "merged", InspectionResult.MergedRegionsResult.class);
    GridGrindResponse.SheetLayoutReport layout =
        read(success, "layout", InspectionResult.SheetLayoutResult.class).layout();
    PrintLayoutReport printLayout =
        read(success, "printLayout", InspectionResult.PrintLayoutResult.class).layout();
    GridGrindResponse.WorkbookSummary workbook =
        read(success, "workbook", InspectionResult.WorkbookSummaryResult.class).workbook();

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(
        "Quarterly",
        cast(GridGrindResponse.CellReport.TextReport.class, cells.cells().getFirst())
            .stringValue());
    assertEquals(
        List.of("A1:B1"),
        merged.mergedRegions().stream().map(GridGrindResponse.MergedRegionReport::range).toList());
    assertInstanceOf(PaneReport.Frozen.class, layout.pane());
    PaneReport.Frozen frozen = cast(PaneReport.Frozen.class, layout.pane());
    assertEquals(1, frozen.splitColumn());
    assertEquals(1, frozen.splitRow());
    assertEquals(125, layout.zoomPercent());
    assertEquals(
        new SheetDisplayReport(false, false, false, true, true), layout.presentation().display());
    assertEquals(rgb("#112233"), layout.presentation().tabColor());
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
    dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayout reopenedLayout =
        XlsxRoundTrip.sheetLayout(workbookPath, "Budget");
    assertEquals(
        new ExcelSheetDisplay(false, false, false, true, true),
        reopenedLayout.presentation().display());
    assertEquals(new ExcelColorSnapshot("#112233"), reopenedLayout.presentation().tabColor());
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
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Layout"), new MutationAction.EnsureSheet()),
                        mutate(
                            new RangeSelector.ByRange("Layout", "A1:F6"),
                            new MutationAction.SetRange(
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
                                        new CellInput.Numeric(42.0),
                                        textCell("Open"),
                                        textCell("Alpha"),
                                        textCell("Ada"),
                                        textCell("Y")),
                                    List.of(
                                        textCell("Support"),
                                        new CellInput.Numeric(84.0),
                                        textCell("Closed"),
                                        textCell("Beta"),
                                        textCell("Lin"),
                                        textCell("N")),
                                    List.of(
                                        textCell("Ops"),
                                        new CellInput.Numeric(168.0),
                                        textCell("Open"),
                                        textCell("Gamma"),
                                        textCell("Bea"),
                                        textCell("Y")),
                                    List.of(
                                        textCell("QA"),
                                        new CellInput.Numeric(21.0),
                                        textCell("Queued"),
                                        textCell("Delta"),
                                        textCell("Kai"),
                                        textCell("N")),
                                    List.of(
                                        textCell("Infra"),
                                        new CellInput.Numeric(7.0),
                                        textCell("Done"),
                                        textCell("Epsilon"),
                                        textCell("Mia"),
                                        textCell("Y"))))),
                        mutate(
                            new RowBandSelector.Span("Layout", 1, 3),
                            new MutationAction.GroupRows(true)),
                        mutate(
                            new RowBandSelector.Span("Layout", 5, 5),
                            new MutationAction.SetRowVisibility(true)),
                        mutate(
                            new ColumnBandSelector.Span("Layout", 1, 3),
                            new MutationAction.GroupColumns(true)),
                        mutate(
                            new ColumnBandSelector.Span("Layout", 5, 5),
                            new MutationAction.SetColumnVisibility(true))),
                    inspect(
                        "layout",
                        new SheetSelector.ByName("Layout"),
                        new InspectionQuery.GetSheetLayout())));

    GridGrindResponse.Success success = success(response);
    GridGrindResponse.SheetLayoutReport layout =
        read(success, "layout", InspectionResult.SheetLayoutResult.class).layout();

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

    dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayout reopenedLayout =
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
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Layout");
      workbook.sheet("Layout").setCell("A1", ExcelCellValue.text("Header"));
      workbook.sheet("Layout").setColumnWidth(0, 0, 16.0d);
      workbook.save(sourceWorkbook);
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
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(malformedWorkbook.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(),
                        inspect(
                            "layout",
                            new SheetSelector.ByName("Layout"),
                            new InspectionQuery.GetSheetLayout()))));

    GridGrindResponse.SheetLayoutReport layout =
        read(success, "layout", InspectionResult.SheetLayoutResult.class).layout();
    assertEquals(999.0d, layout.presentation().sheetDefaults().defaultRowHeightPoints());
    assertEquals(300.0d, layout.columns().getFirst().widthCharacters());
    assertEquals(999.0d, layout.rows().getFirst().heightPoints());

    dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayout reopenedLayout =
        XlsxRoundTrip.sheetLayout(malformedWorkbook, "Layout");
    assertEquals(999.0d, reopenedLayout.presentation().sheetDefaults().defaultRowHeightPoints());
    assertEquals(300.0d, reopenedLayout.columns().getFirst().widthCharacters());
    assertEquals(999.0d, reopenedLayout.rows().getFirst().heightPoints());
  }

  @Test
  void executesB3InsertDeleteAndShiftOperationsAndPersistsMovedCells() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-b3-geometry-");

    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    mutations(
                        mutate(new SheetSelector.ByName("Moves"), new MutationAction.EnsureSheet()),
                        mutate(
                            new RangeSelector.ByRange("Moves", "A1:D3"),
                            new MutationAction.SetRange(
                                List.of(
                                    List.of(
                                        textCell("Item"),
                                        textCell("Qty"),
                                        textCell("Status"),
                                        textCell("Note")),
                                    List.of(
                                        textCell("Hosting"),
                                        new CellInput.Numeric(42.0),
                                        textCell("Open"),
                                        textCell("Alpha")),
                                    List.of(
                                        textCell("Support"),
                                        new CellInput.Numeric(84.0),
                                        textCell("Closed"),
                                        textCell("Beta"))))),
                        mutate(
                            new RowBandSelector.Insertion("Moves", 1, 1),
                            new MutationAction.InsertRows()),
                        mutate(
                            new CellSelector.ByAddress("Moves", "A2"),
                            new MutationAction.SetCell(textCell("Spacer"))),
                        mutate(
                            new RowBandSelector.Span("Moves", 2, 3),
                            new MutationAction.ShiftRows(1)),
                        mutate(
                            new RowBandSelector.Span("Moves", 2, 2),
                            new MutationAction.DeleteRows()),
                        mutate(
                            new ColumnBandSelector.Insertion("Moves", 1, 1),
                            new MutationAction.InsertColumns()),
                        mutate(
                            new CellSelector.ByAddress("Moves", "B1"),
                            new MutationAction.SetCell(textCell("Pad"))),
                        mutate(
                            new ColumnBandSelector.Span("Moves", 2, 4),
                            new MutationAction.ShiftColumns(1)),
                        mutate(
                            new ColumnBandSelector.Span("Moves", 2, 2),
                            new MutationAction.DeleteColumns())),
                    inspect(
                        "cells",
                        new CellSelector.ByAddresses(
                            "Moves", List.of("A2", "B1", "A3", "C3", "E4")),
                        new InspectionQuery.GetCells())));

    GridGrindResponse.Success success = success(response);
    InspectionResult.CellsResult cells = read(success, "cells", InspectionResult.CellsResult.class);

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(
        "Spacer",
        cast(GridGrindResponse.CellReport.TextReport.class, cells.cells().get(0)).stringValue());
    assertEquals(
        "Pad",
        cast(GridGrindResponse.CellReport.TextReport.class, cells.cells().get(1)).stringValue());
    assertEquals(
        "Hosting",
        cast(GridGrindResponse.CellReport.TextReport.class, cells.cells().get(2)).stringValue());
    assertEquals(
        42.0,
        cast(GridGrindResponse.CellReport.NumberReport.class, cells.cells().get(3)).numberValue());
    assertEquals(
        "Beta",
        cast(GridGrindResponse.CellReport.TextReport.class, cells.cells().get(4)).stringValue());

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      assertEquals("Spacer", workbook.sheet("Moves").text("A2"));
      assertEquals("Pad", workbook.sheet("Moves").text("B1"));
      assertEquals("Hosting", workbook.sheet("Moves").text("A3"));
      assertEquals(42.0, workbook.sheet("Moves").number("C3"));
      assertEquals("Beta", workbook.sheet("Moves").text("E4"));
    }
  }

  @Test
  void returnsStructuredFailureForRowStructuralEditsThatWouldMoveTables() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A1:B3"),
                                new MutationAction.SetRange(
                                    List.of(
                                        List.of(textCell("Item"), textCell("Qty")),
                                        List.of(textCell("Hosting"), new CellInput.Numeric(42.0)),
                                        List.of(
                                            textCell("Support"), new CellInput.Numeric(84.0))))),
                            mutate(
                                new MutationAction.SetTable(
                                    new TableInput(
                                        "BudgetTable",
                                        "Budget",
                                        "A1:B3",
                                        false,
                                        new TableStyleInput.None()))),
                            mutate(
                                new RowBandSelector.Insertion("Budget", 1, 1),
                                new MutationAction.InsertRows())))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertTrue(failure.problem().message().contains("table 'BudgetTable'"));
    assertTrue(failure.problem().message().contains("row structural edits"));
  }

  @Test
  void returnsStructuredFailureForColumnStructuralEditsWhenWorkbookHasFormulas() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new CellSelector.ByAddress("Budget", "A1"),
                                new MutationAction.SetCell(textCell("Item"))),
                            mutate(
                                new CellSelector.ByAddress("Budget", "B2"),
                                new MutationAction.SetCell(formulaCell("SUM(1, 1)"))),
                            mutate(
                                new ColumnBandSelector.Insertion("Budget", 1, 1),
                                new MutationAction.InsertColumns())))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertTrue(failure.problem().message().contains("workbook formulas are present"));
    assertTrue(failure.problem().message().contains("column structural edits"));
  }

  @Test
  void returnsAuthoringMetadataAndNamedRangeReadResults() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-authoring-ops-");

    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
                        mutate(
                            new CellSelector.ByAddress("Budget", "A1"),
                            new MutationAction.SetCell(textCell("Report"))),
                        mutate(
                            new CellSelector.ByAddress("Budget", "B4"),
                            new MutationAction.SetCell(new CellInput.Numeric(61.0))),
                        mutate(
                            new CellSelector.ByAddress("Budget", "A1"),
                            new MutationAction.SetHyperlink(
                                new HyperlinkTarget.Url("https://example.com/report"))),
                        mutate(
                            new CellSelector.ByAddress("Budget", "A1"),
                            new MutationAction.SetComment(
                                new CommentInput(text("Review"), "GridGrind", true))),
                        mutate(
                            new MutationAction.SetNamedRange(
                                "BudgetTotal",
                                new NamedRangeScope.Workbook(),
                                new NamedRangeTarget("Budget", "B4"))),
                        mutate(
                            new MutationAction.SetNamedRange(
                                "LocalItem",
                                new NamedRangeScope.Sheet("Budget"),
                                new NamedRangeTarget("Budget", "A1:B2")))),
                    inspect(
                        "cells",
                        new CellSelector.ByAddresses("Budget", List.of("A1", "B4")),
                        new InspectionQuery.GetCells()),
                    inspect(
                        "hyperlinks",
                        new CellSelector.ByAddresses("Budget", List.of("A1")),
                        new InspectionQuery.GetHyperlinks()),
                    inspect(
                        "comments",
                        new CellSelector.ByAddresses("Budget", List.of("A1")),
                        new InspectionQuery.GetComments()),
                    inspect(
                        "ranges",
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                        new InspectionQuery.GetNamedRanges())));

    GridGrindResponse.Success success = success(response);
    GridGrindResponse.CellReport.TextReport linkedCell =
        cast(
            GridGrindResponse.CellReport.TextReport.class,
            read(success, "cells", InspectionResult.CellsResult.class).cells().getFirst());
    InspectionResult.HyperlinksResult hyperlinks =
        read(success, "hyperlinks", InspectionResult.HyperlinksResult.class);
    InspectionResult.CommentsResult comments =
        read(success, "comments", InspectionResult.CommentsResult.class);
    InspectionResult.NamedRangesResult ranges =
        read(success, "ranges", InspectionResult.NamedRangesResult.class);

    assertEquals(new HyperlinkTarget.Url("https://example.com/report"), linkedCell.hyperlink());
    assertEquals("Review", linkedCell.comment().text());
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
                new GridGrindResponse.NamedRangeReport.RangeReport(
                    "BudgetTotal",
                    new NamedRangeScope.Workbook(),
                    ranges.namedRanges().stream()
                        .filter(namedRange -> "BudgetTotal".equals(namedRange.name()))
                        .findFirst()
                        .orElseThrow()
                        .refersToFormula(),
                    new NamedRangeTarget("Budget", "B4"))));
    assertTrue(
        ranges
            .namedRanges()
            .contains(
                new GridGrindResponse.NamedRangeReport.RangeReport(
                    "LocalItem",
                    new NamedRangeScope.Sheet("Budget"),
                    ranges.namedRanges().stream()
                        .filter(namedRange -> "LocalItem".equals(namedRange.name()))
                        .findFirst()
                        .orElseThrow()
                        .refersToFormula(),
                    new NamedRangeTarget("Budget", "A1:B2"))));

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
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
                        mutate(
                            new RangeSelector.ByRange("Budget", "A2:C5"),
                            new MutationAction.SetDataValidation(
                                new DataValidationInput(
                                    new DataValidationRuleInput.ExplicitList(
                                        List.of("Queued", "Done")),
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
                            new MutationAction.ClearDataValidations()),
                        mutate(
                            new RangeSelector.ByRange("Budget", "E2:E5"),
                            new MutationAction.SetDataValidation(
                                new DataValidationInput(
                                    new DataValidationRuleInput.FormulaList("#REF!"),
                                    false,
                                    false,
                                    null,
                                    null)))),
                    inspect(
                        "validations",
                        new RangeSelector.AllOnSheet("Budget"),
                        new InspectionQuery.GetDataValidations()),
                    inspect(
                        "health",
                        new SheetSelector.ByNames(List.of("Budget")),
                        new InspectionQuery.AnalyzeDataValidationHealth())));

    GridGrindResponse.Success success = success(response);
    InspectionResult.DataValidationsResult validations =
        read(success, "validations", InspectionResult.DataValidationsResult.class);
    InspectionResult.DataValidationHealthResult health =
        read(success, "health", InspectionResult.DataValidationHealthResult.class);
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
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
                        mutate(
                            new RangeSelector.ByRange("Budget", "A1:B5"),
                            new MutationAction.SetRange(
                                List.of(
                                    List.of(textCell("Status"), textCell("Amount")),
                                    List.of(textCell("Queued"), new CellInput.Numeric(1.0)),
                                    List.of(textCell("Done"), new CellInput.Numeric(9.0)),
                                    List.of(textCell("Done"), new CellInput.Numeric(11.0)),
                                    List.of(textCell("Queued"), new CellInput.Numeric(4.0))))),
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new MutationAction.SetConditionalFormatting(
                                new ConditionalFormattingBlockInput(
                                    List.of("B2:B5"),
                                    List.of(
                                        new ConditionalFormattingRuleInput.FormulaRule(
                                            "B2>5",
                                            true,
                                            new DifferentialStyleInput(
                                                "0.00", true, null, null, "#102030", null, null,
                                                "#E0F0AA", null)),
                                        new ConditionalFormattingRuleInput.CellValueRule(
                                            ExcelComparisonOperator.BETWEEN,
                                            "1",
                                            "10",
                                            false,
                                            new DifferentialStyleInput(
                                                null, null, true, null, null, null, null, null,
                                                null))))))),
                    inspect(
                        "conditional-formatting",
                        new RangeSelector.AllOnSheet("Budget"),
                        new InspectionQuery.GetConditionalFormatting()),
                    inspect(
                        "conditional-formatting-health",
                        new SheetSelector.ByNames(List.of("Budget")),
                        new InspectionQuery.AnalyzeConditionalFormattingHealth())));

    GridGrindResponse.Success success = success(response);
    InspectionResult.ConditionalFormattingResult conditionalFormatting =
        read(success, "conditional-formatting", InspectionResult.ConditionalFormattingResult.class);
    InspectionResult.ConditionalFormattingHealthResult health =
        read(
            success,
            "conditional-formatting-health",
            InspectionResult.ConditionalFormattingHealthResult.class);

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

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
      dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingResult reopened =
          (dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingResult)
              readExecutor
                  .apply(
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
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
                        mutate(
                            new RangeSelector.ByRange("Budget", "A1:C4"),
                            new MutationAction.SetRange(
                                List.of(
                                    List.of(
                                        textCell("Item"), textCell("Amount"), textCell("Billable")),
                                    List.of(
                                        textCell("Hosting"),
                                        new CellInput.Numeric(49.0),
                                        new CellInput.BooleanValue(true)),
                                    List.of(
                                        textCell("Domain"),
                                        new CellInput.Numeric(12.0),
                                        new CellInput.BooleanValue(false)),
                                    List.of(
                                        textCell("Support"),
                                        new CellInput.Numeric(18.0),
                                        new CellInput.BooleanValue(true))))),
                        mutate(
                            new RangeSelector.ByRange("Budget", "E1:F3"),
                            new MutationAction.SetRange(
                                List.of(
                                    List.of(textCell("Queue"), textCell("Owner")),
                                    List.of(textCell("Late invoices"), textCell("Marta")),
                                    List.of(textCell("Badge orders"), textCell("Rihards"))))),
                        mutate(
                            new RangeSelector.ByRange("Budget", "E1:F3"),
                            new MutationAction.SetAutofilter()),
                        mutate(
                            new MutationAction.SetTable(
                                new TableInput(
                                    "BudgetTable",
                                    "Budget",
                                    "A1:C4",
                                    false,
                                    new TableStyleInput.Named(
                                        "TableStyleMedium2", false, false, true, false))))),
                    inspect(
                        "filters",
                        new SheetSelector.ByName("Budget"),
                        new InspectionQuery.GetAutofilters()),
                    inspect("tables", new TableSelector.All(), new InspectionQuery.GetTables()),
                    inspect(
                        "autofilter-health",
                        new SheetSelector.ByNames(List.of("Budget")),
                        new InspectionQuery.AnalyzeAutofilterHealth()),
                    inspect(
                        "table-health",
                        new TableSelector.All(),
                        new InspectionQuery.AnalyzeTableHealth())));

    GridGrindResponse.Success success = success(response);
    InspectionResult.AutofiltersResult filters =
        read(success, "filters", InspectionResult.AutofiltersResult.class);
    InspectionResult.TablesResult tables =
        read(success, "tables", InspectionResult.TablesResult.class);
    InspectionResult.AutofilterHealthResult autofilterHealth =
        read(success, "autofilter-health", InspectionResult.AutofilterHealthResult.class);
    InspectionResult.TableHealthResult tableHealth =
        read(success, "table-health", InspectionResult.TableHealthResult.class);

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(
        List.of(
            new AutofilterEntryReport.SheetOwned("E1:F3"),
            new AutofilterEntryReport.TableOwned("A1:C4", "BudgetTable")),
        filters.autofilters());
    assertEquals(1, tables.tables().size());
    TableEntryReport table = tables.tables().getFirst();
    assertEquals("BudgetTable", table.name());
    assertEquals(List.of("Item", "Amount", "Billable"), table.columnNames());
    assertEquals(
        new TableStyleReport.Named("TableStyleMedium2", false, false, true, false), table.style());
    assertTrue(table.hasAutofilter());
    assertEquals(2, autofilterHealth.analysis().checkedAutofilterCount());
    assertEquals(List.of(), autofilterHealth.analysis().findings());
    assertEquals(1, tableHealth.analysis().checkedTableCount());
    assertEquals(List.of(), tableHealth.analysis().findings());

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
      dev.erst.gridgrind.excel.WorkbookReadResult.AutofiltersResult reopenedFilters =
          (dev.erst.gridgrind.excel.WorkbookReadResult.AutofiltersResult)
              readExecutor
                  .apply(workbook, new WorkbookReadCommand.GetAutofilters("filters", "Budget"))
                  .getFirst();
      dev.erst.gridgrind.excel.WorkbookReadResult.TablesResult reopenedTables =
          (dev.erst.gridgrind.excel.WorkbookReadResult.TablesResult)
              readExecutor
                  .apply(
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
                  1,
                  0,
                  List.of("Item", "Amount", "Billable"),
                  List.of(
                      new dev.erst.gridgrind.excel.ExcelTableColumnSnapshot(
                          1L, "Item", "", "", "", ""),
                      new dev.erst.gridgrind.excel.ExcelTableColumnSnapshot(
                          2L, "Amount", "", "", "", ""),
                      new dev.erst.gridgrind.excel.ExcelTableColumnSnapshot(
                          3L, "Billable", "", "", "", "")),
                  new dev.erst.gridgrind.excel.ExcelTableStyleSnapshot.Named(
                      "TableStyleMedium2", false, false, true, false),
                  true,
                  "",
                  false,
                  false,
                  false,
                  "",
                  "",
                  "")),
          reopenedTables.tables());
    }
  }

  @Test
  void returnsTableReadResultsForByNameSelectionAndNoneStyle() {
    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
                        mutate(
                            new RangeSelector.ByRange("Budget", "A1:B2"),
                            new MutationAction.SetRange(
                                List.of(
                                    List.of(textCell("Item"), textCell("Amount")),
                                    List.of(textCell("Hosting"), new CellInput.Numeric(49.0))))),
                        mutate(
                            new MutationAction.SetTable(
                                new TableInput(
                                    "BudgetTable",
                                    "Budget",
                                    "A1:B2",
                                    false,
                                    new TableStyleInput.None())))),
                    inspect(
                        "tables",
                        new TableSelector.ByNames(List.of("BudgetTable")),
                        new InspectionQuery.GetTables())));

    InspectionResult.TablesResult tables =
        read(success(response), "tables", InspectionResult.TablesResult.class);

    assertEquals(1, tables.tables().size());
    assertEquals(new TableStyleReport.None(), tables.tables().getFirst().style());
  }

  @Test
  void returnsUnsupportedDataValidationReadEntries() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-data-validation-unsupported-");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Budget");
      CTDataValidation validation =
          workbook
              .getSheet("Budget")
              .getCTWorksheet()
              .addNewDataValidations()
              .addNewDataValidation();
      validation.setType(STDataValidationType.NONE);
      validation.setSqref(List.of("A1"));
      workbook.getSheet("Budget").getCTWorksheet().getDataValidations().setCount(1L);
      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    inspect(
                        "validations",
                        new RangeSelector.AllOnSheet("Budget"),
                        new InspectionQuery.GetDataValidations())));

    InspectionResult.DataValidationsResult validations =
        read(success(response), "validations", InspectionResult.DataValidationsResult.class);
    DataValidationEntryReport.Unsupported unsupported =
        assertInstanceOf(
            DataValidationEntryReport.Unsupported.class, validations.validations().getFirst());

    assertEquals(List.of("A1"), unsupported.ranges());
    assertEquals("ANY", unsupported.kind());
  }

  @Test
  void returnsFileHyperlinksInCanonicalPathShape() throws IOException {
    Path linkedFileDirectory = Files.createTempDirectory("gridgrind file links ");
    Path linkedFile = linkedFileDirectory.resolve("memo final.pdf");
    Files.writeString(linkedFile, "seed");

    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
                        mutate(
                            new CellSelector.ByAddress("Budget", "A1"),
                            new MutationAction.SetCell(textCell("Memo"))),
                        mutate(
                            new CellSelector.ByAddress("Budget", "A1"),
                            new MutationAction.SetHyperlink(
                                new HyperlinkTarget.File(linkedFile.toString())))),
                    inspect(
                        "cells",
                        new CellSelector.ByAddresses("Budget", List.of("A1")),
                        new InspectionQuery.GetCells()),
                    inspect(
                        "hyperlinks",
                        new CellSelector.ByAddresses("Budget", List.of("A1")),
                        new InspectionQuery.GetHyperlinks())));

    GridGrindResponse.Success success = success(response);
    GridGrindResponse.CellReport.TextReport linkedCell =
        cast(
            GridGrindResponse.CellReport.TextReport.class,
            read(success, "cells", InspectionResult.CellsResult.class).cells().getFirst());
    InspectionResult.HyperlinksResult hyperlinks =
        read(success, "hyperlinks", InspectionResult.HyperlinksResult.class);

    assertEquals(new HyperlinkTarget.File(linkedFile.toString()), linkedCell.hyperlink());
    assertEquals(
        new HyperlinkTarget.File(linkedFile.toString()),
        hyperlinks.hyperlinks().getFirst().hyperlink());
  }

  @Test
  void convertsEmailAndDocumentHyperlinksToCanonicalProtocolTargets() {
    assertEquals(
        new HyperlinkTarget.Email("team@example.com"),
        InspectionResultCellReportSupport.toHyperlinkTarget(
            new ExcelHyperlink.Email("team@example.com")));
    assertEquals(
        new HyperlinkTarget.Document("Budget!B4"),
        InspectionResultCellReportSupport.toHyperlinkTarget(
            new ExcelHyperlink.Document("Budget!B4")));
  }

  @Test
  void returnsFactualAndAnalysisReadResults() {
    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new MutationAction.AppendRow(
                                List.of(textCell("Item"), textCell("Amount")))),
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new MutationAction.AppendRow(
                                List.of(textCell("Hosting"), new CellInput.Numeric(49.0)))),
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new MutationAction.AppendRow(
                                List.of(textCell("Domain"), new CellInput.Numeric(12.0)))),
                        mutate(
                            new CellSelector.ByAddress("Budget", "B4"),
                            new MutationAction.SetCell(formulaCell("SUM(B2:B3)"))),
                        mutate(
                            new MutationAction.SetNamedRange(
                                "BudgetTotal",
                                new NamedRangeScope.Workbook(),
                                new NamedRangeTarget("Budget", "B4")))),
                    inspect(
                        "formula",
                        new SheetSelector.All(),
                        new InspectionQuery.GetFormulaSurface()),
                    inspect(
                        "schema",
                        new RangeSelector.RectangularWindow("Budget", "A1", 4, 2),
                        new InspectionQuery.GetSheetSchema()),
                    inspect(
                        "ranges",
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                        new InspectionQuery.GetNamedRangeSurface()),
                    inspect(
                        "formula-health",
                        new SheetSelector.All(),
                        new InspectionQuery.AnalyzeFormulaHealth())));

    GridGrindResponse.Success success = success(response);
    GridGrindResponse.FormulaSurfaceReport formula =
        read(success, "formula", InspectionResult.FormulaSurfaceResult.class).analysis();
    GridGrindResponse.SheetSchemaReport schema =
        read(success, "schema", InspectionResult.SheetSchemaResult.class).analysis();
    GridGrindResponse.NamedRangeSurfaceReport ranges =
        read(success, "ranges", InspectionResult.NamedRangeSurfaceResult.class).analysis();
    GridGrindResponse.FormulaHealthReport formulaHealth =
        read(success, "formula-health", InspectionResult.FormulaHealthResult.class).analysis();

    assertEquals(1, formula.totalFormulaCellCount());
    assertEquals("Budget", formula.sheets().getFirst().sheetName());
    assertEquals("SUM(B2:B3)", formula.sheets().getFirst().formulas().getFirst().formula());
    assertEquals(List.of("B4"), formula.sheets().getFirst().formulas().getFirst().addresses());

    assertEquals("Budget", schema.sheetName());
    assertEquals("A1", schema.topLeftAddress());
    assertEquals(4, schema.rowCount());
    assertEquals(2, schema.columnCount());
    assertEquals(3, schema.dataRowCount());
    assertEquals("Item", schema.columns().getFirst().headerDisplayValue());
    assertEquals("STRING", schema.columns().getFirst().dominantType());

    assertEquals(1, ranges.workbookScopedCount());
    assertEquals(0, ranges.sheetScopedCount());
    assertEquals(1, ranges.rangeBackedCount());
    assertEquals(0, ranges.formulaBackedCount());
    assertEquals("BudgetTotal", ranges.namedRanges().getFirst().name());
    assertEquals(
        GridGrindResponse.NamedRangeBackingKind.RANGE, ranges.namedRanges().getFirst().kind());
    assertEquals(1, formulaHealth.checkedFormulaCellCount());
    assertEquals(0, formulaHealth.summary().totalCount());
  }

  @Test
  void returnsStructuredFailureWhenMoveSheetTargetsMissingSheet() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(
                            mutate(
                                new SheetSelector.ByName("Missing"),
                                new MutationAction.MoveSheet(0))))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals("MOVE_SHEET", failure.problem().context().stepType());
    assertEquals("Missing", failure.problem().context().sheetName());
  }

  @Test
  void returnsStructuredFailureForConflictingRenameTarget() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new SheetSelector.ByName("Summary"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.RenameSheet("Summary"))))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals(2, failure.problem().context().stepIndex());
    assertEquals("RENAME_SHEET", failure.problem().context().stepType());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("Sheet already exists: Summary", failure.problem().message());
  }

  @Test
  void returnsStructuredFailureForInvalidMoveSheetIndex() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.MoveSheet(1))))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("MOVE_SHEET", failure.problem().context().stepType());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals(
        "targetIndex out of range: workbook has 1 sheet(s), valid positions are 0 to 0; got 1",
        failure.problem().message());
  }

  @Test
  void rejectsInvalidMergeCellsRangeAtContractConstructionTime() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                mutate(
                    new RangeSelector.ByRange("Budget", "A1:"), new MutationAction.MergeCells()));

    assertEquals("range address must not be blank", failure.getMessage());
  }

  @Test
  void returnsStructuredFailureForUnmergeCellsWithoutExactMatch() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A1:B2"),
                                new MutationAction.UnmergeCells())))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("UNMERGE_CELLS", failure.problem().context().stepType());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("A1:B2", failure.problem().context().range());
    assertEquals("No merged region matches range: A1:B2", failure.problem().message());
  }

  @Test
  void returnsStructuredFailureWhenSetSheetPaneTargetsMissingSheet() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(
                            mutate(
                                new SheetSelector.ByName("Missing"),
                                new MutationAction.SetSheetPane(
                                    new PaneInput.Frozen(1, 1, 1, 1)))))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("SET_SHEET_PANE", failure.problem().context().stepType());
    assertEquals("Missing", failure.problem().context().sheetName());
  }

  @Test
  void returnsStructuredFailureForMissingNamedRangeDuringRead() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet())),
                        inspect(
                            "ranges",
                            new dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf(
                                List.of(
                                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector
                                        .WorkbookScope("BudgetTotal"))),
                            new InspectionQuery.GetNamedRanges()))));

    assertEquals(GridGrindProblemCode.NAMED_RANGE_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals(1, failure.problem().context().stepIndex());
    assertEquals("GET_NAMED_RANGES", failure.problem().context().stepType());
    assertEquals("ranges", failure.problem().context().stepId());
    assertEquals("BudgetTotal", failure.problem().context().namedRangeName());
  }

  @Test
  void preservesStructuralWorkbookStateAcrossExistingWorkbookRoundTrips() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-round-trip-");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var budget = workbook.createSheet("Budget");
      workbook.createSheet("Summary");

      budget.addMergedRegion(CellRangeAddress.valueOf("A1:B2"));
      budget.setColumnWidth(0, 4096);
      budget.createRow(0).setHeightInPoints(28.5f);
      budget.createFreezePane(1, 2, 3, 4);
      budget.createRow(2).createCell(2).setCellValue("Before");

      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.OverwriteSource(),
                        mutations(
                            mutate(
                                new CellSelector.ByAddress("Budget", "C3"),
                                new MutationAction.SetCell(textCell("After")))),
                        inspect(
                            "cells",
                            new CellSelector.ByAddresses("Budget", List.of("C3")),
                            new InspectionQuery.GetCells()))));

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(
        "After",
        cast(
                GridGrindResponse.CellReport.TextReport.class,
                read(success, "cells", InspectionResult.CellsResult.class).cells().getFirst())
            .stringValue());
    assertEquals(List.of("Budget", "Summary"), XlsxRoundTrip.sheetOrder(workbookPath));
    assertEquals(List.of("A1:B2"), XlsxRoundTrip.mergedRegions(workbookPath, "Budget"));
    assertEquals(4096, XlsxRoundTrip.columnWidth(workbookPath, "Budget", 0));
    assertEquals((short) 570, XlsxRoundTrip.rowHeightTwips(workbookPath, "Budget", 0));
    assertEquals(new ExcelSheetPane.Frozen(1, 2, 3, 4), XlsxRoundTrip.pane(workbookPath, "Budget"));
  }

  @Test
  void returnsStructuredFailureForMissingWorkbookSource() {
    Path workbookPath = Path.of("missing-workbook.xlsx");

    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of())));

    assertEquals(GridGrindProblemCode.WORKBOOK_NOT_FOUND, failure.problem().code());
    assertEquals("OPEN_WORKBOOK", failure.problem().context().stage());
    assertEquals(GridGrindProblemCategory.RESOURCE, failure.problem().category());
    assertEquals(GridGrindProblemRecovery.CHANGE_REQUEST, failure.problem().recovery());
    assertEquals(
        workbookPath.toAbsolutePath().toString(), failure.problem().context().sourceWorkbookPath());
  }

  @Test
  void returnsStructuredFailureWithOperationContext() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.AutoSizeColumns())))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals(0, failure.problem().context().stepIndex());
    assertEquals("AUTO_SIZE_COLUMNS", failure.problem().context().stepType());
    assertEquals("Budget", failure.problem().context().sheetName());
  }

  @Test
  void returnsStructuredFailureWithReadContext() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(),
                        inspect(
                            "cells",
                            new RangeSelector.RectangularWindow("Budget", "A1", 5, 5),
                            new InspectionQuery.GetWindow()))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals("GET_WINDOW", failure.problem().context().stepType());
    assertEquals("cells", failure.problem().context().stepId());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("A1", failure.problem().context().address());
  }

  @Test
  void returnsStructuredFailureWhenReadTargetsMissingSheet() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(),
                        inspect(
                            "sheet",
                            new SheetSelector.ByName("Missing"),
                            new InspectionQuery.GetSheetSummary()))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals("GET_SHEET_SUMMARY", failure.problem().context().stepType());
    assertEquals("sheet", failure.problem().context().stepId());
    assertEquals("Missing", failure.problem().context().sheetName());
  }

  @Test
  void returnsStructuredFailureWhenDeletingLastSheet() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.DeleteSheet())))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals(1, failure.problem().context().stepIndex());
    assertEquals("DELETE_SHEET", failure.problem().context().stepType());
    assertTrue(failure.problem().message().contains("at least one sheet"));
  }

  @Test
  void returnsStructuredFailureWhenDeletingLastVisibleSheet() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(
                            mutate(
                                new SheetSelector.ByName("Alpha"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new SheetSelector.ByName("Beta"), new MutationAction.EnsureSheet()),
                            mutate(
                                new SheetSelector.ByName("Beta"),
                                new MutationAction.SetSheetVisibility(ExcelSheetVisibility.HIDDEN)),
                            mutate(
                                new SheetSelector.ByName("Alpha"),
                                new MutationAction.DeleteSheet())))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals(3, failure.problem().context().stepIndex());
    assertEquals("DELETE_SHEET", failure.problem().context().stepType());
    assertEquals("Alpha", failure.problem().context().sheetName());
    assertTrue(failure.problem().message().contains("last visible sheet"));
  }

  @Test
  void returnsStructuredFailureForInvalidPersistTarget() throws IOException {
    Path parentFile = Files.createTempFile("gridgrind-persist-target-", ".tmp");
    Path workbookPath = parentFile.resolve("book.xlsx");

    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        List.of(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet())))));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("PERSIST_WORKBOOK", failure.problem().context().stage());
    assertEquals(
        workbookPath.toAbsolutePath().toString(), failure.problem().context().persistencePath());
  }

  @Test
  void doesNotPersistWorkbookWhenReadFails() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-analysis-failure-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        List.of(),
                        inspect(
                            "cells",
                            new CellSelector.ByAddresses("Missing", List.of("A1")),
                            new InspectionQuery.GetCells()))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertFalse(Files.exists(workbookPath));
  }

  @Test
  void rejectsMalformedCellAddressesAtContractConstructionTime() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                inspect(
                    "cells",
                    new CellSelector.ByAddresses("Data", List.of("A1", "BADADDR")),
                    new InspectionQuery.GetCells()));

    assertEquals(
        "addresses[1] address must be a single-cell A1-style address", failure.getMessage());
  }

  @Test
  void rejectsZeroRowCellAddressesAtContractConstructionTime() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                inspect(
                    "cells",
                    new CellSelector.ByAddresses("Data", List.of("A0")),
                    new InspectionQuery.GetCells()));

    assertEquals(
        "addresses[0] address must be a single-cell A1-style address", failure.getMessage());
  }

  @Test
  void returnsStructuredFailureForInvalidOverwriteSourceUsage() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.OverwriteSource(),
                        List.of())));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("VALIDATE_REQUEST", failure.problem().context().stage());
  }

  @Test
  void returnsFormulaErrorForInvalidFormulaOperations() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        executionPolicy(calculateAll()),
                        null,
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Data"), new MutationAction.EnsureSheet()),
                            mutate(
                                new CellSelector.ByAddress("Data", "A1"),
                                new MutationAction.SetCell(formulaCell("SUM(")))),
                        inspections())));

    assertEquals(GridGrindProblemCode.INVALID_FORMULA, failure.problem().code());
    assertEquals(GridGrindProblemCategory.FORMULA, failure.problem().category());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals("Invalid formula at Data!A1: SUM(", failure.problem().message());
    assertEquals("A1", failure.problem().context().address());
    assertEquals("SUM(", failure.problem().context().formula());
  }

  @Test
  void returnsFormulaErrorForMalformedParserStateFormulaOperations() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        executionPolicy(calculateAll()),
                        null,
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Data"), new MutationAction.EnsureSheet()),
                            mutate(
                                new CellSelector.ByAddress("Data", "A1"),
                                new MutationAction.SetCell(formulaCell("[^owe_e`ffffff")))))));

    assertEquals(GridGrindProblemCode.INVALID_FORMULA, failure.problem().code());
    assertEquals(GridGrindProblemCategory.FORMULA, failure.problem().category());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals("Invalid formula at Data!A1: [^owe_e`ffffff", failure.problem().message());
    assertEquals("A1", failure.problem().context().address());
    assertEquals("[^owe_e`ffffff", failure.problem().context().formula());
    assertEquals(
        "Invalid formula at Data!A1: [^owe_e`ffffff",
        failure.problem().causes().getFirst().message());
    assertEquals(
        GridGrindProblemCode.INVALID_FORMULA, failure.problem().causes().getFirst().code());
    assertEquals("EXECUTE_STEP", failure.problem().causes().getFirst().stage());
  }

  @Test
  void returnsInvalidFormulaForLambdaAndLetFormulaOperations() {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();

    GridGrindResponse.Failure lambdaFailure =
        failure(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    mutations(
                        mutate(new SheetSelector.ByName("Data"), new MutationAction.EnsureSheet()),
                        mutate(
                            new CellSelector.ByAddress("Data", "A1"),
                            new MutationAction.SetCell(formulaCell("LAMBDA(x,x*2)(5)")))))));
    GridGrindResponse.Failure letFailure =
        failure(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    mutations(
                        mutate(new SheetSelector.ByName("Data"), new MutationAction.EnsureSheet()),
                        mutate(
                            new CellSelector.ByAddress("Data", "A1"),
                            new MutationAction.SetCell(formulaCell("LET(x,5,x*2)")))))));

    assertEquals(GridGrindProblemCode.INVALID_FORMULA, lambdaFailure.problem().code());
    assertEquals("Invalid formula at Data!A1: LAMBDA(x,x*2)(5)", lambdaFailure.problem().message());
    assertEquals("A1", lambdaFailure.problem().context().address());
    assertEquals("LAMBDA(x,x*2)(5)", lambdaFailure.problem().context().formula());

    assertEquals(GridGrindProblemCode.INVALID_FORMULA, letFailure.problem().code());
    assertEquals("Invalid formula at Data!A1: LET(x,5,x*2)", letFailure.problem().message());
    assertEquals("A1", letFailure.problem().context().address());
    assertEquals("LET(x,5,x*2)", letFailure.problem().context().formula());
  }

  @Test
  void surfacesWorkbookFormulaLocationWhenEvaluationFails() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        executionPolicy(calculateAll()),
                        null,
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Data"), new MutationAction.EnsureSheet()),
                            mutate(
                                new CellSelector.ByAddress("Data", "A1"),
                                new MutationAction.SetCell(new CellInput.Numeric(1.0))),
                            mutate(
                                new CellSelector.ByAddress("Data", "B1"),
                                new MutationAction.SetCell(new CellInput.Numeric(2.0))),
                            mutate(
                                new CellSelector.ByAddress("Data", "C1"),
                                new MutationAction.SetCell(
                                    formulaCell("TEXTAFTER(\"a,b\",\",\")")))),
                        inspections())));

    assertEquals(GridGrindProblemCode.UNREGISTERED_USER_DEFINED_FUNCTION, failure.problem().code());
    assertEquals("CALCULATION_PREFLIGHT", failure.problem().context().stage());
    assertEquals(
        "User-defined function TEXTAFTER is not registered at Data!C1: TEXTAFTER(\"a,b\",\",\")",
        failure.problem().message());
    assertEquals("Data", failure.problem().context().sheetName());
    assertEquals("C1", failure.problem().context().address());
    assertEquals("TEXTAFTER(\"a,b\",\",\")", failure.problem().context().formula());
  }

  @Test
  void calculationPolicyFailureRejectsMutationsAfterObservationSteps() {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();

    Optional<String> failure =
        executor.calculationPolicyFailure(
            new WorkbookPlan(
                GridGrindProtocolVersion.current(),
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                executionPolicy(calculateAll()),
                null,
                List.<WorkbookStep>of(
                    new MutationStep(
                        "step-0",
                        new SheetSelector.ByName("Data"),
                        new MutationAction.EnsureSheet()),
                    new InspectionStep(
                        "summary",
                        new WorkbookSelector.Current(),
                        new InspectionQuery.GetWorkbookSummary()),
                    new MutationStep(
                        "step-2",
                        new CellSelector.ByAddress("Data", "A1"),
                        new MutationAction.SetCell(new CellInput.Numeric(1.0))))));

    assertTrue(failure.orElseThrow().contains("mutation-to-observation boundary"));
  }

  @Test
  void executionModeFailureRejectsCalculationForEventReadAndStreamingWrite() {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();

    Optional<String> eventReadFailure =
        executor.executionModeFailure(
            request(
                new WorkbookPlan.WorkbookSource.ExistingFile("/tmp/book.xlsx"),
                new WorkbookPlan.WorkbookPersistence.None(),
                executionPolicy(
                    new ExecutionModeInput(
                        ExecutionModeInput.ReadMode.EVENT_READ,
                        ExecutionModeInput.WriteMode.FULL_XSSF),
                    calculateAll()),
                null,
                List.of(),
                List.of(
                    inspect(
                        "summary",
                        new WorkbookSelector.Current(),
                        new InspectionQuery.GetWorkbookSummary()))));
    Optional<String> streamingFailure =
        executor.executionModeFailure(
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                executionPolicy(
                    new ExecutionModeInput(
                        ExecutionModeInput.ReadMode.FULL_XSSF,
                        ExecutionModeInput.WriteMode.STREAMING_WRITE),
                    calculateAll()),
                null,
                List.of(mutate(new SheetSelector.ByName("Ops"), new MutationAction.EnsureSheet())),
                List.of(),
                List.of()));

    assertTrue(eventReadFailure.orElseThrow().contains("markRecalculateOnOpen=false"));
    assertTrue(streamingFailure.orElseThrow().contains("low-memory streaming writes"));
  }

  @Test
  void returnsCalculationFailureBeforeObservationStepsWhenBoundaryPreflightFails() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        executionPolicy(calculateAll()),
                        null,
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Data"), new MutationAction.EnsureSheet()),
                            mutate(
                                new CellSelector.ByAddress("Data", "C1"),
                                new MutationAction.SetCell(
                                    formulaCell("TEXTAFTER(\"a,b\",\",\")")))),
                        inspect(
                            "summary",
                            new WorkbookSelector.Current(),
                            new InspectionQuery.GetWorkbookSummary()))));

    assertEquals(GridGrindProblemCode.UNREGISTERED_USER_DEFINED_FUNCTION, failure.problem().code());
    assertEquals("CALCULATION_PREFLIGHT", failure.problem().context().stage());
    assertEquals("Data", failure.problem().context().sheetName());
    assertEquals("C1", failure.problem().context().address());
    assertEquals("TEXTAFTER(\"a,b\",\",\")", failure.problem().context().formula());
  }

  @Test
  void returnsStructuredFailureWhenWorkbookCloseFailsAfterSuccess() {
    DefaultGridGrindRequestExecutor executor =
        new DefaultGridGrindRequestExecutor(
            new WorkbookCommandExecutor(),
            new WorkbookReadExecutor(),
            workbook -> {
              throw new IOException("close failed");
            });

    GridGrindResponse.Failure failure =
        failure(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new MutationAction.EnsureSheet())))));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("EXECUTE_REQUEST", failure.problem().context().stage());
    assertEquals("close failed", failure.problem().message());
    assertEquals(GridGrindProblemRecovery.CHECK_ENVIRONMENT, failure.problem().recovery());
  }

  @Test
  void preservesPrimaryFailureWhenWorkbookCloseAlsoFails() {
    DefaultGridGrindRequestExecutor executor =
        new DefaultGridGrindRequestExecutor(
            new WorkbookCommandExecutor(),
            new WorkbookReadExecutor(),
            workbook -> {
              throw new IOException("close failed");
            });

    GridGrindResponse.Failure failure =
        failure(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new MutationAction.AutoSizeColumns())))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals(2, failure.problem().causes().size());
    assertEquals(
        GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().causes().getFirst().code());
    assertTrue(failure.problem().causes().get(1).message().contains("close failed"));
    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().causes().get(1).code());
    assertEquals("EXECUTE_REQUEST", failure.problem().causes().get(1).stage());
  }

  @Test
  void returnsStructuredFailureForNullRequests() {
    GridGrindResponse.Failure failure =
        failure(new DefaultGridGrindRequestExecutor().execute(null));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("VALIDATE_REQUEST", failure.problem().context().stage());
  }

  @Test
  void classifiesProblemCodesAndMessagesDeterministically() {
    assertEquals(
        GridGrindProblemCode.INVALID_CELL_ADDRESS,
        ExecutionResponseSupport.problemCodeFor(
            new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.INVALID_RANGE_ADDRESS,
        ExecutionResponseSupport.problemCodeFor(
            new InvalidRangeAddressException("A1:", new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.INVALID_FORMULA,
        ExecutionResponseSupport.problemCodeFor(
            new InvalidFormulaException(
                "Budget", "B4", "SUM(", "bad formula", new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.UNSUPPORTED_FORMULA,
        ExecutionResponseSupport.problemCodeFor(
            new UnsupportedFormulaException(
                "Budget",
                "C1",
                "LAMBDA(x,x+1)(2)",
                "unsupported",
                new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.WORKBOOK_NOT_FOUND,
        ExecutionResponseSupport.problemCodeFor(
            new WorkbookNotFoundException(Path.of("missing.xlsx"))));
    assertEquals(
        GridGrindProblemCode.SHEET_NOT_FOUND,
        ExecutionResponseSupport.problemCodeFor(new SheetNotFoundException("Budget")));
    assertEquals(
        GridGrindProblemCode.IO_ERROR,
        ExecutionResponseSupport.problemCodeFor(new IOException("disk")));
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST,
        ExecutionResponseSupport.problemCodeFor(new IllegalArgumentException("bad request")));
    assertEquals(
        GridGrindProblemCode.INTERNAL_ERROR,
        ExecutionResponseSupport.problemCodeFor(new UnsupportedOperationException()));
    assertEquals(
        "bad request", GridGrindProblems.messageFor(new IllegalArgumentException("bad request")));
    assertEquals(
        "UnsupportedOperationException",
        GridGrindProblems.messageFor(new UnsupportedOperationException()));
  }

  @Test
  void executesRangeAndStyleWorkflowAndSurfacesStyledCells() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-range-style-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        executionPolicy(calculateAll()),
                        null,
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A1:B2"),
                                new MutationAction.SetRange(
                                    List.of(
                                        List.of(textCell("Item"), textCell("Amount")),
                                        List.of(
                                            textCell("Hosting"), new CellInput.Numeric(49.0))))),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A1:B1"),
                                new MutationAction.ApplyStyle(
                                    new CellStyleInput(
                                        "#,##0.00",
                                        new CellAlignmentInput(
                                            true,
                                            ExcelHorizontalAlignment.CENTER,
                                            ExcelVerticalAlignment.CENTER,
                                            null,
                                            null),
                                        new CellFontInput(true, null, null, null, null, null, null),
                                        null,
                                        null,
                                        null))),
                            mutate(
                                new RangeSelector.ByRange("Budget", "C1"),
                                new MutationAction.ApplyStyle(
                                    new CellStyleInput(
                                        null,
                                        new CellAlignmentInput(
                                            null,
                                            ExcelHorizontalAlignment.RIGHT,
                                            ExcelVerticalAlignment.BOTTOM,
                                            null,
                                            null),
                                        new CellFontInput(null, true, null, null, null, null, null),
                                        null,
                                        null,
                                        null))),
                            mutate(
                                new CellSelector.ByAddress("Budget", "B3"),
                                new MutationAction.SetCell(formulaCell("SUM(B2:B2)"))),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A2"),
                                new MutationAction.ClearRange())),
                        inspect(
                            "cells",
                            new CellSelector.ByAddresses("Budget", List.of("A1", "A2", "B3", "C1")),
                            new InspectionQuery.GetCells()),
                        inspect(
                            "window",
                            new RangeSelector.RectangularWindow("Budget", "A1", 3, 3),
                            new InspectionQuery.GetWindow()))));

    InspectionResult.CellsResult cells = read(success, "cells", InspectionResult.CellsResult.class);
    GridGrindResponse.WindowReport window =
        read(success, "window", InspectionResult.WindowResult.class).window();

    assertTrue(Files.exists(workbookPath));
    assertEquals(
        ExcelHorizontalAlignment.CENTER,
        cells.cells().getFirst().style().alignment().horizontalAlignment());
    assertTrue(cells.cells().getFirst().style().font().bold());
    assertTrue(cells.cells().getFirst().style().alignment().wrapText());
    assertEquals("BLANK", cells.cells().get(1).effectiveType());
    assertEquals("49", cells.cells().get(2).displayValue());
    assertTrue(
        window.rows().getFirst().cells().stream().anyMatch(cell -> "C1".equals(cell.address())));
    assertTrue(cells.cells().get(3).style().font().italic());
  }

  @Test
  void executesFormattingDepthWorkflowAndPersistsReportedStyleState() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-format-depth-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new CellSelector.ByAddress("Budget", "A1"),
                                new MutationAction.SetCell(textCell("Item"))),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A1"),
                                new MutationAction.ApplyStyle(
                                    new CellStyleInput(
                                        null,
                                        new CellAlignmentInput(
                                            true,
                                            ExcelHorizontalAlignment.CENTER,
                                            ExcelVerticalAlignment.TOP,
                                            45,
                                            3),
                                        new CellFontInput(
                                            true,
                                            false,
                                            "Aptos",
                                            new FontHeightInput.Points(new BigDecimal("11.5")),
                                            "#1F4E78",
                                            true,
                                            true),
                                        new CellFillInput(
                                            dev.erst.gridgrind.excel.foundation.ExcelFillPattern
                                                .THIN_HORIZONTAL_BANDS,
                                            "#FFF2CC",
                                            "#DDEBF7"),
                                        new CellBorderInput(
                                            new CellBorderSideInput(
                                                ExcelBorderStyle.THIN, "#102030"),
                                            null,
                                            new CellBorderSideInput(
                                                ExcelBorderStyle.DOUBLE, "#203040"),
                                            null,
                                            null),
                                        new CellProtectionInput(false, true))))),
                        inspect(
                            "cells",
                            new CellSelector.ByAddresses("Budget", List.of("A1")),
                            new InspectionQuery.GetCells()))));

    GridGrindResponse.CellStyleReport style =
        read(success, "cells", InspectionResult.CellsResult.class).cells().getFirst().style();

    assertTrue(Files.exists(workbookPath));
    assertTrue(style.font().bold());
    assertFalse(style.font().italic());
    assertTrue(style.alignment().wrapText());
    assertEquals(ExcelHorizontalAlignment.CENTER, style.alignment().horizontalAlignment());
    assertEquals(ExcelVerticalAlignment.TOP, style.alignment().verticalAlignment());
    assertEquals(45, style.alignment().textRotation());
    assertEquals(3, style.alignment().indentation());
    assertEquals("Aptos", style.font().fontName());
    assertEquals(230, style.font().fontHeight().twips());
    assertEquals(new BigDecimal("11.5"), style.font().fontHeight().points());
    assertEquals(rgb("#1F4E78"), style.font().fontColor());
    assertTrue(style.font().underline());
    assertTrue(style.font().strikeout());
    assertEquals(
        dev.erst.gridgrind.excel.foundation.ExcelFillPattern.THIN_HORIZONTAL_BANDS,
        style.fill().pattern());
    assertEquals(rgb("#FFF2CC"), style.fill().foregroundColor());
    assertEquals(rgb("#DDEBF7"), style.fill().backgroundColor());
    assertEquals(ExcelBorderStyle.THIN, style.border().top().style());
    assertEquals(ExcelBorderStyle.DOUBLE, style.border().right().style());
    assertEquals(ExcelBorderStyle.THIN, style.border().bottom().style());
    assertEquals(ExcelBorderStyle.THIN, style.border().left().style());
    assertEquals(rgb("#102030"), style.border().top().color());
    assertEquals(rgb("#203040"), style.border().right().color());
    assertFalse(style.protection().locked());
    assertTrue(style.protection().hiddenFormula());
    assertEquals(
        style, toResponseStyleReport(XlsxRoundTrip.cellStyle(workbookPath, "Budget", "A1")));
  }

  @Test
  void executesAdvancedStyleWorkflowWithThemeIndexedAndGradientColors() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-advanced-style-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A1:A2"),
                                new MutationAction.SetRange(
                                    List.of(
                                        List.of(textCell("ThemeTintStyle")),
                                        List.of(textCell("GradientFillStyle"))))),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A1"),
                                new MutationAction.ApplyStyle(
                                    new CellStyleInput(
                                        null,
                                        null,
                                        new CellFontInput(
                                            null, true, null, null, null, 6, null, -0.35d, null,
                                            null),
                                        new CellFillInput(
                                            dev.erst.gridgrind.excel.foundation.ExcelFillPattern
                                                .SOLID,
                                            null,
                                            3,
                                            null,
                                            0.30d,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null),
                                        new CellBorderInput(
                                            null,
                                            null,
                                            null,
                                            new CellBorderSideInput(
                                                ExcelBorderStyle.THIN,
                                                null,
                                                null,
                                                Short.toUnsignedInt(
                                                    IndexedColors.DARK_RED.getIndex()),
                                                null),
                                            null),
                                        null))),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A2"),
                                new MutationAction.ApplyStyle(
                                    new CellStyleInput(
                                        null,
                                        null,
                                        null,
                                        new CellFillInput(
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            new CellGradientFillInput(
                                                "LINEAR",
                                                45.0d,
                                                null,
                                                null,
                                                null,
                                                null,
                                                List.of(
                                                    new CellGradientStopInput(
                                                        0.0d, new ColorInput("#1F497D")),
                                                    new CellGradientStopInput(
                                                        1.0d,
                                                        new ColorInput(null, 4, null, 0.45d))))),
                                        null,
                                        new CellProtectionInput(true, true))))),
                        inspect(
                            "cells",
                            new CellSelector.ByAddresses("Budget", List.of("A1", "A2")),
                            new InspectionQuery.GetCells()))));

    InspectionResult.CellsResult cells = read(success, "cells", InspectionResult.CellsResult.class);
    GridGrindResponse.CellStyleReport themedStyle = cells.cells().get(0).style();
    GridGrindResponse.CellStyleReport gradientStyle = cells.cells().get(1).style();

    assertTrue(Files.exists(workbookPath));
    assertEquals(new CellColorReport(null, 6, null, -0.35d), themedStyle.font().fontColor());
    assertEquals(new CellColorReport(null, 3, null, 0.30d), themedStyle.fill().foregroundColor());
    assertEquals(
        new CellColorReport(
            null, null, Short.toUnsignedInt(IndexedColors.DARK_RED.getIndex()), null),
        themedStyle.border().bottom().color());
    assertNotNull(gradientStyle.fill().gradient());
    assertEquals("LINEAR", gradientStyle.fill().gradient().type());
    assertEquals(45.0d, gradientStyle.fill().gradient().degree());
    assertEquals(
        new CellColorReport("#1F497D", null, null, null),
        gradientStyle.fill().gradient().stops().get(0).color());
    assertEquals(
        new CellColorReport(null, 4, null, 0.45d),
        gradientStyle.fill().gradient().stops().get(1).color());
    assertTrue(gradientStyle.protection().locked());
    assertTrue(gradientStyle.protection().hiddenFormula());
  }

  @Test
  void preservesDistinctLinearAndPathGradientStylesInSameRequest() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-distinct-gradients-", ".xlsx");
    assertDoesNotThrow(() -> Files.deleteIfExists(workbookPath));
    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new CellSelector.ByAddress("Budget", "A2"),
                                new MutationAction.SetCell(textCell("Linear gradient"))),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A2"),
                                new MutationAction.ApplyStyle(
                                    new CellStyleInput(
                                        null,
                                        null,
                                        null,
                                        new CellFillInput(
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            new CellGradientFillInput(
                                                "LINEAR",
                                                45.0d,
                                                null,
                                                null,
                                                null,
                                                null,
                                                List.of(
                                                    new CellGradientStopInput(
                                                        0.0d, new ColorInput("#1F497D")),
                                                    new CellGradientStopInput(
                                                        1.0d,
                                                        new ColorInput(null, 4, null, 0.45d))))),
                                        null,
                                        new CellProtectionInput(true, true)))),
                            mutate(
                                new CellSelector.ByAddress("Budget", "A3"),
                                new MutationAction.SetCell(textCell("Path gradient"))),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A3"),
                                new MutationAction.ApplyStyle(
                                    new CellStyleInput(
                                        null,
                                        null,
                                        null,
                                        new CellFillInput(
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            new CellGradientFillInput(
                                                "PATH",
                                                null,
                                                0.1d,
                                                0.2d,
                                                0.3d,
                                                0.4d,
                                                List.of(
                                                    new CellGradientStopInput(
                                                        0.0d, new ColorInput("#112233")),
                                                    new CellGradientStopInput(
                                                        1.0d,
                                                        new ColorInput(
                                                            null,
                                                            null,
                                                            Short.toUnsignedInt(
                                                                IndexedColors.DARK_RED.getIndex()),
                                                            null))))),
                                        null,
                                        new CellProtectionInput(false, true))))),
                        inspect(
                            "cells",
                            new CellSelector.ByAddresses("Budget", List.of("A2", "A3")),
                            new InspectionQuery.GetCells()))));

    InspectionResult.CellsResult cells = read(success, "cells", InspectionResult.CellsResult.class);
    GridGrindResponse.CellStyleReport linearGradientStyle = cells.cells().get(0).style();
    GridGrindResponse.CellStyleReport pathGradientStyle = cells.cells().get(1).style();

    assertEquals("LINEAR", linearGradientStyle.fill().gradient().type());
    assertEquals(45.0d, linearGradientStyle.fill().gradient().degree());
    assertTrue(linearGradientStyle.protection().locked());
    assertTrue(linearGradientStyle.protection().hiddenFormula());
    assertEquals("PATH", pathGradientStyle.fill().gradient().type());
    assertNull(pathGradientStyle.fill().gradient().degree());
    assertEquals(0.1d, pathGradientStyle.fill().gradient().left());
    assertEquals(0.2d, pathGradientStyle.fill().gradient().right());
    assertEquals(0.3d, pathGradientStyle.fill().gradient().top());
    assertEquals(0.4d, pathGradientStyle.fill().gradient().bottom());
    assertFalse(pathGradientStyle.protection().locked());
    assertTrue(pathGradientStyle.protection().hiddenFormula());
  }

  @Test
  void producesErrorReportForCellsWithErrorValues() {
    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        executionPolicy(calculateAll()),
                        null,
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Data"), new MutationAction.EnsureSheet()),
                            mutate(
                                new CellSelector.ByAddress("Data", "A1"),
                                new MutationAction.SetCell(formulaCell("1/0")))),
                        inspect(
                            "cells",
                            new CellSelector.ByAddresses("Data", List.of("A1")),
                            new InspectionQuery.GetCells()))));

    GridGrindResponse.CellReport cell =
        read(success, "cells", InspectionResult.CellsResult.class).cells().getFirst();
    assertInstanceOf(GridGrindResponse.CellReport.FormulaReport.class, cell);
    GridGrindResponse.CellReport evaluation =
        cast(GridGrindResponse.CellReport.FormulaReport.class, cell).evaluation();
    assertInstanceOf(GridGrindResponse.CellReport.ErrorReport.class, evaluation);
    assertEquals("ERROR", evaluation.effectiveType());
  }

  @Test
  void extractsFormulaFromSetCellOperationWhenExceptionCarriesNone() {
    RuntimeException exception = new RuntimeException("test");
    ExecutorTestPlanSupport.PendingMutation operation =
        mutate(
            new CellSelector.ByAddress("Data", "A1"),
            new MutationAction.SetCell(formulaCell("SUM(B1:B2)")));

    assertEquals("SUM(B1:B2)", formulaFor(operation, exception));
    assertEquals("Data", sheetNameFor(operation, exception));
    assertEquals("A1", addressFor(operation, exception));
    assertNull(rangeFor(operation, exception));
  }

  @Test
  void persistencePathResolvesCorrectlyForAllPersistenceAndSourceCombinations() {
    WorkbookPlan.WorkbookSource newSource = new WorkbookPlan.WorkbookSource.New();
    WorkbookPlan.WorkbookSource existingFile =
        new WorkbookPlan.WorkbookSource.ExistingFile("/tmp/source.xlsx");
    WorkbookPlan.WorkbookPersistence none = new WorkbookPlan.WorkbookPersistence.None();
    WorkbookPlan.WorkbookPersistence overwrite =
        new WorkbookPlan.WorkbookPersistence.OverwriteSource();
    WorkbookPlan.WorkbookPersistence saveAs =
        new WorkbookPlan.WorkbookPersistence.SaveAs("/tmp/out.xlsx");

    assertEquals(
        Path.of("/tmp/out.xlsx").toAbsolutePath().normalize().toString(),
        ExecutionRequestPaths.persistencePath(newSource, saveAs));
    assertEquals(
        Path.of("/tmp/source.xlsx").toAbsolutePath().normalize().toString(),
        ExecutionRequestPaths.persistencePath(existingFile, overwrite));
    assertNull(ExecutionRequestPaths.persistencePath(newSource, overwrite));
    assertNull(ExecutionRequestPaths.persistencePath(newSource, none));
    assertNull(ExecutionRequestPaths.persistencePath(existingFile, none));
  }

  @Test
  void persistWorkbookRejectsOverwriteForNewSources() throws Exception {
    ExecutionWorkbookSupport workbookSupport = new ExecutionWorkbookSupport(Files::createTempFile);

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      IllegalArgumentException exception =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  workbookSupport.persistWorkbook(
                      workbook,
                      new WorkbookPlan.WorkbookSource.New(),
                      new WorkbookPlan.WorkbookPersistence.OverwriteSource()));

      assertEquals("OVERWRITE persistence requires an EXISTING source", exception.getMessage());
    }
  }

  @Test
  void persistencePathNormalizesDoubleDotSegments() {
    WorkbookPlan.WorkbookSource newSource = new WorkbookPlan.WorkbookSource.New();
    WorkbookPlan.WorkbookPersistence saveAs =
        new WorkbookPlan.WorkbookPersistence.SaveAs("/tmp/subdir/../out.xlsx");

    assertEquals("/tmp/out.xlsx", ExecutionRequestPaths.persistencePath(newSource, saveAs));
  }

  @Test
  void persistWorkbookSaveAsReportsNormalizedExecutionPath() throws Exception {
    ExecutionWorkbookSupport workbookSupport = new ExecutionWorkbookSupport(Files::createTempFile);
    Path tempDir = Files.createTempDirectory("gridgrind-normalize-test-");
    Path subDir = Files.createDirectory(tempDir.resolve("subdir"));
    String pathWithDotDot = subDir + "/../out.xlsx";

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      GridGrindResponse.PersistenceOutcome outcome =
          workbookSupport.persistWorkbook(
              workbook,
              new WorkbookPlan.WorkbookSource.New(),
              new WorkbookPlan.WorkbookPersistence.SaveAs(pathWithDotDot));

      GridGrindResponse.PersistenceOutcome.SavedAs savedAs =
          assertInstanceOf(GridGrindResponse.PersistenceOutcome.SavedAs.class, outcome);
      assertEquals(pathWithDotDot, savedAs.requestedPath());
      assertEquals(tempDir.resolve("out.xlsx").toString(), savedAs.executionPath());
    } finally {
      Files.deleteIfExists(tempDir.resolve("out.xlsx"));
      Files.deleteIfExists(subDir);
      Files.deleteIfExists(tempDir);
    }
  }

  @Test
  void guardsCatastrophicRuntimeExceptionsAndProducesExecuteRequestFailure() {
    int[] callCount = {0};
    DefaultGridGrindRequestExecutor executor =
        new DefaultGridGrindRequestExecutor(
            new WorkbookCommandExecutor(),
            new WorkbookReadExecutor(),
            workbook -> {
              int count = callCount[0];
              callCount[0] = count + 1;
              if (count == 0) {
                throw new IllegalStateException("catastrophic close failure");
              }
            });

    GridGrindResponse.Failure failure =
        failure(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new MutationAction.EnsureSheet())))));

    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.problem().code());
    assertEquals("EXECUTE_REQUEST", failure.problem().context().stage());
  }

  @Test
  void rejectsInvalidClearRangeSelectorsAtContractConstructionTime() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                mutate(
                    new RangeSelector.ByRange("Budget", "A1:"), new MutationAction.ClearRange()));

    assertEquals("range address must not be blank", failure.getMessage());
  }

  @Test
  void rejectsInvalidSetRangeSelectorsAtContractConstructionTime() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                mutate(
                    new RangeSelector.ByRange("Budget", "A1:"),
                    new MutationAction.SetRange(List.of(List.of(textCell("x"))))));

    assertEquals("range address must not be blank", failure.getMessage());
  }

  @Test
  void rejectsInvalidApplyStyleSelectorsAtContractConstructionTime() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                mutate(
                    new RangeSelector.ByRange("Budget", "A1:"),
                    new MutationAction.ApplyStyle(
                        new CellStyleInput(
                            null,
                            null,
                            new CellFontInput(true, null, null, null, null, null, null),
                            null,
                            null,
                            null))));

    assertEquals("range address must not be blank", failure.getMessage());
  }

  @Test
  void returnsStructuredFailureForAppendRowWithInvalidFormula() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.AppendRow(List.of(formulaCell("SUM("))))))));

    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals("APPEND_ROW", failure.problem().context().stepType());
    assertEquals("Budget", failure.problem().context().sheetName());
  }

  @Test
  void extractsNullContextForOperationsWithNoSheetAddressRangeOrFormula() {
    RuntimeException exception = new RuntimeException("test");
    ExecutorTestPlanSupport.PendingMutation clearWorkbookProtection =
        mutate(new WorkbookSelector.Current(), new MutationAction.ClearWorkbookProtection());
    ExecutorTestPlanSupport.PendingMutation setWorkbookProtection =
        mutate(
            new WorkbookSelector.Current(),
            new MutationAction.SetWorkbookProtection(
                new WorkbookProtectionInput(true, false, false, null, null)));
    ExecutorTestPlanSupport.PendingMutation appendRow =
        mutate(
            new SheetSelector.ByName("Budget"),
            new MutationAction.AppendRow(List.of(textCell("x"))));
    ExecutorTestPlanSupport.PendingMutation ensureSheet =
        mutate(new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet());

    assertNull(formulaFor(clearWorkbookProtection, exception));
    assertNull(formulaFor(setWorkbookProtection, exception));
    assertNull(formulaFor(appendRow, exception));
    assertNull(formulaFor(ensureSheet, exception));
    assertNull(sheetNameFor(clearWorkbookProtection, exception));
    assertNull(sheetNameFor(setWorkbookProtection, exception));
    assertNull(addressFor(clearWorkbookProtection, exception));
    assertNull(addressFor(setWorkbookProtection, exception));
    assertNull(rangeFor(clearWorkbookProtection, exception));
    assertNull(rangeFor(setWorkbookProtection, exception));
  }

  @Test
  void extractsNullFormulaFromSetCellWithNonFormulaValueWhenExceptionCarriesNone() {
    RuntimeException exception = new RuntimeException("test");
    ExecutorTestPlanSupport.PendingMutation operation =
        mutate(
            new CellSelector.ByAddress("Budget", "A1"),
            new MutationAction.SetCell(textCell("hello")));

    assertNull(formulaFor(operation, exception));
    assertEquals("Budget", sheetNameFor(operation, exception));
    assertEquals("A1", addressFor(operation, exception));
    assertNull(rangeFor(operation, exception));
  }

  @Test
  void formulaForSetCellReturnsNullForAllNonFormulaValueTypes() {
    RuntimeException exception = new RuntimeException("test");

    assertNull(
        formulaFor(
            mutate(
                new CellSelector.ByAddress("S", "A1"),
                new MutationAction.SetCell(new CellInput.Blank())),
            exception));
    assertNull(
        formulaFor(
            mutate(
                new CellSelector.ByAddress("S", "A1"),
                new MutationAction.SetCell(textCell("hello"))),
            exception));
    assertNull(
        formulaFor(
            mutate(
                new CellSelector.ByAddress("S", "A1"),
                new MutationAction.SetCell(
                    new CellInput.RichText(
                        List.of(
                            richTextRun("Budget"),
                            new RichTextRunInput(
                                text(" FY26"),
                                new CellFontInput(
                                    true, false, null, null, "#AABBCC", false, false)))))),
            exception));
    assertNull(
        formulaFor(
            mutate(
                new CellSelector.ByAddress("S", "A1"),
                new MutationAction.SetCell(new CellInput.Numeric(1.0))),
            exception));
    assertNull(
        formulaFor(
            mutate(
                new CellSelector.ByAddress("S", "A1"),
                new MutationAction.SetCell(new CellInput.BooleanValue(true))),
            exception));
    assertNull(
        formulaFor(
            mutate(
                new CellSelector.ByAddress("S", "A1"),
                new MutationAction.SetCell(new CellInput.Date(LocalDate.of(2024, 1, 1)))),
            exception));
    assertNull(
        formulaFor(
            mutate(
                new CellSelector.ByAddress("S", "A1"),
                new MutationAction.SetCell(
                    new CellInput.DateTime(LocalDateTime.of(2024, 1, 1, 0, 0)))),
            exception));
  }

  @Test
  void extractsSheetOnlyContextForDeleteSheetOperations() {
    RuntimeException exception = new RuntimeException("test");
    ExecutorTestPlanSupport.PendingMutation ensureSheet =
        mutate(new SheetSelector.ByName("Archive"), new MutationAction.EnsureSheet());
    ExecutorTestPlanSupport.PendingMutation deleteSheet =
        mutate(new SheetSelector.ByName("Archive"), new MutationAction.DeleteSheet());

    assertNull(formulaFor(ensureSheet, exception));
    assertNull(formulaFor(deleteSheet, exception));
    assertEquals("Archive", sheetNameFor(ensureSheet, exception));
    assertEquals("Archive", sheetNameFor(deleteSheet, exception));
    assertNull(addressFor(ensureSheet, exception));
    assertNull(addressFor(deleteSheet, exception));
    assertNull(rangeFor(ensureSheet, exception));
    assertNull(rangeFor(deleteSheet, exception));
  }

  @Test
  void extractsSheetStateContextForB1Operations() {
    RuntimeException exception = new RuntimeException("test");
    ExecutorTestPlanSupport.PendingMutation copySheet =
        mutate(
            new SheetSelector.ByName("Budget"),
            new MutationAction.CopySheet("Budget Copy", new SheetCopyPosition.AppendAtEnd()));
    ExecutorTestPlanSupport.PendingMutation setActiveSheet =
        mutate(new SheetSelector.ByName("Budget Copy"), new MutationAction.SetActiveSheet());
    ExecutorTestPlanSupport.PendingMutation setSelectedSheets =
        mutate(
            new SheetSelector.ByNames(List.of("Budget", "Budget Copy")),
            new MutationAction.SetSelectedSheets());
    ExecutorTestPlanSupport.PendingMutation setSheetVisibility =
        mutate(
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetSheetVisibility(ExcelSheetVisibility.HIDDEN));
    ExecutorTestPlanSupport.PendingMutation setSheetProtection =
        mutate(
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetSheetProtection(protectionSettings()));
    ExecutorTestPlanSupport.PendingMutation clearSheetProtection =
        mutate(new SheetSelector.ByName("Budget"), new MutationAction.ClearSheetProtection());

    assertNull(formulaFor(copySheet, exception));
    assertNull(formulaFor(setActiveSheet, exception));
    assertNull(formulaFor(setSelectedSheets, exception));
    assertNull(formulaFor(setSheetVisibility, exception));
    assertNull(formulaFor(setSheetProtection, exception));
    assertNull(formulaFor(clearSheetProtection, exception));

    assertEquals("Budget", sheetNameFor(copySheet, exception));
    assertEquals("Budget Copy", sheetNameFor(setActiveSheet, exception));
    assertNull(sheetNameFor(setSelectedSheets, exception));
    assertEquals("Budget", sheetNameFor(setSheetVisibility, exception));
    assertEquals("Budget", sheetNameFor(setSheetProtection, exception));
    assertEquals("Budget", sheetNameFor(clearSheetProtection, exception));

    assertNull(addressFor(copySheet, exception));
    assertNull(addressFor(setActiveSheet, exception));
    assertNull(addressFor(setSelectedSheets, exception));
    assertNull(addressFor(setSheetVisibility, exception));
    assertNull(addressFor(setSheetProtection, exception));
    assertNull(addressFor(clearSheetProtection, exception));

    assertNull(rangeFor(copySheet, exception));
    assertNull(rangeFor(setActiveSheet, exception));
    assertNull(rangeFor(setSelectedSheets, exception));
    assertNull(rangeFor(setSheetVisibility, exception));
    assertNull(rangeFor(setSheetProtection, exception));
    assertNull(rangeFor(clearSheetProtection, exception));

    assertNull(namedRangeNameFor(copySheet, exception));
    assertNull(namedRangeNameFor(setActiveSheet, exception));
    assertNull(namedRangeNameFor(setSelectedSheets, exception));
    assertNull(namedRangeNameFor(setSheetVisibility, exception));
    assertNull(namedRangeNameFor(setSheetProtection, exception));
    assertNull(namedRangeNameFor(clearSheetProtection, exception));
  }

  @Test
  @SuppressWarnings("PMD.NcssCount")
  void extractsContextForStructuralLayoutOperations() {
    RuntimeException exception = new RuntimeException("test");
    ExecutorTestPlanSupport.PendingMutation mergeCells =
        mutate(new RangeSelector.ByRange("Budget", "A1:B2"), new MutationAction.MergeCells());
    ExecutorTestPlanSupport.PendingMutation unmergeCells =
        mutate(new RangeSelector.ByRange("Budget", "A1:B2"), new MutationAction.UnmergeCells());
    ExecutorTestPlanSupport.PendingMutation setColumnWidth =
        mutate(
            new ColumnBandSelector.Span("Budget", 0, 1), new MutationAction.SetColumnWidth(16.0));
    ExecutorTestPlanSupport.PendingMutation setRowHeight =
        mutate(new RowBandSelector.Span("Budget", 0, 1), new MutationAction.SetRowHeight(28.5));
    ExecutorTestPlanSupport.PendingMutation insertRows =
        mutate(new RowBandSelector.Insertion("Budget", 1, 2), new MutationAction.InsertRows());
    ExecutorTestPlanSupport.PendingMutation deleteRows =
        mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.DeleteRows());
    ExecutorTestPlanSupport.PendingMutation shiftRows =
        mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.ShiftRows(1));
    ExecutorTestPlanSupport.PendingMutation insertColumns =
        mutate(
            new ColumnBandSelector.Insertion("Budget", 1, 2), new MutationAction.InsertColumns());
    ExecutorTestPlanSupport.PendingMutation deleteColumns =
        mutate(new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.DeleteColumns());
    ExecutorTestPlanSupport.PendingMutation shiftColumns =
        mutate(new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.ShiftColumns(-1));
    ExecutorTestPlanSupport.PendingMutation setRowVisibility =
        mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.SetRowVisibility(true));
    ExecutorTestPlanSupport.PendingMutation setColumnVisibility =
        mutate(
            new ColumnBandSelector.Span("Budget", 1, 2),
            new MutationAction.SetColumnVisibility(false));
    ExecutorTestPlanSupport.PendingMutation groupRows =
        mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.GroupRows(true));
    ExecutorTestPlanSupport.PendingMutation ungroupRows =
        mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.UngroupRows());
    ExecutorTestPlanSupport.PendingMutation groupColumns =
        mutate(new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.GroupColumns(true));
    ExecutorTestPlanSupport.PendingMutation ungroupColumns =
        mutate(new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.UngroupColumns());
    ExecutorTestPlanSupport.PendingMutation setSheetPane =
        mutate(
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetSheetPane(new PaneInput.Frozen(1, 1, 1, 1)));
    ExecutorTestPlanSupport.PendingMutation setSheetZoom =
        mutate(new SheetSelector.ByName("Budget"), new MutationAction.SetSheetZoom(125));
    ExecutorTestPlanSupport.PendingMutation setSheetPresentation =
        mutate(
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetSheetPresentation(
                new SheetPresentationInput(
                    new SheetDisplayInput(false, false, false, true, true),
                    new ColorInput("#112233"),
                    new SheetOutlineSummaryInput(false, false),
                    new SheetDefaultsInput(11, 18.5d),
                    List.of(
                        new IgnoredErrorInput(
                            "A1:B2", List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))))));
    ExecutorTestPlanSupport.PendingMutation setPrintLayout =
        mutate(
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetPrintLayout(
                new PrintLayoutInput(
                    new PrintAreaInput.Range("A1:B12"),
                    ExcelPrintOrientation.LANDSCAPE,
                    new PrintScalingInput.Fit(1, 0),
                    new PrintTitleRowsInput.Band(0, 0),
                    new PrintTitleColumnsInput.Band(0, 0),
                    headerFooter("Budget", "", ""),
                    headerFooter("", "Page &P", ""))));
    ExecutorTestPlanSupport.PendingMutation clearPrintLayout =
        mutate(new SheetSelector.ByName("Budget"), new MutationAction.ClearPrintLayout());

    assertNull(formulaFor(mergeCells, exception));
    assertEquals("Budget", sheetNameFor(mergeCells, exception));
    assertNull(addressFor(mergeCells, exception));
    assertEquals("A1:B2", rangeFor(mergeCells, exception));

    assertNull(formulaFor(unmergeCells, exception));
    assertEquals("Budget", sheetNameFor(unmergeCells, exception));
    assertNull(addressFor(unmergeCells, exception));
    assertEquals("A1:B2", rangeFor(unmergeCells, exception));

    assertNull(formulaFor(setColumnWidth, exception));
    assertEquals("Budget", sheetNameFor(setColumnWidth, exception));
    assertNull(addressFor(setColumnWidth, exception));
    assertNull(rangeFor(setColumnWidth, exception));

    assertNull(formulaFor(setRowHeight, exception));
    assertEquals("Budget", sheetNameFor(setRowHeight, exception));
    assertNull(addressFor(setRowHeight, exception));
    assertNull(rangeFor(setRowHeight, exception));

    assertNull(formulaFor(insertRows, exception));
    assertEquals("Budget", sheetNameFor(insertRows, exception));
    assertNull(addressFor(insertRows, exception));
    assertNull(rangeFor(insertRows, exception));

    assertNull(formulaFor(deleteRows, exception));
    assertEquals("Budget", sheetNameFor(deleteRows, exception));
    assertNull(addressFor(deleteRows, exception));
    assertNull(rangeFor(deleteRows, exception));

    assertNull(formulaFor(shiftRows, exception));
    assertEquals("Budget", sheetNameFor(shiftRows, exception));
    assertNull(addressFor(shiftRows, exception));
    assertNull(rangeFor(shiftRows, exception));

    assertNull(formulaFor(insertColumns, exception));
    assertEquals("Budget", sheetNameFor(insertColumns, exception));
    assertNull(addressFor(insertColumns, exception));
    assertNull(rangeFor(insertColumns, exception));

    assertNull(formulaFor(deleteColumns, exception));
    assertEquals("Budget", sheetNameFor(deleteColumns, exception));
    assertNull(addressFor(deleteColumns, exception));
    assertNull(rangeFor(deleteColumns, exception));

    assertNull(formulaFor(shiftColumns, exception));
    assertEquals("Budget", sheetNameFor(shiftColumns, exception));
    assertNull(addressFor(shiftColumns, exception));
    assertNull(rangeFor(shiftColumns, exception));

    assertNull(formulaFor(setRowVisibility, exception));
    assertEquals("Budget", sheetNameFor(setRowVisibility, exception));
    assertNull(addressFor(setRowVisibility, exception));
    assertNull(rangeFor(setRowVisibility, exception));

    assertNull(formulaFor(setColumnVisibility, exception));
    assertEquals("Budget", sheetNameFor(setColumnVisibility, exception));
    assertNull(addressFor(setColumnVisibility, exception));
    assertNull(rangeFor(setColumnVisibility, exception));

    assertNull(formulaFor(groupRows, exception));
    assertEquals("Budget", sheetNameFor(groupRows, exception));
    assertNull(addressFor(groupRows, exception));
    assertNull(rangeFor(groupRows, exception));

    assertNull(formulaFor(ungroupRows, exception));
    assertEquals("Budget", sheetNameFor(ungroupRows, exception));
    assertNull(addressFor(ungroupRows, exception));
    assertNull(rangeFor(ungroupRows, exception));

    assertNull(formulaFor(groupColumns, exception));
    assertEquals("Budget", sheetNameFor(groupColumns, exception));
    assertNull(addressFor(groupColumns, exception));
    assertNull(rangeFor(groupColumns, exception));

    assertNull(formulaFor(ungroupColumns, exception));
    assertEquals("Budget", sheetNameFor(ungroupColumns, exception));
    assertNull(addressFor(ungroupColumns, exception));
    assertNull(rangeFor(ungroupColumns, exception));

    assertNull(formulaFor(setSheetPane, exception));
    assertEquals("Budget", sheetNameFor(setSheetPane, exception));
    assertNull(addressFor(setSheetPane, exception));
    assertNull(rangeFor(setSheetPane, exception));

    assertNull(formulaFor(setSheetZoom, exception));
    assertEquals("Budget", sheetNameFor(setSheetZoom, exception));
    assertNull(addressFor(setSheetZoom, exception));
    assertNull(rangeFor(setSheetZoom, exception));

    assertNull(formulaFor(setSheetPresentation, exception));
    assertEquals("Budget", sheetNameFor(setSheetPresentation, exception));
    assertNull(addressFor(setSheetPresentation, exception));
    assertNull(rangeFor(setSheetPresentation, exception));

    assertNull(formulaFor(setPrintLayout, exception));
    assertEquals("Budget", sheetNameFor(setPrintLayout, exception));
    assertNull(addressFor(setPrintLayout, exception));
    assertNull(rangeFor(setPrintLayout, exception));

    assertNull(formulaFor(clearPrintLayout, exception));
    assertEquals("Budget", sheetNameFor(clearPrintLayout, exception));
    assertNull(addressFor(clearPrintLayout, exception));
    assertNull(rangeFor(clearPrintLayout, exception));
  }

  @Test
  void convertsWaveThreeOperationsIntoWorkbookCommands() {
    WorkbookCommand setHyperlink =
        command(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new MutationAction.SetHyperlink(
                    new HyperlinkTarget.Url("https://example.com/report"))));
    WorkbookCommand clearHyperlink =
        command(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"), new MutationAction.ClearHyperlink()));
    WorkbookCommand setComment =
        command(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new MutationAction.SetComment(
                    new CommentInput(text("Review"), "GridGrind", false))));
    WorkbookCommand clearComment =
        command(
            mutate(new CellSelector.ByAddress("Budget", "A1"), new MutationAction.ClearComment()));
    WorkbookCommand setNamedRange =
        command(
            mutate(
                new MutationAction.SetNamedRange(
                    "BudgetTotal",
                    new NamedRangeScope.Workbook(),
                    new NamedRangeTarget("Budget", "B4"))));
    WorkbookCommand deleteNamedRange =
        command(
            mutate(
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope(
                    "BudgetTotal", "Budget"),
                new MutationAction.DeleteNamedRange()));

    assertInstanceOf(WorkbookCommand.SetHyperlink.class, setHyperlink);
    assertInstanceOf(WorkbookCommand.ClearHyperlink.class, clearHyperlink);
    assertInstanceOf(WorkbookCommand.SetComment.class, setComment);
    assertInstanceOf(WorkbookCommand.ClearComment.class, clearComment);
    assertInstanceOf(WorkbookCommand.SetNamedRange.class, setNamedRange);
    assertInstanceOf(WorkbookCommand.DeleteNamedRange.class, deleteNamedRange);
    assertEquals(
        new ExcelHyperlink.Url("https://example.com/report"),
        cast(WorkbookCommand.SetHyperlink.class, setHyperlink).target());
    assertEquals(
        new ExcelComment("Review", "GridGrind", false),
        cast(WorkbookCommand.SetComment.class, setComment).comment());
    assertEquals(
        new ExcelNamedRangeDefinition(
            "BudgetTotal",
            new ExcelNamedRangeScope.WorkbookScope(),
            new ExcelNamedRangeTarget("Budget", "B4")),
        cast(WorkbookCommand.SetNamedRange.class, setNamedRange).definition());
    assertEquals(
        new ExcelNamedRangeScope.SheetScope("Budget"),
        cast(WorkbookCommand.DeleteNamedRange.class, deleteNamedRange).scope());
  }

  @Test
  void convertsRemainingWorkbookOperationsIntoWorkbookCommands() {
    assertInstanceOf(
        WorkbookCommand.CreateSheet.class,
        command(mutate(new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet())));
    assertInstanceOf(
        WorkbookCommand.RenameSheet.class,
        command(
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.RenameSheet("Summary"))));
    assertInstanceOf(
        WorkbookCommand.DeleteSheet.class,
        command(mutate(new SheetSelector.ByName("Budget"), new MutationAction.DeleteSheet())));
    assertInstanceOf(
        WorkbookCommand.MoveSheet.class,
        command(mutate(new SheetSelector.ByName("Budget"), new MutationAction.MoveSheet(0))));
    assertInstanceOf(
        WorkbookCommand.MergeCells.class,
        command(
            mutate(new RangeSelector.ByRange("Budget", "A1:B2"), new MutationAction.MergeCells())));
    assertInstanceOf(
        WorkbookCommand.UnmergeCells.class,
        command(
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B2"), new MutationAction.UnmergeCells())));
    assertInstanceOf(
        WorkbookCommand.SetColumnWidth.class,
        command(
            mutate(
                new ColumnBandSelector.Span("Budget", 0, 1),
                new MutationAction.SetColumnWidth(16.0))));
    assertInstanceOf(
        WorkbookCommand.SetRowHeight.class,
        command(
            mutate(
                new RowBandSelector.Span("Budget", 0, 1), new MutationAction.SetRowHeight(28.5))));
    assertInstanceOf(
        WorkbookCommand.InsertRows.class,
        command(
            mutate(
                new RowBandSelector.Insertion("Budget", 1, 2), new MutationAction.InsertRows())));
    assertInstanceOf(
        WorkbookCommand.DeleteRows.class,
        command(mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.DeleteRows())));
    assertInstanceOf(
        WorkbookCommand.ShiftRows.class,
        command(mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.ShiftRows(1))));
    assertInstanceOf(
        WorkbookCommand.InsertColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Insertion("Budget", 1, 2),
                new MutationAction.InsertColumns())));
    assertInstanceOf(
        WorkbookCommand.DeleteColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.DeleteColumns())));
    assertInstanceOf(
        WorkbookCommand.ShiftColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.ShiftColumns(-1))));
    assertInstanceOf(
        WorkbookCommand.SetRowVisibility.class,
        command(
            mutate(
                new RowBandSelector.Span("Budget", 1, 2),
                new MutationAction.SetRowVisibility(true))));
    assertInstanceOf(
        WorkbookCommand.SetColumnVisibility.class,
        command(
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2),
                new MutationAction.SetColumnVisibility(true))));
    assertInstanceOf(
        WorkbookCommand.GroupRows.class,
        command(
            mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.GroupRows(false))));
    assertInstanceOf(
        WorkbookCommand.UngroupRows.class,
        command(
            mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.UngroupRows())));
    assertInstanceOf(
        WorkbookCommand.GroupColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2),
                new MutationAction.GroupColumns(false))));
    assertInstanceOf(
        WorkbookCommand.UngroupColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.UngroupColumns())));
    assertInstanceOf(
        WorkbookCommand.SetSheetPane.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetSheetPane(new PaneInput.Frozen(1, 1, 1, 1)))));
    assertInstanceOf(
        WorkbookCommand.SetSheetZoom.class,
        command(mutate(new SheetSelector.ByName("Budget"), new MutationAction.SetSheetZoom(125))));
    assertInstanceOf(
        WorkbookCommand.SetSheetPresentation.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetSheetPresentation(
                    new SheetPresentationInput(
                        new SheetDisplayInput(false, false, false, true, true),
                        new ColorInput("#112233"),
                        new SheetOutlineSummaryInput(false, false),
                        new SheetDefaultsInput(11, 18.5d),
                        List.of(
                            new IgnoredErrorInput(
                                "A1:B2",
                                List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))))))));
    assertInstanceOf(
        WorkbookCommand.SetPrintLayout.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetPrintLayout(
                    new PrintLayoutInput(
                        new PrintAreaInput.Range("A1:B12"),
                        ExcelPrintOrientation.LANDSCAPE,
                        new PrintScalingInput.Fit(1, 0),
                        new PrintTitleRowsInput.Band(0, 0),
                        new PrintTitleColumnsInput.Band(0, 0),
                        headerFooter("Budget", "", ""),
                        headerFooter("", "Page &P", ""))))));
    assertInstanceOf(
        WorkbookCommand.ClearPrintLayout.class,
        command(mutate(new SheetSelector.ByName("Budget"), new MutationAction.ClearPrintLayout())));
    assertInstanceOf(
        WorkbookCommand.SetCell.class,
        command(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new MutationAction.SetCell(textCell("x")))));
    assertInstanceOf(
        WorkbookCommand.SetRange.class,
        command(
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B1"),
                new MutationAction.SetRange(List.of(List.of(textCell("x")))))));
    assertInstanceOf(
        WorkbookCommand.ClearRange.class,
        command(
            mutate(new RangeSelector.ByRange("Budget", "A1:B1"), new MutationAction.ClearRange())));
    assertInstanceOf(
        WorkbookCommand.SetConditionalFormatting.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetConditionalFormatting(
                    new ConditionalFormattingBlockInput(
                        List.of("A1:A3"),
                        List.of(
                            new ConditionalFormattingRuleInput.FormulaRule(
                                "A1>0",
                                false,
                                new DifferentialStyleInput(
                                    null, true, null, null, null, null, null, null, null))))))));
    assertInstanceOf(
        WorkbookCommand.ClearConditionalFormatting.class,
        command(
            mutate(
                new RangeSelector.AllOnSheet("Budget"),
                new MutationAction.ClearConditionalFormatting())));
    assertInstanceOf(
        WorkbookCommand.SetAutofilter.class,
        command(
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B4"), new MutationAction.SetAutofilter())));
    assertInstanceOf(
        WorkbookCommand.ClearAutofilter.class,
        command(mutate(new SheetSelector.ByName("Budget"), new MutationAction.ClearAutofilter())));
    assertInstanceOf(
        WorkbookCommand.SetTable.class,
        command(
            mutate(
                new MutationAction.SetTable(
                    new TableInput(
                        "BudgetTable", "Budget", "A1:B4", false, new TableStyleInput.None())))));
    assertInstanceOf(
        WorkbookCommand.DeleteTable.class,
        command(
            mutate(
                new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
                new MutationAction.DeleteTable())));
    assertInstanceOf(
        WorkbookCommand.AppendRow.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.AppendRow(List.of(textCell("x"))))));
    assertInstanceOf(
        WorkbookCommand.AutoSizeColumns.class,
        command(mutate(new SheetSelector.ByName("Budget"), new MutationAction.AutoSizeColumns())));

    WorkbookCommand.SetSheetPane setSheetPaneNone =
        cast(
            WorkbookCommand.SetSheetPane.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetSheetPane(new PaneInput.None()))));
    WorkbookCommand.SetSheetPane setSheetPaneSplit =
        cast(
            WorkbookCommand.SetSheetPane.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetSheetPane(
                        new PaneInput.Split(1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT)))));
    WorkbookCommand.SetPrintLayout defaultPrintLayout =
        cast(
            WorkbookCommand.SetPrintLayout.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetPrintLayout(
                        new PrintLayoutInput(null, null, null, null, null, null, null)))));
    WorkbookCommand.SetSheetPresentation defaultSheetPresentation =
        cast(
            WorkbookCommand.SetSheetPresentation.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetSheetPresentation(
                        new SheetPresentationInput(null, null, null, null, null)))));

    assertEquals(new ExcelSheetPane.None(), setSheetPaneNone.pane());
    assertEquals(
        new ExcelSheetPane.Split(1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT),
        setSheetPaneSplit.pane());
    assertEquals(ExcelSheetPresentation.defaults(), defaultSheetPresentation.presentation());
    assertEquals(
        new dev.erst.gridgrind.excel.ExcelPrintLayout(
            new dev.erst.gridgrind.excel.ExcelPrintLayout.Area.None(),
            ExcelPrintOrientation.PORTRAIT,
            new dev.erst.gridgrind.excel.ExcelPrintLayout.Scaling.Automatic(),
            new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleRows.None(),
            new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleColumns.None(),
            new dev.erst.gridgrind.excel.ExcelHeaderFooterText("", "", ""),
            new dev.erst.gridgrind.excel.ExcelHeaderFooterText("", "", "")),
        defaultPrintLayout.printLayout());
  }

  @Test
  void convertsStructuralAndSurfaceReadOperationsIntoWorkbookReadCommands() {
    WorkbookReadCommand workbookSummary =
        readCommand(
            inspect(
                "workbook",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetWorkbookSummary()));
    WorkbookReadCommand workbookProtection =
        readCommand(
            inspect(
                "workbook-protection",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetWorkbookProtection()));
    WorkbookReadCommand namedRanges =
        readCommand(
            inspect(
                "ranges",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf(
                    List.of(
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.ByName(
                            "BudgetTotal"),
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope(
                            "LocalItem", "Budget"))),
                new InspectionQuery.GetNamedRanges()));
    WorkbookReadCommand sheetSummary =
        readCommand(
            inspect(
                "sheet",
                new SheetSelector.ByName("Budget"),
                new InspectionQuery.GetSheetSummary()));
    WorkbookReadCommand cells =
        readCommand(
            inspect(
                "cells",
                new CellSelector.ByAddresses("Budget", List.of("A1")),
                new InspectionQuery.GetCells()));
    WorkbookReadCommand window =
        readCommand(
            inspect(
                "window",
                new RangeSelector.RectangularWindow("Budget", "A1", 2, 2),
                new InspectionQuery.GetWindow()));
    WorkbookReadCommand merged =
        readCommand(
            inspect(
                "merged",
                new SheetSelector.ByName("Budget"),
                new InspectionQuery.GetMergedRegions()));
    WorkbookReadCommand hyperlinks =
        readCommand(
            inspect(
                "hyperlinks",
                new CellSelector.AllUsedInSheet("Budget"),
                new InspectionQuery.GetHyperlinks()));
    WorkbookReadCommand comments =
        readCommand(
            inspect(
                "comments",
                new CellSelector.ByAddresses("Budget", List.of("A1")),
                new InspectionQuery.GetComments()));
    WorkbookReadCommand layout =
        readCommand(
            inspect(
                "layout",
                new SheetSelector.ByName("Budget"),
                new InspectionQuery.GetSheetLayout()));
    WorkbookReadCommand printLayout =
        readCommand(
            inspect(
                "printLayout",
                new SheetSelector.ByName("Budget"),
                new InspectionQuery.GetPrintLayout()));
    WorkbookReadCommand validations =
        readCommand(
            inspect(
                "validations",
                new RangeSelector.AllOnSheet("Budget"),
                new InspectionQuery.GetDataValidations()));
    WorkbookReadCommand conditionalFormatting =
        readCommand(
            inspect(
                "conditionalFormatting",
                new RangeSelector.ByRanges("Budget", List.of("A1:A3")),
                new InspectionQuery.GetConditionalFormatting()));
    WorkbookReadCommand autofilters =
        readCommand(
            inspect(
                "autofilters",
                new SheetSelector.ByName("Budget"),
                new InspectionQuery.GetAutofilters()));
    WorkbookReadCommand tables =
        readCommand(
            inspect(
                "tables",
                new TableSelector.ByNames(List.of("BudgetTable")),
                new InspectionQuery.GetTables()));
    WorkbookReadCommand formulaSurface =
        readCommand(
            inspect(
                "formula",
                new SheetSelector.ByNames(List.of("Budget")),
                new InspectionQuery.GetFormulaSurface()));
    WorkbookReadCommand schema =
        readCommand(
            inspect(
                "schema",
                new RangeSelector.RectangularWindow("Budget", "A1", 3, 2),
                new InspectionQuery.GetSheetSchema()));
    WorkbookReadCommand namedRangeSurface =
        readCommand(
            inspect(
                "surface",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                new InspectionQuery.GetNamedRangeSurface()));

    assertInstanceOf(WorkbookReadCommand.GetWorkbookSummary.class, workbookSummary);
    assertInstanceOf(WorkbookReadCommand.GetWorkbookProtection.class, workbookProtection);
    assertInstanceOf(WorkbookReadCommand.GetNamedRanges.class, namedRanges);
    assertInstanceOf(WorkbookReadCommand.GetSheetSummary.class, sheetSummary);
    assertInstanceOf(WorkbookReadCommand.GetCells.class, cells);
    assertInstanceOf(WorkbookReadCommand.GetWindow.class, window);
    assertInstanceOf(WorkbookReadCommand.GetMergedRegions.class, merged);
    assertInstanceOf(WorkbookReadCommand.GetHyperlinks.class, hyperlinks);
    assertInstanceOf(WorkbookReadCommand.GetComments.class, comments);
    assertInstanceOf(WorkbookReadCommand.GetSheetLayout.class, layout);
    assertInstanceOf(WorkbookReadCommand.GetPrintLayout.class, printLayout);
    assertInstanceOf(WorkbookReadCommand.GetDataValidations.class, validations);
    assertInstanceOf(WorkbookReadCommand.GetConditionalFormatting.class, conditionalFormatting);
    assertInstanceOf(WorkbookReadCommand.GetAutofilters.class, autofilters);
    assertInstanceOf(WorkbookReadCommand.GetTables.class, tables);
    assertInstanceOf(WorkbookReadCommand.GetFormulaSurface.class, formulaSurface);
    assertInstanceOf(WorkbookReadCommand.GetSheetSchema.class, schema);
    assertInstanceOf(WorkbookReadCommand.GetNamedRangeSurface.class, namedRangeSurface);
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelNamedRangeSelection.Selected.class,
        cast(WorkbookReadCommand.GetNamedRanges.class, namedRanges).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelCellSelection.AllUsedCells.class,
        cast(WorkbookReadCommand.GetHyperlinks.class, hyperlinks).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelCellSelection.Selected.class,
        cast(WorkbookReadCommand.GetComments.class, comments).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelRangeSelection.All.class,
        cast(WorkbookReadCommand.GetDataValidations.class, validations).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelRangeSelection.Selected.class,
        cast(WorkbookReadCommand.GetConditionalFormatting.class, conditionalFormatting)
            .selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelTableSelection.ByNames.class,
        cast(WorkbookReadCommand.GetTables.class, tables).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelSheetSelection.Selected.class,
        cast(WorkbookReadCommand.GetFormulaSurface.class, formulaSurface).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelNamedRangeSelection.All.class,
        cast(WorkbookReadCommand.GetNamedRangeSurface.class, namedRangeSurface).selection());
  }

  @Test
  void convertsAnalysisReadOperationsIntoWorkbookReadCommands() {
    WorkbookReadCommand formulaHealth =
        readCommand(
            inspect(
                "formulaHealth",
                new SheetSelector.All(),
                new InspectionQuery.AnalyzeFormulaHealth()));
    WorkbookReadCommand validationHealth =
        readCommand(
            inspect(
                "validationHealth",
                new SheetSelector.All(),
                new InspectionQuery.AnalyzeDataValidationHealth()));
    WorkbookReadCommand conditionalFormattingHealth =
        readCommand(
            inspect(
                "conditionalFormattingHealth",
                new SheetSelector.ByNames(List.of("Budget")),
                new InspectionQuery.AnalyzeConditionalFormattingHealth()));
    WorkbookReadCommand autofilterHealth =
        readCommand(
            inspect(
                "autofilterHealth",
                new SheetSelector.ByNames(List.of("Budget")),
                new InspectionQuery.AnalyzeAutofilterHealth()));
    WorkbookReadCommand tableHealth =
        readCommand(
            inspect(
                "tableHealth",
                new TableSelector.ByNames(List.of("BudgetTable")),
                new InspectionQuery.AnalyzeTableHealth()));
    WorkbookReadCommand hyperlinkHealth =
        readCommand(
            inspect(
                "hyperlinkHealth",
                new SheetSelector.All(),
                new InspectionQuery.AnalyzeHyperlinkHealth()));
    WorkbookReadCommand namedRangeHealth =
        readCommand(
            inspect(
                "namedRangeHealth",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                new InspectionQuery.AnalyzeNamedRangeHealth()));
    WorkbookReadCommand workbookFindings =
        readCommand(
            inspect(
                "workbookFindings",
                new WorkbookSelector.Current(),
                new InspectionQuery.AnalyzeWorkbookFindings()));

    assertInstanceOf(WorkbookReadCommand.AnalyzeFormulaHealth.class, formulaHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeDataValidationHealth.class, validationHealth);
    assertInstanceOf(
        WorkbookReadCommand.AnalyzeConditionalFormattingHealth.class, conditionalFormattingHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeAutofilterHealth.class, autofilterHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeTableHealth.class, tableHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeHyperlinkHealth.class, hyperlinkHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeNamedRangeHealth.class, namedRangeHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeWorkbookFindings.class, workbookFindings);
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelSheetSelection.Selected.class,
        cast(
                WorkbookReadCommand.AnalyzeConditionalFormattingHealth.class,
                conditionalFormattingHealth)
            .selection());
  }

  @Test
  void convertsSheetStateOperationsIntoWorkbookCommandsWithExactFields() {
    WorkbookCommand.CopySheet copySheet =
        cast(
            WorkbookCommand.CopySheet.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.CopySheet(
                        "Budget Copy", new SheetCopyPosition.AtIndex(1)))));
    WorkbookCommand.SetActiveSheet setActiveSheet =
        cast(
            WorkbookCommand.SetActiveSheet.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget Copy"), new MutationAction.SetActiveSheet())));
    WorkbookCommand.SetSelectedSheets setSelectedSheets =
        cast(
            WorkbookCommand.SetSelectedSheets.class,
            command(
                mutate(
                    new SheetSelector.ByNames(List.of("Budget", "Budget Copy")),
                    new MutationAction.SetSelectedSheets())));
    WorkbookCommand.SetSheetVisibility setSheetVisibility =
        cast(
            WorkbookCommand.SetSheetVisibility.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetSheetVisibility(ExcelSheetVisibility.HIDDEN))));
    WorkbookCommand.SetSheetProtection setSheetProtection =
        cast(
            WorkbookCommand.SetSheetProtection.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetSheetProtection(protectionSettings()))));
    WorkbookCommand.ClearSheetProtection clearSheetProtection =
        cast(
            WorkbookCommand.ClearSheetProtection.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.ClearSheetProtection())));

    ExcelSheetCopyPosition.AtIndex position =
        assertInstanceOf(ExcelSheetCopyPosition.AtIndex.class, copySheet.position());

    assertEquals("Budget", copySheet.sourceSheetName());
    assertEquals("Budget Copy", copySheet.newSheetName());
    assertEquals(1, position.targetIndex());
    assertEquals("Budget Copy", setActiveSheet.sheetName());
    assertEquals(List.of("Budget", "Budget Copy"), setSelectedSheets.sheetNames());
    assertEquals(ExcelSheetVisibility.HIDDEN, setSheetVisibility.visibility());
    assertEquals(excelProtectionSettings(), setSheetProtection.protection());
    assertEquals("Budget", clearSheetProtection.sheetName());
  }

  @Test
  void convertsReadResultsIntoProtocolReadResults() {
    ExcelCellSnapshot.BlankSnapshot blank =
        new ExcelCellSnapshot.BlankSnapshot(
            "A1", "BLANK", "", defaultStyle(), ExcelCellMetadataSnapshot.empty());

    InspectionResult workbookSummary =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult(
                "workbook",
                new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary.WithSheets(
                    1, List.of("Budget"), "Budget", List.of("Budget"), 1, true)));
    InspectionResult namedRanges =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangesResult(
                "ranges",
                List.of(
                    new ExcelNamedRangeSnapshot.RangeSnapshot(
                        "BudgetTotal",
                        new ExcelNamedRangeScope.WorkbookScope(),
                        "Budget!$B$4",
                        new ExcelNamedRangeTarget("Budget", "B4")))));
    InspectionResult sheetSummary =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummaryResult(
                "sheet",
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummary(
                    "Budget",
                    dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility.VISIBLE,
                    new dev.erst.gridgrind.excel.WorkbookReadResult.SheetProtection.Unprotected(),
                    4,
                    3,
                    2)));
    InspectionResult cells =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.CellsResult(
                "cells", "Budget", List.of(blank)));
    InspectionResult window =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.WindowResult(
                "window",
                new dev.erst.gridgrind.excel.WorkbookReadResult.Window(
                    "Budget",
                    "A1",
                    1,
                    1,
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.WindowRow(
                            0, List.of(blank))))));
    InspectionResult merged =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.MergedRegionsResult(
                "merged",
                "Budget",
                List.of(new dev.erst.gridgrind.excel.WorkbookReadResult.MergedRegion("A1:B2"))));
    InspectionResult hyperlinks =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.HyperlinksResult(
                "hyperlinks",
                "Budget",
                List.of(
                    new dev.erst.gridgrind.excel.WorkbookReadResult.CellHyperlink(
                        "A1", new ExcelHyperlink.Url("https://example.com/report")))));
    InspectionResult comments =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.CommentsResult(
                "comments",
                "Budget",
                List.of(
                    new dev.erst.gridgrind.excel.WorkbookReadResult.CellComment(
                        "A1",
                        new ExcelCommentSnapshot("Review", "GridGrind", false, null, null)))));
    InspectionResult layout =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayoutResult(
                "layout",
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayout(
                    "Budget",
                    new dev.erst.gridgrind.excel.ExcelSheetPane.Frozen(1, 1, 1, 1),
                    125,
                    defaultSheetPresentationSnapshot(),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.ColumnLayout(
                            0, 12.5, false, 0, false)),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.RowLayout(
                            0, 18.0, false, 0, false)))));
    InspectionResult printLayout =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.PrintLayoutResult(
                "printLayout",
                "Budget",
                new dev.erst.gridgrind.excel.ExcelPrintLayoutSnapshot(
                    new dev.erst.gridgrind.excel.ExcelPrintLayout(
                        new dev.erst.gridgrind.excel.ExcelPrintLayout.Area.Range("A1:B20"),
                        ExcelPrintOrientation.LANDSCAPE,
                        new dev.erst.gridgrind.excel.ExcelPrintLayout.Scaling.Fit(1, 0),
                        new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleRows.Band(0, 0),
                        new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleColumns.Band(0, 0),
                        new dev.erst.gridgrind.excel.ExcelHeaderFooterText("Budget", "", ""),
                        new dev.erst.gridgrind.excel.ExcelHeaderFooterText("", "Page &P", "")),
                    defaultPrintSetupSnapshot())));
    InspectionResult conditionalFormatting =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingResult(
                "conditional-formatting",
                "Budget",
                List.of(
                    new ExcelConditionalFormattingBlockSnapshot(
                        List.of("A2:A5"),
                        List.of(
                            new ExcelConditionalFormattingRuleSnapshot.FormulaRule(
                                1,
                                true,
                                "A2>0",
                                new ExcelDifferentialStyleSnapshot(
                                    "0.00", true, null, null, "#102030", null, null, "#E0F0AA",
                                    null, List.of())),
                            new ExcelConditionalFormattingRuleSnapshot.ColorScaleRule(
                                2,
                                false,
                                List.of(
                                    new ExcelConditionalFormattingThresholdSnapshot(
                                        ExcelConditionalFormattingThresholdType.MIN, null, null),
                                    new ExcelConditionalFormattingThresholdSnapshot(
                                        ExcelConditionalFormattingThresholdType.MAX, null, null)),
                                List.of("#AA0000", "#00AA00")))))));
    InspectionResult formulaSurface =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.FormulaSurfaceResult(
                "formula",
                new dev.erst.gridgrind.excel.WorkbookReadResult.FormulaSurface(
                    1,
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.SheetFormulaSurface(
                            "Budget",
                            1,
                            1,
                            List.of(
                                new dev.erst.gridgrind.excel.WorkbookReadResult.FormulaPattern(
                                    "SUM(B2:B3)", 1, List.of("B4"))))))));
    InspectionResult schema =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSchemaResult(
                "schema",
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSchema(
                    "Budget",
                    "A1",
                    3,
                    2,
                    2,
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.SchemaColumn(
                            0,
                            "A1",
                            "Item",
                            2,
                            0,
                            List.of(
                                new dev.erst.gridgrind.excel.WorkbookReadResult.TypeCount(
                                    "STRING", 2)),
                            "STRING")))));
    InspectionResult namedRangeSurface =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurfaceResult(
                "surface",
                new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurface(
                    1,
                    0,
                    1,
                    0,
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurfaceEntry(
                            "BudgetTotal",
                            new ExcelNamedRangeScope.WorkbookScope(),
                            "Budget!$B$4",
                            dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeBackingKind
                                .RANGE)))));
    InspectionResult formulaHealth =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.FormulaHealthResult(
                "formula-health",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.FormulaHealth(
                    1,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 0, 0, 1),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                                .FORMULA_VOLATILE_FUNCTION,
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.INFO,
                            "Volatile formula",
                            "Formula uses NOW().",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Cell(
                                "Budget", "B4"),
                            List.of("NOW()"))))));

    assertInstanceOf(InspectionResult.WorkbookSummaryResult.class, workbookSummary);
    assertInstanceOf(InspectionResult.NamedRangesResult.class, namedRanges);
    assertInstanceOf(InspectionResult.SheetSummaryResult.class, sheetSummary);
    assertInstanceOf(InspectionResult.CellsResult.class, cells);
    assertInstanceOf(InspectionResult.WindowResult.class, window);
    assertInstanceOf(InspectionResult.MergedRegionsResult.class, merged);
    assertInstanceOf(InspectionResult.HyperlinksResult.class, hyperlinks);
    assertInstanceOf(InspectionResult.CommentsResult.class, comments);
    assertInstanceOf(InspectionResult.SheetLayoutResult.class, layout);
    assertInstanceOf(InspectionResult.PrintLayoutResult.class, printLayout);
    assertInstanceOf(InspectionResult.ConditionalFormattingResult.class, conditionalFormatting);
    assertInstanceOf(InspectionResult.FormulaSurfaceResult.class, formulaSurface);
    assertInstanceOf(InspectionResult.SheetSchemaResult.class, schema);
    assertInstanceOf(InspectionResult.NamedRangeSurfaceResult.class, namedRangeSurface);
    assertInstanceOf(InspectionResult.FormulaHealthResult.class, formulaHealth);
    assertEquals(
        "Budget",
        cast(InspectionResult.WorkbookSummaryResult.class, workbookSummary)
            .workbook()
            .sheetNames()
            .getFirst());
    assertEquals(
        "BudgetTotal",
        cast(InspectionResult.NamedRangesResult.class, namedRanges)
            .namedRanges()
            .getFirst()
            .name());
    assertEquals(
        "Budget",
        cast(InspectionResult.SheetSummaryResult.class, sheetSummary).sheet().sheetName());
    assertEquals(
        "A1", cast(InspectionResult.CellsResult.class, cells).cells().getFirst().address());
    assertEquals(
        "A1",
        cast(InspectionResult.WindowResult.class, window)
            .window()
            .rows()
            .getFirst()
            .cells()
            .getFirst()
            .address());
    assertEquals(
        "A1:B2",
        cast(InspectionResult.MergedRegionsResult.class, merged)
            .mergedRegions()
            .getFirst()
            .range());
    assertEquals(
        new HyperlinkTarget.Url("https://example.com/report"),
        cast(InspectionResult.HyperlinksResult.class, hyperlinks)
            .hyperlinks()
            .getFirst()
            .hyperlink());
    assertEquals(
        "Review",
        cast(InspectionResult.CommentsResult.class, comments)
            .comments()
            .getFirst()
            .comment()
            .text());
    assertInstanceOf(
        PaneReport.Frozen.class,
        cast(InspectionResult.SheetLayoutResult.class, layout).layout().pane());
    assertEquals(
        125, cast(InspectionResult.SheetLayoutResult.class, layout).layout().zoomPercent());
    assertEquals(
        ExcelPrintOrientation.LANDSCAPE,
        cast(InspectionResult.PrintLayoutResult.class, printLayout).layout().orientation());
    assertEquals(
        2,
        cast(InspectionResult.ConditionalFormattingResult.class, conditionalFormatting)
            .conditionalFormattingBlocks()
            .getFirst()
            .rules()
            .size());
    assertEquals(
        1,
        cast(InspectionResult.FormulaSurfaceResult.class, formulaSurface)
            .analysis()
            .totalFormulaCellCount());
    assertEquals(
        "STRING",
        cast(InspectionResult.SheetSchemaResult.class, schema)
            .analysis()
            .columns()
            .getFirst()
            .dominantType());
    assertEquals(
        GridGrindResponse.NamedRangeBackingKind.RANGE,
        cast(InspectionResult.NamedRangeSurfaceResult.class, namedRangeSurface)
            .analysis()
            .namedRanges()
            .getFirst()
            .kind());
    assertEquals(
        1,
        cast(InspectionResult.FormulaHealthResult.class, formulaHealth)
            .analysis()
            .summary()
            .infoCount());
  }

  @Test
  void convertsRemainingAnalysisReadResultsIntoProtocolReadResults() {
    InspectionResult conditionalFormattingHealth =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingHealthResult(
                "conditional-formatting-health",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.ConditionalFormattingHealth(
                    1,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 1, 0, 0),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                                .CONDITIONAL_FORMATTING_PRIORITY_COLLISION,
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.ERROR,
                            "Priority collision",
                            "Conditional-formatting priorities collide.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Sheet(
                                "Budget"),
                            List.of("FORMULA_RULE@Budget!A1:A3"))))));
    InspectionResult hyperlinkHealth =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.HyperlinkHealthResult(
                "hyperlink-health",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.HyperlinkHealth(
                    2,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(2, 1, 1, 0),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                                .HYPERLINK_MALFORMED_TARGET,
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.ERROR,
                            "Malformed hyperlink target",
                            "Hyperlink target is malformed.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation
                                .Workbook(),
                            List.of("mailto:")),
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                                .HYPERLINK_MISSING_DOCUMENT_SHEET,
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.WARNING,
                            "Missing target sheet",
                            "Sheet is missing.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Sheet(
                                "Budget"),
                            List.of("Missing!A1"))))));
    InspectionResult namedRangeHealth =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeHealthResult(
                "named-range-health",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.NamedRangeHealth(
                    1,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 0, 1, 0),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                                .NAMED_RANGE_BROKEN_REFERENCE,
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.WARNING,
                            "Broken named range",
                            "Named range contains #REF!.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Range(
                                "Budget", "A1:B2"),
                            List.of("#REF!"))))));
    InspectionResult workbookFindings =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookFindingsResult(
                "workbook-findings",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.WorkbookFindings(
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 0, 0, 1),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                                .NAMED_RANGE_SCOPE_SHADOWING,
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.INFO,
                            "Scope shadowing",
                            "Name exists in more than one scope.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation
                                .NamedRange(
                                "BudgetTotal", new ExcelNamedRangeScope.SheetScope("Budget")),
                            List.of("Budget!$B$4"))))));

    assertEquals(
        1,
        cast(InspectionResult.ConditionalFormattingHealthResult.class, conditionalFormattingHealth)
            .analysis()
            .checkedConditionalFormattingBlockCount());
    GridGrindResponse.AnalysisLocationReport workbookLocation =
        cast(InspectionResult.HyperlinkHealthResult.class, hyperlinkHealth)
            .analysis()
            .findings()
            .getFirst()
            .location();
    GridGrindResponse.AnalysisLocationReport sheetLocation =
        cast(InspectionResult.HyperlinkHealthResult.class, hyperlinkHealth)
            .analysis()
            .findings()
            .get(1)
            .location();
    GridGrindResponse.AnalysisLocationReport rangeLocation =
        cast(InspectionResult.NamedRangeHealthResult.class, namedRangeHealth)
            .analysis()
            .findings()
            .getFirst()
            .location();
    GridGrindResponse.AnalysisLocationReport namedRangeLocation =
        cast(InspectionResult.WorkbookFindingsResult.class, workbookFindings)
            .analysis()
            .findings()
            .getFirst()
            .location();

    assertInstanceOf(GridGrindResponse.AnalysisLocationReport.Workbook.class, workbookLocation);
    assertEquals(
        "Budget",
        cast(GridGrindResponse.AnalysisLocationReport.Sheet.class, sheetLocation).sheetName());
    assertEquals(
        "A1:B2", cast(GridGrindResponse.AnalysisLocationReport.Range.class, rangeLocation).range());
    assertEquals(
        "BudgetTotal",
        cast(GridGrindResponse.AnalysisLocationReport.NamedRange.class, namedRangeLocation).name());
  }

  @Test
  void convertsNamedRangeFormulaSnapshotsAndFormulaBackedSurfaceEntries() {
    GridGrindResponse.NamedRangeReport formulaReport =
        InspectionResultWorkbookCoreReportSupport.toNamedRangeReport(
            new ExcelNamedRangeSnapshot.FormulaSnapshot(
                "BudgetRollup", new ExcelNamedRangeScope.WorkbookScope(), "SUM(Budget!$B$2:$B$3)"));
    assertInstanceOf(GridGrindResponse.NamedRangeReport.FormulaReport.class, formulaReport);

    InspectionResult.NamedRangeSurfaceResult surface =
        cast(
            InspectionResult.NamedRangeSurfaceResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurfaceResult(
                    "surface",
                    new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurface(
                        0,
                        1,
                        0,
                        1,
                        List.of(
                            new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurfaceEntry(
                                "BudgetRollup",
                                new ExcelNamedRangeScope.SheetScope("Budget"),
                                "SUM(Budget!$B$2:$B$3)",
                                dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeBackingKind
                                    .FORMULA))))));

    assertEquals(
        GridGrindResponse.NamedRangeBackingKind.FORMULA,
        surface.analysis().namedRanges().getFirst().kind());
  }

  @Test
  void convertsPaneNoneIntoProtocolReport() {
    InspectionResult.SheetLayoutResult layout =
        cast(
            InspectionResult.SheetLayoutResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayoutResult(
                    "layout",
                    new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayout(
                        "Budget",
                        new dev.erst.gridgrind.excel.ExcelSheetPane.None(),
                        100,
                        defaultSheetPresentationSnapshot(),
                        List.of(),
                        List.of()))));

    assertInstanceOf(PaneReport.None.class, layout.layout().pane());
  }

  @Test
  void convertsSplitPaneAndDefaultPrintLayoutIntoProtocolReports() {
    InspectionResult.SheetLayoutResult layout =
        cast(
            InspectionResult.SheetLayoutResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayoutResult(
                    "layout",
                    new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayout(
                        "Budget",
                        new dev.erst.gridgrind.excel.ExcelSheetPane.Split(
                            1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT),
                        100,
                        defaultSheetPresentationSnapshot(),
                        List.of(),
                        List.of()))));
    InspectionResult.PrintLayoutResult printLayout =
        cast(
            InspectionResult.PrintLayoutResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.PrintLayoutResult(
                    "print-layout",
                    "Budget",
                    new dev.erst.gridgrind.excel.ExcelPrintLayoutSnapshot(
                        new dev.erst.gridgrind.excel.ExcelPrintLayout(
                            new dev.erst.gridgrind.excel.ExcelPrintLayout.Area.None(),
                            ExcelPrintOrientation.PORTRAIT,
                            new dev.erst.gridgrind.excel.ExcelPrintLayout.Scaling.Automatic(),
                            new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleRows.None(),
                            new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleColumns.None(),
                            new dev.erst.gridgrind.excel.ExcelHeaderFooterText("", "", ""),
                            new dev.erst.gridgrind.excel.ExcelHeaderFooterText("", "", "")),
                        defaultPrintSetupSnapshot()))));

    PaneReport.Split pane = assertInstanceOf(PaneReport.Split.class, layout.layout().pane());
    assertEquals(1200, pane.xSplitPosition());
    assertEquals(2400, pane.ySplitPosition());
    assertEquals(3, pane.leftmostColumn());
    assertEquals(4, pane.topRow());
    assertEquals(ExcelPaneRegion.LOWER_RIGHT, pane.activePane());
    assertInstanceOf(PrintAreaReport.None.class, printLayout.layout().printArea());
    assertInstanceOf(PrintScalingReport.Automatic.class, printLayout.layout().scaling());
    assertInstanceOf(PrintTitleRowsReport.None.class, printLayout.layout().repeatingRows());
    assertInstanceOf(PrintTitleColumnsReport.None.class, printLayout.layout().repeatingColumns());
  }

  @Test
  void convertsSheetStateReadResultsIntoProtocolShapes() {
    InspectionResult emptyWorkbookSummary =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult(
                "workbook",
                new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary.Empty(
                    0, List.of(), 0, false)));
    InspectionResult populatedWorkbookSummary =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult(
                "workbook-2",
                new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary.WithSheets(
                    2,
                    List.of("Budget", "Archive"),
                    "Archive",
                    List.of("Budget", "Archive"),
                    1,
                    true)));
    InspectionResult protectedSheetSummary =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummaryResult(
                "sheet",
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummary(
                    "Budget",
                    dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility.VERY_HIDDEN,
                    new dev.erst.gridgrind.excel.WorkbookReadResult.SheetProtection.Protected(
                        excelProtectionSettings()),
                    4,
                    7,
                    3)));

    GridGrindResponse.WorkbookSummary.Empty empty =
        assertInstanceOf(
            GridGrindResponse.WorkbookSummary.Empty.class,
            cast(InspectionResult.WorkbookSummaryResult.class, emptyWorkbookSummary).workbook());
    GridGrindResponse.WorkbookSummary.WithSheets populated =
        assertInstanceOf(
            GridGrindResponse.WorkbookSummary.WithSheets.class,
            cast(InspectionResult.WorkbookSummaryResult.class, populatedWorkbookSummary)
                .workbook());
    GridGrindResponse.SheetSummaryReport sheet =
        cast(InspectionResult.SheetSummaryResult.class, protectedSheetSummary).sheet();
    GridGrindResponse.SheetProtectionReport.Protected protection =
        assertInstanceOf(
            GridGrindResponse.SheetProtectionReport.Protected.class, sheet.protection());

    assertEquals(0, empty.sheetCount());
    assertEquals("Archive", populated.activeSheetName());
    assertEquals(List.of("Budget", "Archive"), populated.selectedSheetNames());
    assertEquals(ExcelSheetVisibility.VERY_HIDDEN, sheet.visibility());
    assertEquals(protectionSettings(), protection.settings());
  }

  @Test
  void readTypeReturnsDiscriminatorsForAllReadVariants() {
    assertEquals(
        List.of(
            "GET_WORKBOOK_SUMMARY",
            "GET_WORKBOOK_PROTECTION",
            "GET_NAMED_RANGES",
            "GET_SHEET_SUMMARY",
            "GET_CELLS",
            "GET_WINDOW",
            "GET_MERGED_REGIONS",
            "GET_HYPERLINKS",
            "GET_COMMENTS",
            "GET_SHEET_LAYOUT",
            "GET_PRINT_LAYOUT",
            "GET_DATA_VALIDATIONS",
            "GET_CONDITIONAL_FORMATTING",
            "GET_AUTOFILTERS",
            "GET_TABLES",
            "GET_FORMULA_SURFACE",
            "GET_SHEET_SCHEMA",
            "GET_NAMED_RANGE_SURFACE",
            "ANALYZE_FORMULA_HEALTH",
            "ANALYZE_DATA_VALIDATION_HEALTH",
            "ANALYZE_CONDITIONAL_FORMATTING_HEALTH",
            "ANALYZE_AUTOFILTER_HEALTH",
            "ANALYZE_TABLE_HEALTH",
            "ANALYZE_HYPERLINK_HEALTH",
            "ANALYZE_NAMED_RANGE_HEALTH",
            "ANALYZE_WORKBOOK_FINDINGS"),
        List.of(
            readType(
                inspect(
                    "workbook",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookSummary())),
            readType(
                inspect(
                    "workbook-protection",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookProtection())),
            readType(
                inspect(
                    "ranges",
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                    new InspectionQuery.GetNamedRanges())),
            readType(
                inspect(
                    "sheet",
                    new SheetSelector.ByName("Budget"),
                    new InspectionQuery.GetSheetSummary())),
            readType(
                inspect(
                    "cells",
                    new CellSelector.ByAddresses("Budget", List.of("A1")),
                    new InspectionQuery.GetCells())),
            readType(
                inspect(
                    "window",
                    new RangeSelector.RectangularWindow("Budget", "A1", 1, 1),
                    new InspectionQuery.GetWindow())),
            readType(
                inspect(
                    "merged",
                    new SheetSelector.ByName("Budget"),
                    new InspectionQuery.GetMergedRegions())),
            readType(
                inspect(
                    "hyperlinks",
                    new CellSelector.AllUsedInSheet("Budget"),
                    new InspectionQuery.GetHyperlinks())),
            readType(
                inspect(
                    "comments",
                    new CellSelector.AllUsedInSheet("Budget"),
                    new InspectionQuery.GetComments())),
            readType(
                inspect(
                    "layout",
                    new SheetSelector.ByName("Budget"),
                    new InspectionQuery.GetSheetLayout())),
            readType(
                inspect(
                    "print-layout",
                    new SheetSelector.ByName("Budget"),
                    new InspectionQuery.GetPrintLayout())),
            readType(
                inspect(
                    "validations",
                    new RangeSelector.AllOnSheet("Budget"),
                    new InspectionQuery.GetDataValidations())),
            readType(
                inspect(
                    "conditional-formatting",
                    new RangeSelector.AllOnSheet("Budget"),
                    new InspectionQuery.GetConditionalFormatting())),
            readType(
                inspect(
                    "autofilters",
                    new SheetSelector.ByName("Budget"),
                    new InspectionQuery.GetAutofilters())),
            readType(
                inspect(
                    "tables",
                    new TableSelector.ByNames(List.of("BudgetTable")),
                    new InspectionQuery.GetTables())),
            readType(
                inspect(
                    "formula", new SheetSelector.All(), new InspectionQuery.GetFormulaSurface())),
            readType(
                inspect(
                    "schema",
                    new RangeSelector.RectangularWindow("Budget", "A1", 1, 1),
                    new InspectionQuery.GetSheetSchema())),
            readType(
                inspect(
                    "surface",
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                    new InspectionQuery.GetNamedRangeSurface())),
            readType(
                inspect(
                    "formula-health",
                    new SheetSelector.All(),
                    new InspectionQuery.AnalyzeFormulaHealth())),
            readType(
                inspect(
                    "validation-health",
                    new SheetSelector.All(),
                    new InspectionQuery.AnalyzeDataValidationHealth())),
            readType(
                inspect(
                    "conditional-formatting-health",
                    new SheetSelector.All(),
                    new InspectionQuery.AnalyzeConditionalFormattingHealth())),
            readType(
                inspect(
                    "autofilter-health",
                    new SheetSelector.All(),
                    new InspectionQuery.AnalyzeAutofilterHealth())),
            readType(
                inspect(
                    "table-health",
                    new TableSelector.All(),
                    new InspectionQuery.AnalyzeTableHealth())),
            readType(
                inspect(
                    "hyperlink-health",
                    new SheetSelector.All(),
                    new InspectionQuery.AnalyzeHyperlinkHealth())),
            readType(
                inspect(
                    "named-range-health",
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                    new InspectionQuery.AnalyzeNamedRangeHealth())),
            readType(
                inspect(
                    "workbook-findings",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.AnalyzeWorkbookFindings()))));
  }

  @Test
  void extractsContextForReadOperations() {
    NamedRangeNotFoundException missingNamedRange =
        new NamedRangeNotFoundException("BudgetTotal", new ExcelNamedRangeScope.WorkbookScope());
    InvalidCellAddressException invalidAddress =
        new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"));
    RuntimeException runtimeException = new RuntimeException("x");

    InspectionStep workbook =
        inspect(
            "workbook", new WorkbookSelector.Current(), new InspectionQuery.GetWorkbookSummary());
    InspectionStep workbookProtection =
        inspect(
            "workbook-protection",
            new WorkbookSelector.Current(),
            new InspectionQuery.GetWorkbookProtection());
    InspectionStep namedRanges =
        inspect(
            "ranges",
            new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
            new InspectionQuery.GetNamedRanges());
    InspectionStep sheet =
        inspect("sheet", new SheetSelector.ByName("Budget"), new InspectionQuery.GetSheetSummary());
    InspectionStep cells =
        inspect(
            "cells",
            new CellSelector.ByAddresses("Budget", List.of("A1")),
            new InspectionQuery.GetCells());
    InspectionStep window =
        inspect(
            "window",
            new RangeSelector.RectangularWindow("Budget", "B2", 2, 2),
            new InspectionQuery.GetWindow());
    InspectionStep merged =
        inspect(
            "merged", new SheetSelector.ByName("Budget"), new InspectionQuery.GetMergedRegions());
    InspectionStep hyperlinks =
        inspect(
            "hyperlinks",
            new CellSelector.AllUsedInSheet("Budget"),
            new InspectionQuery.GetHyperlinks());
    InspectionStep comments =
        inspect(
            "comments",
            new CellSelector.ByAddresses("Budget", List.of("A1")),
            new InspectionQuery.GetComments());
    InspectionStep layout =
        inspect("layout", new SheetSelector.ByName("Budget"), new InspectionQuery.GetSheetLayout());
    InspectionStep printLayout =
        inspect(
            "print-layout",
            new SheetSelector.ByName("Budget"),
            new InspectionQuery.GetPrintLayout());
    InspectionStep validations =
        inspect(
            "validations",
            new RangeSelector.ByRanges("Budget", List.of("A1:A3")),
            new InspectionQuery.GetDataValidations());
    InspectionStep conditionalFormatting =
        inspect(
            "conditional-formatting",
            new RangeSelector.ByRanges("Budget", List.of("B2:B5")),
            new InspectionQuery.GetConditionalFormatting());
    InspectionStep autofilters =
        inspect(
            "autofilters",
            new SheetSelector.ByName("Budget"),
            new InspectionQuery.GetAutofilters());
    InspectionStep tables =
        inspect(
            "tables",
            new TableSelector.ByNames(List.of("BudgetTable")),
            new InspectionQuery.GetTables());
    InspectionStep formula =
        inspect("formula", new SheetSelector.All(), new InspectionQuery.GetFormulaSurface());
    InspectionStep schema =
        inspect(
            "schema",
            new RangeSelector.RectangularWindow("Budget", "C3", 2, 2),
            new InspectionQuery.GetSheetSchema());
    InspectionStep surface =
        inspect(
            "surface",
            new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
            new InspectionQuery.GetNamedRangeSurface());
    InspectionStep formulaHealth =
        inspect(
            "formula-health",
            new SheetSelector.ByNames(List.of("Budget")),
            new InspectionQuery.AnalyzeFormulaHealth());
    InspectionStep validationHealth =
        inspect(
            "validation-health",
            new SheetSelector.ByNames(List.of("Budget")),
            new InspectionQuery.AnalyzeDataValidationHealth());
    InspectionStep conditionalFormattingHealth =
        inspect(
            "conditional-formatting-health",
            new SheetSelector.ByNames(List.of("Budget")),
            new InspectionQuery.AnalyzeConditionalFormattingHealth());
    InspectionStep autofilterHealth =
        inspect(
            "autofilter-health",
            new SheetSelector.ByNames(List.of("Budget")),
            new InspectionQuery.AnalyzeAutofilterHealth());
    InspectionStep tableHealth =
        inspect(
            "table-health",
            new TableSelector.ByNames(List.of("BudgetTable")),
            new InspectionQuery.AnalyzeTableHealth());
    InspectionStep hyperlinkHealth =
        inspect(
            "hyperlink-health",
            new SheetSelector.All(),
            new InspectionQuery.AnalyzeHyperlinkHealth());
    InspectionStep namedRangeHealth =
        inspect(
            "named-range-health",
            new dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf(
                List.of(
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope(
                        "BudgetTotal"))),
            new InspectionQuery.AnalyzeNamedRangeHealth());
    InspectionStep workbookFindings =
        inspect(
            "workbook-findings",
            new WorkbookSelector.Current(),
            new InspectionQuery.AnalyzeWorkbookFindings());

    assertReadContext(workbook, null, null, null, runtimeException);
    assertReadContext(workbookProtection, null, null, null, runtimeException);
    assertReadContext(namedRanges, null, null, null, runtimeException);
    assertReadContext(sheet, "Budget", null, null, runtimeException);
    assertReadContext(cells, "Budget", null, null, runtimeException);
    assertReadContext(window, "Budget", "B2", null, runtimeException);
    assertReadContext(merged, "Budget", null, null, runtimeException);
    assertReadContext(hyperlinks, "Budget", null, null, runtimeException);
    assertReadContext(comments, "Budget", null, null, runtimeException);
    assertReadContext(layout, "Budget", null, null, runtimeException);
    assertReadContext(printLayout, "Budget", null, null, runtimeException);
    assertReadContext(validations, "Budget", null, null, runtimeException);
    assertReadContext(conditionalFormatting, "Budget", null, null, runtimeException);
    assertReadContext(autofilters, "Budget", null, null, runtimeException);
    assertReadContext(tables, null, null, null, runtimeException);
    assertReadContext(formula, null, null, null, runtimeException);
    assertReadContext(schema, "Budget", "C3", null, runtimeException);
    assertReadContext(surface, null, null, null, runtimeException);
    assertReadContext(formulaHealth, "Budget", null, null, runtimeException);
    assertReadContext(validationHealth, "Budget", null, null, runtimeException);
    assertReadContext(conditionalFormattingHealth, "Budget", null, null, runtimeException);
    assertReadContext(autofilterHealth, "Budget", null, null, runtimeException);
    assertReadContext(tableHealth, null, null, null, runtimeException);
    assertReadContext(hyperlinkHealth, null, null, null, runtimeException);
    assertReadContext(namedRangeHealth, null, null, "BudgetTotal", runtimeException);
    assertReadContext(workbookFindings, null, null, null, runtimeException);

    assertEquals("BudgetTotal", namedRangeNameFor(workbook, missingNamedRange));
    assertEquals("BudgetTotal", namedRangeNameFor(namedRanges, missingNamedRange));
    assertEquals("BAD!", addressFor(cells, invalidAddress));
  }

  @Test
  void extractsSingleSheetAndNamedRangeContextOnlyWhenSelectionsAreUnambiguous() {
    assertEquals(
        "Budget",
        sheetNameFor(
            inspect(
                "formula",
                new SheetSelector.ByNames(List.of("Budget")),
                new InspectionQuery.GetFormulaSurface())));
    assertNull(
        sheetNameFor(
            inspect(
                "hyperlink-health",
                new SheetSelector.ByNames(List.of("Budget", "Forecast")),
                new InspectionQuery.AnalyzeHyperlinkHealth())));

    assertEquals(
        "BudgetTotal",
        namedRangeNameFor(
            inspect(
                "surface",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf(
                    List.of(
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.ByName(
                            "BudgetTotal"))),
                new InspectionQuery.GetNamedRangeSurface()),
            new RuntimeException("x")));
    assertEquals(
        "BudgetTotal",
        namedRangeNameFor(
            inspect(
                "surface",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf(
                    List.of(
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope(
                            "BudgetTotal", "Budget"))),
                new InspectionQuery.GetNamedRangeSurface()),
            new RuntimeException("x")));
    assertNull(
        namedRangeNameFor(
            inspect(
                "surface",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf(
                    List.of(
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope(
                            "BudgetTotal"),
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope(
                            "ForecastTotal"))),
                new InspectionQuery.GetNamedRangeSurface()),
            new RuntimeException("x")));
  }

  @Test
  void extractsReadContextFromExceptionsBeforeFallingBackToReadShape() {
    InvalidCellAddressException invalidAddress =
        new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"));
    NamedRangeNotFoundException missingNamedRange =
        new NamedRangeNotFoundException("BudgetTotal", new ExcelNamedRangeScope.WorkbookScope());

    assertEquals(
        "C3",
        addressFor(
            inspect(
                "schema",
                new RangeSelector.RectangularWindow("Budget", "C3", 2, 2),
                new InspectionQuery.GetSheetSchema()),
            invalidAddress));
    assertEquals(
        "BudgetTotal",
        namedRangeNameFor(
            inspect(
                "surface",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                new InspectionQuery.GetNamedRangeSurface()),
            missingNamedRange));
  }

  @Test
  void extractsContextForAuthoringMetadataAndNamedRangeOperations() {
    RuntimeException exception = new RuntimeException("test");
    assertWriteContext(
        mutate(
            new CellSelector.ByAddress("Budget", "A1"),
            new MutationAction.SetHyperlink(new HyperlinkTarget.Url("https://example.com/report"))),
        exception,
        "Budget",
        "A1",
        null,
        null);
    assertWriteContext(
        mutate(new CellSelector.ByAddress("Budget", "A1"), new MutationAction.ClearHyperlink()),
        exception,
        "Budget",
        "A1",
        null,
        null);
    assertWriteContext(
        mutate(
            new CellSelector.ByAddress("Budget", "A1"),
            new MutationAction.SetComment(new CommentInput(text("Review"), "GridGrind", false))),
        exception,
        "Budget",
        "A1",
        null,
        null);
    assertWriteContext(
        mutate(new CellSelector.ByAddress("Budget", "A1"), new MutationAction.ClearComment()),
        exception,
        "Budget",
        "A1",
        null,
        null);
    assertWriteContext(
        mutate(
            new RangeSelector.ByRange("Budget", "A1:B2"),
            new MutationAction.ApplyStyle(
                new CellStyleInput(
                    null,
                    null,
                    new CellFontInput(true, null, null, null, null, null, null),
                    null,
                    null,
                    null))),
        exception,
        "Budget",
        null,
        "A1:B2",
        null);
    assertWriteContext(
        mutate(
            new RangeSelector.ByRange("Budget", "B2:B5"),
            new MutationAction.SetDataValidation(
                new DataValidationInput(
                    new DataValidationRuleInput.TextLength(
                        ExcelComparisonOperator.LESS_OR_EQUAL, "20", null),
                    true,
                    false,
                    prompt("Reason", "Use 20 characters or fewer.", true),
                    null))),
        exception,
        "Budget",
        null,
        "B2:B5",
        null);
    assertWriteContext(
        mutate(new RangeSelector.AllOnSheet("Budget"), new MutationAction.ClearDataValidations()),
        exception,
        "Budget",
        null,
        null,
        null);
    assertWriteContext(
        mutate(
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetConditionalFormatting(
                new ConditionalFormattingBlockInput(
                    List.of("B2:B5"),
                    List.of(
                        new ConditionalFormattingRuleInput.FormulaRule(
                            "B2>0",
                            true,
                            new DifferentialStyleInput(
                                null, true, null, null, null, null, null, null, null)))))),
        exception,
        "Budget",
        null,
        "B2:B5",
        null);
    assertWriteContext(
        mutate(
            new RangeSelector.AllOnSheet("Budget"),
            new MutationAction.ClearConditionalFormatting()),
        exception,
        "Budget",
        null,
        null,
        null);
    assertWriteContext(
        mutate(new RangeSelector.ByRange("Budget", "A1:C4"), new MutationAction.SetAutofilter()),
        exception,
        "Budget",
        null,
        "A1:C4",
        null);
    assertWriteContext(
        mutate(new SheetSelector.ByName("Budget"), new MutationAction.ClearAutofilter()),
        exception,
        "Budget",
        null,
        null,
        null);
    assertWriteContext(
        mutate(
            new MutationAction.SetTable(
                new TableInput(
                    "BudgetTable", "Budget", "A1:C4", false, new TableStyleInput.None()))),
        exception,
        "Budget",
        null,
        "A1:C4",
        null);
    assertWriteContext(
        mutate(
            new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
            new MutationAction.DeleteTable()),
        exception,
        "Budget",
        null,
        null,
        null);
    assertWriteContext(
        mutate(
            new MutationAction.SetNamedRange(
                "BudgetTotal",
                new NamedRangeScope.Workbook(),
                new NamedRangeTarget("Budget", "B4"))),
        exception,
        "Budget",
        null,
        "B4",
        "BudgetTotal");
    assertWriteContext(
        mutate(
            new dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope(
                "BudgetTotal"),
            new MutationAction.DeleteNamedRange()),
        exception,
        null,
        null,
        null,
        "BudgetTotal");
    assertWriteContext(
        mutate(
            new dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope(
                "LocalItem", "Budget"),
            new MutationAction.DeleteNamedRange()),
        exception,
        "Budget",
        null,
        null,
        "LocalItem");
    assertEquals(
        "BudgetTotal",
        namedRangeNameFor(
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
            new NamedRangeNotFoundException(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
  }

  @Test
  void extractsAddressAndRangeFallbacksAndExhaustsNamedRangeNullArms() {
    InvalidFormulaException invalidFormula =
        new InvalidFormulaException(
            "Budget", "C3", "SUM(", "bad formula", new IllegalArgumentException("bad"));
    NamedRangeNotFoundException missingNamedRange =
        new NamedRangeNotFoundException("BudgetTotal", new ExcelNamedRangeScope.WorkbookScope());
    InvalidCellAddressException invalidAddress =
        new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"));
    InvalidRangeAddressException invalidRange =
        new InvalidRangeAddressException("A1:", new IllegalArgumentException("bad"));

    assertEquals(
        "C3",
        addressFor(
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
            invalidFormula));
    assertEquals(
        "BAD!",
        addressFor(
            mutate(
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope(
                    "BudgetTotal"),
                new MutationAction.DeleteNamedRange()),
            invalidAddress));

    assertEquals(
        "C1:D2",
        rangeFor(
            mutate(
                new RangeSelector.ByRange("Budget", "C1:D2"),
                new MutationAction.SetRange(List.of(List.of(textCell("x"))))),
            invalidFormula));
    assertEquals(
        "E1:E2",
        rangeFor(
            mutate(new RangeSelector.ByRange("Budget", "E1:E2"), new MutationAction.ClearRange()),
            invalidFormula));
    assertEquals(
        "B2:B5",
        rangeFor(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetConditionalFormatting(
                    new ConditionalFormattingBlockInput(
                        List.of("B2:B5"),
                        List.of(
                            new ConditionalFormattingRuleInput.FormulaRule(
                                "B2>0",
                                true,
                                new DifferentialStyleInput(
                                    null, true, null, null, null, null, null, null, null)))))),
            invalidFormula));
    assertNull(
        rangeFor(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetConditionalFormatting(
                    new ConditionalFormattingBlockInput(
                        List.of("B2:B5", "D2:D5"),
                        List.of(
                            new ConditionalFormattingRuleInput.FormulaRule(
                                "B2>0",
                                true,
                                new DifferentialStyleInput(
                                    null, true, null, null, null, null, null, null, null)))))),
            invalidFormula));
    assertNull(
        rangeFor(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new MutationAction.SetHyperlink(
                    new HyperlinkTarget.Url("https://example.com/report"))),
            invalidFormula));
    assertNull(
        rangeFor(
            mutate(new CellSelector.ByAddress("Budget", "A1"), new MutationAction.ClearHyperlink()),
            invalidFormula));
    assertNull(
        rangeFor(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new MutationAction.SetComment(
                    new CommentInput(text("Review"), "GridGrind", false))),
            invalidFormula));
    assertNull(
        rangeFor(
            mutate(new CellSelector.ByAddress("Budget", "A1"), new MutationAction.ClearComment()),
            invalidFormula));
    assertEquals(
        "A1:",
        rangeFor(
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
            invalidRange));

    List<ExecutorTestPlanSupport.PendingMutation> operationsWithoutNamedRanges =
        List.of(
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.RenameSheet("Summary")),
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.DeleteSheet()),
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.MoveSheet(0)),
            mutate(new RangeSelector.ByRange("Budget", "A1:B2"), new MutationAction.MergeCells()),
            mutate(new RangeSelector.ByRange("Budget", "A1:B2"), new MutationAction.UnmergeCells()),
            mutate(
                new ColumnBandSelector.Span("Budget", 0, 1),
                new MutationAction.SetColumnWidth(16.0)),
            mutate(new RowBandSelector.Span("Budget", 0, 1), new MutationAction.SetRowHeight(28.5)),
            mutate(new RowBandSelector.Insertion("Budget", 1, 2), new MutationAction.InsertRows()),
            mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.DeleteRows()),
            mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.ShiftRows(1)),
            mutate(
                new ColumnBandSelector.Insertion("Budget", 1, 2),
                new MutationAction.InsertColumns()),
            mutate(new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.DeleteColumns()),
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.ShiftColumns(-1)),
            mutate(
                new RowBandSelector.Span("Budget", 1, 2),
                new MutationAction.SetRowVisibility(true)),
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2),
                new MutationAction.SetColumnVisibility(false)),
            mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.GroupRows(true)),
            mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.UngroupRows()),
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.GroupColumns(true)),
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.UngroupColumns()),
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetSheetPane(new PaneInput.Frozen(1, 1, 1, 1))),
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.SetSheetZoom(125)),
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetPrintLayout(
                    new PrintLayoutInput(
                        new PrintAreaInput.Range("A1:B12"),
                        ExcelPrintOrientation.LANDSCAPE,
                        new PrintScalingInput.Fit(1, 0),
                        new PrintTitleRowsInput.Band(0, 0),
                        new PrintTitleColumnsInput.Band(0, 0),
                        headerFooter("Budget", "", ""),
                        headerFooter("", "Page &P", "")))),
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.ClearPrintLayout()),
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new MutationAction.SetCell(textCell("x"))),
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B1"),
                new MutationAction.SetRange(List.of(List.of(textCell("x"))))),
            mutate(new RangeSelector.ByRange("Budget", "A1:B1"), new MutationAction.ClearRange()),
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new MutationAction.SetHyperlink(
                    new HyperlinkTarget.Url("https://example.com/report"))),
            mutate(new CellSelector.ByAddress("Budget", "A1"), new MutationAction.ClearHyperlink()),
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new MutationAction.SetComment(
                    new CommentInput(text("Review"), "GridGrind", false))),
            mutate(new CellSelector.ByAddress("Budget", "A1"), new MutationAction.ClearComment()),
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B2"),
                new MutationAction.ApplyStyle(
                    new CellStyleInput(
                        null,
                        null,
                        new CellFontInput(true, null, null, null, null, null, null),
                        null,
                        null,
                        null))),
            mutate(
                new RangeSelector.ByRange("Budget", "B2:B5"),
                new MutationAction.SetDataValidation(
                    new DataValidationInput(
                        new DataValidationRuleInput.TextLength(
                            ExcelComparisonOperator.LESS_OR_EQUAL, "20", null),
                        true,
                        false,
                        null,
                        null))),
            mutate(
                new RangeSelector.AllOnSheet("Budget"), new MutationAction.ClearDataValidations()),
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetConditionalFormatting(
                    new ConditionalFormattingBlockInput(
                        List.of("B2:B5"),
                        List.of(
                            new ConditionalFormattingRuleInput.FormulaRule(
                                "B2>0",
                                true,
                                new DifferentialStyleInput(
                                    null, true, null, null, null, null, null, null, null)))))),
            mutate(
                new RangeSelector.AllOnSheet("Budget"),
                new MutationAction.ClearConditionalFormatting()),
            mutate(
                new RangeSelector.ByRange("Budget", "A1:C4"), new MutationAction.SetAutofilter()),
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.ClearAutofilter()),
            mutate(
                new MutationAction.SetTable(
                    new TableInput(
                        "BudgetTable", "Budget", "A1:C4", false, new TableStyleInput.None()))),
            mutate(
                new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
                new MutationAction.DeleteTable()),
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.AppendRow(List.of(textCell("x")))),
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.AutoSizeColumns()));

    for (ExecutorTestPlanSupport.PendingMutation operation : operationsWithoutNamedRanges) {
      assertNull(namedRangeNameFor(operation, invalidFormula));
    }

    assertEquals(
        "BudgetTotal",
        namedRangeNameFor(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetSheetPane(new PaneInput.Frozen(1, 1, 1, 1))),
            missingNamedRange));
  }

  private static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<ExecutorTestPlanSupport.PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return ExecutorTestPlanSupport.request(
        source, persistence, execution, formulaEnvironment, mutations, inspections);
  }

  private static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<ExecutorTestPlanSupport.PendingMutation> mutations,
      List<ExecutorTestPlanSupport.PendingAssertion> assertions,
      List<InspectionStep> inspections) {
    return ExecutorTestPlanSupport.request(
        source, persistence, execution, formulaEnvironment, mutations, assertions, inspections);
  }

  private static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<ExecutorTestPlanSupport.PendingMutation> mutations,
      InspectionStep... inspections) {
    return ExecutorTestPlanSupport.request(
        source, persistence, execution, formulaEnvironment, mutations, inspections);
  }

  private static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      List<ExecutorTestPlanSupport.PendingMutation> mutations,
      List<ExecutorTestPlanSupport.PendingAssertion> assertions,
      List<InspectionStep> inspections) {
    return ExecutorTestPlanSupport.request(source, persistence, mutations, assertions, inspections);
  }

  private static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      List<ExecutorTestPlanSupport.PendingMutation> mutations,
      InspectionStep... inspections) {
    return ExecutorTestPlanSupport.request(source, persistence, mutations, inspections);
  }

  private static GridGrindResponse.Success success(GridGrindResponse response) {
    return cast(GridGrindResponse.Success.class, response);
  }

  private static GridGrindResponse.Failure failure(GridGrindResponse response) {
    return cast(GridGrindResponse.Failure.class, response);
  }

  private static String savedPath(GridGrindResponse.Success success) {
    return switch (success.persistence()) {
      case GridGrindResponse.PersistenceOutcome.SavedAs savedAs -> savedAs.executionPath();
      case GridGrindResponse.PersistenceOutcome.Overwritten overwritten ->
          overwritten.executionPath();
      case GridGrindResponse.PersistenceOutcome.NotSaved _ ->
          throw new AssertionError("expected persisted workbook");
    };
  }

  private static List<String> stepIds(GridGrindResponse.Success success) {
    return inspectionIds(success);
  }

  private static WorkbookCommand command(ExecutorTestPlanSupport.PendingMutation mutation) {
    return WorkbookCommandConverter.toCommand(mutation.target(), mutation.action());
  }

  private static WorkbookReadCommand readCommand(InspectionStep step) {
    return InspectionCommandConverter.toReadCommand(step);
  }

  private static String readType(InspectionStep step) {
    return step.query().queryType();
  }

  private static <T> T cast(Class<T> type, Object value) {
    return type.cast(assertInstanceOf(type, value));
  }

  private static String sheetNameFor(WorkbookStep step) {
    return ExecutionDiagnosticFields.sheetNameFor(step);
  }

  private static String sheetNameFor(WorkbookStep step, Exception exception) {
    return ExecutionDiagnosticFields.sheetNameFor(step, exception);
  }

  private static String sheetNameFor(
      ExecutorTestPlanSupport.PendingMutation mutation, Exception exception) {
    return sheetNameFor(materializeMutation(mutation, 0), exception);
  }

  private static String addressFor(WorkbookStep step, Exception exception) {
    return ExecutionDiagnosticFields.addressFor(step, exception);
  }

  private static String addressFor(
      ExecutorTestPlanSupport.PendingMutation mutation, Exception exception) {
    return addressFor(materializeMutation(mutation, 0), exception);
  }

  private static String rangeFor(WorkbookStep step, Exception exception) {
    return ExecutionDiagnosticFields.rangeFor(step, exception);
  }

  private static String rangeFor(
      ExecutorTestPlanSupport.PendingMutation mutation, Exception exception) {
    return rangeFor(materializeMutation(mutation, 0), exception);
  }

  private static String formulaFor(WorkbookStep step, Exception exception) {
    return ExecutionDiagnosticFields.formulaFor(step, exception);
  }

  private static String formulaFor(
      ExecutorTestPlanSupport.PendingMutation mutation, Exception exception) {
    return formulaFor(materializeMutation(mutation, 0), exception);
  }

  private static String namedRangeNameFor(
      ExecutorTestPlanSupport.PendingMutation mutation, Exception exception) {
    return namedRangeNameFor(materializeMutation(mutation, 0), exception);
  }

  private static String namedRangeNameFor(WorkbookStep step, Exception exception) {
    return ExecutionDiagnosticFields.namedRangeNameFor(step, exception);
  }

  private static void assertReadContext(
      InspectionStep step,
      String expectedSheetName,
      String expectedRuntimeAddress,
      String expectedNamedRangeName,
      RuntimeException runtimeException) {
    assertEquals(expectedSheetName, sheetNameFor(step));
    assertEquals(expectedRuntimeAddress, addressFor(step, runtimeException));
    assertEquals(expectedNamedRangeName, namedRangeNameFor(step, runtimeException));
  }

  private static void assertWriteContext(
      ExecutorTestPlanSupport.PendingMutation mutation,
      Exception exception,
      String expectedSheetName,
      String expectedAddress,
      String expectedRange,
      String expectedNamedRangeName) {
    MutationStep step = materializeMutation(mutation, 0);
    assertNull(formulaFor(step, exception));
    assertEquals(expectedSheetName, sheetNameFor(step, exception));
    assertEquals(expectedAddress, addressFor(step, exception));
    assertEquals(expectedRange, rangeFor(step, exception));
    assertEquals(expectedNamedRangeName, namedRangeNameFor(step, exception));
  }

  @Test
  void workbookLocationForUsesExistingSourcePathsWhenPersistenceDoesNotSave() {
    Path existingWorkbookPath = Path.of("tmp", "existing-budget.xlsx").toAbsolutePath();

    WorkbookLocation workbookLocation =
        ExecutionRequestPaths.workbookLocationFor(
            new WorkbookPlan.WorkbookSource.ExistingFile(existingWorkbookPath.toString()),
            new WorkbookPlan.WorkbookPersistence.None());

    WorkbookLocation.StoredWorkbook storedWorkbook =
        assertInstanceOf(WorkbookLocation.StoredWorkbook.class, workbookLocation);
    assertEquals(existingWorkbookPath.normalize(), storedWorkbook.workbookPath());
  }

  private static <T extends InspectionResult> T read(
      GridGrindResponse.Success success, String stepId, Class<T> type) {
    return inspection(success, stepId, type);
  }

  private GridGrindResponse.CellStyleReport toResponseStyleReport(
      dev.erst.gridgrind.excel.ExcelCellStyleSnapshot style) {
    return InspectionResultCellReportSupport.toCellStyleReport(style);
  }

  private static ExcelCellStyleSnapshot defaultStyle() {
    return new ExcelCellStyleSnapshot(
        "",
        new ExcelCellAlignmentSnapshot(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new ExcelCellFontSnapshot(
            false,
            false,
            "Aptos",
            ExcelFontHeight.fromPoints(new BigDecimal("11")),
            null,
            false,
            false),
        new ExcelCellFillSnapshot(
            dev.erst.gridgrind.excel.foundation.ExcelFillPattern.NONE, null, null),
        new ExcelBorderSnapshot(
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null)),
        new ExcelCellProtectionSnapshot(true, false));
  }

  private static ExcelPrintSetupSnapshot defaultPrintSetupSnapshot() {
    return new ExcelPrintSetupSnapshot(
        new ExcelPrintMarginsSnapshot(0.0d, 0.0d, 0.0d, 0.0d, 0.0d, 0.0d),
        false,
        false,
        false,
        0,
        false,
        false,
        0,
        false,
        0,
        List.of(),
        List.of());
  }

  private static dev.erst.gridgrind.excel.ExcelSheetPresentationSnapshot
      defaultSheetPresentationSnapshot() {
    return new dev.erst.gridgrind.excel.ExcelSheetPresentationSnapshot(
        ExcelSheetDisplay.defaults(),
        null,
        ExcelSheetOutlineSummary.defaults(),
        ExcelSheetDefaults.defaults(),
        List.of());
  }

  private static SheetProtectionSettings protectionSettings() {
    return new SheetProtectionSettings(
        false, true, false, true, false, true, false, true, false, true, false, true, false, true,
        false);
  }

  private static CellColorReport rgb(String rgb) {
    return new CellColorReport(rgb);
  }

  private static ExcelSheetProtectionSettings excelProtectionSettings() {
    return new ExcelSheetProtectionSettings(
        false, true, false, true, false, true, false, true, false, true, false, true, false, true,
        false);
  }
}
