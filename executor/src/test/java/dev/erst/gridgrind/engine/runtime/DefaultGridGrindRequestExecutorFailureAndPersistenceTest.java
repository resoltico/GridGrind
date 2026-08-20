package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStaticRequestContract;
import dev.erst.gridgrind.contract.step.WorkbookStaticViolation;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.InvalidFormulaException;
import dev.erst.gridgrind.excel.InvalidRangeAddressException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import dev.erst.gridgrind.excel.UnsupportedFormulaException;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.WorkbookLocation;
import dev.erst.gridgrind.excel.WorkbookNotFoundException;
import dev.erst.gridgrind.excel.XlsxRoundTrip;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationType;

/** Failure, persistence, and request-classification tests for DefaultGridGrindRequestExecutor. */
class DefaultGridGrindRequestExecutorFailureAndPersistenceTest
    extends DefaultGridGrindRequestExecutorTestSupport {
  private static List<String> staticValidationMessages(WorkbookPlan request) {
    return WorkbookStaticRequestContract.validate(WorkbookStaticRequestContract.from(request))
        .stream()
        .map(WorkbookStaticViolation::message)
        .toList();
  }

  @Test
  void returnsTableReadResultsForByNameSelectionAndNoneStyle() {
    WorkbookResult response =
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
                        new RangeSelector.ByRange("Budget", "A1:B2"),
                        new CellMutationAction.SetRange(
                            new dev.erst.gridgrind.contract.dto.CellGridInput.Typed(
                                List.of(
                                    List.of(textCell("Item"), textCell("Amount")),
                                    List.of(
                                        textCell("Hosting"), new CellInput.NumberValue(49.0)))))),
                    mutate(
                        new StructuredMutationAction.SetTable(
                            TableInput.withDefaultMetadata(
                                "BudgetTable",
                                "Budget",
                                "A1:B2",
                                false,
                                new TableStyleInput.None())))),
                inspect(
                    "tables",
                    new TableSelector.ByNames(List.of("BudgetTable")),
                    new WorkbookAssetIntrospectionQuery.GetTables())));

    WorkbookAssetInspectionResult.TablesResult tables =
        read(success(response), "tables", WorkbookAssetInspectionResult.TablesResult.class);

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

    WorkbookResult response =
        ExecutionContextFixtureSupport.execute(
            new DefaultGridGrindRequestExecutor(),
            request(
                new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of(),
                inspect(
                    "validations",
                    new RangeSelector.AllOnSheet("Budget"),
                    new SheetIntrospectionQuery.GetDataValidations())));

    SheetInspectionResult.DataValidationsResult validations =
        read(success(response), "validations", SheetInspectionResult.DataValidationsResult.class);
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

    WorkbookResult response =
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
                        new CellMutationAction.SetCell(textCell("Memo"))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "A1"),
                        new CellMutationAction.SetHyperlink(
                            new HyperlinkTarget.File(linkedFile.toString())))),
                inspect(
                    "cells",
                    new CellSelector.ByAddresses("Budget", List.of("A1")),
                    allFacetCellsQuery()),
                inspect(
                    "hyperlinks",
                    new CellSelector.ByAddresses("Budget", List.of("A1")),
                    new SheetIntrospectionQuery.GetHyperlinks())));

    WorkbookResult.Success success = success(response);
    dev.erst.gridgrind.contract.dto.CellReport.TextReport linkedCell =
        cast(
            dev.erst.gridgrind.contract.dto.CellReport.TextReport.class,
            read(success, "cells", SheetInspectionResult.CellsResult.class).cells().getFirst());
    SheetInspectionResult.HyperlinksResult hyperlinks =
        read(success, "hyperlinks", SheetInspectionResult.HyperlinksResult.class);

    assertEquals(
        java.util.Optional.of(new HyperlinkTarget.File(linkedFile.toString())),
        linkedCell.hyperlink());
    assertEquals(
        new HyperlinkTarget.File(linkedFile.toString()),
        hyperlinks.hyperlinks().getFirst().hyperlink());
  }

  @Test
  void convertsEmailAndDocumentHyperlinksToCanonicalProtocolTargets() {
    assertEquals(
        java.util.Optional.of(new HyperlinkTarget.Email("team@example.com")),
        InspectionResultCellReportSupport.toHyperlinkTarget(
            new ExcelHyperlink.Email("team@example.com")));
    assertEquals(
        java.util.Optional.of(new HyperlinkTarget.Document("Budget!B4")),
        InspectionResultCellReportSupport.toHyperlinkTarget(
            new ExcelHyperlink.Document("Budget!B4")));
  }

  @Test
  void returnsFactualAndAnalysisReadResults() {
    WorkbookResult response =
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
                        new SheetSelector.ByName("Budget"),
                        new CellMutationAction.AppendRow(
                            new dev.erst.gridgrind.contract.dto.CellRowInput.Typed(
                                List.of(textCell("Item"), textCell("Amount"))))),
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new CellMutationAction.AppendRow(
                            new dev.erst.gridgrind.contract.dto.CellRowInput.Typed(
                                List.of(textCell("Hosting"), new CellInput.NumberValue(49.0))))),
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new CellMutationAction.AppendRow(
                            new dev.erst.gridgrind.contract.dto.CellRowInput.Typed(
                                List.of(textCell("Domain"), new CellInput.NumberValue(12.0))))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "B4"),
                        new CellMutationAction.SetCell(formulaCell("SUM(B2:B3)"))),
                    mutate(
                        new StructuredMutationAction.SetNamedRange(
                            "BudgetTotal",
                            new NamedRangeScope.Workbook(),
                            NamedRangeTarget.range("Budget", "B4")))),
                inspect(
                    "formula",
                    new SheetSelector.All(),
                    new InspectionSurfaceQuery.GetFormulaSurface()),
                inspect(
                    "schema",
                    new RangeSelector.RectangularWindow("Budget", "A1", 4, 2),
                    new InspectionSurfaceQuery.GetSheetSchema()),
                inspect(
                    "ranges",
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                    new InspectionSurfaceQuery.GetNamedRangeSurface()),
                inspect(
                    "formula-health",
                    new SheetSelector.All(),
                    new InspectionAnalysisQuery.AnalyzeFormulaHealth())));

    WorkbookResult.Success success = success(response);
    FormulaSurfaceReport formula =
        read(success, "formula", WorkbookSurfaceInspectionResult.FormulaSurfaceResult.class)
            .surface();
    SheetSchemaReport schema =
        read(success, "schema", WorkbookSurfaceInspectionResult.SheetSchemaResult.class).surface();
    NamedRangeSurfaceReport ranges =
        read(success, "ranges", WorkbookSurfaceInspectionResult.NamedRangeSurfaceResult.class)
            .surface();
    FormulaHealthReport formulaHealth =
        read(success, "formula-health", WorkbookAnalysisResult.FormulaHealthResult.class)
            .analysis();

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
    assertEquals("TEXT", schema.columns().getFirst().dominantType());

    assertEquals(1, ranges.workbookScopedCount());
    assertEquals(0, ranges.sheetScopedCount());
    assertEquals(1, ranges.rangeBackedCount());
    assertEquals(0, ranges.formulaBackedCount());
    assertEquals("BudgetTotal", ranges.namedRanges().getFirst().name());
    assertEquals(NamedRangeBackingKind.RANGE, ranges.namedRanges().getFirst().kind());
    assertEquals(1, formulaHealth.checkedFormulaCellCount());
    assertEquals(0, formulaHealth.summary().totalCount());
  }

  @Test
  void returnsStructuredFailureWhenMoveSheetTargetsMissingSheet() {
    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Missing"),
                            new WorkbookMutationAction.MoveSheet(0))))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals("MOVE_SHEET", executeStepContext(failure).stepType());
    assertEquals(java.util.Optional.of("Missing"), executeStepContext(failure).sheetName());
  }

  @Test
  void returnsStructuredFailureForConflictingRenameTarget() {
    WorkbookResult.Failure failure =
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
                            new SheetSelector.ByName("Summary"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new WorkbookMutationAction.RenameSheet("Summary"))))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals(2, executeStepContext(failure).stepIndex());
    assertEquals("RENAME_SHEET", executeStepContext(failure).stepType());
    assertEquals(java.util.Optional.of("Budget"), executeStepContext(failure).sheetName());
    assertEquals("Sheet already exists: Summary", failure.problem().message());
  }

  @Test
  void returnsStructuredFailureForInvalidMoveSheetIndex() {
    WorkbookResult.Failure failure =
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
                            new SheetSelector.ByName("Budget"),
                            new WorkbookMutationAction.MoveSheet(1))))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("MOVE_SHEET", executeStepContext(failure).stepType());
    assertEquals(java.util.Optional.of("Budget"), executeStepContext(failure).sheetName());
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
                    new RangeSelector.ByRange("Budget", "A1:"),
                    new WorkbookMutationAction.MergeCells()));

    assertEquals("range must not be blank", failure.getMessage());
  }

  @Test
  void returnsStructuredFailureForUnmergeCellsWithoutExactMatch() {
    WorkbookResult.Failure failure =
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
                            new RangeSelector.ByRange("Budget", "A1:B2"),
                            new WorkbookMutationAction.UnmergeCells())))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("UNMERGE_CELLS", executeStepContext(failure).stepType());
    assertEquals(java.util.Optional.of("Budget"), executeStepContext(failure).sheetName());
    assertEquals(java.util.Optional.of("A1:B2"), executeStepContext(failure).range());
    assertEquals("No merged region matches range: A1:B2", failure.problem().message());
  }

  @Test
  void returnsStructuredFailureWhenSetSheetPaneTargetsMissingSheet() {
    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Missing"),
                            new WorkbookMutationAction.SetSheetPane(
                                new PaneInput.Frozen(1, 1, 1, 1)))))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("SET_SHEET_PANE", executeStepContext(failure).stepType());
    assertEquals(java.util.Optional.of("Missing"), executeStepContext(failure).sheetName());
  }

  @Test
  void returnsStructuredFailureForMissingNamedRangeDuringRead() {
    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new WorkbookMutationAction.EnsureSheet())),
                    inspect(
                        "ranges",
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf(
                            List.of(
                                new dev.erst.gridgrind.contract.selector.NamedRangeSelector
                                    .WorkbookScope("BudgetTotal"))),
                        new WorkbookIntrospectionQuery.GetNamedRanges()))));

    assertEquals(GridGrindProblemCode.NAMED_RANGE_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals(1, executeStepContext(failure).stepIndex());
    assertEquals("GET_NAMED_RANGES", executeStepContext(failure).stepType());
    assertEquals("ranges", executeStepContext(failure).stepId());
    assertEquals(
        java.util.Optional.of("BudgetTotal"), executeStepContext(failure).namedRangeName());
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

    WorkbookResult.Success success =
        success(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.Overwrite(),
                    mutations(
                        mutate(
                            new CellSelector.ByAddress("Budget", "C3"),
                            new CellMutationAction.SetCell(textCell("After")))),
                    inspect(
                        "cells",
                        new CellSelector.ByAddresses("Budget", List.of("C3")),
                        allFacetCellsQuery()))));

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(
        "After",
        cast(
                dev.erst.gridgrind.contract.dto.CellReport.TextReport.class,
                read(success, "cells", SheetInspectionResult.CellsResult.class).cells().getFirst())
            .textValue()
            .orElseThrow());
    assertEquals(List.of("Budget", "Summary"), XlsxRoundTrip.sheetOrder(workbookPath));
    assertEquals(List.of("A1:B2"), XlsxRoundTrip.mergedRegions(workbookPath, "Budget"));
    assertEquals(4096, XlsxRoundTrip.columnWidth(workbookPath, "Budget", 0));
    assertEquals((short) 570, XlsxRoundTrip.rowHeightTwips(workbookPath, "Budget", 0));
    assertEquals(new ExcelSheetPane.Frozen(1, 2, 3, 4), XlsxRoundTrip.pane(workbookPath, "Budget"));
  }

  @Test
  void returnsStructuredFailureForMissingWorkbookSource() throws IOException {
    Path workbookPath = Path.of("missing-workbook.xlsx");
    Path workingDirectory = Files.createTempDirectory("gridgrind-missing-workbook-root-");
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();

    WorkbookResult.Failure failure =
        failure(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of()),
                ExecutionInputBindingsFixtureSupport.bindings(workingDirectory)));

    assertEquals(GridGrindProblemCode.WORKBOOK_NOT_FOUND, failure.problem().code());
    assertEquals("OPEN_WORKBOOK", failure.problem().context().stage());
    assertEquals(GridGrindProblemCategory.RESOURCE, failure.problem().category());
    assertEquals(GridGrindProblemRecovery.CHANGE_REQUEST, failure.problem().recovery());
    assertEquals(
        java.util.Optional.of(workingDirectory.resolve(workbookPath).toAbsolutePath().toString()),
        openWorkbookContext(failure).sourceWorkbookPath());
    ExecutionJournal.Outcome.Failed outcome =
        assertInstanceOf(ExecutionJournal.Outcome.Failed.class, failure.journal().outcome());
    assertEquals(0, outcome.completedStepCount());
    assertTrue(failure.journal().steps().isEmpty());
  }

  @Test
  void returnsStructuredFailureWithOperationContext() {
    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new WorkbookMutationAction.AutoSizeColumns())))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals(0, executeStepContext(failure).stepIndex());
    assertEquals("AUTO_SIZE_COLUMNS", executeStepContext(failure).stepType());
    assertEquals(java.util.Optional.of("Budget"), executeStepContext(failure).sheetName());
  }

  @Test
  void returnsStructuredFailureWithReadContext() {
    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    inspect(
                        "cells",
                        new RangeSelector.RectangularWindow("Budget", "A1", 5, 5),
                        new SheetIntrospectionQuery.GetWindow()))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals("GET_WINDOW", executeStepContext(failure).stepType());
    assertEquals("cells", executeStepContext(failure).stepId());
    assertEquals(java.util.Optional.of("Budget"), executeStepContext(failure).sheetName());
    assertEquals(java.util.Optional.of("A1"), executeStepContext(failure).address());
  }

  @Test
  void returnsStructuredFailureWhenReadTargetsMissingSheet() {
    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    inspect(
                        "sheet",
                        new SheetSelector.ByName("Missing"),
                        new SheetIntrospectionQuery.GetSheetSummary()))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals("GET_SHEET_SUMMARY", executeStepContext(failure).stepType());
    assertEquals("sheet", executeStepContext(failure).stepId());
    assertEquals(java.util.Optional.of("Missing"), executeStepContext(failure).sheetName());
  }

  @Test
  void returnsStructuredFailureWhenDeletingLastSheet() {
    WorkbookResult.Failure failure =
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
                            new SheetSelector.ByName("Budget"),
                            new WorkbookMutationAction.DeleteSheet())))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals(1, executeStepContext(failure).stepIndex());
    assertEquals("DELETE_SHEET", executeStepContext(failure).stepType());
    assertTrue(failure.problem().message().contains("at least one sheet"));
  }

  @Test
  void returnsStructuredFailureWhenDeletingLastVisibleSheet() {
    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Alpha"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new SheetSelector.ByName("Beta"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new SheetSelector.ByName("Beta"),
                            new WorkbookMutationAction.SetSheetVisibility(
                                ExcelSheetVisibility.HIDDEN)),
                        mutate(
                            new SheetSelector.ByName("Alpha"),
                            new WorkbookMutationAction.DeleteSheet())))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals(3, executeStepContext(failure).stepIndex());
    assertEquals("DELETE_SHEET", executeStepContext(failure).stepType());
    assertEquals(java.util.Optional.of("Alpha"), executeStepContext(failure).sheetName());
    assertTrue(failure.problem().message().contains("last visible sheet"));
  }

  @Test
  void returnsStructuredFailureForInvalidPersistTarget() throws IOException {
    Path parentFile = Files.createTempFile("gridgrind-persist-target-", ".tmp");
    Path workbookPath = parentFile.resolve("book.xlsx");

    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(
                        workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new WorkbookMutationAction.EnsureSheet())))));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("PERSIST_WORKBOOK", failure.problem().context().stage());
    WorkbookResultPersistence.PersistenceOutcome.SavedAs persistence =
        assertInstanceOf(
            WorkbookResultPersistence.PersistenceOutcome.SavedAs.class, failure.persistence());
    assertEquals(workbookPath.toString(), persistence.requestedPath());
    assertInstanceOf(WorkbookResultPersistence.WriteResult.NotWritten.class, persistence.write());
    assertEquals(
        java.util.Optional.of(workbookPath.toAbsolutePath().toString()),
        persistWorkbookContext(failure).persistencePath());
  }

  @Test
  void returnsStructuredFailureForSaveAsDestinationCollisions() throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-save-as-collision-");
    Path workbookPath = workspace.resolve("existing-output.xlsx");
    Files.writeString(workbookPath, "occupied");

    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(
                        workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new WorkbookMutationAction.EnsureSheet())))));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("PERSIST_WORKBOOK", failure.problem().context().stage());
    WorkbookResultPersistence.PersistenceOutcome.SavedAs persistence =
        assertInstanceOf(
            WorkbookResultPersistence.PersistenceOutcome.SavedAs.class, failure.persistence());
    assertEquals(workbookPath.toString(), persistence.requestedPath());
    assertInstanceOf(WorkbookResultPersistence.WriteResult.NotWritten.class, persistence.write());
    assertEquals(
        "Could not write workbook to "
            + workbookPath.toAbsolutePath()
            + ": already exists; SAVE_AS.ifExists=REJECT requires a new destination path. Use"
            + " ifExists=REPLACE to allow create-or-replace.",
        failure.problem().message());
  }

  @Test
  void saveAsReplaceCanOverwriteAnExistingDestination() throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-save-as-replace-");
    Path workbookPath = workspace.resolve("existing-output.xlsx");
    Files.writeString(workbookPath, "occupied");

    WorkbookResult.Success success =
        success(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(
                        workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REPLACE),
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new WorkbookMutationAction.EnsureSheet())))));

    WorkbookResultPersistence.PersistenceOutcome.SavedAs persistence =
        assertInstanceOf(
            WorkbookResultPersistence.PersistenceOutcome.SavedAs.class, success.persistence());

    assertEquals(workbookPath.toString(), persistence.requestedPath());
    assertEquals(workbookPath.toAbsolutePath().toString(), writtenExecutionPath(persistence));
    assertTrue(Files.size(workbookPath) > 0L);
    assertDoesNotThrow(
        () -> {
          try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
            assertNotNull(workbook.getSheet("Budget"));
          }
        });
  }

  @Test
  void doesNotPersistWorkbookWhenReadFails() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-analysis-failure-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(
                        workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                    List.of(),
                    inspect(
                        "cells",
                        new CellSelector.ByAddresses("Missing", List.of("A1")),
                        allFacetCellsQuery()))));

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
                    allFacetCellsQuery()));

    assertEquals("addresses[1] must be a single-cell A1-style address", failure.getMessage());
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
                    allFacetCellsQuery()));

    assertEquals("addresses[0] must be a single-cell A1-style address", failure.getMessage());
  }

  @Test
  void returnsStructuredFailureForInvalidOverwriteUsage() {
    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.Overwrite(),
                    List.of())));
    WorkbookResultPersistence.PersistenceOutcome.Overwritten persistence =
        assertInstanceOf(
            WorkbookResultPersistence.PersistenceOutcome.Overwritten.class, failure.persistence());

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("VALIDATE_REQUEST", failure.problem().context().stage());
    assertEquals(java.util.Optional.empty(), persistence.sourcePath());
    assertInstanceOf(WorkbookResultPersistence.WriteResult.NotWritten.class, persistence.write());
  }

  @Test
  void returnsOnlyTheFirstIndependentSemanticValidationProblemDuringExecution() {
    WorkbookResult.Failure failure =
        failure(
            execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.Overwrite(),
                    executionPolicy(ExecutionModeInput.eventRead(), calculateAll()),
                    null,
                    mutations(),
                    inspections())));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("VALIDATE_REQUEST", failure.problem().context().stage());
    assertEquals(
        "execution.mode.type=EVENT_READ requires"
            + " execution.calculation.strategy=DO_NOT_CALCULATE and"
            + " markRecalculateOnOpen=false",
        failure.problem().message());
    assertTrue(
        failure.problem().causes().isEmpty(),
        "execute should report the first validation problem directly instead of batching causes");
    assertFalse(
        failure.problem().message().contains("OVERWRITE persistence requires"),
        "execute must not switch to a later semantic validation problem");
  }

  @Test
  void returnsFormulaErrorForInvalidFormulaOperations() {
    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    executionPolicy(calculateAll()),
                    null,
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Data"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new CellSelector.ByAddress("Data", "A1"),
                            new CellMutationAction.SetCell(formulaCell("SUM(")))),
                    inspections())));

    assertEquals(GridGrindProblemCode.INVALID_FORMULA, failure.problem().code());
    assertEquals(GridGrindProblemCategory.FORMULA, failure.problem().category());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals("Invalid formula at Data!A1: SUM(", failure.problem().message());
    assertEquals(java.util.Optional.of("A1"), executeStepContext(failure).address());
    assertEquals(java.util.Optional.of("SUM("), executeStepContext(failure).formula());
  }

  @Test
  void returnsFormulaErrorForMalformedParserStateFormulaOperations() {
    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    executionPolicy(calculateAll()),
                    null,
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Data"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new CellSelector.ByAddress("Data", "A1"),
                            new CellMutationAction.SetCell(formulaCell("[^owe_e`ffffff")))))));

    assertEquals(GridGrindProblemCode.INVALID_FORMULA, failure.problem().code());
    assertEquals(GridGrindProblemCategory.FORMULA, failure.problem().category());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals("Invalid formula at Data!A1: [^owe_e`ffffff", failure.problem().message());
    assertEquals(java.util.Optional.of("A1"), executeStepContext(failure).address());
    assertEquals(java.util.Optional.of("[^owe_e`ffffff"), executeStepContext(failure).formula());
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

    WorkbookResult.Failure lambdaFailure =
        failure(
            ExecutionContextFixtureSupport.execute(
                executor,
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Data"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new CellSelector.ByAddress("Data", "A1"),
                            new CellMutationAction.SetCell(formulaCell("LAMBDA(x,x*2)(5)")))))));
    WorkbookResult.Failure letFailure =
        failure(
            ExecutionContextFixtureSupport.execute(
                executor,
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Data"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new CellSelector.ByAddress("Data", "A1"),
                            new CellMutationAction.SetCell(formulaCell("LET(x,5,x*2)")))))));

    assertEquals(
        GridGrindProblemCode.UNSUPPORTED_FORMULA_CONSTRUCT, lambdaFailure.problem().code());
    assertEquals(
        "Unsupported formula construct function LAMBDA at Data!A1: LAMBDA(x,x*2)(5)",
        lambdaFailure.problem().message());
    assertEquals(java.util.Optional.of("A1"), executeStepContext(lambdaFailure).address());
    assertEquals(
        java.util.Optional.of("LAMBDA(x,x*2)(5)"), executeStepContext(lambdaFailure).formula());

    assertEquals(GridGrindProblemCode.UNSUPPORTED_FORMULA_CONSTRUCT, letFailure.problem().code());
    assertEquals(
        "Unsupported formula construct function LET at Data!A1: LET(x,5,x*2)",
        letFailure.problem().message());
    assertEquals(java.util.Optional.of("A1"), executeStepContext(letFailure).address());
    assertEquals(java.util.Optional.of("LET(x,5,x*2)"), executeStepContext(letFailure).formula());
  }

  @Test
  void surfacesWorkbookFormulaLocationWhenEvaluationFails() {
    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    executionPolicy(calculateAll()),
                    null,
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Data"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new CellSelector.ByAddress("Data", "A1"),
                            new CellMutationAction.SetCell(new CellInput.NumberValue(1.0))),
                        mutate(
                            new CellSelector.ByAddress("Data", "B1"),
                            new CellMutationAction.SetCell(new CellInput.NumberValue(2.0))),
                        mutate(
                            new CellSelector.ByAddress("Data", "C1"),
                            new CellMutationAction.SetCell(
                                formulaCell("TEXTAFTER(\"a,b\",\",\")")))),
                    inspections())));

    assertEquals(GridGrindProblemCode.UNREGISTERED_USER_DEFINED_FUNCTION, failure.problem().code());
    assertEquals("CALCULATION_PREFLIGHT", failure.problem().context().stage());
    assertEquals(
        "User-defined function TEXTAFTER is not registered at Data!C1: TEXTAFTER(\"a,b\",\",\")",
        failure.problem().message());
    assertEquals(java.util.Optional.of("Data"), calculationPreflightContext(failure).sheetName());
    assertEquals(java.util.Optional.of("C1"), calculationPreflightContext(failure).address());
    assertEquals(
        java.util.Optional.of("TEXTAFTER(\"a,b\",\",\")"),
        calculationPreflightContext(failure).formula());
  }

  @Test
  void calculationPolicyFailureRejectsMutationsAfterObservationSteps() {
    List<String> failures =
        staticValidationMessages(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                executionPolicy(calculateAll()),
                FormulaEnvironmentInput.empty(),
                List.<WorkbookStep>of(
                    new MutationStep(
                        "step-0",
                        new SheetSelector.ByName("Data"),
                        new WorkbookMutationAction.EnsureSheet()),
                    new InspectionStep(
                        "summary",
                        new WorkbookSelector.Current(),
                        new WorkbookIntrospectionQuery.GetWorkbookSummary()),
                    new MutationStep(
                        "step-2",
                        new CellSelector.ByAddress("Data", "A1"),
                        new CellMutationAction.SetCell(new CellInput.NumberValue(1.0))))));

    assertTrue(
        failures.stream()
            .anyMatch(failure -> failure.contains("mutation-to-observation boundary")));
  }

  @Test
  void executionModeFailureRejectsCalculationForEventReadAndStreamingWrite() {
    List<String> eventReadFailures =
        staticValidationMessages(
            request(
                new WorkbookPlan.WorkbookSource.ExistingFile("/tmp/book.xlsx"),
                new WorkbookPlan.WorkbookPersistence.None(),
                executionPolicy(ExecutionModeInput.eventRead(), calculateAll()),
                null,
                List.of(),
                List.of(
                    inspect(
                        "summary",
                        new WorkbookSelector.Current(),
                        new WorkbookIntrospectionQuery.GetWorkbookSummary()))));
    List<String> streamingFailures =
        staticValidationMessages(
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                executionPolicy(ExecutionModeInput.streamingWrite(), calculateAll()),
                null,
                List.of(
                    mutate(
                        new SheetSelector.ByName("Ops"), new WorkbookMutationAction.EnsureSheet())),
                List.of(),
                List.of()));

    assertTrue(
        eventReadFailures.stream()
            .anyMatch(failure -> failure.contains("markRecalculateOnOpen=false")));
    assertTrue(
        streamingFailures.stream()
            .anyMatch(failure -> failure.contains("low-memory streaming writes")));
  }

  @Test
  void returnsCalculationFailureBeforeObservationStepsWhenBoundaryPreflightFails() {
    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    executionPolicy(calculateAll()),
                    null,
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Data"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new CellSelector.ByAddress("Data", "C1"),
                            new CellMutationAction.SetCell(
                                formulaCell("TEXTAFTER(\"a,b\",\",\")")))),
                    inspect(
                        "summary",
                        new WorkbookSelector.Current(),
                        new WorkbookIntrospectionQuery.GetWorkbookSummary()))));

    assertEquals(GridGrindProblemCode.UNREGISTERED_USER_DEFINED_FUNCTION, failure.problem().code());
    assertEquals("CALCULATION_PREFLIGHT", failure.problem().context().stage());
    assertEquals(java.util.Optional.of("Data"), calculationPreflightContext(failure).sheetName());
    assertEquals(java.util.Optional.of("C1"), calculationPreflightContext(failure).address());
    assertEquals(
        java.util.Optional.of("TEXTAFTER(\"a,b\",\",\")"),
        calculationPreflightContext(failure).formula());
  }

  @Test
  void returnsStructuredFailureWhenWorkbookCloseFailsAfterSuccess() {
    DefaultGridGrindRequestExecutor executor =
        new DefaultGridGrindRequestExecutor(
            new DefaultGridGrindRequestExecutorDependencies(
                new WorkbookExecutionEngine(),
                workbook -> {
                  throw new IOException("close failed");
                },
                dev.erst.gridgrind.excel.WorkbookArtifactIo.MaterializedWorkbook::close,
                dev.erst.gridgrind.excel.stream.ExcelStreamingWorkbookWriter
                    ::markRecalculateOnOpen));

    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                executor,
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new WorkbookMutationAction.EnsureSheet())))));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("EXECUTE_REQUEST", failure.problem().context().stage());
    assertEquals("close failed", failure.problem().message());
    assertEquals(GridGrindProblemRecovery.CHECK_ENVIRONMENT, failure.problem().recovery());
  }

  @Test
  void preservesPrimaryFailureWhenWorkbookCloseAlsoFails() {
    DefaultGridGrindRequestExecutor executor =
        new DefaultGridGrindRequestExecutor(
            new DefaultGridGrindRequestExecutorDependencies(
                new WorkbookExecutionEngine(),
                workbook -> {
                  throw new IOException("close failed");
                },
                dev.erst.gridgrind.excel.WorkbookArtifactIo.MaterializedWorkbook::close,
                dev.erst.gridgrind.excel.stream.ExcelStreamingWorkbookWriter
                    ::markRecalculateOnOpen));

    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                executor,
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new WorkbookMutationAction.AutoSizeColumns())))));

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
  void rejectsNullRequestsAtTheBoundary() {
    NullPointerException failure =
        assertThrows(
            NullPointerException.class,
            () ->
                ExecutionContextFixtureSupport.execute(
                    new DefaultGridGrindRequestExecutor(), null));

    assertEquals("request must not be null", failure.getMessage());
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
  void workbookLocationForUsesExistingSourcePathsWhenPersistenceDoesNotSave() {
    Path existingWorkbookPath = Path.of("tmp", "existing-budget.xlsx").toAbsolutePath();
    Path workingDirectory = Path.of("/tmp/gridgrind-workbook-location");

    WorkbookLocation workbookLocation =
        ExecutionRequestPaths.workbookLocationFor(
            new WorkbookPlan.WorkbookSource.ExistingFile(existingWorkbookPath.toString()),
            new WorkbookPlan.WorkbookPersistence.None(),
            workingDirectory);

    WorkbookLocation.StoredWorkbook storedWorkbook =
        assertInstanceOf(WorkbookLocation.StoredWorkbook.class, workbookLocation);
    assertEquals(existingWorkbookPath.normalize(), storedWorkbook.workbookPath());
  }
}
