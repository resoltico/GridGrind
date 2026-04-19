package dev.erst.gridgrind.authoring;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.AnalysisFindingCode;
import dev.erst.gridgrind.contract.dto.AnalysisSeverity;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CellAlignmentReport;
import dev.erst.gridgrind.contract.dto.CellBorderReport;
import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellColorReport;
import dev.erst.gridgrind.contract.dto.CellFillReport;
import dev.erst.gridgrind.contract.dto.CellFontReport;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.DrawingMarkerInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.SheetPresentationInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.ExcelFillPattern;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.excel.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.executor.ExecutionInputBindings;
import dev.erst.gridgrind.executor.ExecutionJournalSink;
import dev.erst.gridgrind.executor.GridGrindRequestExecutor;
import java.io.ByteArrayOutputStream;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.atomic.AtomicReference;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Broad coverage lock-in for the Phase 8 Java authoring API surface. */
class AuthoringApiCoverageTest {
  @TempDir Path tempDir;

  @Test
  void coversChecksQueriesValuesAndPlannedSteps() {
    GridGrindPlan unnamedPlan = GridGrindPlan.newWorkbook();
    assertSame(unnamedPlan, unnamedPlan.planId(null));
    assertChecks();
    assertQueries();
    assertValuesAndSelectorTarget();
    assertPlannedStepBuilders();
  }

  @Test
  void coversTargetFactoriesAcrossWorkbookEntities() {
    assertSelectorFactories();
    assertWorkbookAndSheetTargets();
    assertCellRangeAndWindowTargets();
    assertTableTargets();
    assertNamedRangeTargets();
    assertChartAndPivotTargets();
  }

  @Test
  void targetRecordsRejectNullSelectorsImmediately() {
    assertThrows(NullPointerException.class, () -> new Targets.WorkbookRef(null));
    assertThrows(NullPointerException.class, () -> new Targets.SheetRef(null));
    assertThrows(NullPointerException.class, () -> new Targets.CellRef(null));
    assertThrows(NullPointerException.class, () -> new Targets.RangeRef(null));
    assertThrows(NullPointerException.class, () -> new Targets.WindowRef(null));
    assertThrows(NullPointerException.class, () -> new Targets.TableRef(null));
    assertThrows(NullPointerException.class, () -> new Targets.TableRowRef(null));
    assertThrows(NullPointerException.class, () -> new Targets.TableCellRef(null));
    assertThrows(NullPointerException.class, () -> new Targets.NamedRangeRef(null));
    assertThrows(NullPointerException.class, () -> new Targets.ChartRef(null));
    assertThrows(NullPointerException.class, () -> new Targets.PivotTableRef(null));
  }

  @Test
  void coversGridGrindPlanLifecycleAndRunOverloads() throws Exception {
    assertThrows(IllegalArgumentException.class, () -> GridGrindPlan.newWorkbook().planId(" "));
    assertThrows(NullPointerException.class, () -> GridGrindPlan.open(null));
    assertThrows(NullPointerException.class, () -> GridGrindPlan.from(null));

    GridGrindPlan plan =
        GridGrindPlan.open(tempDir.resolve("source.xlsx"), new OoxmlOpenSecurityInput("secret"))
            .planId("coverage-plan")
            .persistence(null)
            .saveAs(
                tempDir.resolve("copy.xlsx"),
                new OoxmlPersistenceSecurityInput(
                    new dev.erst.gridgrind.contract.dto.OoxmlEncryptionInput("secret", null), null))
            .overwriteSource(
                new OoxmlPersistenceSecurityInput(
                    new dev.erst.gridgrind.contract.dto.OoxmlEncryptionInput("secret", null), null))
            .inMemoryOnly()
            .execution(
                new ExecutionPolicyInput(
                    new ExecutionModeInput(null, null),
                    new dev.erst.gridgrind.contract.dto.ExecutionJournalInput(
                        ExecutionJournalLevel.NORMAL),
                    null))
            .mode(new ExecutionModeInput(null, null))
            .journal(ExecutionJournalLevel.VERBOSE)
            .calculation(
                new CalculationPolicyInput(
                    new dev.erst.gridgrind.contract.dto.CalculationStrategyInput.DoNotCalculate(),
                    true))
            .formulaEnvironment(new FormulaEnvironmentInput(List.of(), null, List.of()))
            .addStep(
                new PlannedMutation(
                        Targets.sheet("Budget").selector(), new MutationAction.EnsureSheet())
                    .toStep("manual"))
            .mutate(Targets.sheet("Budget"), new MutationAction.EnsureSheet())
            .inspect(Targets.workbook(), Queries.workbookSummary())
            .assertThat(
                Targets.workbook(),
                Checks.workbookProtection(
                    new WorkbookProtectionReport(false, false, false, false, false)));

    WorkbookPlan canonical = plan.toPlan();
    assertEquals("coverage-plan", canonical.planId());
    assertEquals(4, canonical.steps().size());
    assertTrue(
        new String(plan.toJsonBytes(), StandardCharsets.UTF_8).contains("\"coverage-plan\""));
    assertTrue(plan.toJsonString().contains("\"coverage-plan\""));
    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    plan.writeJson(outputStream);
    assertTrue(outputStream.toString(StandardCharsets.UTF_8).contains("\"coverage-plan\""));

    AtomicReference<WorkbookPlan> seenPlan = new AtomicReference<>();
    AtomicReference<ExecutionInputBindings> seenBindings = new AtomicReference<>();
    AtomicReference<ExecutionJournalSink> seenSink = new AtomicReference<>();
    GridGrindRequestExecutor executor =
        (request, bindings, sink) -> {
          seenPlan.set(request);
          seenBindings.set(bindings);
          seenSink.set(sink);
          return new GridGrindResponse.Success(
              GridGrindProtocolVersion.current(),
              new GridGrindResponse.PersistenceOutcome.NotSaved(),
              List.of(),
              List.of(),
              List.of());
        };

    GridGrindResponse runWithNullBindings = plan.run(executor, null, ExecutionJournalSink.NOOP);
    assertInstanceOf(GridGrindResponse.Success.class, runWithNullBindings);
    assertEquals(canonical, seenPlan.get());
    assertNotNull(seenBindings.get());
    assertEquals(Path.of("").toAbsolutePath().normalize(), seenBindings.get().workingDirectory());
    assertSame(ExecutionJournalSink.NOOP, seenSink.get());

    ExecutionInputBindings explicitBindings =
        new ExecutionInputBindings(tempDir, "stdin".getBytes(StandardCharsets.UTF_8));
    GridGrindResponse runWithBindings = plan.run(executor, explicitBindings);
    assertInstanceOf(GridGrindResponse.Success.class, runWithBindings);
    assertSame(explicitBindings, seenBindings.get());

    GridGrindResponse runWithExecutorOnly = plan.run(executor);
    assertInstanceOf(GridGrindResponse.Success.class, runWithExecutorOnly);
    assertSame(ExecutionJournalSink.NOOP, seenSink.get());

    GridGrindResponse defaultRun = GridGrindPlan.newWorkbook().planId("default-run").run();
    assertInstanceOf(GridGrindResponse.Success.class, defaultRun);

    WorkbookPlan imported = GridGrindPlan.from(canonical).toPlan();
    assertEquals(canonical, imported);
  }

  @Test
  void coversGridGrindPlanConvenienceOverloads() {
    WorkbookPlan opened = GridGrindPlan.open(tempDir.resolve("source.xlsx")).toPlan();
    assertInstanceOf(WorkbookPlan.WorkbookSource.ExistingFile.class, opened.source());

    WorkbookPlan saved = GridGrindPlan.newWorkbook().saveAs(tempDir.resolve("copy.xlsx")).toPlan();
    assertInstanceOf(WorkbookPlan.WorkbookPersistence.SaveAs.class, saved.persistence());

    WorkbookPlan overwritten =
        GridGrindPlan.open(tempDir.resolve("source.xlsx")).overwriteSource().toPlan();
    assertInstanceOf(
        WorkbookPlan.WorkbookPersistence.OverwriteSource.class, overwritten.persistence());

    WorkbookPlan modeOnly =
        GridGrindPlan.newWorkbook()
            .mode(new ExecutionModeInput(ExecutionModeInput.ReadMode.EVENT_READ, null))
            .toPlan();
    assertEquals(ExecutionModeInput.ReadMode.EVENT_READ, modeOnly.execution().mode().readMode());

    WorkbookPlan journalOnly =
        GridGrindPlan.newWorkbook().journal(ExecutionJournalLevel.VERBOSE).toPlan();
    assertEquals(ExecutionJournalLevel.VERBOSE, journalOnly.execution().journal().level());

    WorkbookPlan calculationOnly =
        GridGrindPlan.newWorkbook().calculation(new CalculationPolicyInput(null, true)).toPlan();
    assertTrue(calculationOnly.execution().calculation().markRecalculateOnOpen());
  }

  private static void assertChecks() {
    assertInstanceOf(Assertion.Present.class, Checks.present());
    assertInstanceOf(Assertion.Absent.class, Checks.absent());
    assertInstanceOf(Assertion.CellValue.class, Checks.cellValue(Values.expectedText("Owner")));
    assertInstanceOf(Assertion.DisplayValue.class, Checks.displayValue("Owner"));
    assertInstanceOf(Assertion.FormulaText.class, Checks.formulaText("SUM(A1:A2)"));
    assertInstanceOf(Assertion.CellStyle.class, Checks.cellStyle(sampleStyle()));
    assertInstanceOf(
        Assertion.WorkbookProtectionFacts.class,
        Checks.workbookProtection(new WorkbookProtectionReport(false, false, false, false, false)));
    assertInstanceOf(
        Assertion.SheetStructureFacts.class,
        Checks.sheetStructure(
            new GridGrindResponse.SheetSummaryReport(
                "Budget",
                ExcelSheetVisibility.VISIBLE,
                new GridGrindResponse.SheetProtectionReport.Unprotected(),
                2,
                1,
                1)));
    assertInstanceOf(Assertion.NamedRangeFacts.class, Checks.namedRanges(List.of()));
    assertInstanceOf(Assertion.TableFacts.class, Checks.tables(List.of()));
    assertInstanceOf(Assertion.PivotTableFacts.class, Checks.pivotTables(List.of()));
    assertInstanceOf(Assertion.ChartFacts.class, Checks.charts(List.of()));
    assertInstanceOf(
        Assertion.AnalysisMaxSeverity.class,
        Checks.analysisMaxSeverity(Queries.formulaHealth(), AnalysisSeverity.WARNING));
    assertInstanceOf(
        Assertion.AnalysisFindingPresent.class,
        Checks.analysisFindingPresent(
            Queries.formulaHealth(),
            AnalysisFindingCode.FORMULA_ERROR_RESULT,
            AnalysisSeverity.ERROR,
            "DIV/0"));
    assertInstanceOf(
        Assertion.AnalysisFindingAbsent.class,
        Checks.analysisFindingAbsent(
            Queries.formulaHealth(),
            AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
            AnalysisSeverity.INFO,
            null));
    assertInstanceOf(Assertion.AllOf.class, Checks.allOf(Checks.present(), Checks.absent()));
    assertInstanceOf(Assertion.AnyOf.class, Checks.anyOf(Checks.present(), Checks.absent()));
    assertInstanceOf(Assertion.Not.class, Checks.not(Checks.present()));
  }

  private static void assertQueries() {
    assertEquals("GET_WORKBOOK_SUMMARY", Queries.workbookSummary().queryType());
    assertEquals("GET_PACKAGE_SECURITY", Queries.packageSecurity().queryType());
    assertEquals("GET_WORKBOOK_PROTECTION", Queries.workbookProtection().queryType());
    assertEquals("GET_NAMED_RANGES", Queries.namedRanges().queryType());
    assertEquals("GET_SHEET_SUMMARY", Queries.sheetSummary().queryType());
    assertEquals("GET_CELLS", Queries.cells().queryType());
    assertEquals("GET_WINDOW", Queries.window().queryType());
    assertEquals("GET_MERGED_REGIONS", Queries.mergedRegions().queryType());
    assertEquals("GET_HYPERLINKS", Queries.hyperlinks().queryType());
    assertEquals("GET_COMMENTS", Queries.comments().queryType());
    assertEquals("GET_DRAWING_OBJECTS", Queries.drawingObjects().queryType());
    assertEquals("GET_CHARTS", Queries.charts().queryType());
    assertEquals("GET_PIVOT_TABLES", Queries.pivotTables().queryType());
    assertEquals("GET_DRAWING_OBJECT_PAYLOAD", Queries.drawingObjectPayload().queryType());
    assertEquals("GET_SHEET_LAYOUT", Queries.sheetLayout().queryType());
    assertEquals("GET_PRINT_LAYOUT", Queries.printLayout().queryType());
    assertEquals("GET_DATA_VALIDATIONS", Queries.dataValidations().queryType());
    assertEquals("GET_CONDITIONAL_FORMATTING", Queries.conditionalFormatting().queryType());
    assertEquals("GET_AUTOFILTERS", Queries.autofilters().queryType());
    assertEquals("GET_TABLES", Queries.tables().queryType());
    assertEquals("GET_FORMULA_SURFACE", Queries.formulaSurface().queryType());
    assertEquals("GET_SHEET_SCHEMA", Queries.sheetSchema().queryType());
    assertEquals("GET_NAMED_RANGE_SURFACE", Queries.namedRangeSurface().queryType());
    assertEquals("ANALYZE_FORMULA_HEALTH", Queries.formulaHealth().queryType());
    assertEquals("ANALYZE_DATA_VALIDATION_HEALTH", Queries.dataValidationHealth().queryType());
    assertEquals(
        "ANALYZE_CONDITIONAL_FORMATTING_HEALTH", Queries.conditionalFormattingHealth().queryType());
    assertEquals("ANALYZE_AUTOFILTER_HEALTH", Queries.autofilterHealth().queryType());
    assertEquals("ANALYZE_TABLE_HEALTH", Queries.tableHealth().queryType());
    assertEquals("ANALYZE_PIVOT_TABLE_HEALTH", Queries.pivotTableHealth().queryType());
    assertEquals("ANALYZE_HYPERLINK_HEALTH", Queries.hyperlinkHealth().queryType());
    assertEquals("ANALYZE_NAMED_RANGE_HEALTH", Queries.namedRangeHealth().queryType());
    assertEquals("ANALYZE_WORKBOOK_FINDINGS", Queries.workbookFindings().queryType());
  }

  private static void assertValuesAndSelectorTarget() {
    assertInstanceOf(CellInput.Blank.class, Values.blank());
    assertInstanceOf(CellInput.Text.class, Values.text("Owner"));
    assertInstanceOf(CellInput.Text.class, Values.text(Values.inlineText("Owner")));
    assertInstanceOf(CellInput.Text.class, Values.textFile(Path.of("note.txt")));
    assertInstanceOf(CellInput.Text.class, Values.textFromStandardInput());
    assertInstanceOf(
        CellInput.RichText.class,
        Values.richText(
            new dev.erst.gridgrind.contract.dto.RichTextRunInput(Values.inlineText("A"), null)));
    assertInstanceOf(CellInput.Numeric.class, Values.number(1.0));
    assertInstanceOf(CellInput.BooleanValue.class, Values.bool(true));
    assertInstanceOf(CellInput.Date.class, Values.date(java.time.LocalDate.of(2026, 4, 18)));
    assertInstanceOf(
        CellInput.DateTime.class, Values.dateTime(java.time.LocalDateTime.of(2026, 4, 18, 12, 0)));
    assertInstanceOf(CellInput.Formula.class, Values.formula("SUM(A1:A2)"));
    assertInstanceOf(CellInput.Formula.class, Values.formulaFile(Path.of("formula.txt")));
    assertInstanceOf(CellInput.Formula.class, Values.formulaFromStandardInput());
    assertInstanceOf(CommentInput.class, Values.comment("Hi", "Ada"));
    assertInstanceOf(CommentInput.class, Values.comment(Values.inlineText("Hi"), "Ada", true));
    assertInstanceOf(TextSourceInput.Inline.class, Values.inlineText("Hi"));
    assertInstanceOf(TextSourceInput.Utf8File.class, Values.textSourceFile(Path.of("note.txt")));
    assertInstanceOf(TextSourceInput.StandardInput.class, Values.textSourceFromStandardInput());
    assertInstanceOf(BinarySourceInput.InlineBase64.class, Values.inlineBase64("SGVsbG8="));
    assertInstanceOf(BinarySourceInput.File.class, Values.binaryFile(Path.of("payload.bin")));
    assertInstanceOf(BinarySourceInput.StandardInput.class, Values.binaryFromStandardInput());
    assertInstanceOf(ExpectedCellValue.Blank.class, Values.expectedBlank());
    assertInstanceOf(ExpectedCellValue.Text.class, Values.expectedText("Owner"));
    assertInstanceOf(ExpectedCellValue.NumericValue.class, Values.expectedNumber(1.0));
    assertInstanceOf(ExpectedCellValue.BooleanValue.class, Values.expectedBoolean(true));
    assertInstanceOf(ExpectedCellValue.ErrorValue.class, Values.expectedError("#DIV/0!"));
    assertEquals(2, Values.row(Values.text("A"), Values.number(1.0)).size());

    SelectorTarget selectorTarget = () -> new WorkbookSelector.Current();
    assertInstanceOf(WorkbookSelector.Current.class, selectorTarget.selector());
  }

  private static void assertPlannedStepBuilders() {
    PlannedMutation plannedMutation =
        new PlannedMutation(new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet());
    assertEquals("named-mutation", plannedMutation.named("named-mutation").stepId());
    assertEquals(
        "named-mutation",
        plannedMutation.named("named-mutation").toStep("generated-mutation").stepId());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PlannedMutation(
                " ", new CellSelector.ByAddress("Budget", "A1"), new MutationAction.EnsureSheet()));
    assertThrows(
        NullPointerException.class,
        () -> new PlannedMutation((Selector) null, new MutationAction.EnsureSheet()));

    PlannedInspection plannedInspection =
        new PlannedInspection(new WorkbookSelector.Current(), Queries.workbookSummary());
    assertEquals("named-inspection", plannedInspection.named("named-inspection").stepId());
    assertEquals(
        "named-inspection",
        plannedInspection.named("named-inspection").toStep("generated-inspection").stepId());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PlannedInspection(" ", new WorkbookSelector.Current(), Queries.workbookSummary()));
    assertThrows(
        NullPointerException.class,
        () -> new PlannedInspection(new WorkbookSelector.Current(), null));

    PlannedAssertion plannedAssertion =
        new PlannedAssertion(
            new WorkbookSelector.Current(),
            Checks.workbookProtection(
                new WorkbookProtectionReport(false, false, false, false, false)));
    assertEquals("named-assertion", plannedAssertion.named("named-assertion").stepId());
    assertEquals(
        "named-assertion",
        plannedAssertion.named("named-assertion").toStep("generated-assertion").stepId());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PlannedAssertion(
                " ",
                new WorkbookSelector.Current(),
                Checks.workbookProtection(
                    new WorkbookProtectionReport(false, false, false, false, false))));
    assertThrows(
        NullPointerException.class,
        () -> new PlannedAssertion(new WorkbookSelector.Current(), null));
  }

  private static void assertSelectorFactories() {
    assertInstanceOf(Targets.WorkbookRef.class, Targets.workbook());
    assertInstanceOf(SheetSelector.All.class, Targets.allSheets());
    assertEquals(2, Targets.sheets("A", "B").names().size());
    assertInstanceOf(Targets.CellRef.class, Targets.cell("Budget", "A1"));
    assertEquals(2, Targets.cells("Budget", "A1", "B2").addresses().size());
    assertEquals("Budget", Targets.allUsedCells("Budget").sheetName());
    assertEquals(1, Targets.qualifiedCells(Targets.qualifiedCell("Budget", "A1")).cells().size());
    assertEquals("A1", Targets.qualifiedCell("Budget", "A1").address());
    assertInstanceOf(Targets.RangeRef.class, Targets.range("Budget", "A1:B2"));
    assertEquals(2, Targets.ranges("Budget", "A1:B2", "C1:D2").ranges().size());
    assertEquals("Budget", Targets.allRanges("Budget").sheetName());
    assertEquals("A1", Targets.window("Budget", "A1", 2, 2).selector().topLeftAddress());
    assertEquals(0, Targets.rows("Budget", 0, 2).firstRowIndex());
    assertEquals(2, Targets.insertRowsBefore("Budget", 3, 2).rowCount());
    assertEquals(1, Targets.columns("Budget", 1, 3).firstColumnIndex());
    assertEquals(4, Targets.insertColumnsBefore("Budget", 4, 2).beforeColumnIndex());
    assertInstanceOf(Targets.TableRef.class, Targets.table("BudgetTable"));
    assertInstanceOf(Targets.TableRef.class, Targets.tableOnSheet("BudgetTable", "Budget"));
    assertInstanceOf(TableSelector.All.class, Targets.allTables());
    assertEquals(2, Targets.tables("A", "B").names().size());
    assertInstanceOf(Targets.NamedRangeRef.class, Targets.namedRange("Total"));
    assertInstanceOf(Targets.NamedRangeRef.class, Targets.namedRangeOnSheet("Local", "Budget"));
    assertInstanceOf(Targets.NamedRangeRef.class, Targets.workbookNamedRange("Global"));
    assertInstanceOf(NamedRangeSelector.All.class, Targets.allNamedRanges());
    assertEquals(2, Targets.namedRanges("A", "B").names().size());
    assertEquals(
        1,
        Targets.anyNamedRange(new NamedRangeSelector.WorkbookScope("Global")).selectors().size());
    assertInstanceOf(Targets.ChartRef.class, Targets.chart("Budget", "Trend"));
    assertEquals("Budget", Targets.chartsOnSheet("Budget").sheetName());
    assertEquals("Logo", Targets.drawingObject("Budget", "Logo").objectName());
    assertEquals("Budget", Targets.drawingObjectsOnSheet("Budget").sheetName());
    assertInstanceOf(Targets.PivotTableRef.class, Targets.pivotTable("Pivot"));
    assertInstanceOf(Targets.PivotTableRef.class, Targets.pivotTableOnSheet("Pivot", "Budget"));
    assertInstanceOf(PivotTableSelector.All.class, Targets.allPivotTables());
    assertEquals(2, Targets.pivotTables("A", "B").names().size());
  }

  private static void assertWorkbookAndSheetTargets() {
    assertEquals("GET_WORKBOOK_SUMMARY", Targets.workbook().summary().query().queryType());
    assertEquals("GET_PACKAGE_SECURITY", Targets.workbook().packageSecurity().query().queryType());
    assertEquals("GET_WORKBOOK_PROTECTION", Targets.workbook().protection().query().queryType());
    assertEquals("ANALYZE_WORKBOOK_FINDINGS", Targets.workbook().findings().query().queryType());
    assertEquals(
        "SET_WORKBOOK_PROTECTION",
        Targets.workbook()
            .protect(new WorkbookProtectionInput(true, true, true, "secret", "revisions"))
            .action()
            .actionType());
    assertEquals(
        "CLEAR_WORKBOOK_PROTECTION", Targets.workbook().clearProtection().action().actionType());

    Targets.SheetRef sheet = Targets.sheet("Budget");
    assertEquals("ENSURE_SHEET", sheet.ensureExists().action().actionType());
    assertEquals("RENAME_SHEET", sheet.renameTo("Renamed").action().actionType());
    assertEquals("DELETE_SHEET", sheet.delete().action().actionType());
    assertEquals("SET_SHEET_ZOOM", sheet.setZoom(125).action().actionType());
    assertEquals(
        "SET_SHEET_PRESENTATION",
        sheet.setPresentation(SheetPresentationInput.defaults()).action().actionType());
    assertEquals(
        "SET_PRINT_LAYOUT",
        sheet
            .setPrintLayout(new PrintLayoutInput(null, null, null, null, null, null, null))
            .action()
            .actionType());
    assertEquals("CLEAR_PRINT_LAYOUT", sheet.clearPrintLayout().action().actionType());
    assertEquals("GET_SHEET_SUMMARY", sheet.summary().query().queryType());
    assertEquals("GET_SHEET_LAYOUT", sheet.layout().query().queryType());
    assertEquals("GET_PRINT_LAYOUT", sheet.printLayout().query().queryType());
    assertEquals("GET_MERGED_REGIONS", sheet.mergedRegions().query().queryType());
    assertEquals("GET_AUTOFILTERS", sheet.autofilters().query().queryType());
    assertEquals("GET_CHARTS", sheet.charts().query().queryType());
    assertEquals("GET_DRAWING_OBJECTS", sheet.drawingObjects().query().queryType());
    assertEquals("GET_FORMULA_SURFACE", sheet.formulaSurface().query().queryType());
    assertEquals("ANALYZE_FORMULA_HEALTH", sheet.formulaHealth().query().queryType());
  }

  private static void assertCellRangeAndWindowTargets() {
    Targets.CellRef cell = Targets.cell("Budget", "A1");
    assertEquals("SET_CELL", cell.set(Values.text("Owner")).action().actionType());
    assertEquals(
        "SET_HYPERLINK",
        cell.setHyperlink(new HyperlinkTarget.Url("https://example.com")).action().actionType());
    assertEquals("CLEAR_HYPERLINK", cell.clearHyperlink().action().actionType());
    assertEquals("SET_COMMENT", cell.setComment(Values.comment("Hi", "Ada")).action().actionType());
    assertEquals("CLEAR_COMMENT", cell.clearComment().action().actionType());
    assertEquals("GET_CELLS", cell.read().query().queryType());
    assertEquals("GET_HYPERLINKS", cell.hyperlinks().query().queryType());
    assertEquals("GET_COMMENTS", cell.comments().query().queryType());
    assertEquals(
        "EXPECT_CELL_VALUE",
        cell.valueEquals(Values.expectedText("Owner")).assertion().assertionType());
    assertEquals(
        "EXPECT_DISPLAY_VALUE", cell.displayValueEquals("Owner").assertion().assertionType());
    assertEquals(
        "EXPECT_FORMULA_TEXT", cell.formulaEquals("SUM(A1:A2)").assertion().assertionType());
    assertEquals("EXPECT_CELL_STYLE", cell.styleEquals(sampleStyle()).assertion().assertionType());

    Targets.RangeRef range = Targets.range("Budget", "A1:B2");
    assertEquals(
        "SET_RANGE", range.setRows(List.of(Values.row(Values.text("A")))).action().actionType());
    assertEquals("CLEAR_RANGE", range.clear().action().actionType());
    assertEquals("MERGE_CELLS", range.merge().action().actionType());
    assertEquals("UNMERGE_CELLS", range.unmerge().action().actionType());
    assertEquals(
        "APPLY_STYLE",
        range
            .applyStyle(
                new dev.erst.gridgrind.contract.dto.CellStyleInput(
                    "0.00", null, null, null, null, null))
            .action()
            .actionType());
    assertEquals(
        "SET_DATA_VALIDATION",
        range.setDataValidation(sampleDataValidation()).action().actionType());
    assertEquals("GET_DATA_VALIDATIONS", range.dataValidations().query().queryType());
    assertEquals("GET_CONDITIONAL_FORMATTING", range.conditionalFormatting().query().queryType());

    Targets.WindowRef window = Targets.window("Budget", "A1", 2, 2);
    assertEquals("GET_WINDOW", window.read().query().queryType());
    assertEquals("GET_SHEET_SCHEMA", window.schema().query().queryType());
  }

  private static void assertTableTargets() {
    Targets.TableRef table = Targets.table("BudgetTable");
    assertInstanceOf(TableRowSelector.ByIndex.class, table.row(1).selector());
    assertInstanceOf(
        TableRowSelector.ByKeyCell.class,
        table.rowByKey("Item", Values.text("Hosting")).selector());
    assertEquals("SET_TABLE", table.define(sampleTable()).action().actionType());
    assertEquals("DELETE_TABLE", table.delete().action().actionType());
    assertEquals("GET_TABLES", table.inspect().query().queryType());
    assertEquals("ANALYZE_TABLE_HEALTH", table.analyzeHealth().query().queryType());
    assertEquals("EXPECT_PRESENT", table.present().assertion().assertionType());
    assertEquals("EXPECT_ABSENT", table.absent().assertion().assertionType());

    Targets.TableCellRef tableCell = table.rowByKey("Item", Values.text("Hosting")).cell("Amount");
    assertEquals("SET_CELL", tableCell.set(Values.number(1.0)).action().actionType());
    assertEquals(
        "SET_HYPERLINK",
        tableCell
            .setHyperlink(new HyperlinkTarget.Url("https://example.com"))
            .action()
            .actionType());
    assertEquals("CLEAR_HYPERLINK", tableCell.clearHyperlink().action().actionType());
    assertEquals(
        "SET_COMMENT", tableCell.setComment(Values.comment("Hi", "Ada")).action().actionType());
    assertEquals("CLEAR_COMMENT", tableCell.clearComment().action().actionType());
    assertEquals("GET_CELLS", tableCell.read().query().queryType());
    assertEquals("GET_HYPERLINKS", tableCell.hyperlinks().query().queryType());
    assertEquals("GET_COMMENTS", tableCell.comments().query().queryType());
    assertEquals(
        "EXPECT_CELL_VALUE",
        tableCell.valueEquals(Values.expectedNumber(1.0)).assertion().assertionType());
    assertEquals(
        "EXPECT_DISPLAY_VALUE", tableCell.displayValueEquals("1").assertion().assertionType());
    assertEquals(
        "EXPECT_FORMULA_TEXT", tableCell.formulaEquals("SUM(A1:A2)").assertion().assertionType());
    assertEquals(
        "EXPECT_CELL_STYLE", tableCell.styleEquals(sampleStyle()).assertion().assertionType());
  }

  private static void assertNamedRangeTargets() {
    Targets.NamedRangeRef namedRange = Targets.namedRange("Total");
    assertEquals(
        "SET_NAMED_RANGE",
        namedRange
            .define(new NamedRangeScope.Workbook(), new NamedRangeTarget("Budget", "A1"))
            .action()
            .actionType());
    assertEquals(
        "SET_NAMED_RANGE",
        new Targets.NamedRangeRef(new NamedRangeSelector.ByName("Another"))
            .define(new NamedRangeScope.Workbook(), new NamedRangeTarget("Budget", "A1"))
            .action()
            .actionType());
    assertEquals(
        "SET_NAMED_RANGE",
        Targets.workbookNamedRange("Global")
            .define(new NamedRangeScope.Workbook(), new NamedRangeTarget("Budget", "A1"))
            .action()
            .actionType());
    assertEquals(
        "SET_NAMED_RANGE",
        new Targets.NamedRangeRef(new NamedRangeSelector.WorkbookScope("WorkbookScoped"))
            .define(new NamedRangeScope.Workbook(), new NamedRangeTarget("Budget", "A1"))
            .action()
            .actionType());
    assertEquals(
        "SET_NAMED_RANGE",
        Targets.namedRangeOnSheet("Local", "Budget")
            .define(new NamedRangeScope.Sheet("Budget"), new NamedRangeTarget("Budget", "A1"))
            .action()
            .actionType());
    assertEquals(
        "SET_NAMED_RANGE",
        new Targets.NamedRangeRef(new NamedRangeSelector.SheetScope("LocalDirect", "Budget"))
            .define(new NamedRangeScope.Sheet("Budget"), new NamedRangeTarget("Budget", "A1"))
            .action()
            .actionType());
    assertThrows(
        IllegalStateException.class,
        () ->
            new Targets.NamedRangeRef(Targets.namedRanges("A", "B"))
                .define(new NamedRangeScope.Workbook(), new NamedRangeTarget("Budget", "A1")));
    assertThrows(
        IllegalStateException.class,
        () ->
            new Targets.NamedRangeRef(Targets.allNamedRanges())
                .define(new NamedRangeScope.Workbook(), new NamedRangeTarget("Budget", "A1")));
    assertThrows(
        IllegalStateException.class,
        () ->
            new Targets.NamedRangeRef(
                    Targets.anyNamedRange(new NamedRangeSelector.WorkbookScope("Global")))
                .define(new NamedRangeScope.Workbook(), new NamedRangeTarget("Budget", "A1")));
    assertThrows(
        IllegalStateException.class,
        () ->
            new Targets.NamedRangeRef(
                    new NamedRangeSelector.AnyOf(
                        List.of(new NamedRangeSelector.WorkbookScope("X"))))
                .define(new NamedRangeScope.Workbook(), new NamedRangeTarget("Budget", "A1")));
    assertEquals("DELETE_NAMED_RANGE", namedRange.delete().action().actionType());
    assertEquals("GET_NAMED_RANGES", namedRange.inspect().query().queryType());
    assertEquals("GET_NAMED_RANGE_SURFACE", namedRange.surface().query().queryType());
    assertEquals("ANALYZE_NAMED_RANGE_HEALTH", namedRange.analyzeHealth().query().queryType());
    assertEquals("EXPECT_PRESENT", namedRange.present().assertion().assertionType());
    assertEquals("EXPECT_ABSENT", namedRange.absent().assertion().assertionType());
  }

  private static void assertChartAndPivotTargets() {
    Targets.ChartRef exactChart = Targets.chart("Budget", "Trend");
    assertEquals(
        "SET_CHART", exactChart.defineOnSheet("Budget", sampleChart()).action().actionType());
    assertEquals("GET_CHARTS", exactChart.inspectOnSheet().query().queryType());
    assertEquals("EXPECT_PRESENT", exactChart.present().assertion().assertionType());
    assertEquals("EXPECT_ABSENT", exactChart.absent().assertion().assertionType());
    assertEquals(
        "GET_CHARTS", Targets.chart("Budget", "Trend").inspectOnSheet().query().queryType());
    assertEquals(
        "GET_CHARTS",
        new Targets.ChartRef(new ChartSelector.AllOnSheet("Budget"))
            .inspectOnSheet()
            .query()
            .queryType());

    Targets.PivotTableRef pivot = Targets.pivotTable("BudgetPivot");
    assertEquals("SET_PIVOT_TABLE", pivot.define(samplePivot()).action().actionType());
    assertEquals("DELETE_PIVOT_TABLE", pivot.delete().action().actionType());
    assertEquals("GET_PIVOT_TABLES", pivot.inspect().query().queryType());
    assertEquals("ANALYZE_PIVOT_TABLE_HEALTH", pivot.analyzeHealth().query().queryType());
    assertEquals("EXPECT_PRESENT", pivot.present().assertion().assertionType());
    assertEquals("EXPECT_ABSENT", pivot.absent().assertion().assertionType());
  }

  private static GridGrindResponse.CellStyleReport sampleStyle() {
    return new GridGrindResponse.CellStyleReport(
        "General",
        new CellAlignmentReport(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new CellFontReport(
            false,
            false,
            "Calibri",
            new FontHeightReport(220, BigDecimal.valueOf(11)),
            new CellColorReport("#000000"),
            false,
            false),
        new CellFillReport(ExcelFillPattern.SOLID, new CellColorReport("#FFFFFF"), null),
        new CellBorderReport(
            new CellBorderSideReport(ExcelBorderStyle.NONE, null),
            new CellBorderSideReport(ExcelBorderStyle.NONE, null),
            new CellBorderSideReport(ExcelBorderStyle.NONE, null),
            new CellBorderSideReport(ExcelBorderStyle.NONE, null)),
        new CellProtectionReport(true, false));
  }

  private static DataValidationInput sampleDataValidation() {
    return new DataValidationInput(
        new DataValidationRuleInput.ExplicitList(List.of("Open", "Closed")),
        false,
        false,
        null,
        null);
  }

  private static TableInput sampleTable() {
    return new TableInput("BudgetTable", "Budget", "A1:B3", false, new TableStyleInput.None());
  }

  private static PivotTableInput samplePivot() {
    return new PivotTableInput(
        "BudgetPivot",
        "Budget",
        new PivotTableInput.Source.Table("BudgetTable"),
        new PivotTableInput.Anchor("D4"),
        List.of("Item"),
        List.of(),
        List.of(),
        List.of(
            new PivotTableInput.DataField(
                "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total Amount", null)));
  }

  private static ChartInput sampleChart() {
    return new ChartInput.Bar(
        "BudgetChart",
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(0, 0),
            new DrawingMarkerInput(6, 12),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
        new ChartInput.Title.Text(Values.inlineText("Budget Trend")),
        null,
        ExcelChartDisplayBlanksAs.GAP,
        true,
        false,
        ExcelChartBarDirection.COLUMN,
        List.of(
            new ChartInput.Series(
                new ChartInput.Title.Text(Values.inlineText("Amounts")),
                new ChartInput.DataSource("Budget!$A$2:$A$3"),
                new ChartInput.DataSource("Budget!$B$2:$B$3"))));
  }
}
