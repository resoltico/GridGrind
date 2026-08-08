package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResults;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicReference;
import org.junit.jupiter.api.Test;

/** Edge coverage for executor guards, explicit bindings, and error mapping. */
class ExecutorGuardCoverageTest {
  @Test
  void defaultExecutorMethodsForwardAndRejectNullInputs() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
            dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
            java.util.List.of());
    WorkbookResult.Success expected =
        WorkbookResults.success(java.util.List.of(), java.util.List.of(), java.util.List.of());
    AtomicReference<ExecutionInputBindings> seenBindings = new AtomicReference<>();
    AtomicReference<ExecutionJournalSink> seenSink = new AtomicReference<>();
    GridGrindRequestExecutor executor =
        (ignoredRequest, bindings, sink) -> {
          seenBindings.set(bindings);
          seenSink.set(sink);
          return expected;
        };

    ExecutionInputBindings explicitBindings =
        ExecutionInputBindingsFixtureSupport.bindings(
            Path.of("/tmp"), "stdin".getBytes(StandardCharsets.UTF_8));
    assertTrue(explicitBindings.hasStandardInput());
    assertArrayEquals(
        "stdin".getBytes(StandardCharsets.UTF_8),
        explicitBindings.standardInputBytes().orElseThrow());
    assertSame(expected, executor.execute(request, explicitBindings));
    assertSame(explicitBindings, seenBindings.get());
    assertSame(ExecutionJournalSink.NOOP, seenSink.get());
    assertThrows(
        NullPointerException.class, () -> executor.execute(request, (ExecutionInputBindings) null));

    AtomicReference<ExecutionInputBindings> forwardedBindings = new AtomicReference<>();
    ExecutionJournalSink sink = event -> {};
    GridGrindRequestExecutor journalExecutor =
        (ignoredRequest, bindings, actualSink) -> {
          forwardedBindings.set(bindings);
          seenSink.set(actualSink);
          return expected;
        };
    assertSame(expected, journalExecutor.execute(request, explicitBindings, sink));
    assertSame(explicitBindings, forwardedBindings.get());
    assertSame(sink, seenSink.get());
  }

  @Test
  void executionInputBindingsAndInputSourceExceptionsCoverNullAndValidationBranches() {
    ExecutionInputBindings withoutStandardInput =
        ExecutionInputBindingsFixtureSupport.bindings(Path.of("/tmp"));
    assertFalse(withoutStandardInput.hasStandardInput());
    assertTrue(withoutStandardInput.standardInputBytes().isEmpty());
    assertThrows(
        NullPointerException.class,
        () ->
            new ExecutionInputBindings(
                Path.of("/tmp"),
                Path.of("/tmp/.gridgrind/tmp"),
                (ExecutionInputBindings.StandardInputBinding) null));

    InputSourceReadException exception =
        new InputSourceReadException("bad file", "cell text", "/tmp/cell.txt", null);
    assertEquals("/tmp/cell.txt", exception.inputPath());
    assertEquals("cell text", exception.inputKind());
    assertEquals(
        "inputKind must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new InputSourceReadException("bad", " ", "/tmp/x", null))
            .getMessage());
  }

  @Test
  void workbookCommandConverterRejectsUnresolvedSourceBackedValues() {
    assertEquals(
        "cell text must be resolved to INLINE before conversion",
        assertThrows(
                IllegalStateException.class,
                () ->
                    WorkbookCommandCellInputConverter.toExcelCellValue(
                        new CellInput.Text(TextSourceInput.utf8File("title.txt"))))
            .getMessage());

    PictureInput unresolvedPicture =
        picture(
            "Logo",
            new PictureDataInput(ExcelPictureFormat.PNG, BinarySourceInput.file("logo.png")),
            new dev.erst.gridgrind.contract.dto.DrawingAnchorInput.TwoCell(
                new dev.erst.gridgrind.contract.dto.DrawingMarkerInput(0, 0, 0, 0),
                new dev.erst.gridgrind.contract.dto.DrawingMarkerInput(1, 1, 0, 0),
                dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
            TextSourceInput.inline("Logo"));
    assertEquals(
        "picture payload must be resolved to INLINE_BASE64 before conversion",
        assertThrows(
                IllegalStateException.class,
                () ->
                    WorkbookCommandDrawingInputConverter.toExcelPictureDefinition(
                        unresolvedPicture))
            .getMessage());
  }

  @Test
  void defaultExecutorSurfacesResolveInputContextForSourceFailures() throws Exception {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    assertThrows(
        NullPointerException.class,
        () ->
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    mutations()),
                null,
                ExecutionJournalSink.NOOP));
    WorkbookPlan standardInputRequest =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Budget", "A1"),
                    new CellMutationAction.SetCell(
                        new CellInput.Text(TextSourceInput.standardInput())))));

    WorkbookResult.Failure unavailableFailure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            ExecutionContextFixtureSupport.execute(executor, standardInputRequest));
    dev.erst.gridgrind.contract.dto.ProblemContext.ResolveInputs unavailableContext =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.ProblemContext.ResolveInputs.class,
            unavailableFailure.problem().context());
    assertEquals(
        GridGrindProblemCode.INPUT_SOURCE_UNAVAILABLE, unavailableFailure.problem().code());
    assertEquals(java.util.Optional.of("cell text"), unavailableContext.inputKind());
    assertEquals(java.util.Optional.empty(), unavailableContext.inputPath());

    Path workingDirectory = Files.createTempDirectory("gridgrind-phase7-default-executor-");
    Files.writeString(workingDirectory.resolve("blank.txt"), "   ", StandardCharsets.UTF_8);
    WorkbookPlan blankFileRequest =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Budget", "A1"),
                    new CellMutationAction.SetCell(
                        new CellInput.Text(TextSourceInput.utf8File("blank.txt"))))));
    WorkbookResult.Failure blankFailure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            executor.execute(
                blankFileRequest,
                ExecutionInputBindingsFixtureSupport.bindings(workingDirectory),
                ExecutionJournalSink.NOOP));
    dev.erst.gridgrind.contract.dto.ProblemContext.ResolveInputs blankContext =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.ProblemContext.ResolveInputs.class,
            blankFailure.problem().context());
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, blankFailure.problem().code());
    assertEquals(java.util.Optional.empty(), blankContext.inputKind());
    assertEquals(java.util.Optional.empty(), blankContext.inputPath());
  }

  @Test
  void formulaDiagnosticsExposeInlineFormulasOnly() {
    CellMutationAction.SetCell fileBackedFormula =
        new CellMutationAction.SetCell(
            new CellInput.Formula(TextSourceInput.utf8File("formula.txt")));
    CellMutationAction.SetCell inlineFormula =
        new CellMutationAction.SetCell(new CellInput.Formula(TextSourceInput.inline("SUM(A1:A2)")));

    assertEquals(
        java.util.Optional.empty(), ExecutionActionDiagnosticFields.formulaFor(fileBackedFormula));
    assertEquals(
        java.util.Optional.of("SUM(A1:A2)"),
        ExecutionActionDiagnosticFields.formulaFor(inlineFormula));
  }

  @Test
  void mutationAndRangeSelectorDiagnosticHelpersCoverNullAndSingleSheetForwarders() {
    dev.erst.gridgrind.contract.dto.DrawingAnchorInput.TwoCell drawingAnchor =
        new dev.erst.gridgrind.contract.dto.DrawingAnchorInput.TwoCell(
            new dev.erst.gridgrind.contract.dto.DrawingMarkerInput(0, 0),
            new dev.erst.gridgrind.contract.dto.DrawingMarkerInput(1, 1),
            dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    dev.erst.gridgrind.contract.dto.PictureDataInput imageData =
        new dev.erst.gridgrind.contract.dto.PictureDataInput(
            ExcelPictureFormat.PNG, BinarySourceInput.inlineBase64("AQ=="));
    dev.erst.gridgrind.contract.dto.ChartInput simpleChart =
        new dev.erst.gridgrind.contract.dto.ChartInput(
            "Revenue",
            drawingAnchor,
            new dev.erst.gridgrind.contract.dto.ChartTitleInput.None(),
            new dev.erst.gridgrind.contract.dto.ChartLegendInput.Visible(
                dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition.RIGHT),
            dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs.GAP,
            true,
            java.util.List.of(
                new dev.erst.gridgrind.contract.dto.ChartPlotInput.Bar(
                    false,
                    dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection.COLUMN,
                    dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping.CLUSTERED,
                    java.util.Optional.empty(),
                    java.util.Optional.empty(),
                    java.util.List.of(
                        new dev.erst.gridgrind.contract.dto.ChartSeriesInput(
                            new dev.erst.gridgrind.contract.dto.ChartTitleInput.None(),
                            new dev.erst.gridgrind.contract.dto.ChartDataSourceInput.Reference(
                                "Budget!$A$1:$A$2"),
                            new dev.erst.gridgrind.contract.dto.ChartDataSourceInput.Reference(
                                "Budget!$B$1:$B$2"),
                            java.util.Optional.empty(),
                            java.util.Optional.empty(),
                            java.util.Optional.empty(),
                            java.util.Optional.empty())))));

    for (MutationAction action :
        java.util.List.of(
            new WorkbookMutationAction.EnsureSheet(),
            new WorkbookMutationAction.RenameSheet("Budget Copy"),
            new WorkbookMutationAction.DeleteSheet(),
            new WorkbookMutationAction.MoveSheet(0),
            new WorkbookMutationAction.CopySheet(
                "Budget Copy", new dev.erst.gridgrind.contract.dto.SheetCopyPosition.AppendAtEnd()),
            new WorkbookMutationAction.SetActiveSheet(),
            new WorkbookMutationAction.SetSelectedSheets(),
            new WorkbookMutationAction.SetSheetVisibility(
                dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility.HIDDEN),
            new WorkbookMutationAction.SetSheetProtection(
                new dev.erst.gridgrind.contract.dto.SheetProtectionSettings(
                    false, false, false, false, false, false, false, false, false, false, false,
                    false, false, false, false)),
            new WorkbookMutationAction.ClearSheetProtection(),
            new WorkbookMutationAction.SetWorkbookProtection(
                workbookProtection(true, false, false, null, null)),
            new WorkbookMutationAction.ClearWorkbookProtection(),
            new WorkbookMutationAction.MergeCells(),
            new WorkbookMutationAction.UnmergeCells(),
            new WorkbookMutationAction.SetColumnWidth(8.5d),
            new WorkbookMutationAction.SetRowHeight(12.0d),
            new WorkbookMutationAction.InsertRows(),
            new WorkbookMutationAction.DeleteRows(),
            new WorkbookMutationAction.ShiftRows(1),
            new WorkbookMutationAction.InsertColumns(),
            new WorkbookMutationAction.DeleteColumns(),
            new WorkbookMutationAction.ShiftColumns(1),
            new WorkbookMutationAction.SetRowVisibility(true),
            new WorkbookMutationAction.SetColumnVisibility(true),
            new WorkbookMutationAction.GroupRows(true),
            new WorkbookMutationAction.UngroupRows(),
            new WorkbookMutationAction.GroupColumns(true),
            new WorkbookMutationAction.UngroupColumns(),
            new WorkbookMutationAction.SetSheetPane(
                new dev.erst.gridgrind.contract.dto.PaneInput.Frozen(1, 1, 1, 1)),
            new WorkbookMutationAction.SetSheetZoom(125),
            new WorkbookMutationAction.SetSheetPresentation(
                dev.erst.gridgrind.contract.dto.SheetPresentationInput.defaults()),
            new WorkbookMutationAction.SetPrintLayout(
                dev.erst.gridgrind.contract.dto.PrintLayoutInput.defaults()),
            new CellMutationAction.SetRange(
                new dev.erst.gridgrind.contract.dto.CellGridInput.Typed(
                    java.util.List.of(
                        java.util.List.of(new CellInput.Text(TextSourceInput.inline("budget")))))),
            new CellMutationAction.ClearRange(),
            new CellMutationAction.SetHyperlink(
                new dev.erst.gridgrind.contract.dto.HyperlinkTarget.Url("https://example.com")),
            new CellMutationAction.ClearHyperlink(),
            new CellMutationAction.SetComment(
                dev.erst.gridgrind.contract.dto.CommentInput.plain(
                    TextSourceInput.inline("Reviewed"), "GridGrind", true)),
            new CellMutationAction.ClearComment(),
            new DrawingMutationAction.SetPicture(
                picture("Logo", imageData, drawingAnchor, TextSourceInput.inline("Budget logo"))),
            new DrawingMutationAction.SetChart(simpleChart),
            new DrawingMutationAction.SetShape(
                simpleShape("BudgetBox", drawingAnchor, "rect", TextSourceInput.inline("Budget"))),
            new DrawingMutationAction.SetEmbeddedObject(
                new dev.erst.gridgrind.contract.dto.EmbeddedObjectInput(
                    "Attachment",
                    "Attachment",
                    "attachment.bin",
                    "open",
                    BinarySourceInput.inlineBase64("AQ=="),
                    imageData,
                    drawingAnchor)),
            new DrawingMutationAction.SetDrawingObjectAnchor(drawingAnchor),
            new DrawingMutationAction.DeleteDrawingObject(),
            new CellMutationAction.ApplyStyle(styleInput("0.00", null, null, null, null, null)),
            new StructuredMutationAction.SetDataValidation(
                dataValidation(
                    new dev.erst.gridgrind.contract.dto.DataValidationRuleInput.ExplicitList(
                        java.util.List.of("Open", "Closed")),
                    true,
                    false,
                    null,
                    null)),
            new StructuredMutationAction.SetConditionalFormatting(
                new dev.erst.gridgrind.contract.dto.ConditionalFormattingDefinitionInput(
                    java.util.List.of(
                        new dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput
                            .FormulaRule("A1>0", false, Optional.empty())))),
            new StructuredMutationAction.SetAutofilter(),
            new WorkbookMutationAction.ClearPrintLayout(),
            new StructuredMutationAction.ClearConditionalFormatting(),
            new StructuredMutationAction.ClearDataValidations(),
            new StructuredMutationAction.ClearAutofilter(),
            new StructuredMutationAction.DeleteTable(),
            new StructuredMutationAction.DeletePivotTable(),
            new StructuredMutationAction.DeleteNamedRange(),
            new CellMutationAction.AppendRow(
                new dev.erst.gridgrind.contract.dto.CellRowInput.Typed(
                    java.util.List.of(new CellInput.Text(TextSourceInput.inline("budget"))))),
            new WorkbookMutationAction.AutoSizeColumns())) {
      assertEquals(Optional.empty(), ExecutionActionDiagnosticFields.sheetNameFor(action));
      assertEquals(Optional.empty(), ExecutionActionDiagnosticFields.rangeFor(action));
      assertEquals(Optional.empty(), ExecutionActionDiagnosticFields.formulaFor(action));
      assertEquals(Optional.empty(), ExecutionActionDiagnosticFields.namedRangeNameFor(action));
    }

    StructuredMutationAction.SetNamedRange rangeDefinedNamedRange =
        new StructuredMutationAction.SetNamedRange(
            "BudgetWindow",
            new dev.erst.gridgrind.contract.dto.NamedRangeScope.Sheet("Budget"),
            dev.erst.gridgrind.contract.dto.NamedRangeTarget.range("Budget", "A1:B2"));
    assertEquals(
        Optional.of("Budget"),
        ExecutionActionDiagnosticFields.sheetNameFor(rangeDefinedNamedRange));
    assertEquals(
        Optional.of("A1:B2"), ExecutionActionDiagnosticFields.rangeFor(rangeDefinedNamedRange));
    assertEquals(
        Optional.empty(), ExecutionActionDiagnosticFields.formulaFor(rangeDefinedNamedRange));
    assertEquals(
        Optional.of("BudgetWindow"),
        ExecutionActionDiagnosticFields.namedRangeNameFor(rangeDefinedNamedRange));

    StructuredMutationAction.SetTable setTable =
        new StructuredMutationAction.SetTable(
            dev.erst.gridgrind.contract.dto.TableInput.withDefaultMetadata(
                "BudgetTable",
                "Budget",
                "A1:B5",
                false,
                new dev.erst.gridgrind.contract.dto.TableStyleInput.None()));
    assertEquals(Optional.of("Budget"), ExecutionActionDiagnosticFields.sheetNameFor(setTable));
    assertEquals(Optional.of("A1:B5"), ExecutionActionDiagnosticFields.rangeFor(setTable));
    assertEquals(Optional.empty(), ExecutionActionDiagnosticFields.formulaFor(setTable));
    assertEquals(Optional.empty(), ExecutionActionDiagnosticFields.namedRangeNameFor(setTable));

    assertEquals(
        Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (dev.erst.gridgrind.contract.selector.RangeSelector)
                new dev.erst.gridgrind.contract.selector.RangeSelector.ByRange("Budget", "A1:B2")));
  }

  @Test
  void inputSourceProblemsMapToDedicatedCodesAndPassthroughContexts() {
    assertEquals(
        GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND,
        GridGrindProblems.codeFor(
            new InputSourceNotFoundException("missing", "cell text", "/tmp/missing.txt", null)));
    assertEquals(
        GridGrindProblemCode.INPUT_SOURCE_UNAVAILABLE,
        GridGrindProblems.codeFor(
            new InputSourceUnavailableException("stdin missing", "cell text")));
    assertEquals(
        GridGrindProblemCode.INPUT_SOURCE_IO_ERROR,
        GridGrindProblems.codeFor(
            new InputSourceReadException("io failed", "cell text", "/tmp/cell.txt", null)));

    dev.erst.gridgrind.contract.dto.ProblemContext.ResolveInputs context =
        new dev.erst.gridgrind.contract.dto.ProblemContext.ResolveInputs(
            dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape.known(
                "NEW", "NONE"),
            dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.InputReference.path(
                "cell text", "/tmp/cell.txt"));
    GridGrindProblemDetail.Problem problem =
        GridGrindProblems.fromException(
            new InputSourceReadException("io failed", "cell text", "/tmp/cell.txt", null), context);
    assertSame(context, problem.context());
    assertEquals(GridGrindProblemCode.INPUT_SOURCE_IO_ERROR, problem.code());
  }

  @Test
  void summaryTargetDescribesSourceBackedTextKindsInsideTableKeySelectors() {
    ExecutionJournal.Target blankTarget =
        ExecutionJournalTargetResolver.summaryTarget(
            new dev.erst.gridgrind.contract.selector.TableRowSelector.ByKeyCell(
                new dev.erst.gridgrind.contract.selector.TableSelector.ByName("BudgetTable"),
                "Item",
                new CellInput.Blank()));
    assertEquals("Row where Item=Blank[] in Table BudgetTable", blankTarget.label());

    ExecutionJournal.Target fileTextTarget =
        ExecutionJournalTargetResolver.summaryTarget(
            new dev.erst.gridgrind.contract.selector.TableRowSelector.ByKeyCell(
                new dev.erst.gridgrind.contract.selector.TableSelector.ByName("BudgetTable"),
                "Item",
                new CellInput.Text(TextSourceInput.utf8File("title.txt"))));
    assertTrue(fileTextTarget.label().contains("Text[path=title.txt]"));

    ExecutionJournal.Target stdinTextTarget =
        ExecutionJournalTargetResolver.summaryTarget(
            new dev.erst.gridgrind.contract.selector.TableRowSelector.ByKeyCell(
                new dev.erst.gridgrind.contract.selector.TableSelector.ByName("BudgetTable"),
                "Item",
                new CellInput.Formula(TextSourceInput.standardInput())));
    assertTrue(stdinTextTarget.label().contains("Formula[source=STANDARD_INPUT]"));

    assertTrue(
        ExecutionJournalTargetResolver.summaryTarget(
                new dev.erst.gridgrind.contract.selector.TableRowSelector.ByKeyCell(
                    new dev.erst.gridgrind.contract.selector.TableSelector.ByName("BudgetTable"),
                    "Item",
                    new CellInput.NumberValue(42.0d)))
            .label()
            .contains("Number[number=42.0]"));
    assertTrue(
        ExecutionJournalTargetResolver.summaryTarget(
                new dev.erst.gridgrind.contract.selector.TableRowSelector.ByKeyCell(
                    new dev.erst.gridgrind.contract.selector.TableSelector.ByName("BudgetTable"),
                    "Item",
                    new CellInput.BooleanValue(true)))
            .label()
            .contains("Boolean[value=true]"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new dev.erst.gridgrind.contract.selector.TableRowSelector.ByKeyCell(
                new dev.erst.gridgrind.contract.selector.TableSelector.ByName("BudgetTable"),
                "Item",
                new CellInput.RichText(
                    java.util.List.of(
                        new dev.erst.gridgrind.contract.dto.RichTextRunInput(
                            TextSourceInput.inline("Ada"), Optional.empty())))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new dev.erst.gridgrind.contract.selector.TableRowSelector.ByKeyCell(
                new dev.erst.gridgrind.contract.selector.TableSelector.ByName("BudgetTable"),
                "Item",
                new CellInput.Date(LocalDate.of(2026, 4, 18))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new dev.erst.gridgrind.contract.selector.TableRowSelector.ByKeyCell(
                new dev.erst.gridgrind.contract.selector.TableSelector.ByName("BudgetTable"),
                "Item",
                new CellInput.DateTime(LocalDateTime.of(2026, 4, 18, 12, 30))));
  }
}
