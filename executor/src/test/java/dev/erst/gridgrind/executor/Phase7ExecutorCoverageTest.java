package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.mutate;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.mutations;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.request;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelPictureFormat;
import java.lang.reflect.Method;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.concurrent.atomic.AtomicReference;
import org.junit.jupiter.api.Test;

/** Focused Phase 7 edge coverage for executor guards, transport defaults, and error mapping. */
class Phase7ExecutorCoverageTest {
  @Test
  void defaultExecutorMethodsForwardAndRejectNullInputs() {
    WorkbookPlan request =
        new WorkbookPlan(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            java.util.List.of());
    GridGrindResponse.Success expected =
        new GridGrindResponse.Success(
            null,
            new GridGrindResponse.PersistenceOutcome.NotSaved(),
            java.util.List.of(),
            java.util.List.of(),
            java.util.List.of());
    AtomicReference<ExecutionInputBindings> seenBindings = new AtomicReference<>();
    AtomicReference<ExecutionJournalSink> seenSink = new AtomicReference<>();
    GridGrindRequestExecutor executor =
        (ignoredRequest, bindings, sink) -> {
          seenBindings.set(bindings);
          seenSink.set(sink);
          return expected;
        };

    ExecutionInputBindings explicitBindings =
        new ExecutionInputBindings(Path.of("/tmp"), "stdin".getBytes(StandardCharsets.UTF_8));
    assertSame(expected, executor.execute(request, explicitBindings));
    assertSame(explicitBindings, seenBindings.get());
    assertSame(ExecutionJournalSink.NOOP, seenSink.get());
    assertThrows(
        NullPointerException.class, () -> executor.execute(request, (ExecutionInputBindings) null));

    AtomicReference<ExecutionInputBindings> processDefaultBindings = new AtomicReference<>();
    ExecutionJournalSink sink = event -> {};
    GridGrindRequestExecutor journalExecutor =
        (ignoredRequest, bindings, actualSink) -> {
          processDefaultBindings.set(bindings);
          seenSink.set(actualSink);
          return expected;
        };
    assertSame(expected, journalExecutor.execute(request, sink));
    assertNotNull(processDefaultBindings.get());
    assertSame(sink, seenSink.get());
  }

  @Test
  void executionInputBindingsAndInputSourceExceptionsCoverNullAndValidationBranches() {
    ExecutionInputBindings withoutStandardInput =
        new ExecutionInputBindings(Path.of("/tmp"), (byte[]) null);
    assertNull(withoutStandardInput.standardInputBytes());

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
                    WorkbookCommandConverter.toExcelCellValue(
                        new CellInput.Text(TextSourceInput.utf8File("title.txt"))))
            .getMessage());

    PictureInput unresolvedPicture =
        new PictureInput(
            "Logo",
            new PictureDataInput(ExcelPictureFormat.PNG, BinarySourceInput.file("logo.png")),
            new dev.erst.gridgrind.contract.dto.DrawingAnchorInput.TwoCell(
                new dev.erst.gridgrind.contract.dto.DrawingMarkerInput(0, 0, 0, 0),
                new dev.erst.gridgrind.contract.dto.DrawingMarkerInput(1, 1, 0, 0),
                dev.erst.gridgrind.excel.ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
            TextSourceInput.inline("Logo"));
    assertEquals(
        "picture payload must be resolved to INLINE_BASE64 before conversion",
        assertThrows(
                IllegalStateException.class,
                () -> WorkbookCommandConverter.toExcelPictureDefinition(unresolvedPicture))
            .getMessage());
  }

  @Test
  void defaultExecutorSurfacesResolveInputContextForSourceFailures() throws Exception {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    WorkbookPlan standardInputRequest =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Budget", "A1"),
                    new MutationAction.SetCell(
                        new CellInput.Text(TextSourceInput.standardInput())))));

    GridGrindResponse.Failure unavailableFailure =
        assertInstanceOf(
            GridGrindResponse.Failure.class,
            executor.execute(
                standardInputRequest, (ExecutionInputBindings) null, ExecutionJournalSink.NOOP));
    GridGrindResponse.ProblemContext.ResolveInputs unavailableContext =
        assertInstanceOf(
            GridGrindResponse.ProblemContext.ResolveInputs.class,
            unavailableFailure.problem().context());
    assertEquals(
        GridGrindProblemCode.INPUT_SOURCE_UNAVAILABLE, unavailableFailure.problem().code());
    assertEquals("cell text", unavailableContext.inputKind());
    assertNull(unavailableContext.inputPath());

    Path workingDirectory = Files.createTempDirectory("gridgrind-phase7-default-executor-");
    Files.writeString(workingDirectory.resolve("blank.txt"), "   ", StandardCharsets.UTF_8);
    WorkbookPlan blankFileRequest =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Budget", "A1"),
                    new MutationAction.SetCell(
                        new CellInput.Text(TextSourceInput.utf8File("blank.txt"))))));
    GridGrindResponse.Failure blankFailure =
        assertInstanceOf(
            GridGrindResponse.Failure.class,
            executor.execute(
                blankFileRequest,
                new ExecutionInputBindings(workingDirectory, (byte[]) null),
                ExecutionJournalSink.NOOP));
    GridGrindResponse.ProblemContext.ResolveInputs blankContext =
        assertInstanceOf(
            GridGrindResponse.ProblemContext.ResolveInputs.class, blankFailure.problem().context());
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, blankFailure.problem().code());
    assertNull(blankContext.inputKind());
    assertNull(blankContext.inputPath());
  }

  @Test
  @SuppressWarnings("PMD.AvoidAccessibilityAlteration")
  void privateInlineFormulaReturnsNullForNonInlineSources() throws Exception {
    Method inlineFormula =
        DefaultGridGrindRequestExecutor.class.getDeclaredMethod(
            "inlineFormula", CellInput.Formula.class);
    inlineFormula.setAccessible(true);

    assertNull(
        inlineFormula.invoke(null, new CellInput.Formula(TextSourceInput.utf8File("formula.txt"))));
    assertEquals(
        "SUM(A1:A2)",
        inlineFormula.invoke(null, new CellInput.Formula(TextSourceInput.inline("SUM(A1:A2)"))));
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

    GridGrindResponse.ProblemContext.ResolveInputs context =
        new GridGrindResponse.ProblemContext.ResolveInputs(
            "NEW", "NONE", "cell text", "/tmp/cell.txt");
    GridGrindResponse.Problem problem =
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
                    new CellInput.Numeric(42.0d)))
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
                            TextSourceInput.inline("Ada"), null)))));
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
