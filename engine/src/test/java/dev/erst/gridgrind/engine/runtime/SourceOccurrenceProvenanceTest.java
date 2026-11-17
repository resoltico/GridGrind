package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.CellAssertion;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.PreparedMutationStep;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.engine.poi.NativeInvocationContext;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.ValueSource;

/** One resource snapshot can serve distinct authored occurrences without merging their origins. */
class SourceOccurrenceProvenanceTest {
  @TempDir Path root;

  @Test
  void sameSourceObjectDiagnosesEverySelectorAndActionOccurrenceAtItsOwnPath() throws Exception {
    Files.writeString(root.resolve("key.txt"), "\u0000");
    TextSourceSpec shared = TextSourceSpec.utf8File("key.txt");
    WorkbookPlan request = plan(shared);
    var report =
        new GridGrindRequestDoctor(
                NativeInvocationContext.enterOwned(
                    dev.erst.gridgrind.core.semantics.ExecutionConventions.STANDARD))
            .diagnose(request, bindings(1));
    assertFalse(report.valid());
    assertEquals(
        List.of(
            "steps[0].action.value.value.path",
            "steps[0].target.row.expectedValue.value.path",
            "steps[1].target.row.expectedValue.value.path",
            "steps[2].target.row.expectedValue.value.path"),
        report.problems().stream()
            .map(problem -> assertInstanceOf(ProblemContext.ResolveInputs.class, problem.context()))
            .map(context -> context.json().orElseThrow().jsonPathValue().orElseThrow())
            .sorted()
            .toList());
    assertSame(shared, authoredKey(request, 0));
    assertSame(shared, authoredKey(request, 1));
    assertSame(shared, authoredKey(request, 2));
    assertEquals("\u0000", Files.readString(root.resolve("key.txt")));
  }

  @Test
  void materializationRetainsTheExactOriginalPlanAndSeparatesResolvedSelectorValues()
      throws Exception {
    Files.writeString(root.resolve("key.txt"), "Hosting");
    TextSourceSpec shared = TextSourceSpec.utf8File("key.txt");
    WorkbookPlan request = plan(shared);
    var inputs = bindings(7);
    try {
      var result = RequestPreflight.verify(request, inputs);
      assertTrue(result.problems().isEmpty(), result.problems().toString());
      var prepared = result.preparedRequest().orElseThrow().plan();
      assertSame(request, prepared.request());
      for (int index = 0; index < 3; index++) {
        assertSame(shared, authoredKey(request, index));
        assertSame(request.steps().get(index), prepared.steps().get(index).authoredStep());
        var boundTarget =
            switch (prepared.steps().get(index)) {
              case PreparedMutationStep mutation -> mutation.target();
              case dev.erst.gridgrind.core.plan.PreparedAssertionStep assertion ->
                  assertion.target();
              case dev.erst.gridgrind.core.plan.PreparedInspectionStep inspection ->
                  inspection.target();
            };
        var boundRow =
            assertInstanceOf(
                TableRowSelector.ByKeyCell.class,
                assertInstanceOf(TableCellSelector.ByColumnName.class, boundTarget).row());
        assertEquals(
            "Hosting",
            assertInstanceOf(
                    TextSourceSpec.Inline.class,
                    assertInstanceOf(CellValueSpec.Text.class, boundRow.expectedValue()).value())
                .text());
      }
    } finally {
      inputs.closeRunResources();
    }
  }

  @ParameterizedTest
  @CsvSource({"FILE,false", "STDIN,false", "FILE,true", "STDIN,true"})
  void failedTableKeyAssertionKeepsItsAuthoredSourceAndResolvedObservation(
      String kind, boolean hostRequired) throws Exception {
    Files.writeString(root.resolve("key.txt"), "Hosting");
    TextSourceSpec source =
        "FILE".equals(kind) ? TextSourceSpec.utf8File("key.txt") : new TextSourceSpec.Stdin();
    var sheet = new SheetName("Budget");
    var table = new TableName("Orders");
    var target =
        new TableCellSelector.ByColumnName(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName(table), "Item", new CellValueSpec.Text(source)),
            "Amount");
    var assertion =
        new AssertionStep(
            new StepId("assert"), target, new CellAssertion.Value(new CellValue.NumberValue(99)));
    var request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of(
                new MutationStep(
                    new StepId("sheet"),
                    new dev.erst.gridgrind.core.plan.SheetSelector.ByName(sheet),
                    new dev.erst.gridgrind.core.plan.WorkbookMutationAction.EnsureSheet()),
                new MutationStep(
                    new StepId("values"),
                    new dev.erst.gridgrind.core.plan.RangeSelector.ByRange(
                        sheet, new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B2")),
                    new CellMutationAction.SetRange(
                        new dev.erst.gridgrind.core.model.value.CellGridSpec.Typed(
                            List.of(
                                List.of(
                                    new CellValueSpec.Text(TextSourceSpec.inline("Item")),
                                    new CellValueSpec.Text(TextSourceSpec.inline("Amount"))),
                                List.of(
                                    new CellValueSpec.Text(TextSourceSpec.inline("Hosting")),
                                    new CellValueSpec.NumberValue(100)))))),
                new MutationStep(
                    new StepId("table"),
                    new TableSelector.ByNameOnSheet(table, sheet),
                    new dev.erst.gridgrind.core.plan.StructuredMutationAction.SetTable(
                        dev.erst.gridgrind.core.model.table.TableSpec.withDefaultMetadata(
                            table,
                            sheet,
                            new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B2"),
                            false,
                            new dev.erst.gridgrind.core.model.table.TableStyleSpec.None()))),
                assertion));
    if (hostRequired)
      request =
          WorkbookPlan.standard(
              request.source(),
              request.persistence(),
              request.execution(),
              request.formulaEnvironment(),
              request.steps().subList(0, 3));
    var acceptance =
        hostRequired
            ? new AcceptancePolicy.RequireAll(
                List.of(new AcceptancePolicy.Requirement.TerminalAssertions(List.of(assertion))))
            : AcceptancePolicy.minimum();
    var grant =
        new ExecutionGrant(
            7,
            List.of(
                "FILE".equals(kind)
                    ? new ExecutionGrant.ReadAuthority.File(root.resolve("key.txt"))
                    : new ExecutionGrant.ReadAuthority.Stdin()),
            List.of("ENSURE_SHEET", "SET_RANGE", "SET_TABLE", "EXPECT_CELL_VALUE"),
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new ExecutionGrant.PublicationAuthority.None(),
            List.of(),
            acceptance);
    var inputs =
        "FILE".equals(kind)
            ? new ExecutionInputBindings(root, root.resolve("scratch"), grant)
            : new ExecutionInputBindings(
                root,
                root.resolve("scratch"),
                "Hosting".getBytes(java.nio.charset.StandardCharsets.UTF_8),
                grant);
    var result =
        assertInstanceOf(
            dev.erst.gridgrind.core.result.WorkbookResult.Failure.class,
            new DefaultGridGrindRequestExecutor(
                    NativeInvocationContext.enterOwned(
                        dev.erst.gridgrind.core.semantics.ExecutionConventions.STANDARD))
                .execute(request, inputs, ExecutionProgressSink.NOOP));
    assertEquals(
        dev.erst.gridgrind.core.problem.GridGrindProblemCode.ASSERTION_FAILED,
        result.problem().code(),
        result.problem().toString());
    var taskEvidence =
        assertInstanceOf(
                dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.TaskSpecific.Failed.class,
                result.evidence().taskSpecific())
            .assertions()
            .getFirst();
    assertEquals(
        hostRequired
            ? dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.TaskSpecific.Origin
                .HOST_REQUIRED
            : dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.TaskSpecific.Origin
                .PLAN_AUTHORED,
        taskEvidence.origin());
    var failure =
        assertInstanceOf(
                dev.erst.gridgrind.core.plan.AssertionResult.Failed.class, taskEvidence.result())
            .failure();
    assertSame(target, failure.target());
    var failedRow =
        assertInstanceOf(
            TableRowSelector.ByKeyCell.class,
            assertInstanceOf(TableCellSelector.ByColumnName.class, failure.target()).row());
    assertSame(
        source, assertInstanceOf(CellValueSpec.Text.class, failedRow.expectedValue()).value());
    assertEquals(assertion.stepId(), failure.stepId());
    if (!hostRequired) {
      var context = assertInstanceOf(ProblemContext.ExecuteStep.class, result.problem().context());
      assertEquals(3, context.stepIndex());
      assertEquals(assertion.stepId(), context.stepId());
    }
    var observed =
        assertInstanceOf(
                dev.erst.gridgrind.core.plan.SheetInspectionResult.CellsResult.class,
                failure.observations().getFirst())
            .cells()
            .getFirst();
    assertEquals(new CellAddress("B2"), observed.address());
    assertEquals(new CellValue.NumberValue(100), observed.value().orElseThrow());
  }

  @ParameterizedTest
  @CsvSource({"FILE,STDIN,false", "STDIN,FILE,false", "FILE,FILE,true", "STDIN,STDIN,true"})
  void commentCrossFieldMismatchRetainsThePlainTextOriginAndContinuesIndependentSteps(
      String plainKind, String runKind, boolean shared) throws Exception {
    Files.writeString(root.resolve("plain.txt"), "A");
    Files.writeString(root.resolve("runs.txt"), "B");
    TextSourceSpec plain =
        "STDIN".equals(plainKind)
            ? new TextSourceSpec.Stdin()
            : TextSourceSpec.utf8File("plain.txt");
    TextSourceSpec run =
        shared
            ? plain
            : "STDIN".equals(runKind)
                ? new TextSourceSpec.Stdin()
                : TextSourceSpec.utf8File("runs.txt");
    var runs =
        shared
            ? List.of(
                new RichTextRunSpec(run, Optional.empty()),
                new RichTextRunSpec(run, Optional.empty()))
            : List.of(new RichTextRunSpec(run, Optional.empty()));
    var comment =
        new CommentSpec(plain, "Synthetic Author", false, Optional.of(runs), Optional.empty());
    var request = commentPlan(comment);
    var inputs = commentBindings(plainKind, runKind, shared ? 1 : 2, shared);
    var visited = new ArrayList<StepId>();
    try {
      var result =
          RequestPreflight.verify(
              request,
              inputs,
              (step, bindings) -> {
                visited.add(step.stepId());
                return SourceBackedPlanResolver.resolveStep(step, bindings);
              });
      assertEquals(List.of(new StepId("comment"), new StepId("independent")), visited);
      assertEquals(1, result.problems().size(), result.problems().toString());
      var failure = result.problems().getFirst();
      assertTrue(
          failure.message().contains("comment runs must concatenate to the plain text"),
          failure::message);
      var context = assertInstanceOf(ProblemContext.ResolveInputs.class, failure.context());
      assertEquals(
          "steps[0].action.comment.text." + ("STDIN".equals(plainKind) ? "type" : "path"),
          context.json().orElseThrow().jsonPathValue().orElseThrow());
    } finally {
      inputs.closeRunResources();
    }
  }

  @ParameterizedTest
  @ValueSource(strings = {"FILE", "STDIN"})
  void aSharedCommentSourceResolvesBothPlainAndRichTextWithoutDoubleCharging(String kind)
      throws Exception {
    Files.writeString(root.resolve("plain.txt"), "A");
    TextSourceSpec shared =
        "STDIN".equals(kind) ? new TextSourceSpec.Stdin() : TextSourceSpec.utf8File("plain.txt");
    var comment =
        new CommentSpec(
            shared,
            "Synthetic Author",
            false,
            Optional.of(List.of(new RichTextRunSpec(shared, Optional.empty()))),
            Optional.empty());
    var request = commentPlan(comment);
    var inputs = commentBindings(kind, kind, 1, true);
    try {
      var result = RequestPreflight.verify(request, inputs);
      assertTrue(result.problems().isEmpty(), result.problems().toString());
      var prepared =
          assertInstanceOf(
              PreparedMutationStep.class,
              result.preparedRequest().orElseThrow().plan().steps().getFirst());
      var resolved =
          assertInstanceOf(ResolvedMutationAction.SetComment.class, prepared.action()).comment();
      assertEquals("A", resolved.text());
      assertEquals("A", resolved.runs().orElseThrow().plainText());
      assertSame(shared, comment.text());
    } finally {
      inputs.closeRunResources();
    }
  }

  private WorkbookPlan commentPlan(CommentSpec comment) {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExecutionPolicySpec.defaults(),
        FormulaEnvironmentSpec.empty(),
        List.of(
            new MutationStep(
                new StepId("comment"),
                new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                new CellMutationAction.SetComment(comment)),
            new MutationStep(
                new StepId("independent"),
                new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("B1")),
                new CellMutationAction.SetCell(new CellValueSpec.NumberValue(1)))));
  }

  private ExecutionInputBindings commentBindings(
      String plainKind, String runKind, long budget, boolean shared) {
    var reads = new ArrayList<ExecutionGrant.ReadAuthority>();
    if ("FILE".equals(plainKind))
      reads.add(new ExecutionGrant.ReadAuthority.File(root.resolve("plain.txt")));
    if ("FILE".equals(runKind) && !shared)
      reads.add(new ExecutionGrant.ReadAuthority.File(root.resolve("runs.txt")));
    boolean stdin = "STDIN".equals(plainKind) || "STDIN".equals(runKind);
    if (stdin) reads.add(new ExecutionGrant.ReadAuthority.Stdin());
    var grant =
        new ExecutionGrant(
            budget,
            reads,
            List.of("SET_COMMENT", "SET_CELL"),
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new ExecutionGrant.PublicationAuthority.None(),
            List.of(),
            AcceptancePolicy.minimum());
    return stdin
        ? new ExecutionInputBindings(
            root,
            root.resolve("scratch"),
            (shared || "STDIN".equals(plainKind) ? "A" : "B")
                .getBytes(java.nio.charset.StandardCharsets.UTF_8),
            grant)
        : new ExecutionInputBindings(root, root.resolve("scratch"), grant);
  }

  private WorkbookPlan plan(TextSourceSpec shared) {
    var target =
        new TableCellSelector.ByColumnName(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName(new TableName("Orders")),
                "Item",
                new CellValueSpec.Text(shared)),
            "Amount");
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExecutionPolicySpec.defaults(),
        FormulaEnvironmentSpec.empty(),
        List.of(
            new MutationStep(
                new StepId("write"),
                target,
                new CellMutationAction.SetCell(new CellValueSpec.Text(shared))),
            new InspectionStep(
                new StepId("inspect"), target, new SheetIntrospectionQuery.GetCells()),
            new AssertionStep(
                new StepId("assert"),
                target,
                new CellAssertion.Value(new CellValue.Text("Hosting")))));
  }

  private ExecutionInputBindings bindings(long budget) {
    return new ExecutionInputBindings(
        root,
        root.resolve("scratch"),
        new ExecutionGrant(
            budget,
            List.of(new ExecutionGrant.ReadAuthority.File(root.resolve("key.txt"))),
            List.of("SET_CELL", "GET_CELLS", "EXPECT_CELL_VALUE"),
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new ExecutionGrant.PublicationAuthority.None(),
            List.of(),
            AcceptancePolicy.minimum()));
  }

  private TextSourceSpec authoredKey(WorkbookPlan request, int index) {
    var target = (TableCellSelector.ByColumnName) request.steps().get(index).target();
    var row = (TableRowSelector.ByKeyCell) target.row();
    return ((CellValueSpec.Text) row.expectedValue()).value();
  }
}
