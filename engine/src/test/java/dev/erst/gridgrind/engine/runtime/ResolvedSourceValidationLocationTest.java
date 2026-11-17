package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.api.GridGrindEngine;
import dev.erst.gridgrind.engine.api.GridGrindRequestInputs;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/** Validation after reading a source retains the original source-use location. */
class ResolvedSourceValidationLocationTest {
  @TempDir Path root;

  @Test
  void selectorComparisonsDoNotAcquireTheCellWriteLengthRestriction() throws Exception {
    byte[] bytes = "x".repeat(32768).getBytes(StandardCharsets.UTF_8);
    var grant =
        new ExecutionGrant(
            bytes.length,
            List.of(new ExecutionGrant.ReadAuthority.Stdin()),
            List.of(),
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new ExecutionGrant.PublicationAuthority.None(),
            List.of(),
            AcceptancePolicy.minimum());
    var bindings = new ExecutionInputBindings(root, root.resolve("scratch"), bytes, grant);
    var original = new CellValueSpec.Text(new TextSourceSpec.Stdin());
    try {
      var comparison =
          assertInstanceOf(
              CellValueSpec.Text.class,
              SourceBackedPlanResolver.resolveCellInput(original, bindings));
      assertEquals(
          32768, assertInstanceOf(TextSourceSpec.Inline.class, comparison.value()).text().length());
      assertThrows(
          IllegalArgumentException.class,
          () -> SourceBackedPlanResolver.resolveCellWriteInput(original, bindings));
    } finally {
      bindings.closeRunResources();
    }
  }

  @ParameterizedTest
  @CsvSource({
    "FILE,TEXT",
    "STDIN,TEXT",
    "FILE,FORMULA",
    "STDIN,FORMULA",
    "FILE,RAW_FORMULA",
    "STDIN,RAW_FORMULA",
    "FILE,RICH_XML",
    "STDIN,RICH_XML",
    "FILE,LONG_TEXT",
    "STDIN,LONG_TEXT",
    "FILE,LONG_RICH",
    "STDIN,LONG_RICH"
  })
  void invalidResolvedContentRetainsEachSharedSourceOccurrence(String transport, String kind)
      throws Exception {
    byte[] bytes = invalidContent(kind).getBytes(StandardCharsets.UTF_8);
    boolean stdin = "STDIN".equals(transport);
    TextSourceSpec source =
        stdin ? new TextSourceSpec.Stdin() : TextSourceSpec.utf8File("input.txt");
    if (!stdin) Files.write(root.resolve("input.txt"), bytes);
    var request = requestWithRepeatedValue(inputValue(kind, source));
    var report = GridGrindEngine.requestDoctor().diagnose(request, inputsFor(bytes, stdin));
    assertFalse(report.valid());
    String field = stdin ? "type" : "path";
    String operand = expectedOperand(kind);
    assertEquals(
        List.of("steps[0]." + operand + field, "steps[1]." + operand + field),
        report.problems().stream()
            .map(problem -> assertInstanceOf(ProblemContext.ResolveInputs.class, problem.context()))
            .map(context -> context.json().orElseThrow().jsonPathValue().orElseThrow())
            .sorted()
            .toList());
  }

  private static String invalidContent(String kind) {
    return switch (kind) {
      case "LONG_TEXT" -> "x".repeat(32768);
      case "LONG_RICH" -> "x".repeat(20000);
      case "FORMULA" -> "=1+1";
      default -> "\u0000";
    };
  }

  private static CellValueSpec inputValue(String kind, TextSourceSpec source) {
    return switch (kind) {
      case "TEXT", "LONG_TEXT" -> new CellValueSpec.Text(source);
      case "RICH_XML" ->
          new CellValueSpec.RichText(List.of(new RichTextRunSpec(source, Optional.empty())));
      case "LONG_RICH" ->
          new CellValueSpec.RichText(
              List.of(
                  new RichTextRunSpec(source, Optional.empty()),
                  new RichTextRunSpec(source, Optional.empty())));
      case "FORMULA" -> new CellValueSpec.Formula(source);
      case "RAW_FORMULA" -> new CellValueSpec.RawFormula(source);
      default -> throw new AssertionError(kind);
    };
  }

  private static WorkbookPlan requestWithRepeatedValue(CellValueSpec value) {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExecutionPolicySpec.defaults(),
        FormulaEnvironmentSpec.empty(),
        List.of(
            new MutationStep(
                new StepId("first"),
                new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                new CellMutationAction.SetCell(value)),
            new MutationStep(
                new StepId("second"),
                new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("B1")),
                new CellMutationAction.SetCell(value))));
  }

  private GridGrindRequestInputs inputsFor(byte[] bytes, boolean stdin) {
    var grant =
        new ExecutionGrant(
            bytes.length,
            stdin
                ? List.of(new ExecutionGrant.ReadAuthority.Stdin())
                : List.of(new ExecutionGrant.ReadAuthority.File(root.resolve("input.txt"))),
            List.of("SET_CELL"),
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new ExecutionGrant.PublicationAuthority.None(),
            List.of(),
            AcceptancePolicy.minimum());
    return stdin
        ? new GridGrindRequestInputs(root, root.resolve("scratch"), bytes, grant)
        : new GridGrindRequestInputs(root, root.resolve("scratch"), grant);
  }

  private static String expectedOperand(String kind) {
    return switch (kind) {
      case "RICH_XML" -> "action.value.runs[0].text.";
      case "LONG_RICH" -> "action.value.runs[1].text.";
      default -> "action.value.value.";
    };
  }
}
