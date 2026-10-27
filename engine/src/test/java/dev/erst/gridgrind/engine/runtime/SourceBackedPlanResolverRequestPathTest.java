package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/**
 * Verifies source-backed resolution consumes the phase-four path capability when one is present.
 */
class SourceBackedPlanResolverRequestPathTest {
  @TempDir Path root;

  @Test
  void usesThePreparedRequestPathCapabilityWithoutOpeningAnotherLifecycle() throws Exception {
    WorkbookPlan plan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of());
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            root,
            Files.createDirectory(root.resolve("temp")),
            ExecutionGrantTestSupport.noPublication());

    try (RequestPathAccess access =
        new RequestPathAccess(root, bindings.tempFileFactory(), bindings.executionGrant())) {
      assertEquals(
          plan, SourceBackedPlanResolver.resolve(plan, bindings.withRequestPathAccess(access)));
    }
  }

  @Test
  void reportsInvalidBinaryFilePathsWithoutBypassingPathMaterialization() throws Exception {
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            root,
            Files.createDirectory(root.resolve("temp")),
            ExecutionGrantTestSupport.noPublication());

    InputSourceReadException failure =
        assertThrows(
            InputSourceReadException.class,
            () ->
                SourceBackedPlanResolver.resolveBinarySource(
                    new BinarySourceSpec.File("\u0000"), bindings, "picture payload"));

    assertEquals("picture payload", failure.inputKind());
  }

  @Test
  void preservesTheDedicatedEscapeDiagnosticForSourceBackedFiles() throws Exception {
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            root,
            Files.createDirectory(root.resolve("temp")),
            ExecutionGrantTestSupport.noPublication());

    assertThrows(
        RequestPathEscapeException.class,
        () ->
            SourceBackedPlanResolver.resolveBinarySource(
                new BinarySourceSpec.File("../outside.bin"), bindings, "picture payload"));
  }

  @Test
  void preservesUnresolvedFormulaSourcesWhileCollectingTheirIndependentFailures() throws Exception {
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            root,
            Files.createDirectory(root.resolve("temp")),
            ExecutionGrantTestSupport.noPublication());
    InputResolutionFailures failures = new InputResolutionFailures();
    TextSourceSpec.Utf8File missingFormula = TextSourceSpec.utf8File("missing-formula.txt");

    assertSame(
        missingFormula,
        SourceBackedPlanResolver.resolveFormulaSource(
            missingFormula, bindings.collectingInputResolutionFailures(failures)));

    InputResolutionBatchException failure =
        assertThrows(InputResolutionBatchException.class, failures::throwIfAny);
    assertInstanceOf(InputSourceNotFoundException.class, failure.failures().getFirst().exception());
  }
}
