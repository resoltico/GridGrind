package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
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
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(root, Files.createDirectory(root.resolve("temp")));

    try (RequestPathAccess access = new RequestPathAccess(root, bindings.tempFileFactory())) {
      assertEquals(
          plan, SourceBackedPlanResolver.resolve(plan, bindings.withRequestPathAccess(access)));
    }
  }

  @Test
  void reportsInvalidBinaryFilePathsWithoutBypassingPathMaterialization() throws Exception {
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(root, Files.createDirectory(root.resolve("temp")));

    InputSourceReadException failure =
        assertThrows(
            InputSourceReadException.class,
            () ->
                SourceBackedPlanResolver.resolveBinarySource(
                    new BinarySourceInput.File("\u0000"), bindings, "picture payload"));

    assertEquals("picture payload", failure.inputKind());
  }

  @Test
  void preservesTheDedicatedEscapeDiagnosticForSourceBackedFiles() throws Exception {
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(root, Files.createDirectory(root.resolve("temp")));

    assertThrows(
        RequestPathEscapeException.class,
        () ->
            SourceBackedPlanResolver.resolveBinarySource(
                new BinarySourceInput.File("../outside.bin"), bindings, "picture payload"));
  }

  @Test
  void preservesUnresolvedFormulaSourcesWhileCollectingTheirIndependentFailures() throws Exception {
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(root, Files.createDirectory(root.resolve("temp")));
    InputResolutionFailures failures = new InputResolutionFailures();
    TextSourceInput.Utf8File missingFormula = TextSourceInput.utf8File("missing-formula.txt");

    assertSame(
        missingFormula,
        SourceBackedPlanResolver.resolveFormulaSource(
            missingFormula, bindings.collectingInputResolutionFailures(failures)));

    InputResolutionBatchException failure =
        assertThrows(InputResolutionBatchException.class, failures::throwIfAny);
    assertInstanceOf(InputSourceNotFoundException.class, failure.failures().getFirst().exception());
  }
}
