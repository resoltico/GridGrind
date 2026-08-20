package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
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
}
