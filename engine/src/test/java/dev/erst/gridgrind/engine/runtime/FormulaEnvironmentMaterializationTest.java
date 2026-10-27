package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.formula.FormulaExternalWorkbookSpec;
import dev.erst.gridgrind.core.model.formula.FormulaMissingWorkbookPolicy;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Tests request-scoped formula workbooks are materialized through prepared file capabilities. */
class FormulaEnvironmentMaterializationTest {
  @TempDir Path temporaryDirectory;

  @Test
  void materializesRelativeExternalWorkbookBindingsThroughTheExecutionRoot() throws Exception {
    Path references = Files.createDirectory(temporaryDirectory.resolve("refs"));
    Path externalWorkbook = Files.write(references.resolve("q1.xlsx"), new byte[] {1, 2, 3});
    FormulaEnvironmentSpec input =
        new FormulaEnvironmentSpec(
            List.of(new FormulaExternalWorkbookSpec("referenced.xlsx", "refs/q1.xlsx")),
            FormulaMissingWorkbookPolicy.ERROR,
            List.of());

    try (var prepared =
        ExecutionInputBindingsFixtureSupport.preparedBindings(
            temporaryDirectory, List.of(externalWorkbook))) {
      Path materialized =
          WorkbookLifecycle.materializeFormulaEnvironment(input, prepared.bindings())
              .externalWorkbooks()
              .getFirst()
              .path();
      assertArrayEquals(Files.readAllBytes(externalWorkbook), Files.readAllBytes(materialized));
      assertTrue(materialized.startsWith(temporaryDirectory.resolve(".gridgrind/tmp")));
    }
  }
}
