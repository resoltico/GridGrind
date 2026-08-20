package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.FormulaExternalWorkbookInput;
import dev.erst.gridgrind.contract.dto.FormulaMissingWorkbookPolicy;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Tests request-scoped formula workbooks are materialized through prepared file capabilities. */
class FormulaEnvironmentConverterTest {
  @TempDir Path temporaryDirectory;

  @Test
  void materializesRelativeExternalWorkbookBindingsThroughTheExecutionRoot() throws Exception {
    Path references = Files.createDirectory(temporaryDirectory.resolve("refs"));
    Path externalWorkbook = Files.write(references.resolve("q1.xlsx"), new byte[] {1, 2, 3});
    FormulaEnvironmentInput input =
        new FormulaEnvironmentInput(
            List.of(new FormulaExternalWorkbookInput("referenced.xlsx", "refs/q1.xlsx")),
            FormulaMissingWorkbookPolicy.ERROR,
            List.of());

    try (var prepared = ExecutionInputBindingsFixtureSupport.preparedBindings(temporaryDirectory)) {
      Path materialized =
          FormulaEnvironmentConverter.toExcelFormulaEnvironment(input, prepared.bindings())
              .externalWorkbooks()
              .getFirst()
              .path();
      assertArrayEquals(Files.readAllBytes(externalWorkbook), Files.readAllBytes(materialized));
      assertTrue(materialized.startsWith(temporaryDirectory.resolve(".gridgrind/tmp")));
    }
  }
}
