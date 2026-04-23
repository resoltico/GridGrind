package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.FormulaExternalWorkbookInput;
import dev.erst.gridgrind.contract.dto.FormulaMissingWorkbookPolicy;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for request-scoped formula-environment path resolution. */
class FormulaEnvironmentConverterTest {
  @Test
  void rootsRelativeExternalWorkbookBindingsInTheProvidedWorkingDirectory() {
    FormulaEnvironmentInput input =
        new FormulaEnvironmentInput(
            List.of(new FormulaExternalWorkbookInput("referenced.xlsx", "refs/q1.xlsx")),
            FormulaMissingWorkbookPolicy.ERROR,
            List.of());
    Path workingDirectory = Path.of("/tmp/gridgrind-request-bundle");

    assertEquals(
        workingDirectory.resolve("refs/q1.xlsx").normalize(),
        FormulaEnvironmentConverter.toExcelFormulaEnvironment(input, workingDirectory)
            .externalWorkbooks()
            .getFirst()
            .path());
  }
}
