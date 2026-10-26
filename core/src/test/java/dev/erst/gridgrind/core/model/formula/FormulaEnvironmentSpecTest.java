package dev.erst.gridgrind.core.model.formula;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.problem.InvalidRequestException;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Validation coverage for formula-environment request DTOs. */
class FormulaEnvironmentSpecTest {
  @Test
  void normalizesAndCopiesFormulaEnvironmentInputs() {
    FormulaExternalWorkbookSpec externalWorkbook =
        new FormulaExternalWorkbookSpec("Rates.xlsx", "tmp/rates.xlsx");
    FormulaUdfFunctionSpec udfFunction =
        new FormulaUdfFunctionSpec("_pkg.DOUBLE2", 1, 1, " = ARG1 * 2 ");
    List<FormulaExternalWorkbookSpec> externalWorkbooks =
        new java.util.ArrayList<>(List.of(externalWorkbook));
    List<FormulaUdfToolpackSpec> udfToolpacks =
        new java.util.ArrayList<>(
            List.of(new FormulaUdfToolpackSpec("math", List.of(udfFunction))));

    FormulaEnvironmentSpec environment =
        new FormulaEnvironmentSpec(
            externalWorkbooks, FormulaMissingWorkbookPolicy.ERROR, udfToolpacks);
    externalWorkbooks.clear();
    udfToolpacks.clear();

    assertFalse(environment.isEmpty());
    assertEquals(FormulaMissingWorkbookPolicy.ERROR, environment.missingWorkbookPolicy());
    assertEquals(List.of(externalWorkbook), environment.externalWorkbooks());
    assertEquals(
        "_pkg.DOUBLE2", environment.udfToolpacks().getFirst().functions().getFirst().name());
    assertEquals(
        1, environment.udfToolpacks().getFirst().functions().getFirst().maximumArgumentCount());
    assertEquals(
        "ARG1 * 2", environment.udfToolpacks().getFirst().functions().getFirst().formulaTemplate());
    assertThrows(
        UnsupportedOperationException.class,
        () ->
            environment
                .externalWorkbooks()
                .add(new FormulaExternalWorkbookSpec("Other.xlsx", "tmp/other.xlsx")));

    FormulaEnvironmentSpec defaults = FormulaEnvironmentSpec.empty();
    assertTrue(defaults.isEmpty());
    assertEquals(List.of(), defaults.externalWorkbooks());
    assertEquals(List.of(), defaults.udfToolpacks());
  }

  @Test
  void rejectsInvalidFormulaEnvironmentInputs() {
    FormulaUdfFunctionSpec validFunction = new FormulaUdfFunctionSpec("DOUBLE", 1, 1, "ARG1");

    assertThrows(
        IllegalArgumentException.class,
        () -> new FormulaExternalWorkbookSpec(" ", "tmp/rates.xlsx"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new FormulaExternalWorkbookSpec("[Rates.xlsx]", "tmp/rates.xlsx"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new FormulaExternalWorkbookSpec("Rates].xlsx", "tmp/rates.xlsx"));
    InvalidRequestException invalidExternalWorkbookPath =
        assertThrows(
            InvalidRequestException.class,
            () -> new FormulaExternalWorkbookSpec("Rates.xlsx", "tmp/rates.txt"));
    assertEquals(Optional.of("path"), invalidExternalWorkbookPath.jsonPath());

    assertThrows(NullPointerException.class, () -> new FormulaUdfFunctionSpec(null, 1, 1, "ARG1"));
    assertThrows(
        IllegalArgumentException.class, () -> new FormulaUdfFunctionSpec(" ", 1, 1, "ARG1"));
    assertThrows(
        IllegalArgumentException.class, () -> new FormulaUdfFunctionSpec("1BAD", 1, 1, "ARG1"));
    assertEquals(1, new FormulaUdfFunctionSpec("DOUBLE", 1, "ARG1").maximumArgumentCount());
    assertThrows(
        NullPointerException.class,
        () -> new FormulaUdfFunctionSpec("DOUBLE", 1, (Integer) null, "ARG1"));
    assertThrows(
        IllegalArgumentException.class, () -> new FormulaUdfFunctionSpec("DOUBLE", -1, 1, "ARG1"));
    assertThrows(
        IllegalArgumentException.class, () -> new FormulaUdfFunctionSpec("DOUBLE", 2, 1, "ARG1"));
    assertThrows(
        NullPointerException.class, () -> new FormulaUdfFunctionSpec("DOUBLE", 1, 1, null));
    assertThrows(
        IllegalArgumentException.class, () -> new FormulaUdfFunctionSpec("DOUBLE", 1, 1, " "));
    assertThrows(
        IllegalArgumentException.class, () -> new FormulaUdfFunctionSpec("DOUBLE", 1, 1, "ARG2"));

    assertThrows(
        NullPointerException.class, () -> new FormulaUdfToolpackSpec(null, List.of(validFunction)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new FormulaUdfToolpackSpec(" ", List.of(validFunction)));
    assertThrows(NullPointerException.class, () -> new FormulaUdfToolpackSpec("math", null));
    assertThrows(
        IllegalArgumentException.class, () -> new FormulaUdfToolpackSpec("math", List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new FormulaUdfToolpackSpec("math", List.of((FormulaUdfFunctionSpec) null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new FormulaUdfToolpackSpec(
                "math",
                List.of(validFunction, new FormulaUdfFunctionSpec("double", 1, 1, "ARG1"))));

    assertThrows(
        NullPointerException.class,
        () ->
            new FormulaEnvironmentSpec(
                List.of((FormulaExternalWorkbookSpec) null),
                FormulaMissingWorkbookPolicy.ERROR,
                List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new FormulaEnvironmentSpec(
                List.of(
                    new FormulaExternalWorkbookSpec("Rates.xlsx", "tmp/rates.xlsx"),
                    new FormulaExternalWorkbookSpec("rates.xlsx", "tmp/rates-2.xlsx")),
                FormulaMissingWorkbookPolicy.ERROR,
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new FormulaEnvironmentSpec(
                List.of(),
                FormulaMissingWorkbookPolicy.ERROR,
                List.of((FormulaUdfToolpackSpec) null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new FormulaEnvironmentSpec(
                List.of(),
                FormulaMissingWorkbookPolicy.ERROR,
                List.of(
                    new FormulaUdfToolpackSpec("math", List.of(validFunction)),
                    new FormulaUdfToolpackSpec(
                        "Math", List.of(new FormulaUdfFunctionSpec("TRIPLE", 1, 1, "ARG1*3"))))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new FormulaEnvironmentSpec(
                List.of(),
                FormulaMissingWorkbookPolicy.ERROR,
                List.of(
                    new FormulaUdfToolpackSpec("math", List.of(validFunction)),
                    new FormulaUdfToolpackSpec(
                        "stats", List.of(new FormulaUdfFunctionSpec("double", 1, 1, "ARG1"))))));
  }
}
