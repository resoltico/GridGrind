package dev.erst.gridgrind.protocol.dto;

import static org.junit.jupiter.api.Assertions.*;

import java.util.List;
import org.junit.jupiter.api.Test;

/** Validation coverage for formula-environment request DTOs. */
class FormulaEnvironmentInputTest {
  @Test
  void normalizesAndCopiesFormulaEnvironmentInputs() {
    FormulaExternalWorkbookInput externalWorkbook =
        new FormulaExternalWorkbookInput("Rates.xlsx", "tmp/rates.xlsx");
    FormulaUdfFunctionInput udfFunction =
        new FormulaUdfFunctionInput("_pkg.DOUBLE2", 1, null, " = ARG1 * 2 ");
    List<FormulaExternalWorkbookInput> externalWorkbooks =
        new java.util.ArrayList<>(List.of(externalWorkbook));
    List<FormulaUdfToolpackInput> udfToolpacks =
        new java.util.ArrayList<>(
            List.of(new FormulaUdfToolpackInput("math", List.of(udfFunction))));

    FormulaEnvironmentInput environment =
        new FormulaEnvironmentInput(externalWorkbooks, null, udfToolpacks);
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
                .add(new FormulaExternalWorkbookInput("Other.xlsx", "tmp/other.xlsx")));

    FormulaEnvironmentInput defaults = new FormulaEnvironmentInput(null, null, null);
    assertTrue(defaults.isEmpty());
    assertEquals(List.of(), defaults.externalWorkbooks());
    assertEquals(List.of(), defaults.udfToolpacks());
  }

  @Test
  void rejectsInvalidFormulaEnvironmentInputs() {
    FormulaUdfFunctionInput validFunction = new FormulaUdfFunctionInput("DOUBLE", 1, 1, "ARG1");

    assertThrows(NullPointerException.class, () -> new FormulaCellTargetInput("Budget", null));
    assertThrows(IllegalArgumentException.class, () -> new FormulaCellTargetInput("Budget", " "));

    assertThrows(
        IllegalArgumentException.class,
        () -> new FormulaExternalWorkbookInput(" ", "tmp/rates.xlsx"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new FormulaExternalWorkbookInput("[Rates.xlsx]", "tmp/rates.xlsx"));

    assertThrows(NullPointerException.class, () -> new FormulaUdfFunctionInput(null, 1, 1, "ARG1"));
    assertThrows(
        IllegalArgumentException.class, () -> new FormulaUdfFunctionInput(" ", 1, 1, "ARG1"));
    assertThrows(
        IllegalArgumentException.class, () -> new FormulaUdfFunctionInput("1BAD", 1, 1, "ARG1"));
    assertThrows(
        NullPointerException.class, () -> new FormulaUdfFunctionInput("DOUBLE", null, 1, "ARG1"));
    assertThrows(
        IllegalArgumentException.class, () -> new FormulaUdfFunctionInput("DOUBLE", -1, 1, "ARG1"));
    assertThrows(
        IllegalArgumentException.class, () -> new FormulaUdfFunctionInput("DOUBLE", 2, 1, "ARG1"));
    assertThrows(
        NullPointerException.class, () -> new FormulaUdfFunctionInput("DOUBLE", 1, 1, null));
    assertThrows(
        IllegalArgumentException.class, () -> new FormulaUdfFunctionInput("DOUBLE", 1, 1, " "));
    assertThrows(
        IllegalArgumentException.class, () -> new FormulaUdfFunctionInput("DOUBLE", 1, 1, "ARG2"));

    assertThrows(
        NullPointerException.class,
        () -> new FormulaUdfToolpackInput(null, List.of(validFunction)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new FormulaUdfToolpackInput(" ", List.of(validFunction)));
    assertThrows(NullPointerException.class, () -> new FormulaUdfToolpackInput("math", null));
    assertThrows(
        IllegalArgumentException.class, () -> new FormulaUdfToolpackInput("math", List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new FormulaUdfToolpackInput("math", List.of((FormulaUdfFunctionInput) null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new FormulaUdfToolpackInput(
                "math",
                List.of(validFunction, new FormulaUdfFunctionInput("double", 1, 1, "ARG1"))));

    assertThrows(
        NullPointerException.class,
        () ->
            new FormulaEnvironmentInput(
                List.of((FormulaExternalWorkbookInput) null),
                FormulaMissingWorkbookPolicy.ERROR,
                List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new FormulaEnvironmentInput(
                List.of(
                    new FormulaExternalWorkbookInput("Rates.xlsx", "tmp/rates.xlsx"),
                    new FormulaExternalWorkbookInput("rates.xlsx", "tmp/rates-2.xlsx")),
                FormulaMissingWorkbookPolicy.ERROR,
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new FormulaEnvironmentInput(
                List.of(),
                FormulaMissingWorkbookPolicy.ERROR,
                List.of((FormulaUdfToolpackInput) null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new FormulaEnvironmentInput(
                List.of(),
                FormulaMissingWorkbookPolicy.ERROR,
                List.of(
                    new FormulaUdfToolpackInput("math", List.of(validFunction)),
                    new FormulaUdfToolpackInput(
                        "Math", List.of(new FormulaUdfFunctionInput("TRIPLE", 1, 1, "ARG1*3"))))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new FormulaEnvironmentInput(
                List.of(),
                FormulaMissingWorkbookPolicy.ERROR,
                List.of(
                    new FormulaUdfToolpackInput("math", List.of(validFunction)),
                    new FormulaUdfToolpackInput(
                        "stats", List.of(new FormulaUdfFunctionInput("double", 1, 1, "ARG1"))))));
  }
}
