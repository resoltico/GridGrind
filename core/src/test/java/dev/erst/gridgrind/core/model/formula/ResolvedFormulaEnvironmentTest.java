package dev.erst.gridgrind.core.model.formula;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Materialized environments own path lists and share source-free UDF definitions. */
class ResolvedFormulaEnvironmentTest {
  @Test
  void ownsBindingsAndSharesCanonicalToolpacks() {
    var bindings =
        new ArrayList<>(List.of(new ResolvedExternalWorkbook("Rates.xlsx", Path.of("rates.xlsx"))));
    var function = new FormulaUdfFunctionSpec("DOUBLE", 1, "ARG1*2");
    var pack = new FormulaUdfToolpackSpec("math", List.of(function));
    var environment =
        new ResolvedFormulaEnvironment(bindings, FormulaMissingWorkbookPolicy.ERROR, List.of(pack));
    bindings.clear();
    assertEquals(1, environment.externalWorkbooks().size());
    assertSame(pack, environment.udfToolpacks().getFirst());
    assertSame(function, environment.udfToolpacks().getFirst().functions().getFirst());
    assertFalse(environment.isDefault());
    assertTrue(ResolvedFormulaEnvironment.defaults().isDefault());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedFormulaEnvironment(
                List.of(
                    new ResolvedExternalWorkbook("a", Path.of("a.xlsx")),
                    new ResolvedExternalWorkbook("A", Path.of("b.xlsx"))),
                FormulaMissingWorkbookPolicy.ERROR,
                List.of()));
  }
}
