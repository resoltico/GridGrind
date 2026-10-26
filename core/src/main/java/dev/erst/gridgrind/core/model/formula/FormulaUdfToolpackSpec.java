package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.vocabulary.FormulaEnvironmentSupport;
import java.util.List;
import java.util.Objects;

/** One named UDF toolpack registered for workbook formula evaluation. */
@CatalogType(
    id = "FormulaUdfToolpackSpec",
    summary = "One named collection of template-backed user-defined functions.",
    displayOrder = 27)
public record FormulaUdfToolpackSpec(String name, List<FormulaUdfFunctionSpec> functions) {
  public FormulaUdfToolpackSpec {
    Objects.requireNonNull(name, "name must not be null");
    if (name.isBlank()) {
      throw new IllegalArgumentException("name must not be blank");
    }
    functions =
        FormulaEnvironmentSupport.copyRequiredDistinctNamedValues(
            functions,
            "functions",
            "functions must not be empty",
            "functions must not contain duplicate names: ",
            FormulaUdfFunctionSpec::name);
  }
}
