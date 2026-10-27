package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.vocabulary.FormulaEnvironmentSupport;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;

/** One named UDF toolpack registered for workbook formula evaluation. */
@CatalogType(
    id = "FormulaUdfToolpackSpec",
    summary = "One named collection of template-backed user-defined functions.",
    displayOrder = 27)
public record FormulaUdfToolpackSpec(String name, List<FormulaUdfFunctionSpec> functions) {
  public FormulaUdfToolpackSpec {
    name = TextValidation.requireNonBlank(name, "name");
    functions =
        FormulaEnvironmentSupport.copyRequiredDistinctNamedValues(
            functions,
            "functions",
            "functions must not be empty",
            "functions must not contain duplicate names: ",
            FormulaUdfFunctionSpec::name);
  }
}
