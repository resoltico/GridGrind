package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.FormulaEnvironmentSupport;
import java.util.List;
import java.util.Objects;

/** One named UDF toolpack registered for workbook formula evaluation. */
public record FormulaUdfToolpackInput(String name, List<FormulaUdfFunctionInput> functions) {
  public FormulaUdfToolpackInput {
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
            FormulaUdfFunctionInput::name);
  }
}
