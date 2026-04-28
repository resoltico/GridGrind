package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.FormulaEnvironmentSupport;
import java.util.List;
import java.util.Objects;

/** One named collection of template-backed user-defined functions. */
public record ExcelFormulaUdfToolpack(String name, List<ExcelFormulaUdfFunction> functions) {
  public ExcelFormulaUdfToolpack {
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
            ExcelFormulaUdfFunction::name);
  }
}
