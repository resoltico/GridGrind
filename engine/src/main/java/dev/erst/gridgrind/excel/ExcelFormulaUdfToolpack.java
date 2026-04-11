package dev.erst.gridgrind.excel;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;

/** One named collection of template-backed user-defined functions. */
public record ExcelFormulaUdfToolpack(String name, List<ExcelFormulaUdfFunction> functions) {
  public ExcelFormulaUdfToolpack {
    Objects.requireNonNull(name, "name must not be null");
    if (name.isBlank()) {
      throw new IllegalArgumentException("name must not be blank");
    }
    Objects.requireNonNull(functions, "functions must not be null");
    functions = List.copyOf(functions);
    if (functions.isEmpty()) {
      throw new IllegalArgumentException("functions must not be empty");
    }
    Set<String> seen = new LinkedHashSet<>();
    for (ExcelFormulaUdfFunction function : functions) {
      Objects.requireNonNull(function, "functions must not contain nulls");
      String normalized = function.name().toUpperCase(Locale.ROOT);
      if (!seen.add(normalized)) {
        throw new IllegalArgumentException(
            "functions must not contain duplicate names: " + function.name());
      }
    }
  }
}
