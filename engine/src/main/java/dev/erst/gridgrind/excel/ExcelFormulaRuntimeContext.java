package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.formula.FormulaMissingWorkbookPolicy;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;

/** Immutable snapshot of the evaluator configuration used by readback and analysis code. */
record ExcelFormulaRuntimeContext(
    Set<String> externalWorkbookNames,
    FormulaMissingWorkbookPolicy missingWorkbookPolicy,
    Set<String> userDefinedFunctionNames) {
  ExcelFormulaRuntimeContext {
    externalWorkbookNames = Set.copyOf(Objects.requireNonNull(externalWorkbookNames));
    Objects.requireNonNull(missingWorkbookPolicy, "missingWorkbookPolicy must not be null");
    userDefinedFunctionNames = Set.copyOf(Objects.requireNonNull(userDefinedFunctionNames));
  }

  static ExcelFormulaRuntimeContext from(
      dev.erst.gridgrind.core.model.formula.ResolvedFormulaEnvironment environment) {
    return new ExcelFormulaRuntimeContext(
        environment.externalWorkbooks().stream()
            .map(dev.erst.gridgrind.core.model.formula.ResolvedExternalWorkbook::workbookName)
            .collect(java.util.stream.Collectors.toUnmodifiableSet()),
        environment.missingWorkbookPolicy(),
        environment.udfToolpacks().stream()
            .flatMap(pack -> pack.functions().stream())
            .map(dev.erst.gridgrind.core.model.formula.FormulaUdfFunctionSpec::name)
            .collect(java.util.stream.Collectors.toUnmodifiableSet()));
  }

  boolean hasExternalWorkbookBinding(String workbookName) {
    Objects.requireNonNull(workbookName, "workbookName must not be null");
    return externalWorkbookNames.stream()
        .anyMatch(candidate -> candidate.equalsIgnoreCase(workbookName));
  }

  boolean hasUserDefinedFunction(String functionName) {
    Objects.requireNonNull(functionName, "functionName must not be null");
    String normalized = functionName.toUpperCase(Locale.ROOT);
    return userDefinedFunctionNames.stream()
        .map(name -> name.toUpperCase(Locale.ROOT))
        .anyMatch(normalized::equals);
  }
}
