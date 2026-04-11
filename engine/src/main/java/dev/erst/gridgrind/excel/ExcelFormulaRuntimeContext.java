package dev.erst.gridgrind.excel;

import java.util.Locale;
import java.util.Objects;
import java.util.Set;

/** Immutable snapshot of the evaluator configuration used by readback and analysis code. */
record ExcelFormulaRuntimeContext(
    Set<String> externalWorkbookNames,
    ExcelFormulaMissingWorkbookPolicy missingWorkbookPolicy,
    Set<String> userDefinedFunctionNames) {
  ExcelFormulaRuntimeContext {
    externalWorkbookNames = Set.copyOf(Objects.requireNonNull(externalWorkbookNames));
    Objects.requireNonNull(missingWorkbookPolicy, "missingWorkbookPolicy must not be null");
    userDefinedFunctionNames = Set.copyOf(Objects.requireNonNull(userDefinedFunctionNames));
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
