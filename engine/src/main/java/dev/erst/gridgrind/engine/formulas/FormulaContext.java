package dev.erst.gridgrind.engine.formulas;

import dev.erst.gridgrind.core.model.formula.FormulaMissingWorkbookPolicy;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;

/** Immutable snapshot of the evaluator configuration used by readback and analysis code. */
public record FormulaContext(
    Set<String> externalWorkbookNames,
    FormulaMissingWorkbookPolicy missingWorkbookPolicy,
    Set<String> userDefinedFunctionNames) {
  public FormulaContext {
    externalWorkbookNames = Set.copyOf(Objects.requireNonNull(externalWorkbookNames));
    Objects.requireNonNull(missingWorkbookPolicy, "missingWorkbookPolicy must not be null");
    userDefinedFunctionNames = Set.copyOf(Objects.requireNonNull(userDefinedFunctionNames));
  }

  /** Captures evaluator bindings and policies from the resolved formula environment. */
  public static FormulaContext from(
      dev.erst.gridgrind.core.model.formula.ResolvedFormulaEnvironment environment) {
    return new FormulaContext(
        environment.externalWorkbooks().stream()
            .map(dev.erst.gridgrind.core.model.formula.ResolvedExternalWorkbook::workbookName)
            .collect(java.util.stream.Collectors.toUnmodifiableSet()),
        environment.missingWorkbookPolicy(),
        environment.udfToolpacks().stream()
            .flatMap(pack -> pack.functions().stream())
            .map(dev.erst.gridgrind.core.model.formula.FormulaUdfFunctionSpec::name)
            .collect(java.util.stream.Collectors.toUnmodifiableSet()));
  }

  /** Checks external workbook bindings without case sensitivity. */
  public boolean hasExternalWorkbookBinding(String workbookName) {
    Objects.requireNonNull(workbookName, "workbookName must not be null");
    return externalWorkbookNames.stream()
        .anyMatch(candidate -> candidate.equalsIgnoreCase(workbookName));
  }

  /** Checks registered function names without case sensitivity. */
  public boolean hasUserDefinedFunction(String functionName) {
    Objects.requireNonNull(functionName, "functionName must not be null");
    String normalized = functionName.toUpperCase(Locale.ROOT);
    return userDefinedFunctionNames.stream()
        .map(name -> name.toUpperCase(Locale.ROOT))
        .anyMatch(normalized::equals);
  }
}
