package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.FormulaEnvironmentSupport;
import java.util.List;

/** Workbook-scoped formula-evaluation environment with external workbook bindings and UDFs. */
public record ExcelFormulaEnvironment(
    List<ExcelFormulaExternalWorkbookBinding> externalWorkbooks,
    ExcelFormulaMissingWorkbookPolicy missingWorkbookPolicy,
    List<ExcelFormulaUdfToolpack> udfToolpacks) {
  public ExcelFormulaEnvironment {
    externalWorkbooks =
        FormulaEnvironmentSupport.copyOptionalDistinctNamedValues(
            externalWorkbooks,
            "externalWorkbooks",
            "externalWorkbooks must not contain duplicate workbookName values: ",
            ExcelFormulaExternalWorkbookBinding::workbookName);
    missingWorkbookPolicy =
        missingWorkbookPolicy == null
            ? ExcelFormulaMissingWorkbookPolicy.ERROR
            : missingWorkbookPolicy;
    udfToolpacks =
        FormulaEnvironmentSupport.copyOptionalDistinctNamedValues(
            udfToolpacks,
            "udfToolpacks",
            "udfToolpacks must not contain duplicate names: ",
            ExcelFormulaUdfToolpack::name);
    FormulaEnvironmentSupport.requireDistinctNestedNames(
        udfToolpacks,
        ExcelFormulaUdfToolpack::functions,
        ExcelFormulaUdfFunction::name,
        "udfToolpacks must not define duplicate function names across toolpacks: ");
  }

  /**
   * Returns the default environment with no external bindings, no UDFs, and strict missing refs.
   */
  public static ExcelFormulaEnvironment defaults() {
    return new ExcelFormulaEnvironment(
        List.of(), ExcelFormulaMissingWorkbookPolicy.ERROR, List.of());
  }

  /** Returns whether this environment is the default evaluator state. */
  public boolean isDefault() {
    return externalWorkbooks.isEmpty()
        && missingWorkbookPolicy == ExcelFormulaMissingWorkbookPolicy.ERROR
        && udfToolpacks.isEmpty();
  }

  /** Returns the runtime context snapshot exposed to analysis code. */
  ExcelFormulaRuntimeContext runtimeContext() {
    return new ExcelFormulaRuntimeContext(
        externalWorkbooks.stream()
            .map(ExcelFormulaExternalWorkbookBinding::workbookName)
            .collect(java.util.stream.Collectors.toUnmodifiableSet()),
        missingWorkbookPolicy,
        udfToolpacks.stream()
            .flatMap(toolpack -> toolpack.functions().stream())
            .map(ExcelFormulaUdfFunction::name)
            .collect(java.util.stream.Collectors.toUnmodifiableSet()));
  }
}
