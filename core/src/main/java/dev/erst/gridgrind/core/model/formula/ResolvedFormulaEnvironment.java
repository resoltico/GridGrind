package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.vocabulary.FormulaEnvironmentSupport;
import java.util.List;

/** Workbook-scoped formula-evaluation environment with external workbook bindings and UDFs. */
public record ResolvedFormulaEnvironment(
    List<ResolvedExternalWorkbook> externalWorkbooks,
    FormulaMissingWorkbookPolicy missingWorkbookPolicy,
    List<FormulaUdfToolpackSpec> udfToolpacks) {
  public ResolvedFormulaEnvironment {
    externalWorkbooks =
        FormulaEnvironmentSupport.copyOptionalDistinctNamedValues(
            externalWorkbooks,
            "externalWorkbooks",
            "externalWorkbooks must not contain duplicate workbookName values: ",
            ResolvedExternalWorkbook::workbookName);
    missingWorkbookPolicy =
        missingWorkbookPolicy == null ? FormulaMissingWorkbookPolicy.ERROR : missingWorkbookPolicy;
    udfToolpacks =
        FormulaEnvironmentSupport.copyOptionalDistinctNamedValues(
            udfToolpacks,
            "udfToolpacks",
            "udfToolpacks must not contain duplicate names: ",
            FormulaUdfToolpackSpec::name);
    FormulaEnvironmentSupport.requireDistinctNestedNames(
        udfToolpacks,
        FormulaUdfToolpackSpec::functions,
        FormulaUdfFunctionSpec::name,
        "udfToolpacks must not define duplicate function names across toolpacks: ");
  }

  /**
   * Returns the default environment with no external bindings, no UDFs, and strict missing refs.
   */
  public static ResolvedFormulaEnvironment defaults() {
    return new ResolvedFormulaEnvironment(List.of(), FormulaMissingWorkbookPolicy.ERROR, List.of());
  }

  /** Returns whether this environment is the default evaluator state. */
  public boolean isDefault() {
    return externalWorkbooks.isEmpty()
        && missingWorkbookPolicy == FormulaMissingWorkbookPolicy.ERROR
        && udfToolpacks.isEmpty();
  }
}
