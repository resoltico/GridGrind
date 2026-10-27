package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.limits.Limit;

/** Shared dangerous-formula rejection applied to all formula-bearing input types. */
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_CHART_TITLE_FORMULA_INJECTION,
  dev.erst.gridgrind.core.limits.Limit.ID_DDE_REJECTION_IN_ALL_FORMULA_INPUTS,
  dev.erst.gridgrind.core.limits.Limit.ID_NAMED_RANGE_FORMULA_INJECTION,
  dev.erst.gridgrind.core.limits.Limit.ID_UDF_FORMULA_TEMPLATE_INJECTION_DEFENSE_IN_DEPTH,
  dev.erst.gridgrind.core.limits.Limit.ID_WEBSERVICE_FORMULA_REJECTION
})
public final class FormulaInputSecurity {
  private FormulaInputSecurity() {}

  /** Rejects formula functions that cross the permitted execution boundary. */
  public static void rejectDde(String formula) {
    String toCheck = formula;
    if (toCheck.startsWith("{=") && toCheck.endsWith("}")) {
      toCheck = toCheck.substring(2, toCheck.length() - 1);
    } else if (toCheck.startsWith("=")) {
      toCheck = toCheck.substring(1);
    }
    if (toCheck.regionMatches(
        true, 0, Limit.DDE_FUNCTION_PREFIX, 0, Limit.DDE_FUNCTION_PREFIX.length())) {
      throw new IllegalArgumentException(Limit.DDE_REJECTION);
    }
    if (toCheck.regionMatches(
        true, 0, Limit.WEBSERVICE_FUNCTION_PREFIX, 0, Limit.WEBSERVICE_FUNCTION_PREFIX.length())) {
      throw new IllegalArgumentException(Limit.WEBSERVICE_REJECTION);
    }
  }
}
