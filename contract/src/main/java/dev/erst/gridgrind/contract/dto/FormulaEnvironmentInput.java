package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonIgnore;
import dev.erst.gridgrind.excel.foundation.FormulaEnvironmentSupport;
import java.util.List;
import java.util.Objects;

/** Request-scoped formula-evaluation environment for external workbooks and UDFs. */
public record FormulaEnvironmentInput(
    List<FormulaExternalWorkbookInput> externalWorkbooks,
    FormulaMissingWorkbookPolicy missingWorkbookPolicy,
    List<FormulaUdfToolpackInput> udfToolpacks) {
  /** Returns the default empty formula environment with no external workbooks or UDFs. */
  public static FormulaEnvironmentInput empty() {
    return new FormulaEnvironmentInput(List.of(), FormulaMissingWorkbookPolicy.ERROR, List.of());
  }

  public FormulaEnvironmentInput {
    externalWorkbooks =
        List.copyOf(
            Objects.requireNonNull(externalWorkbooks, "externalWorkbooks must not be null"));
    externalWorkbooks =
        FormulaEnvironmentSupport.copyOptionalDistinctNamedValues(
            externalWorkbooks,
            "externalWorkbooks",
            "externalWorkbooks must not contain duplicate workbookName values: ",
            FormulaExternalWorkbookInput::workbookName);
    Objects.requireNonNull(missingWorkbookPolicy, "missingWorkbookPolicy must not be null");
    udfToolpacks =
        List.copyOf(Objects.requireNonNull(udfToolpacks, "udfToolpacks must not be null"));
    udfToolpacks =
        FormulaEnvironmentSupport.copyOptionalDistinctNamedValues(
            udfToolpacks,
            "udfToolpacks",
            "udfToolpacks must not contain duplicate names: ",
            FormulaUdfToolpackInput::name);
    FormulaEnvironmentSupport.requireDistinctNestedNames(
        udfToolpacks,
        FormulaUdfToolpackInput::functions,
        FormulaUdfFunctionInput::name,
        "udfToolpacks must not define duplicate function names across toolpacks: ");
  }

  /** Returns whether this environment carries any behavior beyond the default evaluator state. */
  @JsonIgnore
  public boolean isEmpty() {
    return externalWorkbooks.isEmpty()
        && missingWorkbookPolicy == FormulaMissingWorkbookPolicy.ERROR
        && udfToolpacks.isEmpty();
  }

  /** Custom Jackson inclusion filter that omits an empty formula-environment object. */
  public static final class EmptyFilter {
    @Override
    public boolean equals(Object other) {
      return other == null || (other instanceof FormulaEnvironmentInput input && input.isEmpty());
    }

    @Override
    public int hashCode() {
      return EmptyFilter.class.hashCode();
    }
  }
}
