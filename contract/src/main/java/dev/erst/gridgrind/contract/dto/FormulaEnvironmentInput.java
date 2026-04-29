package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.excel.foundation.FormulaEnvironmentSupport;
import java.util.List;

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
        FormulaEnvironmentSupport.copyOptionalDistinctNamedValues(
            externalWorkbooks,
            "externalWorkbooks",
            "externalWorkbooks must not contain duplicate workbookName values: ",
            FormulaExternalWorkbookInput::workbookName);
    java.util.Objects.requireNonNull(
        missingWorkbookPolicy, "missingWorkbookPolicy must not be null");
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

  @JsonCreator
  static FormulaEnvironmentInput create(
      @JsonProperty("externalWorkbooks") List<FormulaExternalWorkbookInput> externalWorkbooks,
      @JsonProperty("missingWorkbookPolicy") FormulaMissingWorkbookPolicy missingWorkbookPolicy,
      @JsonProperty("udfToolpacks") List<FormulaUdfToolpackInput> udfToolpacks) {
    return new FormulaEnvironmentInput(
        externalWorkbooks == null ? List.of() : externalWorkbooks,
        missingWorkbookPolicy == null ? FormulaMissingWorkbookPolicy.ERROR : missingWorkbookPolicy,
        udfToolpacks == null ? List.of() : udfToolpacks);
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
      return 0;
    }
  }
}
