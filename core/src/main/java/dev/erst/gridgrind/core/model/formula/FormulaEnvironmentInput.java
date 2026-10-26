package dev.erst.gridgrind.core.model.formula;

import com.fasterxml.jackson.annotation.JsonIgnore;
import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.vocabulary.FormulaEnvironmentSupport;
import java.util.List;
import java.util.Objects;

/** Request-scoped formula-evaluation environment for external workbooks and UDFs. */
@CatalogType(
    id = "FormulaEnvironmentInput",
    summary =
        "Optional request-scoped formula-evaluation environment covering external workbook bindings, missing-workbook policy, and template-backed UDF toolpacks. Omit formulaEnvironment when the default evaluator is intended.",
    displayOrder = 18)
public record FormulaEnvironmentInput(
    @ProtocolField(optional = true) @WireDefault("[]")
        List<FormulaExternalWorkbookInput> externalWorkbooks,
    @ProtocolField(optional = true) @WireDefault("ERROR")
        FormulaMissingWorkbookPolicy missingWorkbookPolicy,
    @ProtocolField(optional = true) @WireDefault("[]") List<FormulaUdfToolpackInput> udfToolpacks) {
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
}
