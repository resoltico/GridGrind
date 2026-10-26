package dev.erst.gridgrind.core.model.validation;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.WireDefault;
import java.util.Objects;
import java.util.Optional;

/** Complete protocol-facing definition of one supported data-validation rule. */
@CatalogType(
    id = "DataValidationSpec",
    summary = "Supported data-validation definition attached to one sheet range.",
    displayOrder = 52)
public record DataValidationSpec(
    DataValidationRuleSpec rule,
    @ProtocolField(optional = true) @WireDefault("false") boolean allowBlank,
    @ProtocolField(optional = true) @WireDefault("false") boolean suppressDropDownArrow,
    Optional<DataValidationPromptSpec> prompt,
    Optional<DataValidationErrorAlertSpec> errorAlert) {
  public DataValidationSpec {
    Objects.requireNonNull(rule, "rule must not be null");
    Objects.requireNonNull(prompt, "prompt must not be null");
    Objects.requireNonNull(errorAlert, "errorAlert must not be null");
  }
}
