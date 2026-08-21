package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;
import java.util.Optional;

/** Complete protocol-facing definition of one supported data-validation rule. */
public record DataValidationInput(
    DataValidationRuleInput rule,
    @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
        boolean allowBlank,
    @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
        boolean suppressDropDownArrow,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<DataValidationPromptInput> prompt,
    @JsonInclude(JsonInclude.Include.NON_ABSENT)
        Optional<DataValidationErrorAlertInput> errorAlert) {
  public DataValidationInput {
    Objects.requireNonNull(rule, "rule must not be null");
    Objects.requireNonNull(prompt, "prompt must not be null");
    Objects.requireNonNull(errorAlert, "errorAlert must not be null");
  }

  /** Creates one authored data-validation definition with explicit wire booleans. */
  @JsonCreator
  public DataValidationInput(
      @JsonProperty("rule") DataValidationRuleInput rule,
      @JsonProperty("allowBlank") Boolean allowBlank,
      @JsonProperty("suppressDropDownArrow") Boolean suppressDropDownArrow,
      @JsonProperty("prompt") Optional<DataValidationPromptInput> prompt,
      @JsonProperty("errorAlert") Optional<DataValidationErrorAlertInput> errorAlert) {
    this(
        rule,
        ProtocolBooleanDefault.FALSE.resolve(allowBlank),
        ProtocolBooleanDefault.FALSE.resolve(suppressDropDownArrow),
        emptyIfNull(prompt),
        emptyIfNull(errorAlert));
  }

  private static <T> Optional<T> emptyIfNull(Optional<T> value) {
    return value == null ? Optional.empty() : value;
  }
}
