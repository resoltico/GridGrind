package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;

/** Complete protocol-facing definition of one supported data-validation rule. */
public record DataValidationInput(
    DataValidationRuleInput rule,
    boolean allowBlank,
    boolean suppressDropDownArrow,
    DataValidationPromptInput prompt,
    DataValidationErrorAlertInput errorAlert) {
  public DataValidationInput {
    Objects.requireNonNull(rule, "rule must not be null");
  }

  /** Creates one authored data-validation definition with explicit wire booleans. */
  @JsonCreator
  public DataValidationInput(
      @JsonProperty("rule") DataValidationRuleInput rule,
      @JsonProperty("allowBlank") Boolean allowBlank,
      @JsonProperty("suppressDropDownArrow") Boolean suppressDropDownArrow,
      @JsonProperty("prompt") DataValidationPromptInput prompt,
      @JsonProperty("errorAlert") DataValidationErrorAlertInput errorAlert) {
    this(
        rule,
        Objects.requireNonNull(allowBlank, "allowBlank must not be null").booleanValue(),
        Objects.requireNonNull(suppressDropDownArrow, "suppressDropDownArrow must not be null")
            .booleanValue(),
        prompt,
        errorAlert);
  }
}
