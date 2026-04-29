package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;

/** Complete protocol-facing definition of one supported data-validation rule. */
public record DataValidationInput(
    DataValidationRuleInput rule,
    Boolean allowBlank,
    Boolean suppressDropDownArrow,
    DataValidationPromptInput prompt,
    DataValidationErrorAlertInput errorAlert) {
  public DataValidationInput {
    Objects.requireNonNull(rule, "rule must not be null");
    Objects.requireNonNull(allowBlank, "allowBlank must not be null");
    Objects.requireNonNull(suppressDropDownArrow, "suppressDropDownArrow must not be null");
  }

  @JsonCreator
  static DataValidationInput create(
      @JsonProperty("rule") DataValidationRuleInput rule,
      @JsonProperty("allowBlank") Boolean allowBlank,
      @JsonProperty("suppressDropDownArrow") Boolean suppressDropDownArrow,
      @JsonProperty("prompt") DataValidationPromptInput prompt,
      @JsonProperty("errorAlert") DataValidationErrorAlertInput errorAlert) {
    return new DataValidationInput(
        rule,
        Boolean.TRUE.equals(allowBlank),
        Boolean.TRUE.equals(suppressDropDownArrow),
        prompt,
        errorAlert);
  }
}
