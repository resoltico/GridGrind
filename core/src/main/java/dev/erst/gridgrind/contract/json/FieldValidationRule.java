package dev.erst.gridgrind.contract.json;

/** One typed request-field validation rule with public wording and remediation guidance. */
public sealed interface FieldValidationRule
    permits FieldValidationBasicRule,
        FieldValidationBoundRule,
        FieldValidationNamingRule,
        FieldValidationAddressRule,
        FieldValidationLayoutRule {
  /** Renders the canonical public message for one field-validation problem. */
  String message(FieldValidationProblem problem);

  /** Renders the canonical public resolution for one field-validation problem. */
  String resolution(FieldValidationProblem problem);
}
