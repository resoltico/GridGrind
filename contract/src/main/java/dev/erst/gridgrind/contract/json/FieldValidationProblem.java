package dev.erst.gridgrind.contract.json;

import java.util.List;
import java.util.Optional;

/**
 * Typed request-field validation failure whose public wording derives from one rule plus operands.
 */
public record FieldValidationProblem(
    String fieldName, Optional<String> jsonPath, FieldValidationRule rule, List<String> operands)
    implements RequestProblemDescriptor.Invariant {
  public FieldValidationProblem {
    fieldName = RequestProblemDescriptorSupport.requireNonBlank(fieldName, "fieldName");
    jsonPath = RequestProblemDescriptorSupport.copyJsonPath(jsonPath);
    java.util.Objects.requireNonNull(rule, "rule must not be null");
    operands = RequestProblemDescriptorSupport.copyStrings(operands, "operands");
  }

  /** Creates one field-owned validation problem whose default JSON path is the field name. */
  public static FieldValidationProblem atField(
      String fieldName, FieldValidationRule rule, String... operands) {
    return new FieldValidationProblem(fieldName, Optional.of(fieldName), rule, List.of(operands));
  }

  /** Creates one validation problem with no default JSON path because the issue spans fields. */
  public static FieldValidationProblem detached(
      String fieldName, FieldValidationRule rule, String... operands) {
    return new FieldValidationProblem(fieldName, Optional.empty(), rule, List.of(operands));
  }

  /** Renders the canonical public message for this field-validation problem. */
  public String message() {
    return rule.message(this);
  }

  /** Renders the canonical remediation sentence for this field-validation problem. */
  public String resolution() {
    return rule.resolution(this);
  }

  String operand(int index) {
    return operands.get(index);
  }
}
