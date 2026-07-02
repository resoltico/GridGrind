package dev.erst.gridgrind.contract.json;

import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Shared validation helpers for typed request-problem records. */
final class RequestProblemDescriptorSupport {
  private RequestProblemDescriptorSupport() {}

  static RequestProblemDescriptor withJsonPath(
      RequestProblemDescriptor requestProblem, Optional<String> jsonPath) {
    Objects.requireNonNull(requestProblem, "requestProblem must not be null");
    Objects.requireNonNull(jsonPath, "jsonPath must not be null");
    if (jsonPath.isEmpty()) {
      return requestProblem;
    }
    if (requestProblem instanceof RequestProblemDescriptor.Shape shape) {
      return withJsonPath(shape, jsonPath);
    }
    return withJsonPath((RequestProblemDescriptor.Invariant) requestProblem, jsonPath);
  }

  private static RequestProblemDescriptor.Shape withJsonPath(
      RequestProblemDescriptor.Shape requestProblem, Optional<String> jsonPath) {
    return switch (requestProblem) {
      case MissingRequiredField _ -> new MissingRequiredField(jsonPath.orElseThrow());
      case MissingTypeDiscriminator _ -> new MissingTypeDiscriminator(jsonPath.orElseThrow());
      case UnknownField _ -> new UnknownField(jsonPath.orElseThrow());
      case UnknownTypeValue unknownTypeValue ->
          new UnknownTypeValue(
              unknownTypeValue.typeId(),
              jsonPath,
              unknownTypeValue.similarValues(),
              unknownTypeValue.specificGuidance());
      case UnsupportedValue unsupportedValue ->
          new UnsupportedValue(
              unsupportedValue.value(), jsonPath, unsupportedValue.allowedValues());
      case ExplicitNullField _ -> new ExplicitNullField(jsonPath.orElseThrow());
      case ActionableShapeMessage actionableShapeMessage ->
          new ActionableShapeMessage(
              actionableShapeMessage.message(), actionableShapeMessage.resolutionValue(), jsonPath);
      case MessageShape messageShape -> new MessageShape(messageShape.message(), jsonPath);
    };
  }

  private static RequestProblemDescriptor.Invariant withJsonPath(
      RequestProblemDescriptor.Invariant requestProblem, Optional<String> jsonPath) {
    return switch (requestProblem) {
      case DuplicateStepId duplicateStepId ->
          new DuplicateStepId(duplicateStepId.value(), jsonPath.orElseThrow());
      case NonXlsxPath nonXlsxPath -> new NonXlsxPath(nonXlsxPath.actualExtension(), jsonPath);
      case FieldValidationProblem fieldValidationProblem ->
          new FieldValidationProblem(
              fieldValidationProblem.fieldName(),
              jsonPath,
              fieldValidationProblem.rule(),
              fieldValidationProblem.operands());
      case ActionableInvariantMessage actionableInvariantMessage ->
          new ActionableInvariantMessage(
              actionableInvariantMessage.message(),
              actionableInvariantMessage.resolutionValue(),
              jsonPath);
      case MessageInvariant messageInvariant ->
          new MessageInvariant(messageInvariant.message(), jsonPath);
    };
  }

  static Optional<String> copyJsonPath(Optional<String> jsonPath) {
    Objects.requireNonNull(jsonPath, "jsonPath must not be null");
    return jsonPath.map(path -> requireNonBlank(path, "jsonPath"));
  }

  static Optional<String> copyOptionalText(Optional<String> text, String fieldName) {
    Objects.requireNonNull(text, fieldName + " must not be null");
    return text.map(value -> requireNonBlank(value, fieldName));
  }

  static List<String> copyStrings(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    return values.stream().map(value -> requireNonBlank(value, fieldName)).toList();
  }

  static String requireJsonPath(String jsonPath) {
    return requireNonBlank(jsonPath, "jsonPath");
  }

  static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
