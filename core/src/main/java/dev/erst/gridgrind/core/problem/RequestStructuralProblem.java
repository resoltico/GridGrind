package dev.erst.gridgrind.core.problem;

import java.util.Objects;
import java.util.Optional;

/** One structural defect found while analysing a request without stopping at the first defect. */
public sealed interface RequestStructuralProblem
    permits RequestInvalidEncoding,
        RequestInvalidJson,
        RequestDuplicateKey,
        RequestNumberNotRepresentable,
        RequestShapeStructuralProblem {

  /** Returns the JSON path when one unambiguously identifies the malformed value. */
  Optional<String> jsonPath();

  /** Returns the zero-based UTF-8 byte offset of the offending token when one exists. */
  Optional<Long> byteOffset();

  /** Returns the one-based JSON line when syntax analysis captured it. */
  default Optional<Integer> jsonLine() {
    return Optional.empty();
  }

  /** Returns the one-based JSON column when syntax analysis captured it. */
  default Optional<Integer> jsonColumn() {
    return Optional.empty();
  }

  /** Returns the stable, product-owned diagnostic message. */
  String message();

  /** Returns the classified validation exception for this finding. */
  default IllegalArgumentException exception() {
    return switch (this) {
      case RequestInvalidEncoding invalidEncoding -> invalidEncodingException(invalidEncoding);
      case RequestInvalidJson invalidJson -> invalidJsonException(invalidJson);
      case RequestDuplicateKey duplicateKey -> invalidJsonException(duplicateKey);
      case RequestNumberNotRepresentable number ->
          new NumberNotRepresentableException(number.message(), number.jsonPath().orElseThrow());
      case RequestShapeStructuralProblem shapeProblem -> requestShapeException(shapeProblem);
    };
  }

  private static IllegalArgumentException requestShapeException(
      RequestShapeStructuralProblem problem) {
    return switch (Objects.requireNonNull(problem, "problem must not be null")) {
      case RequestUnknownField unknownField -> {
        Optional<String> jsonPath = unknownField.jsonPath();
        if (jsonPath.isEmpty()) {
          yield invalidRequestShape(
              new MessageShape(unknownField.message(), Optional.empty()), jsonPath);
        }
        yield invalidRequestShape(new UnknownField(jsonPath.orElseThrow()), jsonPath);
      }
      case RequestMissingRequiredField missingRequiredField ->
          invalidRequestShape(
              new MissingRequiredField(missingRequiredField.jsonPath().orElseThrow()),
              missingRequiredField.jsonPath());
      case RequestExplicitNullField explicitNullField ->
          invalidRequestShape(
              new ExplicitNullField(explicitNullField.jsonPath().orElseThrow()),
              explicitNullField.jsonPath());
      case RequestMissingTypeDiscriminator missingTypeDiscriminator ->
          invalidRequestShape(
              new MissingTypeDiscriminator(missingTypeDiscriminator.jsonPath().orElseThrow()),
              missingTypeDiscriminator.jsonPath());
      case RequestUnknownTypeDiscriminator unknownTypeDiscriminator -> {
        if (unknownTypeDiscriminator.value().isBlank()) {
          yield invalidRequestShape(
              new MessageShape(
                  unknownTypeDiscriminator.message(), unknownTypeDiscriminator.jsonPath()),
              unknownTypeDiscriminator.jsonPath());
        }
        yield invalidRequestShape(
            new UnknownTypeValue(
                unknownTypeDiscriminator.value(),
                unknownTypeDiscriminator.jsonPath(),
                unknownTypeDiscriminator.similarValues(),
                unknownTypeDiscriminator.specificGuidance()),
            unknownTypeDiscriminator.jsonPath());
      }
      case RequestUnsupportedEnumValue unsupportedEnumValue -> {
        if (unsupportedEnumValue.value().isBlank()) {
          yield invalidRequestShape(
              new MessageShape(unsupportedEnumValue.message(), unsupportedEnumValue.jsonPath()),
              unsupportedEnumValue.jsonPath());
        }
        yield invalidRequestShape(
            new UnsupportedValue(
                unsupportedEnumValue.value(),
                unsupportedEnumValue.jsonPath(),
                unsupportedEnumValue.allowedValues()),
            unsupportedEnumValue.jsonPath());
      }
      case RequestMalformedScalar malformedScalar ->
          invalidRequestShape(shapeFor(malformedScalar), malformedScalar.jsonPath());
    };
  }

  private static InvalidEncodingException invalidEncodingException(
      RequestInvalidEncoding invalidEncoding) {
    return new InvalidEncodingException(
        invalidEncoding.message(), Optional.empty(), Optional.empty(), Optional.empty(), null);
  }

  private static InvalidJsonException invalidJsonException(RequestStructuralProblem problem) {
    return new InvalidJsonException(
        problem.message(), Optional.empty(), Optional.empty(), Optional.empty(), null);
  }

  private static InvalidRequestShapeException invalidRequestShape(
      RequestProblemDescriptor.Shape shape, Optional<String> jsonPath) {
    return new InvalidRequestShapeException(
        shape, jsonPath, Optional.empty(), Optional.empty(), null);
  }

  private static RequestProblemDescriptor.Shape shapeFor(RequestMalformedScalar malformedScalar) {
    if ("a JSON string type id".equals(malformedScalar.expected())) {
      return new ActionableShapeMessage(
          malformedScalar.message(),
          "Replace field '"
              + malformedScalar.jsonPath().orElseThrow()
              + "' with a JSON string type id.",
          malformedScalar.jsonPath());
    }
    return new MessageShape(malformedScalar.message(), malformedScalar.jsonPath());
  }
}
