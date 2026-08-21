package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.nio.charset.StandardCharsets;
import java.util.Objects;
import java.util.Optional;

/** Owns request-specific analysis, early structural rejection, and typed DTO decoding. */
final class GridGrindRequestDecoder {
  private GridGrindRequestDecoder() {}

  static WorkbookPlan read(byte[] bytes) {
    return completePlan(analyze(bytes));
  }

  static WorkbookPlan read(String json) {
    return completePlan(analyze(json.getBytes(StandardCharsets.UTF_8)));
  }

  private static WorkbookPlan completePlan(RequestAnalysis analysis) {
    analysis.structuralProblems().stream()
        .findFirst()
        .ifPresent(
            problem -> {
              throw structuralException(problem);
            });
    return analysis.requireCompletePlan();
  }

  static RequestAnalysis analyze(byte[] bytes) {
    GridGrindJsonMapperSupport.requireSupportedRequestLength(bytes.length);
    return RequestStructuralAnalyzer.analyze(bytes);
  }

  static IllegalArgumentException structuralException(RequestStructuralProblem problem) {
    return switch (Objects.requireNonNull(problem, "problem must not be null")) {
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
      case RequestUnknownField unknownField ->
          invalidRequestShape(
              new UnknownField(unknownField.jsonPath().orElseThrow()), unknownField.jsonPath());
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
      case RequestUnknownTypeDiscriminator unknownTypeDiscriminator ->
          invalidRequestShape(
              new UnknownTypeValue(
                  unknownTypeDiscriminator.value(),
                  unknownTypeDiscriminator.jsonPath(),
                  unknownTypeDiscriminator.similarValues(),
                  unknownTypeDiscriminator.specificGuidance()),
              unknownTypeDiscriminator.jsonPath());
      case RequestUnsupportedEnumValue unsupportedEnumValue ->
          invalidRequestShape(
              new UnsupportedValue(
                  unsupportedEnumValue.value(),
                  unsupportedEnumValue.jsonPath(),
                  unsupportedEnumValue.allowedValues()),
              unsupportedEnumValue.jsonPath());
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
