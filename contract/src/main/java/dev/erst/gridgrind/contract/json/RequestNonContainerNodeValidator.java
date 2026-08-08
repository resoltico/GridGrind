package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.selector.Selector;
import java.util.List;

/** Validates the non-container request creator families after optional and collection handling. */
final class RequestNonContainerNodeValidator {
  private RequestNonContainerNodeValidator() {}

  static void validate(
      RequestJsonNode node,
      Class<?> rawType,
      String jsonPath,
      long diagnosticByteOffset,
      List<RequestStructuralProblem> problems) {
    if (rawType.isEnum()) {
      RequestNodeValidator.validateEnum(node, rawType, jsonPath, diagnosticByteOffset, problems);
      return;
    }
    if (rawType == String.class || CharSequence.class.isAssignableFrom(rawType)) {
      requireNodeKind(
          node, jsonPath, "a JSON string", RequestJsonString.class, diagnosticByteOffset, problems);
      return;
    }
    if (rawType == boolean.class || rawType == Boolean.class) {
      requireNodeKind(
          node,
          jsonPath,
          "a JSON boolean",
          RequestJsonBoolean.class,
          diagnosticByteOffset,
          problems);
      return;
    }
    if (RequestTypeSupport.isNumberType(rawType)) {
      validateNumber(node, rawType, jsonPath, diagnosticByteOffset, problems);
      return;
    }
    if (RequestTypeSupport.isTemporalType(rawType)) {
      validateTemporal(node, rawType, jsonPath, diagnosticByteOffset, problems);
      return;
    }
    if (rawType == Selector.class) {
      RequestUnionValidator.validateSelector(node, jsonPath, diagnosticByteOffset, problems);
      return;
    }
    if (rawType.isSealed()) {
      RequestUnionValidator.validateUnion(node, rawType, jsonPath, diagnosticByteOffset, problems);
    } else if (rawType.isRecord()) {
      RequestRecordValidator.validate(
          node, rawType.asSubclass(Record.class), jsonPath, diagnosticByteOffset, problems);
    }
  }

  private static void requireNodeKind(
      RequestJsonNode node,
      String jsonPath,
      String expected,
      Class<? extends RequestJsonNode> expectedNodeType,
      long diagnosticByteOffset,
      List<RequestStructuralProblem> problems) {
    if (!expectedNodeType.isInstance(node)) {
      problems.add(new RequestMalformedScalar(jsonPath, expected, diagnosticByteOffset));
    }
  }

  private static void validateNumber(
      RequestJsonNode node,
      Class<?> rawType,
      String jsonPath,
      long diagnosticByteOffset,
      List<RequestStructuralProblem> problems) {
    if (!(node instanceof RequestJsonNumber number)) {
      requireNodeKind(
          node, jsonPath, "a JSON number", RequestJsonNumber.class, diagnosticByteOffset, problems);
      return;
    }
    if (!canBindScalar(number, rawType)) {
      problems.add(
          new RequestMalformedScalar(
              jsonPath, RequestTypeSupport.numericExpectation(rawType), diagnosticByteOffset));
    }
  }

  private static void validateTemporal(
      RequestJsonNode node,
      Class<?> rawType,
      String jsonPath,
      long diagnosticByteOffset,
      List<RequestStructuralProblem> problems) {
    if (!(node instanceof RequestJsonString string)) {
      requireNodeKind(
          node, jsonPath, "a JSON string", RequestJsonString.class, diagnosticByteOffset, problems);
      return;
    }
    if (!canBindScalar(string, rawType)) {
      problems.add(
          new RequestMalformedScalar(
              jsonPath, RequestTypeSupport.temporalExpectation(rawType), diagnosticByteOffset));
    }
  }

  private static boolean canBindScalar(RequestJsonNode node, Class<?> rawType) {
    try {
      Object value =
          GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER.treeToValue(
              RequestFragmentBinder.toJsonNode(node), rawType);
      return switch (value) {
        case Float floatValue -> Float.isFinite(floatValue);
        case Double doubleValue -> Double.isFinite(doubleValue);
        default -> true;
      };
    } catch (RuntimeException exception) {
      return false;
    }
  }
}
