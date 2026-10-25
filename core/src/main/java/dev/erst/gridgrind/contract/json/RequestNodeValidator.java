package dev.erst.gridgrind.contract.json;

import java.lang.reflect.Type;
import java.util.Arrays;
import java.util.List;

/** Validates token kinds and recursively dispatches request creators to focused validators. */
final class RequestNodeValidator {
  private RequestNodeValidator() {}

  static void validateNode(
      RequestJsonNode node,
      Type expectedType,
      String jsonPath,
      long diagnosticByteOffset,
      List<RequestStructuralProblem> problems) {
    if (node instanceof RequestJsonNull) {
      problems.add(new RequestExplicitNullField(jsonPath, diagnosticByteOffset));
      return;
    }
    if (RequestNodeTypeClassifier.isOptional(expectedType)) {
      validateNode(
          node,
          RequestTypeSupport.typeArgument(expectedType),
          jsonPath,
          diagnosticByteOffset,
          problems);
      return;
    }
    if (RequestNodeTypeClassifier.isCollection(expectedType)) {
      RequestCollectionValidator.validate(
          node,
          RequestTypeSupport.typeArgument(expectedType),
          jsonPath,
          diagnosticByteOffset,
          problems);
      return;
    }
    RequestNonContainerNodeValidator.validate(
        node, RequestTypeSupport.rawType(expectedType), jsonPath, diagnosticByteOffset, problems);
  }

  static void validateEnum(
      RequestJsonNode node,
      Class<?> enumType,
      String jsonPath,
      long diagnosticByteOffset,
      List<RequestStructuralProblem> problems) {
    if (node instanceof RequestJsonNull) {
      problems.add(new RequestExplicitNullField(jsonPath, diagnosticByteOffset));
      return;
    }
    if (!(node instanceof RequestJsonString value)) {
      problems.add(new RequestMalformedScalar(jsonPath, "a JSON string", diagnosticByteOffset));
      return;
    }
    List<String> allowedValues =
        Arrays.stream(enumType.getEnumConstants())
            .map(constant -> ((Enum<?>) constant).name())
            .toList();
    if (!allowedValues.contains(value.value())) {
      problems.add(
          new RequestUnsupportedEnumValue(
              jsonPath, value.value(), allowedValues, diagnosticByteOffset));
    }
  }
}
