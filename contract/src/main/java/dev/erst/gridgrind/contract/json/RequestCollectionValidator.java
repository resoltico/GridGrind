package dev.erst.gridgrind.contract.json;

import java.lang.reflect.Type;
import java.util.List;

/** Validates the shape and each element of one request collection creator component. */
final class RequestCollectionValidator {
  private RequestCollectionValidator() {}

  static void validate(
      RequestJsonNode node,
      Type elementType,
      String jsonPath,
      long diagnosticByteOffset,
      List<RequestStructuralProblem> problems) {
    if (!(node instanceof RequestJsonArray array)) {
      problems.add(new RequestMalformedScalar(jsonPath, "a JSON array", diagnosticByteOffset));
      return;
    }
    for (int index = 0; index < array.elements().size(); index++) {
      RequestJsonNode element = array.elements().get(index);
      RequestNodeValidator.validateNode(
          element, elementType, jsonPath + "[" + index + "]", element.byteOffset(), problems);
    }
  }
}
