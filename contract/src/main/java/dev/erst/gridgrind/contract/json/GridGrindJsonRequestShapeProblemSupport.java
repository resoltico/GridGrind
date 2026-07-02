package dev.erst.gridgrind.contract.json;

import java.util.Optional;
import tools.jackson.core.JsonToken;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.exc.MismatchedInputException;
import tools.jackson.databind.exc.ValueInstantiationException;

/** Message-shape branches for structural request-problem detection. */
final class GridGrindJsonRequestShapeProblemSupport {
  private GridGrindJsonRequestShapeProblemSupport() {}

  static Optional<RequestProblemDescriptor.Shape> mismatchedInputProblem(
      JsonNode rootNode, Class<?> fallbackTargetType, MismatchedInputException exception) {
    if (GridGrindJsonValueProblemSupport.isFloatingPointIntoInteger(exception)) {
      return Optional.of(
          new MessageShape(
              GridGrindJsonValueProblemSupport.mismatchedInputMessage(exception),
              GridGrindJsonRequestTypeProblemSupport.payloadPath(exception)));
    }
    Class<?> targetType =
        exception.getTargetType() == null ? fallbackTargetType : exception.getTargetType();
    Optional<RequestProblemDescriptor.Shape> missingRequiredComponent =
        GridGrindJsonRequestContractSupport.missingRequiredComponentProblem(
            rootNode, targetType, GridGrindJsonRequestTypeProblemSupport.renderedPath(exception));
    if (missingRequiredComponent.isPresent()) {
      return missingRequiredComponent;
    }
    String jsonPath = GridGrindJsonRequestTypeProblemSupport.renderedPath(exception);
    JsonToken currentToken = exception.getCurrentToken();
    String genericMessage =
        currentToken == JsonToken.START_OBJECT
            ? GridGrindJsonValueProblemSupport.genericObjectShapeMessage()
            : GridGrindJsonValueProblemSupport.genericValueShapeMessage();
    return Optional.of(
        new MessageShape(
            genericMessage, GridGrindJsonRequestTypeProblemSupport.optionalPath(jsonPath)));
  }

  static Optional<RequestProblemDescriptor.Shape> valueInstantiationProblem(
      JsonNode rootNode, Class<?> fallbackTargetType, ValueInstantiationException exception) {
    Class<?> targetType =
        exception.getType() == null ? fallbackTargetType : exception.getType().getRawClass();
    Optional<RequestProblemDescriptor.Shape> missingRequiredComponent =
        GridGrindJsonRequestContractSupport.missingRequiredComponentProblem(
            rootNode, targetType, GridGrindJsonRequestTypeProblemSupport.renderedPath(exception));
    if (missingRequiredComponent.isPresent()) {
      return missingRequiredComponent;
    }
    String jsonPath = GridGrindJsonRequestTypeProblemSupport.renderedPath(exception);
    return Optional.of(
        new MessageShape(
            GridGrindJsonValueProblemSupport.genericObjectShapeMessage(),
            GridGrindJsonRequestTypeProblemSupport.optionalPath(jsonPath)));
  }
}
