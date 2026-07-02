package dev.erst.gridgrind.contract.json;

import java.util.Objects;
import java.util.Optional;
import tools.jackson.core.JacksonException;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.exc.InvalidFormatException;
import tools.jackson.databind.exc.InvalidTypeIdException;
import tools.jackson.databind.exc.MismatchedInputException;
import tools.jackson.databind.exc.UnrecognizedPropertyException;
import tools.jackson.databind.exc.ValueInstantiationException;

/** Structural request-problem detection at the Jackson intake boundary. */
public final class GridGrindJsonRequestProblemDetector {
  private GridGrindJsonRequestProblemDetector() {}

  /** Detects one typed request-shape problem from the parsed JSON tree plus the Jackson failure. */
  public static Optional<RequestProblemDescriptor.Shape> detect(
      JsonNode rootNode, Class<?> fallbackTargetType, JacksonException exception) {
    Objects.requireNonNull(rootNode, "rootNode must not be null");
    Objects.requireNonNull(fallbackTargetType, "fallbackTargetType must not be null");
    Objects.requireNonNull(exception, "exception must not be null");
    return switch (exception) {
      case UnrecognizedPropertyException unknownField ->
          Optional.of(
              new UnknownField(
                  GridGrindJsonRequestTypeProblemSupport.fullUnknownFieldPath(unknownField)));
      case InvalidTypeIdException invalidTypeId ->
          Optional.of(GridGrindJsonRequestTypeProblemSupport.typeProblem(rootNode, invalidTypeId));
      case InvalidFormatException invalidFormat
          when invalidFormat.getTargetType() != null && invalidFormat.getTargetType().isEnum() ->
          Optional.of(GridGrindJsonRequestTypeProblemSupport.enumValueProblem(invalidFormat));
      case ValueInstantiationException valueInstantiation ->
          GridGrindJsonRequestShapeProblemSupport.valueInstantiationProblem(
              rootNode, fallbackTargetType, valueInstantiation);
      case MismatchedInputException mismatchedInput ->
          GridGrindJsonRequestShapeProblemSupport.mismatchedInputProblem(
              rootNode, fallbackTargetType, mismatchedInput);
      default -> Optional.empty();
    };
  }
}
