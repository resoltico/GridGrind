package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.dto.GridGrindRequestProblemSupport;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Signals that valid JSON did not match the GridGrind request or response payload shape. */
public final class InvalidRequestShapeException extends IllegalArgumentException
    implements PayloadException, RequestProblemSource {
  private static final long serialVersionUID = 1L;

  private final RequestProblemDescriptor.Shape requestProblem;
  private final PayloadLocation jsonLocation;

  /** Creates the exception with the given typed request problem, JSON location, and cause. */
  public InvalidRequestShapeException(
      RequestProblemDescriptor.Shape requestProblem,
      Optional<String> jsonPath,
      Optional<Integer> jsonLine,
      Optional<Integer> jsonColumn,
      @Nullable Throwable cause) {
    super(
        GridGrindRequestProblemSupport.message(normalizeRequestProblem(requestProblem, jsonPath)),
        cause);
    this.requestProblem = normalizeRequestProblem(requestProblem, jsonPath);
    this.jsonLocation =
        PayloadLocation.from(
            java.util.Objects.requireNonNull(jsonPath, "jsonPath must not be null")
                .or(requestProblem::jsonPath),
            jsonLine,
            jsonColumn);
  }

  @Override
  public PayloadLocation jsonLocation() {
    return jsonLocation;
  }

  @Override
  public RequestProblemDescriptor requestProblem() {
    return requestProblem;
  }

  private static RequestProblemDescriptor.Shape normalizeRequestProblem(
      RequestProblemDescriptor.Shape requestProblem, Optional<String> jsonPath) {
    return (RequestProblemDescriptor.Shape)
        RequestProblemDescriptorSupport.withJsonPath(
            java.util.Objects.requireNonNull(requestProblem, "requestProblem must not be null"),
            java.util.Objects.requireNonNull(jsonPath, "jsonPath must not be null"));
  }
}
