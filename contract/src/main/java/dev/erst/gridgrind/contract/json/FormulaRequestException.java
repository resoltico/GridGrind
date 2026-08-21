package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindRequestProblemSupport;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Preserves an owned formula-input code while attaching request-path metadata. */
public final class FormulaRequestException extends IllegalArgumentException
    implements PayloadException, RequestProblemSource {
  private static final long serialVersionUID = 1L;

  private final GridGrindProblemCode problemCode;
  private final RequestProblemDescriptor.Invariant requestProblem;
  private final PayloadLocation jsonLocation;

  /** Creates one classified formula-input failure at an optional request location. */
  public FormulaRequestException(
      GridGrindProblemCode problemCode,
      String message,
      Optional<String> jsonPath,
      Optional<Integer> jsonLine,
      Optional<Integer> jsonColumn,
      @Nullable Throwable cause) {
    super(message, cause);
    this.problemCode = requireFormulaProblemCode(problemCode);
    this.requestProblem =
        new MessageInvariant(
            Objects.requireNonNull(message, "message must not be null"),
            Objects.requireNonNullElseGet(jsonPath, Optional::empty));
    this.jsonLocation =
        PayloadLocation.from(
            Objects.requireNonNullElseGet(jsonPath, Optional::empty),
            Objects.requireNonNullElseGet(jsonLine, Optional::empty),
            Objects.requireNonNullElseGet(jsonColumn, Optional::empty));
  }

  /** Returns the formula-specific public code retained through request binding. */
  public GridGrindProblemCode problemCode() {
    return problemCode;
  }

  @Override
  public PayloadLocation jsonLocation() {
    return jsonLocation;
  }

  @Override
  public RequestProblemDescriptor requestProblem() {
    return requestProblem;
  }

  @Override
  public String getMessage() {
    return GridGrindRequestProblemSupport.message(requestProblem);
  }

  private static GridGrindProblemCode requireFormulaProblemCode(GridGrindProblemCode problemCode) {
    return switch (Objects.requireNonNull(problemCode, "problemCode must not be null")) {
      case INVALID_FORMULA, INVALID_FORMULA_TEXT -> problemCode;
      default -> throw new IllegalArgumentException("problemCode must classify formula input");
    };
  }
}
