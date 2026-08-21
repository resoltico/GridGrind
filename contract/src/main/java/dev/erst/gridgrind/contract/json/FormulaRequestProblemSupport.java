package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.dto.FormulaInputException;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.InvalidFormulaInputException;
import dev.erst.gridgrind.contract.dto.InvalidRawFormulaTextException;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Retains the distinct public problem codes for normal and opaque formula-input validation. */
final class FormulaRequestProblemSupport {
  private FormulaRequestProblemSupport() {}

  static Optional<FormulaRequestException> inputFailure(
      Throwable failure,
      Optional<String> jsonPath,
      Optional<Integer> jsonLine,
      Optional<Integer> jsonColumn,
      @Nullable Throwable cause) {
    return inputFailure(failure)
        .map(
            input ->
                new FormulaRequestException(
                    codeFor(input), input.publicMessage(), jsonPath, jsonLine, jsonColumn, cause));
  }

  static Optional<FormulaRequestException> requestFailure(Throwable failure) {
    Throwable current = failure;
    while (current != null) {
      if (current instanceof FormulaRequestException formulaRequestException) {
        return Optional.of(formulaRequestException);
      }
      current = current.getCause();
    }
    return Optional.empty();
  }

  static FormulaRequestException normalizeRequestFailure(
      FormulaRequestException failure,
      RuntimeException cause,
      RequestJsonNode fragment,
      String fragmentPath) {
    Optional<String> preciseInnerPath =
        RequestBindingPathSupport.preciseInnerPath(
            fragment,
            RequestBindingPathSupport.jacksonFailure(cause)
                .flatMap(
                    exception ->
                        GridGrindJsonPayloadMetadataSupport.payloadMetadata(exception).jsonPath()),
            failure.requestProblem().jsonPath());
    return new FormulaRequestException(
        failure.problemCode(),
        failure.getMessage(),
        GridGrindJsonPathSupport.qualifyPath(Optional.of(fragmentPath), preciseInnerPath),
        Optional.empty(),
        Optional.empty(),
        cause);
  }

  private static Optional<FormulaInputException> inputFailure(Throwable failure) {
    Throwable current = failure;
    while (current != null) {
      if (current instanceof FormulaInputException formulaInputException) {
        return Optional.of(formulaInputException);
      }
      current = current.getCause();
    }
    return Optional.empty();
  }

  private static GridGrindProblemCode codeFor(FormulaInputException failure) {
    return switch (failure) {
      case InvalidFormulaInputException _ -> GridGrindProblemCode.INVALID_FORMULA;
      case InvalidRawFormulaTextException _ -> GridGrindProblemCode.INVALID_FORMULA_TEXT;
    };
  }
}
