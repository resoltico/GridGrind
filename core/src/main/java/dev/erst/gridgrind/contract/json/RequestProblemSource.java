package dev.erst.gridgrind.contract.json;

/** Structured request-problem carrier exposed by the request exception family. */
public sealed interface RequestProblemSource
    permits FormulaRequestException, InvalidRequestException, InvalidRequestShapeException {
  /** Typed request problem owned by the exception. */
  RequestProblemDescriptor requestProblem();
}
