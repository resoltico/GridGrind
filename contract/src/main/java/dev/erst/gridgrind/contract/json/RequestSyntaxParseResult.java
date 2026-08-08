package dev.erst.gridgrind.contract.json;

import java.util.List;
import java.util.Objects;

/** The tolerant syntax tree and every JSON-level problem found while building it. */
record RequestSyntaxParseResult(RequestJsonNode root, List<RequestStructuralProblem> problems) {
  RequestSyntaxParseResult {
    Objects.requireNonNull(root, "root must not be null");
    problems = List.copyOf(Objects.requireNonNull(problems, "problems must not be null"));
  }
}
