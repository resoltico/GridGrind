package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.dto.RequestWarning;
import java.util.List;
import java.util.Objects;

/** The tolerant syntax tree and every JSON-level problem found while building it. */
record RequestSyntaxParseResult(
    RequestJsonNode root, List<RequestStructuralProblem> problems, List<RequestWarning> warnings) {
  RequestSyntaxParseResult(RequestJsonNode root, List<RequestStructuralProblem> problems) {
    this(root, problems, List.of());
  }

  RequestSyntaxParseResult {
    Objects.requireNonNull(root, "root must not be null");
    problems = List.copyOf(Objects.requireNonNull(problems, "problems must not be null"));
    warnings = List.copyOf(Objects.requireNonNull(warnings, "warnings must not be null"));
  }
}
