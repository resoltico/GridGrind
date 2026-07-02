package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Optional;
import java.util.Set;

/** Cross-field request validation that sits above record-level contract checks. */
final class ExecutionValidationSupport {
  List<GridGrindProblemDetail.Problem> validateRequest(WorkbookPlan request) {
    Set<String> messages = new LinkedHashSet<>();
    messages.addAll(ExecutionModeRules.calculationPolicyFailures(request));
    messages.addAll(
        ExecutionModeRules.executionModeFailures(
            request, ExecutionModeRules.executionMode(request)));
    persistenceFailureMessage(request).ifPresent(messages::add);
    List<GridGrindProblemDetail.Problem> problems = new ArrayList<>(messages.size());
    for (String message : messages) {
      problems.add(validationProblem(request, message));
    }
    return List.copyOf(problems);
  }

  Optional<GridGrindProblemDetail.Problem> firstValidationProblem(WorkbookPlan request) {
    return validateRequest(request).stream().findFirst();
  }

  private static Optional<String> persistenceFailureMessage(WorkbookPlan request) {
    return switch (request.persistence()) {
      case WorkbookPlan.WorkbookPersistence.Overwrite _ ->
          switch (request.source()) {
            case WorkbookPlan.WorkbookSource.New _ ->
                Optional.of(
                    "OVERWRITE persistence requires an EXISTING source; a NEW workbook has no"
                        + " source file to overwrite");
            case WorkbookPlan.WorkbookSource.ExistingFile _ -> Optional.empty();
          };
      case WorkbookPlan.WorkbookPersistence.None _ -> Optional.empty();
      case WorkbookPlan.WorkbookPersistence.SaveAs _ -> Optional.empty();
    };
  }

  private static GridGrindProblemDetail.Problem validationProblem(
      WorkbookPlan request, String message) {
    return GridGrindProblems.problem(
        GridGrindProblemCode.INVALID_REQUEST,
        message,
        new dev.erst.gridgrind.contract.dto.ProblemContext.ValidateRequest(
            ExecutionRequestPaths.requestShape(request)),
        (Throwable) null);
  }
}
