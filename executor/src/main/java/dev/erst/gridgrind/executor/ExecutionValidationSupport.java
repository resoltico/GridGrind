package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.Optional;

/** Cross-field request validation that sits above record-level contract checks. */
final class ExecutionValidationSupport {
  Optional<GridGrindResponse.Problem> validateRequest(WorkbookPlan request) {
    Optional<String> calculationPolicyFailure =
        ExecutionModeRules.calculationPolicyFailure(request);
    if (calculationPolicyFailure.isPresent()) {
      return Optional.of(
          GridGrindProblems.problem(
              GridGrindProblemCode.INVALID_REQUEST,
              calculationPolicyFailure.get(),
              new GridGrindResponse.ProblemContext.ValidateRequest(
                  ExecutionRequestPaths.reqSourceType(request),
                  ExecutionRequestPaths.reqPersistenceType(request)),
              (Throwable) null));
    }

    Optional<String> executionModeFailure =
        ExecutionModeRules.executionModeFailure(
            request, ExecutionModeRules.executionModes(request));
    if (executionModeFailure.isPresent()) {
      return Optional.of(
          GridGrindProblems.problem(
              GridGrindProblemCode.INVALID_REQUEST,
              executionModeFailure.get(),
              new GridGrindResponse.ProblemContext.ValidateRequest(
                  ExecutionRequestPaths.reqSourceType(request),
                  ExecutionRequestPaths.reqPersistenceType(request)),
              (Throwable) null));
    }

    return switch (request.persistence()) {
      case WorkbookPlan.WorkbookPersistence.OverwriteSource _ ->
          switch (request.source()) {
            case WorkbookPlan.WorkbookSource.New _ ->
                Optional.of(
                    GridGrindProblems.problem(
                        GridGrindProblemCode.INVALID_REQUEST,
                        "OVERWRITE persistence requires an EXISTING source; "
                            + "a NEW workbook has no source file to overwrite",
                        new GridGrindResponse.ProblemContext.ValidateRequest(
                            ExecutionRequestPaths.reqSourceType(request),
                            ExecutionRequestPaths.reqPersistenceType(request)),
                        (Throwable) null));
            case WorkbookPlan.WorkbookSource.ExistingFile _ -> Optional.empty();
          };
      case WorkbookPlan.WorkbookPersistence.None _ -> Optional.empty();
      case WorkbookPlan.WorkbookPersistence.SaveAs _ -> Optional.empty();
    };
  }
}
