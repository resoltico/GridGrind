package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionResult;
import java.util.List;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** All execution facts required to construct one failed workbook result. */
record ExecutionFailure(Context context, Artifacts artifacts, Detail detail) {
  ExecutionFailure {
    Objects.requireNonNull(context, "context must not be null");
    Objects.requireNonNull(artifacts, "artifacts must not be null");
    Objects.requireNonNull(detail, "detail must not be null");
  }

  record Context(
      GridGrindProtocolVersion protocolVersion,
      ExecutionJournalRecorder journal,
      WorkbookPlan request,
      CalculationReport calculation) {
    Context {
      Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
      Objects.requireNonNull(journal, "journal must not be null");
      Objects.requireNonNull(request, "request must not be null");
      Objects.requireNonNull(calculation, "calculation must not be null");
    }
  }

  record Artifacts(
      List<RequestWarning> warnings,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections) {
    Artifacts {
      warnings = List.copyOf(warnings);
      assertions = List.copyOf(assertions);
      inspections = List.copyOf(inspections);
    }

    static Artifacts empty() {
      return new Artifacts(List.of(), List.of(), List.of());
    }
  }

  record Detail(
      GridGrindProblemDetail.Problem problem,
      @Nullable Integer failedStepIndex,
      @Nullable String failedStepId) {
    Detail {
      Objects.requireNonNull(problem, "problem must not be null");
    }
  }
}
