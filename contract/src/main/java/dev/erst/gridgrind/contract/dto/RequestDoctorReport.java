package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Machine-readable lint report for one authored request before workbook execution begins. */
public record RequestDoctorReport(
    GridGrindProtocolVersion protocolVersion,
    AnalysisSeverity severity,
    boolean valid,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Summary> summary,
    List<RequestWarning> warnings,
    List<GridGrindProblemDetail.Problem> problems) {
  private static final String VALID_SUMMARY_MESSAGE = "valid doctor report requires a summary";
  private static final String VALID_PROBLEMS_MESSAGE =
      "valid doctor report cannot contain problems";
  private static final String INVALID_PROBLEMS_MESSAGE =
      "invalid doctor report requires at least one problem";
  private static final String VALID_WARNING_SEVERITY_MESSAGE =
      "valid doctor report warnings require WARNING severity";
  private static final String VALID_INFO_SEVERITY_MESSAGE =
      "clean valid doctor report requires INFO severity";
  private static final String INVALID_ERROR_SEVERITY_MESSAGE =
      "invalid doctor report requires ERROR severity";

  public RequestDoctorReport {
    Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    Objects.requireNonNull(severity, "severity must not be null");
    summary = Objects.requireNonNullElseGet(summary, Optional::empty);
    warnings = copyWarnings(warnings);
    problems = copyProblems(problems);
    if (valid && summary.isEmpty()) {
      throw new IllegalArgumentException(VALID_SUMMARY_MESSAGE);
    }
    if (valid && !problems.isEmpty()) {
      throw new IllegalArgumentException(VALID_PROBLEMS_MESSAGE);
    }
    if (!valid && problems.isEmpty()) {
      throw new IllegalArgumentException(INVALID_PROBLEMS_MESSAGE);
    }
    if (valid && !warnings.isEmpty() && severity != AnalysisSeverity.WARNING) {
      throw new IllegalArgumentException(VALID_WARNING_SEVERITY_MESSAGE);
    }
    if (valid && warnings.isEmpty() && severity != AnalysisSeverity.INFO) {
      throw new IllegalArgumentException(VALID_INFO_SEVERITY_MESSAGE);
    }
    if (!valid && severity != AnalysisSeverity.ERROR) {
      throw new IllegalArgumentException(INVALID_ERROR_SEVERITY_MESSAGE);
    }
  }

  /** Returns one clean doctor report with no warnings or blocking problems. */
  public static RequestDoctorReport clean(Summary summary) {
    return new RequestDoctorReport(
        GridGrindProtocolVersion.current(),
        AnalysisSeverity.INFO,
        true,
        Optional.of(summary),
        List.of(),
        List.of());
  }

  /** Returns one valid doctor report that surfaces non-fatal warnings. */
  public static RequestDoctorReport warnings(Summary summary, List<RequestWarning> warnings) {
    return new RequestDoctorReport(
        GridGrindProtocolVersion.current(),
        AnalysisSeverity.WARNING,
        true,
        Optional.of(summary),
        warnings,
        List.of());
  }

  /** Returns one invalid doctor report with blocking problems and any derived warnings. */
  public static RequestDoctorReport invalid(
      Summary summary,
      List<RequestWarning> warnings,
      List<GridGrindProblemDetail.Problem> problems) {
    return invalid(Optional.ofNullable(summary), warnings, problems);
  }

  /** Returns one invalid doctor report with one blocking problem and any derived warnings. */
  public static RequestDoctorReport invalid(
      Summary summary, List<RequestWarning> warnings, GridGrindProblemDetail.Problem problem) {
    return invalid(Optional.ofNullable(summary), warnings, List.of(problem));
  }

  /** Returns one invalid doctor report with an optional summary and blocking problems. */
  public static RequestDoctorReport invalid(
      Optional<Summary> summary,
      List<RequestWarning> warnings,
      List<GridGrindProblemDetail.Problem> problems) {
    return new RequestDoctorReport(
        GridGrindProtocolVersion.current(),
        AnalysisSeverity.ERROR,
        false,
        summary,
        warnings,
        problems);
  }

  /** Returns one invalid doctor report with an optional summary and one blocking problem. */
  public static RequestDoctorReport invalid(
      Optional<Summary> summary,
      List<RequestWarning> warnings,
      GridGrindProblemDetail.Problem problem) {
    return invalid(summary, warnings, List.of(problem));
  }

  /** Returns the first blocking problem when the report is invalid. */
  public Optional<GridGrindProblemDetail.Problem> primaryProblem() {
    return problems.isEmpty() ? Optional.empty() : Optional.of(problems.getFirst());
  }

  /** Derived summary of the authored request shape and execution posture. */
  public record Summary(
      String sourceType,
      String persistenceType,
      String executionMode,
      String calculationStrategy,
      boolean markRecalculateOnOpen,
      boolean requiresStandardInputBinding,
      int stepCount,
      int mutationStepCount,
      int assertionStepCount,
      int inspectionStepCount) {
    public Summary {
      sourceType = requireNonBlank(sourceType, "sourceType");
      persistenceType = requireNonBlank(persistenceType, "persistenceType");
      executionMode = requireNonBlank(executionMode, "executionMode");
      calculationStrategy = requireNonBlank(calculationStrategy, "calculationStrategy");
      requireNonNegative(stepCount, "stepCount");
      requireNonNegative(mutationStepCount, "mutationStepCount");
      requireNonNegative(assertionStepCount, "assertionStepCount");
      requireNonNegative(inspectionStepCount, "inspectionStepCount");
      if (mutationStepCount + assertionStepCount + inspectionStepCount != stepCount) {
        throw new IllegalArgumentException(
            "mutationStepCount + assertionStepCount + inspectionStepCount must equal stepCount");
      }
    }
  }

  private static List<RequestWarning> copyWarnings(List<RequestWarning> warnings) {
    Objects.requireNonNull(warnings, "warnings must not be null");
    List<RequestWarning> copy = new ArrayList<>(warnings.size());
    for (RequestWarning warning : warnings) {
      copy.add(Objects.requireNonNull(warning, "warnings must not contain nulls"));
    }
    return List.copyOf(copy);
  }

  private static List<GridGrindProblemDetail.Problem> copyProblems(
      List<GridGrindProblemDetail.Problem> problems) {
    Objects.requireNonNull(problems, "problems must not be null");
    List<GridGrindProblemDetail.Problem> copy = new ArrayList<>(problems.size());
    for (GridGrindProblemDetail.Problem problem : problems) {
      copy.add(Objects.requireNonNull(problem, "problems must not contain nulls"));
    }
    return List.copyOf(copy);
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static void requireNonNegative(int value, String fieldName) {
    if (value < 0) {
      throw new IllegalArgumentException(fieldName + " must not be negative");
    }
  }
}
