package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Machine-readable lint report for one authored request before workbook execution begins. */
public record RequestDoctorReport(
    GridGrindProtocolVersion protocolVersion,
    AnalysisSeverity severity,
    boolean valid,
    @JsonInclude(JsonInclude.Include.NON_NULL) Summary summary,
    List<RequestWarning> warnings,
    @JsonInclude(JsonInclude.Include.NON_NULL) GridGrindResponse.Problem problem) {
  public RequestDoctorReport {
    protocolVersion =
        protocolVersion == null ? GridGrindProtocolVersion.current() : protocolVersion;
    Objects.requireNonNull(severity, "severity must not be null");
    warnings = copyWarnings(warnings);
    if (valid && summary == null) {
      throw new IllegalArgumentException("summary must not be null for a valid doctor report");
    }
    if (valid && problem != null) {
      throw new IllegalArgumentException("problem must be null for a valid doctor report");
    }
    if (!valid && problem == null) {
      throw new IllegalArgumentException("problem must not be null for an invalid doctor report");
    }
    if (valid && !warnings.isEmpty() && severity != AnalysisSeverity.WARNING) {
      throw new IllegalArgumentException(
          "severity must be WARNING when a valid doctor report contains warnings");
    }
    if (valid && warnings.isEmpty() && severity != AnalysisSeverity.INFO) {
      throw new IllegalArgumentException(
          "severity must be INFO when a valid doctor report contains no warnings");
    }
    if (!valid && severity != AnalysisSeverity.ERROR) {
      throw new IllegalArgumentException("severity must be ERROR when a doctor report is invalid");
    }
  }

  /** Returns one clean doctor report with no warnings or blocking problem. */
  public static RequestDoctorReport clean(Summary summary) {
    return new RequestDoctorReport(
        GridGrindProtocolVersion.current(), AnalysisSeverity.INFO, true, summary, List.of(), null);
  }

  /** Returns one valid doctor report that still surfaces non-fatal warnings. */
  public static RequestDoctorReport warnings(Summary summary, List<RequestWarning> warnings) {
    return new RequestDoctorReport(
        GridGrindProtocolVersion.current(),
        AnalysisSeverity.WARNING,
        true,
        summary,
        warnings,
        null);
  }

  /** Returns one invalid doctor report with a blocking problem and any derived warnings. */
  public static RequestDoctorReport invalid(
      Summary summary, List<RequestWarning> warnings, GridGrindResponse.Problem problem) {
    return new RequestDoctorReport(
        GridGrindProtocolVersion.current(),
        AnalysisSeverity.ERROR,
        false,
        summary,
        warnings,
        problem);
  }

  /** One derived summary of the authored request shape and execution posture. */
  public record Summary(
      String sourceType,
      String persistenceType,
      String readMode,
      String writeMode,
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
      readMode = requireNonBlank(readMode, "readMode");
      writeMode = requireNonBlank(writeMode, "writeMode");
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
    if (warnings == null) {
      return List.of();
    }
    List<RequestWarning> copy = new ArrayList<>(warnings.size());
    for (RequestWarning warning : warnings) {
      copy.add(Objects.requireNonNull(warning, "warnings must not contain nulls"));
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
