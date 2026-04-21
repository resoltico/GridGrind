package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;

/**
 * Complete GridGrind workbook plan for source, execution settings, ordered steps, and persistence.
 */
public record WorkbookPlan(
    GridGrindProtocolVersion protocolVersion,
    @JsonInclude(JsonInclude.Include.NON_NULL) String planId,
    WorkbookSource source,
    WorkbookPersistence persistence,
    @JsonInclude(JsonInclude.Include.NON_NULL) ExecutionPolicyInput execution,
    @JsonInclude(JsonInclude.Include.NON_NULL) FormulaEnvironmentInput formulaEnvironment,
    List<WorkbookStep> steps) {
  public WorkbookPlan {
    protocolVersion =
        protocolVersion == null ? GridGrindProtocolVersion.current() : protocolVersion;
    if (planId != null) {
      requireNonBlank(planId, "planId");
    }
    Objects.requireNonNull(source, "source must not be null");
    persistence = persistence == null ? new WorkbookPersistence.None() : persistence;
    execution = execution == null || execution.isDefault() ? null : execution;
    formulaEnvironment =
        formulaEnvironment == null || formulaEnvironment.isEmpty() ? null : formulaEnvironment;
    steps = copySteps(steps);
  }

  /** Creates a plan with the current protocol version and an explicit execution policy. */
  public WorkbookPlan(
      GridGrindProtocolVersion protocolVersion,
      WorkbookSource source,
      WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<WorkbookStep> steps) {
    this(protocolVersion, null, source, persistence, execution, formulaEnvironment, steps);
  }

  /** Creates a plan with the current protocol version and the given parameters. */
  public WorkbookPlan(
      GridGrindProtocolVersion protocolVersion,
      WorkbookSource source,
      WorkbookPersistence persistence,
      FormulaEnvironmentInput formulaEnvironment,
      List<WorkbookStep> steps) {
    this(protocolVersion, null, source, persistence, null, formulaEnvironment, steps);
  }

  /** Creates a plan with the current protocol version and the given parameters. */
  public WorkbookPlan(
      GridGrindProtocolVersion protocolVersion,
      WorkbookSource source,
      WorkbookPersistence persistence,
      ExecutionModeInput executionMode,
      FormulaEnvironmentInput formulaEnvironment,
      List<WorkbookStep> steps) {
    this(
        protocolVersion,
        null,
        source,
        persistence,
        executionMode == null ? null : new ExecutionPolicyInput(executionMode),
        formulaEnvironment,
        steps);
  }

  /** Creates a plan with the current protocol version and the given parameters. */
  public WorkbookPlan(
      WorkbookSource source,
      WorkbookPersistence persistence,
      ExecutionModeInput executionMode,
      FormulaEnvironmentInput formulaEnvironment,
      List<WorkbookStep> steps) {
    this(
        GridGrindProtocolVersion.current(),
        source,
        persistence,
        executionMode,
        formulaEnvironment,
        steps);
  }

  /** Creates a plan with the current protocol version and the given parameters. */
  public WorkbookPlan(
      WorkbookSource source,
      WorkbookPersistence persistence,
      FormulaEnvironmentInput formulaEnvironment,
      List<WorkbookStep> steps) {
    this(source, persistence, null, formulaEnvironment, steps);
  }

  /** Creates a plan with the current protocol version and no explicit formula environment. */
  public WorkbookPlan(
      WorkbookSource source, WorkbookPersistence persistence, List<WorkbookStep> steps) {
    this(GridGrindProtocolVersion.current(), source, persistence, null, steps);
  }

  /** Returns the effective execution policy after default normalization. */
  public ExecutionPolicyInput effectiveExecution() {
    return execution == null ? new ExecutionPolicyInput(null, null) : execution;
  }

  /** Returns the explicitly configured execution mode family, or null when omitted. */
  public ExecutionModeInput executionMode() {
    return execution == null ? null : execution.mode();
  }

  /** Returns the effective execution mode family after default normalization. */
  public ExecutionModeInput effectiveExecutionMode() {
    return effectiveExecution().effectiveMode();
  }

  /** Returns the effective execution journal level after default normalization. */
  public ExecutionJournalLevel journalLevel() {
    return effectiveExecution().effectiveJournalLevel();
  }

  /** Returns the effective calculation policy after default normalization. */
  public CalculationPolicyInput calculationPolicy() {
    return effectiveExecution().effectiveCalculation();
  }

  /** Returns the mutation steps in authored order. */
  public List<MutationStep> mutationSteps() {
    return steps.stream()
        .filter(MutationStep.class::isInstance)
        .map(MutationStep.class::cast)
        .toList();
  }

  /** Returns the assertion steps in authored order. */
  public List<AssertionStep> assertionSteps() {
    return steps.stream()
        .filter(AssertionStep.class::isInstance)
        .map(AssertionStep.class::cast)
        .toList();
  }

  /** Returns the inspection steps in authored order. */
  public List<InspectionStep> inspectionSteps() {
    return steps.stream()
        .filter(InspectionStep.class::isInstance)
        .map(InspectionStep.class::cast)
        .toList();
  }

  /** Describes where the input workbook comes from. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = WorkbookSource.New.class, name = "NEW"),
    @JsonSubTypes.Type(value = WorkbookSource.ExistingFile.class, name = "EXISTING")
  })
  public sealed interface WorkbookSource {
    /** Creates a brand-new empty workbook in memory. */
    record New() implements WorkbookSource {}

    /** Opens an existing workbook file from disk. */
    record ExistingFile(
        String path, @JsonInclude(JsonInclude.Include.NON_NULL) OoxmlOpenSecurityInput security)
        implements WorkbookSource {
      public ExistingFile {
        requireXlsxWorkbookPath(path);
        security = security == null || security.isEmpty() ? null : security;
      }

      /**
       * Opens the existing workbook at the supplied path with no explicit package-open settings.
       */
      public ExistingFile(String path) {
        this(path, null);
      }
    }
  }

  /** Describes whether and where the resulting workbook should be saved. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = WorkbookPersistence.None.class, name = "NONE"),
    @JsonSubTypes.Type(value = WorkbookPersistence.OverwriteSource.class, name = "OVERWRITE"),
    @JsonSubTypes.Type(value = WorkbookPersistence.SaveAs.class, name = "SAVE_AS")
  })
  public sealed interface WorkbookPersistence {
    /** Leaves the workbook in memory only and does not persist it. */
    record None() implements WorkbookPersistence {}

    /** Saves the workbook back to the exact path it was opened from. */
    record OverwriteSource(
        @JsonInclude(JsonInclude.Include.NON_NULL) OoxmlPersistenceSecurityInput security)
        implements WorkbookPersistence {
      public OverwriteSource {
        security = security == null ? null : security;
      }

      /** Overwrites the source workbook with no explicit package-security persistence settings. */
      public OverwriteSource() {
        this(null);
      }
    }

    /** Saves the workbook to a new `.xlsx` path. */
    record SaveAs(
        String path,
        @JsonInclude(JsonInclude.Include.NON_NULL) OoxmlPersistenceSecurityInput security)
        implements WorkbookPersistence {
      public SaveAs {
        requireXlsxWorkbookPath(path);
        security = security == null ? null : security;
      }

      /** Saves the workbook to the supplied path with no explicit package-security settings. */
      public SaveAs(String path) {
        this(path, null);
      }
    }
  }

  private static List<WorkbookStep> copySteps(List<WorkbookStep> steps) {
    if (steps == null) {
      return List.of();
    }
    List<WorkbookStep> copy = new java.util.ArrayList<>(steps.size());
    Set<String> seen = new HashSet<>();
    for (WorkbookStep step : steps) {
      Objects.requireNonNull(step, "steps must not contain nulls");
      copy.add(step);
      // LIM-006
      if (!seen.add(step.stepId())) {
        throw new IllegalArgumentException(
            "steps must not contain duplicate stepId values: " + step.stepId());
      }
    }
    return List.copyOf(copy);
  }

  static void requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
  }

  static void requireXlsxWorkbookPath(String path) { // LIM-002
    requireNonBlank(path, "path");
    if (!path.toLowerCase(Locale.ROOT).endsWith(".xlsx")) {
      throw new IllegalArgumentException(
          "path must point to a .xlsx workbook; .xls, .xlsm, and .xlsb are not supported: " + path);
    }
  }
}
