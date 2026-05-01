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
import java.util.Optional;
import java.util.Set;

/**
 * Complete GridGrind workbook plan for source, execution settings, ordered steps, and persistence.
 */
public record WorkbookPlan(
    GridGrindProtocolVersion protocolVersion,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> planId,
    WorkbookSource source,
    WorkbookPersistence persistence,
    ExecutionPolicyInput execution,
    FormulaEnvironmentInput formulaEnvironment,
    List<WorkbookStep> steps) {
  /** Normalizes one authored plan instance. */
  public WorkbookPlan {
    Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    Optional<String> normalizedPlanId = Objects.requireNonNull(planId, "planId must not be null");
    if (normalizedPlanId.isPresent()) {
      normalizedPlanId = Optional.of(requireNonBlank(normalizedPlanId.orElseThrow(), "planId"));
    }
    planId = normalizedPlanId;
    Objects.requireNonNull(source, "source must not be null");
    Objects.requireNonNull(persistence, "persistence must not be null");
    Objects.requireNonNull(execution, "execution must not be null");
    Objects.requireNonNull(formulaEnvironment, "formulaEnvironment must not be null");
    steps = copySteps(steps);
  }

  /**
   * Creates one canonical plan on the current protocol version with explicit execution settings.
   */
  public static WorkbookPlan standard(
      WorkbookSource source,
      WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<WorkbookStep> steps) {
    return new WorkbookPlan(
        GridGrindProtocolVersion.current(),
        Optional.empty(),
        source,
        persistence,
        execution,
        formulaEnvironment,
        steps);
  }

  /** Creates one caller-identified canonical plan on the current protocol version. */
  public static WorkbookPlan identified(
      String planId,
      WorkbookSource source,
      WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<WorkbookStep> steps) {
    return new WorkbookPlan(
        GridGrindProtocolVersion.current(),
        Optional.of(requireNonBlank(planId, "planId")),
        source,
        persistence,
        execution,
        formulaEnvironment,
        steps);
  }

  /** Creates one caller-identified canonical plan on an explicit protocol version. */
  public static WorkbookPlan identified(
      GridGrindProtocolVersion protocolVersion,
      String planId,
      WorkbookSource source,
      WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<WorkbookStep> steps) {
    return new WorkbookPlan(
        protocolVersion,
        Optional.of(requireNonBlank(planId, "planId")),
        source,
        persistence,
        execution,
        formulaEnvironment,
        steps);
  }

  /** Returns the effective execution policy after default normalization. */
  public ExecutionPolicyInput effectiveExecution() {
    return execution;
  }

  /** Returns the normalized execution mode family after request-default expansion. */
  public ExecutionModeInput executionMode() {
    return execution.mode();
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

  /** Returns the authored steps partitioned by family in authored order. */
  public StepPartition stepPartition() {
    return partitionSteps(steps);
  }

  /** One authored-step partition with each family preserved in original authored order. */
  public record StepPartition(
      List<MutationStep> mutations,
      List<AssertionStep> assertions,
      List<InspectionStep> inspections) {
    public StepPartition {
      mutations = List.copyOf(Objects.requireNonNull(mutations, "mutations must not be null"));
      assertions = List.copyOf(Objects.requireNonNull(assertions, "assertions must not be null"));
      inspections =
          List.copyOf(Objects.requireNonNull(inspections, "inspections must not be null"));
    }
  }

  private static StepPartition partitionSteps(List<WorkbookStep> steps) {
    List<MutationStep> mutationSteps = new java.util.ArrayList<>();
    List<AssertionStep> assertionSteps = new java.util.ArrayList<>();
    List<InspectionStep> inspectionSteps = new java.util.ArrayList<>();
    for (WorkbookStep step : steps) {
      switch (step) {
        case MutationStep mutationStep -> mutationSteps.add(mutationStep);
        case AssertionStep assertionStep -> assertionSteps.add(assertionStep);
        case InspectionStep inspectionStep -> inspectionSteps.add(inspectionStep);
      }
    }
    return new StepPartition(mutationSteps, assertionSteps, inspectionSteps);
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
        String path,
        @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<OoxmlOpenSecurityInput> security)
        implements WorkbookSource {
      public ExistingFile {
        requireXlsxWorkbookPath(path);
        security = normalizeOpenSecurity(security);
      }

      /**
       * Opens the existing workbook at the supplied path with no explicit package-open settings.
       */
      public ExistingFile(String path) {
        this(path, Optional.empty());
      }

      /** Opens the existing workbook at the supplied path with explicit package-open settings. */
      public ExistingFile(String path, OoxmlOpenSecurityInput security) {
        this(path, Optional.of(Objects.requireNonNull(security, "security must not be null")));
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
        @JsonInclude(JsonInclude.Include.NON_ABSENT)
            Optional<OoxmlPersistenceSecurityInput> security)
        implements WorkbookPersistence {
      public OverwriteSource {
        security = normalizePersistenceSecurity(security);
      }

      /** Overwrites the source workbook with no explicit package-security persistence settings. */
      public OverwriteSource() {
        this(Optional.empty());
      }

      /** Overwrites the source workbook with explicit package-security persistence settings. */
      public OverwriteSource(OoxmlPersistenceSecurityInput security) {
        this(Optional.of(Objects.requireNonNull(security, "security must not be null")));
      }
    }

    /** Saves the workbook to a new `.xlsx` path. */
    record SaveAs(
        String path,
        @JsonInclude(JsonInclude.Include.NON_ABSENT)
            Optional<OoxmlPersistenceSecurityInput> security)
        implements WorkbookPersistence {
      public SaveAs {
        requireXlsxWorkbookPath(path);
        security = normalizePersistenceSecurity(security);
      }

      /** Saves the workbook to the supplied path with no explicit package-security settings. */
      public SaveAs(String path) {
        this(path, Optional.empty());
      }

      /** Saves the workbook to the supplied path with explicit package-security settings. */
      public SaveAs(String path, OoxmlPersistenceSecurityInput security) {
        this(path, Optional.of(Objects.requireNonNull(security, "security must not be null")));
      }
    }
  }

  private static List<WorkbookStep> copySteps(List<WorkbookStep> steps) {
    Objects.requireNonNull(steps, "steps must not be null");
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

  static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  static void requireXlsxWorkbookPath(String path) { // LIM-002
    requireNonBlank(path, "path");
    if (!path.toLowerCase(Locale.ROOT).endsWith(".xlsx")) {
      throw new IllegalArgumentException(
          "path must point to a .xlsx workbook; .xls, .xlsm, and .xlsb are not supported: " + path);
    }
  }

  private static Optional<OoxmlOpenSecurityInput> normalizeOpenSecurity(
      Optional<OoxmlOpenSecurityInput> security) {
    Optional<OoxmlOpenSecurityInput> normalized =
        Objects.requireNonNull(security, "security must not be null");
    return normalized.filter(value -> !value.isEmpty());
  }

  private static Optional<OoxmlPersistenceSecurityInput> normalizePersistenceSecurity(
      Optional<OoxmlPersistenceSecurityInput> security) {
    return Objects.requireNonNull(security, "security must not be null");
  }
}
