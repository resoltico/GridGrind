package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.json.DuplicateStepId;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.MessageInvariant;
import dev.erst.gridgrind.contract.json.NonXlsxPath;
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
    @ProtocolField(optional = true)
        @JsonInclude(
            value = JsonInclude.Include.CUSTOM,
            valueFilter = ExecutionPolicyInput.DefaultFilter.class)
        ExecutionPolicyInput execution,
    @ProtocolField(optional = true)
        @JsonInclude(
            value = JsonInclude.Include.CUSTOM,
            valueFilter = FormulaEnvironmentInput.EmptyFilter.class)
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
    execution = execution == null ? ExecutionPolicyInput.defaults() : execution;
    formulaEnvironment =
        formulaEnvironment == null ? FormulaEnvironmentInput.empty() : formulaEnvironment;
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

  /** Returns the effective assertion policy after default normalization. */
  public AssertionModeInput assertionMode() {
    return effectiveExecution().effectiveAssertionMode();
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
    @JsonSubTypes.Type(value = WorkbookPersistence.Overwrite.class, name = "OVERWRITE"),
    @JsonSubTypes.Type(value = WorkbookPersistence.SaveAs.class, name = "SAVE_AS")
  })
  public sealed interface WorkbookPersistence {
    /** Leaves the workbook in memory only and does not persist it. */
    record None() implements WorkbookPersistence {}

    /** Explicit collision policy for SAVE_AS persistence. */
    enum IfExists {
      REJECT,
      REPLACE
    }

    /** Saves the workbook back to the exact path it was opened from. */
    record Overwrite(
        @JsonInclude(JsonInclude.Include.NON_ABSENT)
            Optional<OoxmlPersistenceSecurityInput> security)
        implements WorkbookPersistence {
      public Overwrite {
        security = normalizePersistenceSecurity(security);
      }

      /** Overwrites the source workbook with explicit package-security persistence settings. */
      public Overwrite(OoxmlPersistenceSecurityInput security) {
        this(Optional.of(Objects.requireNonNull(security, "security must not be null")));
      }
    }

    /** Saves the workbook to one `.xlsx` path with an explicit collision policy. */
    record SaveAs(
        String path,
        IfExists ifExists,
        @JsonInclude(JsonInclude.Include.NON_ABSENT)
            Optional<OoxmlPersistenceSecurityInput> security)
        implements WorkbookPersistence {
      public SaveAs {
        requireXlsxWorkbookPath(path);
        Objects.requireNonNull(ifExists, "ifExists must not be null");
        security = normalizePersistenceSecurity(security);
      }

      /** Saves the workbook to the supplied path with no explicit package-security settings. */
      public SaveAs(String path, IfExists ifExists) {
        this(path, ifExists, Optional.empty());
      }

      /** Saves the workbook to the supplied path with explicit package-security settings. */
      public SaveAs(String path, IfExists ifExists, OoxmlPersistenceSecurityInput security) {
        this(
            path,
            ifExists,
            Optional.of(Objects.requireNonNull(security, "security must not be null")));
      }
    }
  }

  /** Maximum number of steps permitted in one plan (LIM-024). */
  public static final int MAX_STEPS = 10_000;

  private static List<WorkbookStep> copySteps(List<WorkbookStep> steps) {
    Objects.requireNonNull(steps, "steps must not be null");
    if (steps.size() > MAX_STEPS) { // LIM-024
      throw new IllegalArgumentException(
          "steps must not exceed " + MAX_STEPS + " entries but was " + steps.size());
    }
    List<WorkbookStep> copy = new java.util.ArrayList<>(steps.size());
    Set<String> seen = new HashSet<>();
    for (int index = 0; index < steps.size(); index++) {
      WorkbookStep step = steps.get(index);
      String stepPath = "steps[" + index + "]";
      if (step == null) {
        throw nullStepException(stepPath);
      }
      copy.add(step);
      // LIM-006
      if (!seen.add(step.stepId())) {
        throw duplicateStepIdException(step.stepId(), stepPath + ".stepId");
      }
    }
    return List.copyOf(copy);
  }

  private static InvalidRequestException nullStepException(String stepPath) {
    return new InvalidRequestException(
        new MessageInvariant("steps must not contain nulls", Optional.of(stepPath)),
        Optional.of(stepPath),
        Optional.empty(),
        Optional.empty(),
        null);
  }

  private static InvalidRequestException duplicateStepIdException(String stepId, String jsonPath) {
    return new InvalidRequestException(
        new DuplicateStepId(stepId, jsonPath),
        Optional.of(jsonPath),
        Optional.empty(),
        Optional.empty(),
        null);
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
      // Preserve the local operand path so request decoding can qualify it precisely against
      // source.path, persistence.path, and nested formula-environment path owners.
      throw new InvalidRequestException(
          new NonXlsxPath(actualExtension(path), Optional.of("path")),
          Optional.empty(),
          Optional.empty(),
          Optional.empty(),
          null);
    }
  }

  private static String actualExtension(String path) {
    int lastSeparator = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'));
    int lastDot = path.lastIndexOf('.');
    if (lastDot <= lastSeparator || lastDot == path.length() - 1) {
      return "<none>";
    }
    return path.substring(lastDot);
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
