package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.limits.Limit;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.model.formula.CalculationPolicySpec;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec;
import dev.erst.gridgrind.core.problem.DuplicateStepId;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.problem.MessageInvariant;
import dev.erst.gridgrind.core.problem.NonXlsxPath;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import dev.erst.gridgrind.core.vocabulary.PlanId;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;

/**
 * Complete GridGrind workbook plan for source, execution settings, ordered steps, and persistence.
 */
@dev.erst.gridgrind.core.metadata.CatalogType(
    id = "WorkbookPlan",
    summary =
        "Complete GridGrind plan for workbook source, execution policy, formula environment, ordered mutation, assertion, and inspection steps, and persistence.")
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_DUPLICATE_STEP_IDS,
  dev.erst.gridgrind.core.limits.Limit.ID_FILE_FORMAT,
  dev.erst.gridgrind.core.limits.Limit.ID_STEP_COUNT
})
public record WorkbookPlan(
    ProtocolVersion protocolVersion,
    Optional<PlanId> planId,
    WorkbookSource source,
    WorkbookPersistence persistence,
    @ProtocolField(optional = true)
        @WireDefault(factory = ExecutionPolicySpec.class, method = "defaults", omit = true)
        ExecutionPolicySpec execution,
    @ProtocolField(optional = true)
        @WireDefault(factory = FormulaEnvironmentSpec.class, method = "empty", omit = true)
        FormulaEnvironmentSpec formulaEnvironment,
    List<WorkbookStep> steps) {
  /** Normalizes one authored plan instance. */
  public WorkbookPlan {
    Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    Objects.requireNonNull(planId, "planId must not be null");
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
      ExecutionPolicySpec execution,
      FormulaEnvironmentSpec formulaEnvironment,
      List<WorkbookStep> steps) {
    return new WorkbookPlan(
        ProtocolVersion.CURRENT,
        Optional.empty(),
        source,
        persistence,
        execution,
        formulaEnvironment,
        steps);
  }

  /** Creates one caller-identified canonical plan on the current protocol version. */
  public static WorkbookPlan identified(
      PlanId planId,
      WorkbookSource source,
      WorkbookPersistence persistence,
      ExecutionPolicySpec execution,
      FormulaEnvironmentSpec formulaEnvironment,
      List<WorkbookStep> steps) {
    return new WorkbookPlan(
        ProtocolVersion.CURRENT,
        Optional.of(planId),
        source,
        persistence,
        execution,
        formulaEnvironment,
        steps);
  }

  /** Creates one caller-identified canonical plan on an explicit protocol version. */
  public static WorkbookPlan identified(
      ProtocolVersion protocolVersion,
      PlanId planId,
      WorkbookSource source,
      WorkbookPersistence persistence,
      ExecutionPolicySpec execution,
      FormulaEnvironmentSpec formulaEnvironment,
      List<WorkbookStep> steps) {
    return new WorkbookPlan(
        protocolVersion,
        Optional.of(planId),
        source,
        persistence,
        execution,
        formulaEnvironment,
        steps);
  }

  /** Returns the effective execution policy after default normalization. */
  public ExecutionPolicySpec effectiveExecution() {
    return execution;
  }

  /** Returns the normalized execution mode family after request-default expansion. */
  public ExecutionModeSpec executionMode() {
    return execution.mode();
  }

  /** Returns the effective execution mode family after default normalization. */
  public ExecutionModeSpec effectiveExecutionMode() {
    return effectiveExecution().effectiveMode();
  }

  /** Returns the effective execution journal level after default normalization. */
  public ExecutionJournalLevel journalLevel() {
    return effectiveExecution().effectiveJournalLevel();
  }

  /** Returns the effective calculation policy after default normalization. */
  public CalculationPolicySpec calculationPolicy() {
    return effectiveExecution().effectiveCalculation();
  }

  /** Returns the effective assertion policy after default normalization. */
  public AssertionModeSpec assertionMode() {
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
  @Union(discriminator = "type")
  public sealed interface WorkbookSource {
    /** Creates a brand-new empty workbook in memory. */
    @Variant(
        id = "NEW",
        summary =
            "Create a brand-new empty workbook. A new workbook starts with zero sheets; use ENSURE_SHEET to create the first sheet.")
    record New() implements WorkbookSource {}

    /** Opens an existing workbook file from disk. */
    @Variant(
        id = "EXISTING",
        summary =
            "Open an existing .xlsx workbook from disk. source.security.password unlocks encrypted OOXML packages.",
        noteRefs = {"requestOwnedPathRule"})
    record ExistingFile(
        @dev.erst.gridgrind.core.metadata.FieldRule(
                kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.PATH_SUFFIX,
                text = dev.erst.gridgrind.core.limits.Limit.WORKBOOK_PATH_SUFFIX)
            String path,
        Optional<OoxmlOpenSecuritySpec> security)
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
      public ExistingFile(String path, OoxmlOpenSecuritySpec security) {
        this(path, Optional.of(Objects.requireNonNull(security, "security must not be null")));
      }
    }
  }

  /** Describes whether and where the resulting workbook should be saved. */
  @Union(discriminator = "type")
  public sealed interface WorkbookPersistence {
    /** Leaves the workbook in memory only and does not persist it. */
    @Variant(
        id = "NONE",
        summary = "Keep the workbook in memory only. The response persistence.type echoes NONE.")
    record None() implements WorkbookPersistence {}

    /** Explicit collision policy for SAVE_AS persistence. */
    enum IfExists {
      REJECT,
      REPLACE
    }

    /** Saves the workbook back to the exact path it was opened from. */
    @Variant(
        id = "OVERWRITE",
        summary =
            "Overwrite the opened source workbook at source.path. No path field is accepted on OVERWRITE; the write target is the same path opened by the EXISTING source. persistence.security is required and declares both final OOXML encryption and signature dispositions. The response persistence.type echoes OVERWRITE, includes sourcePath (the original source path string) whenever an EXISTING source path was available, and otherwise omits sourcePath instead of inventing one. It carries write.status=WRITTEN with executionPath after a successful save or write.status=NOT_WRITTEN when the run fails before any file is updated.",
        noteRefs = {"requestOwnedPathRule"})
    record Overwrite(Optional<OoxmlPersistenceSecuritySpec> security)
        implements WorkbookPersistence {
      public Overwrite {
        security = normalizePersistenceSecurity(security);
      }

      /** Overwrites the source workbook with explicit package-security persistence settings. */
      public Overwrite(OoxmlPersistenceSecuritySpec security) {
        this(Optional.of(Objects.requireNonNull(security, "security must not be null")));
      }
    }

    /** Saves the workbook to one `.xlsx` path with an explicit collision policy. */
    @Variant(
        id = "SAVE_AS",
        summary =
            "Save the workbook to one .xlsx path with an explicit ifExists collision policy. ifExists=REJECT requires the target path to be absent. ifExists=REPLACE enables create-or-replace. Existing-source writes require persistence.security to declare both final OOXML encryption and signature dispositions; NEW writes default to NONE on both axes when security is omitted. The response persistence.type echoes SAVE_AS, includes requestedPath (the literal path from the request), and carries write.status=WRITTEN with executionPath (the absolute normalized path where the file was written) or write.status=NOT_WRITTEN when the run fails before the target file is created. Missing parent directories are created automatically.",
        noteRefs = {"requestOwnedPathRule"})
    record SaveAs(
        @dev.erst.gridgrind.core.metadata.FieldRule(
                kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.PATH_SUFFIX,
                text = dev.erst.gridgrind.core.limits.Limit.WORKBOOK_PATH_SUFFIX)
            String path,
        IfExists ifExists,
        Optional<OoxmlPersistenceSecuritySpec> security)
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
      public SaveAs(String path, IfExists ifExists, OoxmlPersistenceSecuritySpec security) {
        this(
            path,
            ifExists,
            Optional.of(Objects.requireNonNull(security, "security must not be null")));
      }
    }
  }

  /** Maximum number of steps permitted in one plan. */
  private static List<WorkbookStep> copySteps(List<WorkbookStep> steps) {
    Objects.requireNonNull(steps, "steps must not be null");
    if (steps.size() > Limit.MAX_STEPS) {
      throw new IllegalArgumentException(
          "steps must not exceed " + Limit.MAX_STEPS + " entries but was " + steps.size());
    }
    List<WorkbookStep> copy = new java.util.ArrayList<>(steps.size());
    Set<StepId> seen = new HashSet<>();
    for (int index = 0; index < steps.size(); index++) {
      WorkbookStep step = steps.get(index);
      String stepPath = "steps[" + index + "]";
      if (step == null) {
        throw nullStepException(stepPath);
      }
      copy.add(step);
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

  private static InvalidRequestException duplicateStepIdException(StepId stepId, String jsonPath) {
    return new InvalidRequestException(
        new DuplicateStepId(stepId, jsonPath),
        Optional.of(jsonPath),
        Optional.empty(),
        Optional.empty(),
        null);
  }

  /** Validates the required workbook file suffix. */
  public static void requireXlsxWorkbookPath(String path) {
    TextValidation.requireNonBlank(path, "path");
    if (!path.toLowerCase(Locale.ROOT).endsWith(Limit.WORKBOOK_PATH_SUFFIX)) {
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

  private static Optional<OoxmlOpenSecuritySpec> normalizeOpenSecurity(
      Optional<OoxmlOpenSecuritySpec> security) {
    Optional<OoxmlOpenSecuritySpec> normalized =
        Objects.requireNonNull(security, "security must not be null");
    return normalized.filter(value -> value.passwordRef().isPresent());
  }

  private static Optional<OoxmlPersistenceSecuritySpec> normalizePersistenceSecurity(
      Optional<OoxmlPersistenceSecuritySpec> security) {
    return Objects.requireNonNull(security, "security must not be null");
  }
}
