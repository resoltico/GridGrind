package dev.erst.gridgrind.core.authoring;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentInput;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.ExecutionJournalInput;
import dev.erst.gridgrind.core.plan.ExecutionPolicyInput;
import dev.erst.gridgrind.core.plan.GridGrindProtocolVersion;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import dev.erst.gridgrind.core.vocabulary.PlanId;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/**
 * Primary fluent Java entrypoint that authors focused workbook workflows and compiles them to the
 * canonical {@link WorkbookPlan}.
 *
 * <p>This module intentionally stops at the contract boundary. Callers that need in-process
 * execution should pass {@link #toPlan()} to an executor explicitly instead of relying on hidden
 * runtime coupling from the authoring surface.
 */
public final class GridGrindPlan {
  private GridGrindProtocolVersion protocolVersion;
  private Optional<PlanId> planId;
  private WorkbookPlan.WorkbookSource source;
  private WorkbookPlan.WorkbookPersistence persistence;
  private ExecutionPolicyInput execution;
  private FormulaEnvironmentInput formulaEnvironment;
  private final List<WorkbookStep> steps;
  private int mutationCount;
  private int inspectionCount;
  private int assertionCount;

  private GridGrindPlan(WorkbookPlan.WorkbookSource source) {
    this(
        GridGrindProtocolVersion.current(),
        Optional.empty(),
        source,
        new WorkbookPlan.WorkbookPersistence.None(),
        ExecutionPolicyInput.defaults(),
        FormulaEnvironmentInput.empty(),
        List.of());
  }

  private GridGrindPlan(
      GridGrindProtocolVersion protocolVersion,
      Optional<PlanId> planId,
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<WorkbookStep> steps) {
    this.protocolVersion =
        Objects.requireNonNullElse(protocolVersion, GridGrindProtocolVersion.current());
    this.planId = planId;
    this.source = Objects.requireNonNull(source, "source must not be null");
    this.persistence =
        Objects.requireNonNullElseGet(persistence, WorkbookPlan.WorkbookPersistence.None::new);
    this.execution = Objects.requireNonNullElseGet(execution, ExecutionPolicyInput::defaults);
    this.formulaEnvironment =
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentInput::empty);
    this.steps = new ArrayList<>(Objects.requireNonNull(steps, "steps must not be null"));
    recountStepIds(this.steps);
  }

  /** Starts one plan against a brand-new empty workbook. */
  public static GridGrindPlan newWorkbook() {
    return new GridGrindPlan(new WorkbookPlan.WorkbookSource.New());
  }

  /** Starts one plan against an existing workbook path. */
  public static GridGrindPlan open(Path path) {
    Objects.requireNonNull(path, "path must not be null");
    return new GridGrindPlan(new WorkbookPlan.WorkbookSource.ExistingFile(path.toString()));
  }

  /** Continues authoring from an existing canonical workbook plan. */
  public static GridGrindPlan from(WorkbookPlan plan) {
    Objects.requireNonNull(plan, "plan must not be null");
    return new GridGrindPlan(
        plan.protocolVersion(),
        plan.planId(),
        plan.source(),
        plan.persistence(),
        plan.execution(),
        plan.formulaEnvironment(),
        plan.steps());
  }

  /** Sets an explicit plan id for journaling and external correlation. */
  public GridGrindPlan planId(PlanId newPlanId) {
    this.planId = Optional.of(newPlanId);
    return this;
  }

  /** Clears any previously assigned plan id. */
  public GridGrindPlan clearPlanId() {
    this.planId = Optional.empty();
    return this;
  }

  /** Keeps the workbook in memory only and skips persistence. */
  public GridGrindPlan inMemoryOnly() {
    this.persistence = new WorkbookPlan.WorkbookPersistence.None();
    return this;
  }

  /** Saves the resulting workbook to a new `.xlsx` path. */
  public GridGrindPlan saveAs(Path path, WorkbookPlan.WorkbookPersistence.IfExists ifExists) {
    Objects.requireNonNull(path, "path must not be null");
    Objects.requireNonNull(ifExists, "ifExists must not be null");
    if (source instanceof WorkbookPlan.WorkbookSource.ExistingFile) {
      throw new IllegalStateException(
          "EXISTING-workbook SAVE_AS requires an explicit OOXML persistence security policy");
    }
    this.persistence = new WorkbookPlan.WorkbookPersistence.SaveAs(path.toString(), ifExists);
    return this;
  }

  /** Saves the resulting workbook with an explicit OOXML package-security policy. */
  public GridGrindPlan saveAs(
      Path path,
      WorkbookPlan.WorkbookPersistence.IfExists ifExists,
      OoxmlPersistenceSecurityInput security) {
    Objects.requireNonNull(path, "path must not be null");
    Objects.requireNonNull(ifExists, "ifExists must not be null");
    this.persistence =
        new WorkbookPlan.WorkbookPersistence.SaveAs(
            path.toString(),
            ifExists,
            Objects.requireNonNull(security, "security must not be null"));
    return this;
  }

  /** Overwrites the source workbook path with an explicit OOXML package-security policy. */
  public GridGrindPlan overwrite(OoxmlPersistenceSecurityInput security) {
    this.persistence =
        new WorkbookPlan.WorkbookPersistence.Overwrite(
            Objects.requireNonNull(security, "security must not be null"));
    return this;
  }

  /** Sets the execution journal level while preserving any existing execution mode or policy. */
  public GridGrindPlan journal(ExecutionJournalLevel level) {
    this.execution =
        new ExecutionPolicyInput(
            execution.mode(),
            new ExecutionJournalInput(level),
            execution.calculation(),
            execution.assertionMode());
    return this;
  }

  /** Appends one authored mutation with an auto-generated step id when needed. */
  public GridGrindPlan mutate(PlannedMutation mutation) {
    Objects.requireNonNull(mutation, "mutation must not be null");
    mutationCount++;
    steps.add(mutation.toStep(nextStepId("mutation", mutationCount)));
    return this;
  }

  /** Appends one authored inspection with an auto-generated step id when needed. */
  public GridGrindPlan inspect(PlannedInspection inspection) {
    Objects.requireNonNull(inspection, "inspection must not be null");
    inspectionCount++;
    steps.add(inspection.toStep(nextStepId("inspection", inspectionCount)));
    return this;
  }

  /** Appends one authored assertion with an auto-generated step id when needed. */
  public GridGrindPlan assertThat(PlannedAssertion assertion) {
    Objects.requireNonNull(assertion, "assertion must not be null");
    assertionCount++;
    steps.add(assertion.toStep(nextStepId("assertion", assertionCount)));
    return this;
  }

  /** Returns the immutable canonical {@link WorkbookPlan} emitted by this authoring builder. */
  public WorkbookPlan toPlan() {
    return new WorkbookPlan(
        protocolVersion, planId, source, persistence, execution, formulaEnvironment, steps);
  }

  private String nextStepId(String prefix, int index) {
    return prefix + "-" + String.format(java.util.Locale.ROOT, "%03d", index);
  }

  private void recountStepIds(List<WorkbookStep> existingSteps) {
    mutationCount = 0;
    inspectionCount = 0;
    assertionCount = 0;
    for (WorkbookStep step : existingSteps) {
      switch (step) {
        case MutationStep _ -> mutationCount++;
        case InspectionStep _ -> inspectionCount++;
        case AssertionStep _ -> assertionCount++;
      }
    }
  }
}
