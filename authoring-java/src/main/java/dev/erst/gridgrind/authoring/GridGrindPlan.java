package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.dto.ExecutionJournalInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
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
  private Optional<String> planId;
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
      Optional<String> planId,
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<WorkbookStep> steps) {
    this.protocolVersion =
        Objects.requireNonNullElse(protocolVersion, GridGrindProtocolVersion.current());
    this.planId = Objects.requireNonNullElseGet(planId, Optional::empty);
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
  public GridGrindPlan planId(String newPlanId) {
    Objects.requireNonNull(newPlanId, "newPlanId must not be null");
    if (newPlanId.isBlank()) {
      throw new IllegalArgumentException("planId must not be blank");
    }
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
  public GridGrindPlan saveAs(Path path) {
    Objects.requireNonNull(path, "path must not be null");
    this.persistence = new WorkbookPlan.WorkbookPersistence.SaveAs(path.toString());
    return this;
  }

  /** Overwrites the source workbook path. */
  public GridGrindPlan overwriteSource() {
    this.persistence = new WorkbookPlan.WorkbookPersistence.OverwriteSource();
    return this;
  }

  /** Sets the execution journal level while preserving any existing execution mode or policy. */
  public GridGrindPlan journal(ExecutionJournalLevel level) {
    this.execution =
        new ExecutionPolicyInput(
            execution.mode(), new ExecutionJournalInput(level), execution.calculation());
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

  /** Serializes the canonical workbook plan as UTF-8 JSON bytes. */
  public byte[] toJsonBytes() throws IOException {
    return GridGrindJson.writeRequestBytes(toPlan());
  }

  /** Serializes the canonical workbook plan as an indented UTF-8 JSON string. */
  public String toJsonString() throws IOException {
    return new String(toJsonBytes(), StandardCharsets.UTF_8);
  }

  /** Writes the canonical workbook-plan JSON to one caller-owned output stream. */
  public void writeJson(OutputStream outputStream) throws IOException {
    GridGrindJson.writeRequest(outputStream, toPlan());
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
