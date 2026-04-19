package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournalInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.executor.DefaultGridGrindRequestExecutor;
import dev.erst.gridgrind.executor.ExecutionInputBindings;
import dev.erst.gridgrind.executor.ExecutionJournalSink;
import dev.erst.gridgrind.executor.GridGrindRequestExecutor;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * Primary fluent Java entrypoint that compiles authored workflows to canonical WorkbookPlan data.
 */
public final class GridGrindPlan {
  private GridGrindProtocolVersion protocolVersion;
  private String planId;
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
        null,
        source,
        new WorkbookPlan.WorkbookPersistence.None(),
        null,
        null,
        List.of());
  }

  private GridGrindPlan(
      GridGrindProtocolVersion protocolVersion,
      String planId,
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
    this.execution = execution;
    this.formulaEnvironment = formulaEnvironment;
    this.steps = new ArrayList<>(Objects.requireNonNull(steps, "steps must not be null"));
    recountStepIds(this.steps);
  }

  /** Starts one plan against a brand-new empty workbook. */
  public static GridGrindPlan newWorkbook() {
    return new GridGrindPlan(new WorkbookPlan.WorkbookSource.New());
  }

  /** Starts one plan against an existing workbook path with no explicit open security. */
  public static GridGrindPlan open(Path path) {
    return open(path, null);
  }

  /** Starts one plan against an existing workbook path with explicit open security. */
  public static GridGrindPlan open(Path path, OoxmlOpenSecurityInput security) {
    Objects.requireNonNull(path, "path must not be null");
    return new GridGrindPlan(
        new WorkbookPlan.WorkbookSource.ExistingFile(path.toString(), security));
  }

  /** Continues authoring from an existing canonical WorkbookPlan. */
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
    if (newPlanId != null && newPlanId.isBlank()) {
      throw new IllegalArgumentException("planId must not be blank");
    }
    this.planId = newPlanId;
    return this;
  }

  /** Replaces the persistence policy directly with the canonical contract object. */
  public GridGrindPlan persistence(WorkbookPlan.WorkbookPersistence newPersistence) {
    this.persistence =
        Objects.requireNonNullElseGet(newPersistence, WorkbookPlan.WorkbookPersistence.None::new);
    return this;
  }

  /** Keeps the workbook in memory only and skips persistence. */
  public GridGrindPlan inMemoryOnly() {
    return persistence(new WorkbookPlan.WorkbookPersistence.None());
  }

  /** Saves the resulting workbook to a new `.xlsx` path. */
  public GridGrindPlan saveAs(Path path) {
    return saveAs(path, null);
  }

  /** Saves the resulting workbook to a new `.xlsx` path with explicit persistence security. */
  public GridGrindPlan saveAs(Path path, OoxmlPersistenceSecurityInput security) {
    Objects.requireNonNull(path, "path must not be null");
    return persistence(new WorkbookPlan.WorkbookPersistence.SaveAs(path.toString(), security));
  }

  /** Overwrites the source workbook path. */
  public GridGrindPlan overwriteSource() {
    return overwriteSource(null);
  }

  /** Overwrites the source workbook path with explicit persistence security. */
  public GridGrindPlan overwriteSource(OoxmlPersistenceSecurityInput security) {
    return persistence(new WorkbookPlan.WorkbookPersistence.OverwriteSource(security));
  }

  /** Replaces the full execution policy directly with the canonical contract object. */
  public GridGrindPlan execution(ExecutionPolicyInput newExecution) {
    this.execution = newExecution;
    return this;
  }

  /** Replaces the execution mode while preserving existing journal and calculation settings. */
  public GridGrindPlan mode(ExecutionModeInput mode) {
    this.execution =
        new ExecutionPolicyInput(
            mode,
            execution == null ? null : execution.journal(),
            execution == null ? null : execution.calculation());
    return this;
  }

  /** Replaces the execution journal level while preserving existing mode and calculation. */
  public GridGrindPlan journal(ExecutionJournalLevel level) {
    this.execution =
        new ExecutionPolicyInput(
            execution == null ? null : execution.mode(),
            new ExecutionJournalInput(level),
            execution == null ? null : execution.calculation());
    return this;
  }

  /** Replaces the calculation policy while preserving existing mode and journal settings. */
  public GridGrindPlan calculation(CalculationPolicyInput calculationPolicy) {
    this.execution =
        new ExecutionPolicyInput(
            execution == null ? null : execution.mode(),
            execution == null ? null : execution.journal(),
            calculationPolicy);
    return this;
  }

  /** Replaces the request-scoped formula environment. */
  public GridGrindPlan formulaEnvironment(FormulaEnvironmentInput newFormulaEnvironment) {
    this.formulaEnvironment = newFormulaEnvironment;
    return this;
  }

  /** Appends one already-built canonical workbook step. */
  public GridGrindPlan addStep(WorkbookStep step) {
    this.steps.add(Objects.requireNonNull(step, "step must not be null"));
    recountStepIds(this.steps);
    return this;
  }

  /** Appends one authored mutation with an auto-generated step id when needed. */
  public GridGrindPlan mutate(PlannedMutation mutation) {
    Objects.requireNonNull(mutation, "mutation must not be null");
    mutationCount++;
    steps.add(mutation.toStep(nextStepId("mutation", mutationCount)));
    return this;
  }

  /** Appends one mutation step from a canonical selector and action. */
  public GridGrindPlan mutate(Selector target, MutationAction action) {
    return mutate(new PlannedMutation(target, action));
  }

  /** Appends one mutation step from a fluent authoring target and canonical action. */
  public GridGrindPlan mutate(SelectorTarget target, MutationAction action) {
    Objects.requireNonNull(target, "target must not be null");
    return mutate(target.selector(), action);
  }

  /** Appends one authored inspection with an auto-generated step id when needed. */
  public GridGrindPlan inspect(PlannedInspection inspection) {
    Objects.requireNonNull(inspection, "inspection must not be null");
    inspectionCount++;
    steps.add(inspection.toStep(nextStepId("inspection", inspectionCount)));
    return this;
  }

  /** Appends one inspection step from a canonical selector and query. */
  public GridGrindPlan inspect(Selector target, InspectionQuery query) {
    return inspect(new PlannedInspection(target, query));
  }

  /** Appends one inspection step from a fluent authoring target and canonical query. */
  public GridGrindPlan inspect(SelectorTarget target, InspectionQuery query) {
    Objects.requireNonNull(target, "target must not be null");
    return inspect(target.selector(), query);
  }

  /** Appends one authored assertion with an auto-generated step id when needed. */
  public GridGrindPlan assertThat(PlannedAssertion assertion) {
    Objects.requireNonNull(assertion, "assertion must not be null");
    assertionCount++;
    steps.add(assertion.toStep(nextStepId("assertion", assertionCount)));
    return this;
  }

  /** Appends one assertion step from a canonical selector and assertion. */
  public GridGrindPlan assertThat(Selector target, Assertion assertion) {
    return assertThat(new PlannedAssertion(target, assertion));
  }

  /** Appends one assertion step from a fluent authoring target and canonical assertion. */
  public GridGrindPlan assertThat(SelectorTarget target, Assertion assertion) {
    Objects.requireNonNull(target, "target must not be null");
    return assertThat(target.selector(), assertion);
  }

  /** Returns the immutable canonical WorkbookPlan emitted by the authoring builder. */
  public WorkbookPlan toPlan() {
    return new WorkbookPlan(
        protocolVersion, planId, source, persistence, execution, formulaEnvironment, steps);
  }

  /** Serializes the canonical WorkbookPlan as UTF-8 JSON bytes. */
  public byte[] toJsonBytes() throws IOException {
    return GridGrindJson.writeRequestBytes(toPlan());
  }

  /** Serializes the canonical WorkbookPlan as an indented UTF-8 JSON string. */
  public String toJsonString() throws IOException {
    return new String(toJsonBytes(), StandardCharsets.UTF_8);
  }

  /** Writes the canonical WorkbookPlan JSON to one caller-owned output stream. */
  public void writeJson(OutputStream outputStream) throws IOException {
    GridGrindJson.writeRequest(outputStream, toPlan());
  }

  /** Executes the authored plan through the production in-process executor and default bindings. */
  public GridGrindResponse run() {
    return run(new DefaultGridGrindRequestExecutor());
  }

  /** Executes the authored plan through the provided executor and default process bindings. */
  public GridGrindResponse run(GridGrindRequestExecutor executor) {
    return run(executor, ExecutionInputBindings.processDefault(), ExecutionJournalSink.NOOP);
  }

  /** Executes the authored plan with explicit input bindings and no live journal sink. */
  public GridGrindResponse run(GridGrindRequestExecutor executor, ExecutionInputBindings bindings) {
    return run(executor, bindings, ExecutionJournalSink.NOOP);
  }

  /** Executes the authored plan with explicit input bindings and live journal emission. */
  public GridGrindResponse run(
      GridGrindRequestExecutor executor,
      ExecutionInputBindings bindings,
      ExecutionJournalSink sink) {
    return GridGrindRequestExecutor.requireNonNull(executor)
        .execute(
            toPlan(), bindings == null ? ExecutionInputBindings.processDefault() : bindings, sink);
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
        case dev.erst.gridgrind.contract.step.MutationStep _ -> mutationCount++;
        case dev.erst.gridgrind.contract.step.InspectionStep _ -> inspectionCount++;
        case dev.erst.gridgrind.contract.step.AssertionStep _ -> assertionCount++;
      }
    }
  }
}
