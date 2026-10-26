package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.drawing.PictureSpec;
import dev.erst.gridgrind.core.model.drawing.ShapeSpec;
import dev.erst.gridgrind.core.model.formula.CalculationPolicySpec;
import dev.erst.gridgrind.core.model.formula.CalculationStrategySpec;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextSpec;
import dev.erst.gridgrind.core.model.protection.CellProtectionPatch;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionSpec;
import dev.erst.gridgrind.core.model.style.CellAlignmentPatch;
import dev.erst.gridgrind.core.model.style.CellBorderPatch;
import dev.erst.gridgrind.core.model.style.CellFillSpec;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.FontHeightSpec;
import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationAction;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;

/** Shared step-native request builders for executor tests. */
final class ExecutorTestPlanSupport {
  private ExecutorTestPlanSupport() {}

  record PendingMutation(Selector target, MutationAction action) {
    PendingMutation {
      Objects.requireNonNull(target, "target must not be null");
      Objects.requireNonNull(action, "action must not be null");
    }
  }

  record PendingAssertion(String stepId, Selector target, Assertion assertion) {
    PendingAssertion {
      Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(target, "target must not be null");
      Objects.requireNonNull(assertion, "assertion must not be null");
    }
  }

  static PendingMutation mutate(Selector target, MutationAction action) {
    return new PendingMutation(target, action);
  }

  static PendingMutation mutate(StructuredMutationAction.SetTable action) {
    TableSpec table = action.table();
    return mutate(new TableSelector.ByNameOnSheet(table.name(), table.sheetName()), action);
  }

  static PendingMutation mutate(StructuredMutationAction.SetPivotTable action) {
    PivotTableSpec pivotTable = action.pivotTable();
    return mutate(
        new PivotTableSelector.ByNameOnSheet(pivotTable.name(), pivotTable.sheetName()), action);
  }

  static PendingAssertion assertThat(String stepId, Selector target, Assertion assertion) {
    return new PendingAssertion(stepId, target, assertion);
  }

  static InspectionStep inspect(String stepId, Selector target, InspectionQuery query) {
    return new InspectionStep(new StepId(stepId), target, query);
  }

  static List<PendingMutation> mutations(PendingMutation... mutations) {
    return mutations == null ? List.of() : List.of(mutations);
  }

  static List<PendingAssertion> assertions(PendingAssertion... assertions) {
    return assertions == null ? List.of() : List.of(assertions);
  }

  static List<InspectionStep> inspections(InspectionStep... inspections) {
    return inspections == null ? List.of() : List.of(inspections);
  }

  static CalculationPolicySpec calculateAll() {
    return CalculationPolicySpec.strategy(new CalculationStrategySpec.EvaluateAll());
  }

  static CalculationPolicySpec calculateAllAndMarkRecalculateOnOpen() {
    return new CalculationPolicySpec(new CalculationStrategySpec.EvaluateAll(), true);
  }

  static CalculationPolicySpec deferCalculation() {
    return CalculationPolicySpec.strategy(new CalculationStrategySpec.DeferredCalculation());
  }

  static CalculationPolicySpec requireEvaluation() {
    return CalculationPolicySpec.strategy(new CalculationStrategySpec.RequireEvaluation());
  }

  static CalculationPolicySpec calculateTargets(
      dev.erst.gridgrind.core.model.formula.FormulaCellTarget... cells) {
    return CalculationPolicySpec.strategy(
        new CalculationStrategySpec.EvaluateTargets(List.of(cells)));
  }

  static CalculationPolicySpec clearFormulaCaches() {
    return CalculationPolicySpec.strategy(new CalculationStrategySpec.ClearCachesOnly());
  }

  static CalculationPolicySpec markRecalculateOnOpen() {
    return new CalculationPolicySpec(new CalculationStrategySpec.DoNotCalculate(), true);
  }

  static ExecutionPolicySpec executionPolicy(CalculationPolicySpec calculation) {
    return ExecutionPolicySpec.calculation(calculation);
  }

  static ExecutionPolicySpec executionPolicy(
      ExecutionModeSpec mode, CalculationPolicySpec calculation) {
    return ExecutionPolicySpec.modeAndCalculation(mode, calculation);
  }

  static TextSourceSpec text(String value) {
    return TextSourceSpec.inline(value);
  }

  static BinarySourceSpec binary(String value) {
    return BinarySourceSpec.inlineBase64(value);
  }

  static <T> Optional<T> maybe(T value) {
    return Optional.ofNullable(value);
  }

  static Optional<dev.erst.gridgrind.core.source.SecretReference> maybeReference(String value) {
    return value == null || value.isBlank()
        ? Optional.empty()
        : Optional.of(new dev.erst.gridgrind.core.source.SecretReference(value));
  }

  static CellFontPatch fontInput(
      Boolean bold,
      Boolean italic,
      String fontName,
      FontHeightSpec fontHeight,
      ColorSpec fontColor,
      Boolean underline,
      Boolean strikeout) {
    return new CellFontPatch(
        maybe(bold),
        maybe(italic),
        maybe(fontName),
        maybe(fontHeight),
        maybe(fontColor),
        maybe(underline),
        maybe(strikeout));
  }

  static CellStylePatch styleInput(
      String numberFormat,
      CellAlignmentPatch alignment,
      CellFontPatch font,
      CellFillSpec fill,
      CellBorderPatch border,
      CellProtectionPatch protection) {
    return new CellStylePatch(
        maybe(numberFormat),
        maybe(alignment),
        maybe(font),
        maybe(fill),
        maybe(border),
        maybe(protection));
  }

  static CellValueSpec.Text textCell(String value) {
    return new CellValueSpec.Text(TextSourceSpec.inline(value));
  }

  static CellValueSpec.Formula formulaCell(String value) {
    return new CellValueSpec.Formula(TextSourceSpec.inline(value));
  }

  static RichTextRunSpec richTextRun(String value) {
    return new RichTextRunSpec(text(value), Optional.empty());
  }

  static RichTextRunSpec richTextRun(String value, CellFontPatch font) {
    return new RichTextRunSpec(text(value), maybe(font));
  }

  static RichTextRunSpec richTextRun(TextSourceSpec source, CellFontPatch font) {
    return new RichTextRunSpec(source, maybe(font));
  }

  static HeaderFooterTextSpec headerFooter(String left, String center, String right) {
    return new HeaderFooterTextSpec(text(left), text(center), text(right));
  }

  static DataValidationPromptSpec prompt(String title, String body, Boolean showPromptBox) {
    return new DataValidationPromptSpec(
        text(title), text(body), Objects.requireNonNullElse(showPromptBox, true));
  }

  static DataValidationErrorAlertSpec errorAlert(
      ExcelDataValidationErrorStyle style, String title, String body, Boolean showErrorBox) {
    return new DataValidationErrorAlertSpec(
        style, text(title), text(body), Objects.requireNonNullElse(showErrorBox, true));
  }

  static DataValidationSpec dataValidation(
      DataValidationRuleSpec rule,
      boolean allowBlank,
      boolean suppressDropDownArrow,
      DataValidationPromptSpec prompt,
      DataValidationErrorAlertSpec errorAlert) {
    return new DataValidationSpec(
        rule, allowBlank, suppressDropDownArrow, maybe(prompt), maybe(errorAlert));
  }

  static PictureSpec picture(
      String name, PictureDataSpec image, DrawingAnchorSpec anchor, TextSourceSpec description) {
    return new PictureSpec(name, image, anchor, maybe(description));
  }

  static ShapeSpec simpleShape(
      String name, DrawingAnchorSpec anchor, String presetGeometryToken, TextSourceSpec text) {
    return new ShapeSpec.SimpleShape(name, anchor, presetGeometryToken, maybe(text));
  }

  static ShapeSpec connectorShape(String name, DrawingAnchorSpec anchor) {
    return new ShapeSpec.Connector(name, anchor);
  }

  static WorkbookProtectionSpec workbookProtection(
      boolean structureLocked,
      boolean windowsLocked,
      boolean revisionsLocked,
      String workbookPassword,
      String revisionsPassword) {
    return new WorkbookProtectionSpec(
        structureLocked,
        windowsLocked,
        revisionsLocked,
        maybeReference(workbookPassword),
        maybeReference(revisionsPassword));
  }

  static MutationStep materializeMutation(PendingMutation mutation, int stepIndex) {
    Objects.requireNonNull(mutation, "mutation must not be null");
    if (stepIndex < 0) {
      throw new IllegalArgumentException("stepIndex must be non-negative");
    }
    return new MutationStep(
        new StepId(stepIdFor(stepIndex, mutation.action())), mutation.target(), mutation.action());
  }

  static AssertionStep materializeAssertion(PendingAssertion assertion) {
    Objects.requireNonNull(assertion, "assertion must not be null");
    return new AssertionStep(
        new StepId(assertion.stepId()), assertion.target(), assertion.assertion());
  }

  static List<WorkbookStep> steps(
      List<PendingMutation> mutations,
      List<PendingAssertion> assertions,
      List<InspectionStep> inspections) {
    List<PendingMutation> mutationCopy = mutations == null ? List.of() : List.copyOf(mutations);
    List<PendingAssertion> assertionCopy = assertions == null ? List.of() : List.copyOf(assertions);
    List<InspectionStep> inspectionCopy =
        inspections == null ? List.of() : List.copyOf(inspections);
    List<WorkbookStep> steps =
        new ArrayList<>(mutationCopy.size() + assertionCopy.size() + inspectionCopy.size());
    for (int i = 0; i < mutationCopy.size(); i++) {
      steps.add(materializeMutation(mutationCopy.get(i), i));
    }
    assertionCopy.stream().map(ExecutorTestPlanSupport::materializeAssertion).forEach(steps::add);
    steps.addAll(inspectionCopy);
    return List.copyOf(steps);
  }

  static List<WorkbookStep> steps(
      List<PendingMutation> mutations, List<InspectionStep> inspections) {
    return steps(mutations, List.of(), inspections);
  }

  static List<WorkbookStep> steps(List<PendingMutation> mutations, InspectionStep... inspections) {
    return steps(mutations, inspections == null ? List.of() : List.of(inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      List<PendingMutation> mutations,
      List<PendingAssertion> assertions,
      List<InspectionStep> inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicySpec.defaults(),
        FormulaEnvironmentSpec.empty(),
        steps(mutations, assertions, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      List<PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicySpec.defaults(),
        FormulaEnvironmentSpec.empty(),
        steps(mutations, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicySpec.defaults(),
        FormulaEnvironmentSpec.empty(),
        steps(mutations, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicySpec execution,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      List<PendingAssertion> assertions,
      List<InspectionStep> inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        Objects.requireNonNullElseGet(execution, ExecutionPolicySpec::defaults),
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentSpec::empty),
        steps(mutations, assertions, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicySpec execution,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return request(
        source, persistence, execution, formulaEnvironment, mutations, List.of(), inspections);
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicySpec execution,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return request(
        source,
        persistence,
        execution,
        formulaEnvironment,
        mutations,
        inspections == null ? List.of() : List.of(inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      List<PendingAssertion> assertions,
      List<InspectionStep> inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicySpec.defaults(),
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentSpec::empty),
        steps(mutations, assertions, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicySpec.defaults(),
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentSpec::empty),
        steps(mutations, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicySpec.defaults(),
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentSpec::empty),
        steps(mutations, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionModeSpec executionMode,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      List<PendingAssertion> assertions,
      List<InspectionStep> inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicySpec.mode(
            Objects.requireNonNullElseGet(executionMode, ExecutionModeSpec::defaults)),
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentSpec::empty),
        steps(mutations, assertions, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionModeSpec executionMode,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicySpec.mode(
            Objects.requireNonNullElseGet(executionMode, ExecutionModeSpec::defaults)),
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentSpec::empty),
        steps(mutations, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionModeSpec executionMode,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicySpec.mode(
            Objects.requireNonNullElseGet(executionMode, ExecutionModeSpec::defaults)),
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentSpec::empty),
        steps(mutations, inspections));
  }

  static List<String> inspectionIds(WorkbookResult.Success success) {
    return success.inspections().stream().map(result -> result.stepId().value()).toList();
  }

  static List<String> assertionIds(WorkbookResult.Success success) {
    return success.assertions().stream().map(result -> result.stepId().value()).toList();
  }

  static <T extends InspectionResult> T inspection(
      WorkbookResult.Success success, String stepId, Class<T> type) {
    return type.cast(
        success.inspections().stream()
            .filter(result -> result.stepId().value().equals(stepId))
            .findFirst()
            .orElseThrow(
                () -> new AssertionError("Missing inspection result for stepId " + stepId)));
  }

  static Path createWorkbookFile(String prefix) throws IOException {
    Path workbookPath = Files.createTempFile(prefix, ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").cells().setCell("A1", new CellValue.Text("Header"));
      ExecutionContextFixtureSupport.saveWorkbook(workbook, workbookPath);
    }
    return workbookPath;
  }

  private static String stepIdFor(int stepIndex, MutationAction action) {
    String normalizedType = action.actionType().toLowerCase(Locale.ROOT).replace('_', '-');
    return "step-" + String.format(Locale.ROOT, "%02d", stepIndex + 1) + "-" + normalizedType;
  }
}
