package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.assertion.*;
import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.CellAlignmentInput;
import dev.erst.gridgrind.contract.dto.CellBorderInput;
import dev.erst.gridgrind.contract.dto.CellFillInput;
import dev.erst.gridgrind.contract.dto.CellFontInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellProtectionInput;
import dev.erst.gridgrind.contract.dto.CellStylePatchInput;
import dev.erst.gridgrind.contract.dto.ColorInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FontHeightInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.HeaderFooterTextInput;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
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
    TableInput table = action.table();
    return mutate(new TableSelector.ByNameOnSheet(table.name(), table.sheetName()), action);
  }

  static PendingMutation mutate(StructuredMutationAction.SetPivotTable action) {
    PivotTableInput pivotTable = action.pivotTable();
    return mutate(
        new PivotTableSelector.ByNameOnSheet(pivotTable.name(), pivotTable.sheetName()), action);
  }

  static PendingMutation mutate(StructuredMutationAction.SetNamedRange action) {
    return mutate(namedRangeSelector(action.name(), action.scope()), action);
  }

  static PendingAssertion assertThat(String stepId, Selector target, Assertion assertion) {
    return new PendingAssertion(stepId, target, assertion);
  }

  static InspectionStep inspect(String stepId, Selector target, InspectionQuery query) {
    return new InspectionStep(stepId, target, query);
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

  static CalculationPolicyInput calculateAll() {
    return CalculationPolicyInput.strategy(new CalculationStrategyInput.EvaluateAll());
  }

  static CalculationPolicyInput calculateAllAndMarkRecalculateOnOpen() {
    return new CalculationPolicyInput(new CalculationStrategyInput.EvaluateAll(), true);
  }

  static CalculationPolicyInput calculateTargets(CellSelector.QualifiedAddress... cells) {
    return CalculationPolicyInput.strategy(
        new CalculationStrategyInput.EvaluateTargets(List.of(cells)));
  }

  static CalculationPolicyInput clearFormulaCaches() {
    return CalculationPolicyInput.strategy(new CalculationStrategyInput.ClearCachesOnly());
  }

  static CalculationPolicyInput markRecalculateOnOpen() {
    return new CalculationPolicyInput(new CalculationStrategyInput.DoNotCalculate(), true);
  }

  static ExecutionPolicyInput executionPolicy(CalculationPolicyInput calculation) {
    return ExecutionPolicyInput.calculation(calculation);
  }

  static ExecutionPolicyInput executionPolicy(
      ExecutionModeInput mode, CalculationPolicyInput calculation) {
    return ExecutionPolicyInput.modeAndCalculation(mode, calculation);
  }

  static TextSourceInput text(String value) {
    return TextSourceInput.inline(value);
  }

  static BinarySourceInput binary(String value) {
    return BinarySourceInput.inlineBase64(value);
  }

  static <T> Optional<T> maybe(T value) {
    return Optional.ofNullable(value);
  }

  static CellFontInput fontInput(
      Boolean bold,
      Boolean italic,
      String fontName,
      FontHeightInput fontHeight,
      ColorInput fontColor,
      Boolean underline,
      Boolean strikeout) {
    return new CellFontInput(
        maybe(bold),
        maybe(italic),
        maybe(fontName),
        maybe(fontHeight),
        maybe(fontColor),
        maybe(underline),
        maybe(strikeout));
  }

  static CellStylePatchInput styleInput(
      String numberFormat,
      CellAlignmentInput alignment,
      CellFontInput font,
      CellFillInput fill,
      CellBorderInput border,
      CellProtectionInput protection) {
    return new CellStylePatchInput(
        maybe(numberFormat),
        maybe(alignment),
        maybe(font),
        maybe(fill),
        maybe(border),
        maybe(protection));
  }

  static CellInput.Text textCell(String value) {
    return new CellInput.Text(TextSourceInput.inline(value));
  }

  static CellInput.Formula formulaCell(String value) {
    return new CellInput.Formula(TextSourceInput.inline(value));
  }

  static RichTextRunInput richTextRun(String value) {
    return new RichTextRunInput(text(value), Optional.empty());
  }

  static RichTextRunInput richTextRun(String value, CellFontInput font) {
    return new RichTextRunInput(text(value), maybe(font));
  }

  static RichTextRunInput richTextRun(TextSourceInput source, CellFontInput font) {
    return new RichTextRunInput(source, maybe(font));
  }

  static HeaderFooterTextInput headerFooter(String left, String center, String right) {
    return new HeaderFooterTextInput(text(left), text(center), text(right));
  }

  static DataValidationPromptInput prompt(String title, String body, Boolean showPromptBox) {
    return new DataValidationPromptInput(
        text(title), text(body), Objects.requireNonNullElse(showPromptBox, true));
  }

  static DataValidationErrorAlertInput errorAlert(
      ExcelDataValidationErrorStyle style, String title, String body, Boolean showErrorBox) {
    return new DataValidationErrorAlertInput(
        style, text(title), text(body), Objects.requireNonNullElse(showErrorBox, true));
  }

  static DataValidationInput dataValidation(
      DataValidationRuleInput rule,
      boolean allowBlank,
      boolean suppressDropDownArrow,
      DataValidationPromptInput prompt,
      DataValidationErrorAlertInput errorAlert) {
    return new DataValidationInput(
        rule, allowBlank, suppressDropDownArrow, maybe(prompt), maybe(errorAlert));
  }

  static PictureInput picture(
      String name, PictureDataInput image, DrawingAnchorInput anchor, TextSourceInput description) {
    return new PictureInput(name, image, anchor, maybe(description));
  }

  static ShapeInput simpleShape(
      String name, DrawingAnchorInput anchor, String presetGeometryToken, TextSourceInput text) {
    return new ShapeInput.SimpleShape(name, anchor, presetGeometryToken, maybe(text));
  }

  static ShapeInput connectorShape(String name, DrawingAnchorInput anchor) {
    return new ShapeInput.Connector(name, anchor);
  }

  static WorkbookProtectionInput workbookProtection(
      boolean structureLocked,
      boolean windowsLocked,
      boolean revisionsLocked,
      String workbookPassword,
      String revisionsPassword) {
    return new WorkbookProtectionInput(
        structureLocked,
        windowsLocked,
        revisionsLocked,
        maybe(workbookPassword),
        maybe(revisionsPassword));
  }

  static MutationStep materializeMutation(PendingMutation mutation, int stepIndex) {
    Objects.requireNonNull(mutation, "mutation must not be null");
    if (stepIndex < 0) {
      throw new IllegalArgumentException("stepIndex must be non-negative");
    }
    return new MutationStep(
        stepIdFor(stepIndex, mutation.action()), mutation.target(), mutation.action());
  }

  static AssertionStep materializeAssertion(PendingAssertion assertion) {
    Objects.requireNonNull(assertion, "assertion must not be null");
    return new AssertionStep(assertion.stepId(), assertion.target(), assertion.assertion());
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
        ExecutionPolicyInput.defaults(),
        FormulaEnvironmentInput.empty(),
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
        ExecutionPolicyInput.defaults(),
        FormulaEnvironmentInput.empty(),
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
        ExecutionPolicyInput.defaults(),
        FormulaEnvironmentInput.empty(),
        steps(mutations, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      List<PendingAssertion> assertions,
      List<InspectionStep> inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        Objects.requireNonNullElseGet(execution, ExecutionPolicyInput::defaults),
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentInput::empty),
        steps(mutations, assertions, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return request(
        source, persistence, execution, formulaEnvironment, mutations, List.of(), inspections);
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
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
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      List<PendingAssertion> assertions,
      List<InspectionStep> inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicyInput.defaults(),
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentInput::empty),
        steps(mutations, assertions, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicyInput.defaults(),
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentInput::empty),
        steps(mutations, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicyInput.defaults(),
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentInput::empty),
        steps(mutations, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionModeInput executionMode,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      List<PendingAssertion> assertions,
      List<InspectionStep> inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicyInput.mode(
            Objects.requireNonNullElseGet(executionMode, ExecutionModeInput::defaults)),
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentInput::empty),
        steps(mutations, assertions, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionModeInput executionMode,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicyInput.mode(
            Objects.requireNonNullElseGet(executionMode, ExecutionModeInput::defaults)),
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentInput::empty),
        steps(mutations, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionModeInput executionMode,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicyInput.mode(
            Objects.requireNonNullElseGet(executionMode, ExecutionModeInput::defaults)),
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentInput::empty),
        steps(mutations, inspections));
  }

  static List<String> inspectionIds(WorkbookResult.Success success) {
    return success.inspections().stream().map(InspectionResult::stepId).toList();
  }

  static List<String> assertionIds(WorkbookResult.Success success) {
    return success.assertions().stream().map(AssertionResult::stepId).toList();
  }

  static <T extends InspectionResult> T inspection(
      WorkbookResult.Success success, String stepId, Class<T> type) {
    return type.cast(
        success.inspections().stream()
            .filter(result -> result.stepId().equals(stepId))
            .findFirst()
            .orElseThrow(
                () -> new AssertionError("Missing inspection result for stepId " + stepId)));
  }

  static Path createWorkbookFile(String prefix) throws IOException {
    Path workbookPath = Files.createTempFile(prefix, ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").cells().setCell("A1", ExcelCellValue.text("Header"));
      ExecutionContextFixtureSupport.saveWorkbook(workbook, workbookPath);
    }
    return workbookPath;
  }

  private static String stepIdFor(int stepIndex, MutationAction action) {
    String normalizedType = action.actionType().toLowerCase(Locale.ROOT).replace('_', '-');
    return "step-" + String.format(Locale.ROOT, "%02d", stepIndex + 1) + "-" + normalizedType;
  }

  private static Selector namedRangeSelector(String name, NamedRangeScope scope) {
    return switch (scope) {
      case NamedRangeScope.Workbook _ -> new NamedRangeSelector.WorkbookScope(name);
      case NamedRangeScope.Sheet sheet ->
          new NamedRangeSelector.SheetScope(name, sheet.sheetName());
    };
  }
}
