package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.dto.ColorInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingThresholdInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.ProtocolConstraintValues;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import java.util.ArrayList;
import java.util.List;

/** Maps validator-owned scalar restrictions onto the protocol catalog's constraint vocabulary. */
final class CatalogFieldConstraints {
  private static final List<ConstraintRule> RULES = rules();

  private CatalogFieldConstraints() {}

  static List<FieldConstraint> forComponent(Class<?> owner, String fieldName) {
    return RULES.stream()
        .filter(rule -> rule.owner() == owner && rule.fieldName().equals(fieldName))
        .findFirst()
        .map(ConstraintRule::constraints)
        .orElseGet(List::of);
  }

  private static List<ConstraintRule> rules() {
    List<ConstraintRule> rules = new ArrayList<>();
    addStepIdRules(rules);
    rules.add(rule(ColorInput.Rgb.class, "rgb", rgbConstraints()));
    addDefinedNameRules(rules);
    addDataBarWidthRules(rules);
    addFormulaRules(rules);
    rules.add(
        rule(WorkbookPlan.WorkbookSource.ExistingFile.class, "path", workbookPathConstraints()));
    rules.add(
        rule(WorkbookPlan.WorkbookPersistence.SaveAs.class, "path", workbookPathConstraints()));
    return List.copyOf(rules);
  }

  private static void addStepIdRules(List<ConstraintRule> rules) {
    List<FieldConstraint> constraints =
        List.of(
            new FieldConstraint.NonBlank(),
            new FieldConstraint.StringPattern(ProtocolConstraintValues.STEP_ID_PATTERN));
    rules.add(rule(MutationStep.class, "stepId", constraints));
    rules.add(rule(AssertionStep.class, "stepId", constraints));
    rules.add(rule(InspectionStep.class, "stepId", constraints));
  }

  private static void addDefinedNameRules(List<ConstraintRule> rules) {
    List<FieldConstraint> constraints =
        List.of(
            new FieldConstraint.LengthRange(
                1, ProtocolConstraintValues.DEFINED_NAME_MAX_CODE_POINTS),
            new FieldConstraint.StringPattern(ProtocolConstraintValues.DEFINED_NAME_PATTERN));
    rules.add(rule(StructuredMutationAction.SetNamedRange.class, "name", constraints));
    rules.add(rule(TableInput.class, "name", constraints));
    rules.add(rule(NamedRangeTarget.class, "name", constraints));
  }

  private static void addDataBarWidthRules(List<ConstraintRule> rules) {
    List<FieldConstraint> constraints =
        List.of(
            new FieldConstraint.Integral(),
            new FieldConstraint.NumberRange(
                ProtocolConstraintValues.DATA_BAR_WIDTH_MIN,
                ProtocolConstraintValues.DATA_BAR_WIDTH_MAX));
    rules.add(rule(ConditionalFormattingRuleInput.DataBarRule.class, "widthMin", constraints));
    rules.add(rule(ConditionalFormattingRuleInput.DataBarRule.class, "widthMax", constraints));
  }

  private static void addFormulaRules(List<ConstraintRule> rules) {
    List<FieldConstraint> constraints = List.of(new FieldConstraint.NonBlank());
    rules.add(rule(ConditionalFormattingRuleInput.FormulaRule.class, "formula", constraints));
    rules.add(rule(ConditionalFormattingThresholdInput.Formula.class, "formula", constraints));
    rules.add(rule(DataValidationRuleInput.FormulaList.class, "formula", constraints));
    rules.add(rule(DataValidationRuleInput.CustomFormula.class, "formula", constraints));
    addComparisonFormulaRules(rules, DataValidationRuleInput.WholeNumber.class, constraints);
    addComparisonFormulaRules(rules, DataValidationRuleInput.DecimalNumber.class, constraints);
    addComparisonFormulaRules(rules, DataValidationRuleInput.DateRule.class, constraints);
    addComparisonFormulaRules(rules, DataValidationRuleInput.TimeRule.class, constraints);
    addComparisonFormulaRules(rules, DataValidationRuleInput.TextLength.class, constraints);
  }

  private static void addComparisonFormulaRules(
      List<ConstraintRule> rules, Class<?> owner, List<FieldConstraint> constraints) {
    rules.add(rule(owner, "formula1", constraints));
    rules.add(rule(owner, "formula2", constraints));
  }

  private static List<FieldConstraint> rgbConstraints() {
    return List.of(new FieldConstraint.StringPattern(ProtocolConstraintValues.RGB_HEX_PATTERN));
  }

  private static List<FieldConstraint> workbookPathConstraints() {
    return List.of(new FieldConstraint.PathSuffix(ProtocolConstraintValues.WORKBOOK_PATH_SUFFIX));
  }

  private static ConstraintRule rule(
      Class<?> owner, String fieldName, List<FieldConstraint> constraints) {
    return new ConstraintRule(owner, fieldName, constraints);
  }

  private record ConstraintRule(
      Class<?> owner, String fieldName, List<FieldConstraint> constraints) {
    private ConstraintRule {
      constraints = List.copyOf(constraints);
    }
  }
}
