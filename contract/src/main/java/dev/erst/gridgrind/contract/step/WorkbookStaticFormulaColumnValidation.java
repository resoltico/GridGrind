package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.ArrayList;
import java.util.List;

/** Detects ordered formula-authoring and column-edit conflicts that are provable from a request. */
final class WorkbookStaticFormulaColumnValidation {
  static final String PRECONDITION_ID = "NO_COLUMN_EDITS_AFTER_FORMULA_AUTHORING";
  static final String VIOLATION_MESSAGE =
      "Column insert, delete, and shift operations must appear before formula authoring because"
          + " formula-bearing workbook column edits are unsupported.";

  private WorkbookStaticFormulaColumnValidation() {}

  static List<WorkbookStaticViolation> validate(WorkbookStaticRequest request) {
    return validate(request.steps(), initialFormulaPresence(request));
  }

  static List<WorkbookStaticViolation> validate(
      List<WorkbookStaticStep> steps, boolean initialFormulaPresence) {
    return validate(
        steps, initialFormulaPresence ? FormulaPresence.PRESENT : FormulaPresence.ABSENT);
  }

  private static List<WorkbookStaticViolation> validate(
      List<WorkbookStaticStep> steps, FormulaPresence formulaPresence) {
    List<WorkbookStaticViolation> violations = new ArrayList<>();
    FormulaPresence currentFormulaPresence = formulaPresence;
    for (WorkbookStaticStep staticStep : steps) {
      if (staticStep.value().isEmpty()) {
        currentFormulaPresence = FormulaPresence.UNKNOWN;
        continue;
      }
      if (!(staticStep.value().orElseThrow() instanceof MutationStep mutationStep)) {
        continue;
      }
      MutationAction action = mutationStep.action();
      if (isColumnEdit(action) && currentFormulaPresence == FormulaPresence.PRESENT) {
        violations.add(
            new WorkbookStaticViolation(
                "steps[" + staticStep.index() + "].action", VIOLATION_MESSAGE));
      }
      currentFormulaPresence = transition(currentFormulaPresence, action);
    }
    return List.copyOf(violations);
  }

  private static FormulaPresence initialFormulaPresence(WorkbookStaticRequest request) {
    return request
        .source()
        .map(
            source ->
                source instanceof WorkbookPlan.WorkbookSource.New
                    ? FormulaPresence.ABSENT
                    : FormulaPresence.UNKNOWN)
        .orElse(FormulaPresence.UNKNOWN);
  }

  private static FormulaPresence transition(FormulaPresence current, MutationAction action) {
    if (authorsFormula(action)) {
      return FormulaPresence.PRESENT;
    }
    if (mayReplaceFormula(action) && current == FormulaPresence.PRESENT) {
      return FormulaPresence.UNKNOWN;
    }
    return current;
  }

  private static boolean isColumnEdit(MutationAction action) {
    return action instanceof WorkbookMutationAction.InsertColumns
        || action instanceof WorkbookMutationAction.DeleteColumns
        || action instanceof WorkbookMutationAction.ShiftColumns;
  }

  private static boolean authorsFormula(MutationAction action) {
    return switch (action) {
      case CellMutationAction.SetCell setCell -> isFormula(setCell.value());
      case CellMutationAction.SetRange setRange ->
          setRange.rows().toCellInputRows().stream()
              .flatMap(List::stream)
              .anyMatch(WorkbookStaticFormulaColumnValidation::isFormula);
      case CellMutationAction.AppendRow appendRow ->
          appendRow.values().toCellInputs().stream()
              .anyMatch(WorkbookStaticFormulaColumnValidation::isFormula);
      case CellMutationAction.SetArrayFormula _ -> true;
      default -> false;
    };
  }

  private static boolean mayReplaceFormula(MutationAction action) {
    return action instanceof CellMutationAction.SetCell
        || action instanceof CellMutationAction.SetRange
        || action instanceof CellMutationAction.ClearRange
        || action instanceof CellMutationAction.ClearArrayFormula;
  }

  private static boolean isFormula(CellInput cell) {
    return cell instanceof CellInput.Formula || cell instanceof CellInput.RawFormula;
  }

  /** The knownness of formula-bearing workbook state during ordered static validation. */
  private enum FormulaPresence {
    /** No preceding action has established a formula-bearing workbook state. */
    ABSENT,
    /** At least one preceding action has established a formula-bearing workbook state. */
    PRESENT,
    /** A preceding unbound or replacing action prevents a static conclusion. */
    UNKNOWN
  }
}
