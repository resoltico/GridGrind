package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.selector.CellSelector;
import java.util.Objects;

/** Exact A1 cell fluent target. */
public final class CellTarget {
  private final CellSelector.ByAddress selector;

  CellTarget(CellSelector.ByAddress selector) {
    this.selector = Objects.requireNonNull(selector, "selector must not be null");
  }

  CellSelector.ByAddress selector() {
    return selector;
  }

  /** Returns one set-cell mutation step. */
  public PlannedMutation set(Values.CellValue value) {
    return new PlannedMutation(selector, new CellMutationAction.SetCell(Values.toCellInput(value)));
  }

  /** Returns one set-hyperlink mutation step. */
  public PlannedMutation setHyperlink(Links.Target hyperlink) {
    return new PlannedMutation(
        selector, new CellMutationAction.SetHyperlink(Links.toHyperlinkTarget(hyperlink)));
  }

  /** Returns one clear-hyperlink mutation step. */
  public PlannedMutation clearHyperlink() {
    return new PlannedMutation(selector, new CellMutationAction.ClearHyperlink());
  }

  /** Returns one set-comment mutation step. */
  public PlannedMutation setComment(Values.Comment comment) {
    return new PlannedMutation(
        selector, new CellMutationAction.SetComment(Values.toCommentInput(comment)));
  }

  /** Returns one clear-comment mutation step. */
  public PlannedMutation clearComment() {
    return new PlannedMutation(selector, new CellMutationAction.ClearComment());
  }

  /** Returns one exact-cell inspection step. */
  public PlannedInspection read() {
    return new PlannedInspection(selector, Queries.cells());
  }

  /** Returns one hyperlink inspection step for this cell. */
  public PlannedInspection hyperlinks() {
    return new PlannedInspection(selector, Queries.hyperlinks());
  }

  /** Returns one comment inspection step for this cell. */
  public PlannedInspection comments() {
    return new PlannedInspection(selector, Queries.comments());
  }

  /** Returns one effective-value assertion step. */
  public PlannedAssertion valueEquals(Values.ExpectedValue expectedValue) {
    return new PlannedAssertion(selector, Checks.cellValue(expectedValue));
  }

  /** Returns one rendered display-value assertion step. */
  public PlannedAssertion displayValueEquals(String displayValue) {
    return new PlannedAssertion(selector, Checks.displayValue(displayValue));
  }

  /** Returns one formula-text assertion step. */
  public PlannedAssertion formulaEquals(String formula) {
    return new PlannedAssertion(selector, Checks.formulaText(formula));
  }
}
