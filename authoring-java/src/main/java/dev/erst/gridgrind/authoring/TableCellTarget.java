package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import java.util.Objects;

/** Table-cell fluent target compiled to one canonical table-cell selector. */
public final class TableCellTarget {
  private final TableCellSelector.ByColumnName selector;

  TableCellTarget(TableCellSelector.ByColumnName selector) {
    this.selector = Objects.requireNonNull(selector, "selector must not be null");
  }

  TableCellSelector.ByColumnName selector() {
    return selector;
  }

  /** Returns one set-cell mutation step against the logical table cell. */
  public PlannedMutation set(Values.CellValue value) {
    return new PlannedMutation(selector, new CellMutationAction.SetCell(Values.toCellInput(value)));
  }

  /** Returns one set-hyperlink mutation step against the logical table cell. */
  public PlannedMutation setHyperlink(Links.Target hyperlink) {
    return new PlannedMutation(
        selector, new CellMutationAction.SetHyperlink(Links.toHyperlinkTarget(hyperlink)));
  }

  /** Returns one clear-hyperlink mutation step against the logical table cell. */
  public PlannedMutation clearHyperlink() {
    return new PlannedMutation(selector, new CellMutationAction.ClearHyperlink());
  }

  /** Returns one set-comment mutation step against the logical table cell. */
  public PlannedMutation setComment(Values.Comment comment) {
    return new PlannedMutation(
        selector, new CellMutationAction.SetComment(Values.toCommentInput(comment)));
  }

  /** Returns one clear-comment mutation step against the logical table cell. */
  public PlannedMutation clearComment() {
    return new PlannedMutation(selector, new CellMutationAction.ClearComment());
  }

  /** Returns one exact-cell inspection step against the logical table cell. */
  public PlannedInspection read() {
    return new PlannedInspection(selector, Queries.cells());
  }

  /** Returns one hyperlink inspection step against the logical table cell. */
  public PlannedInspection hyperlinks() {
    return new PlannedInspection(selector, Queries.hyperlinks());
  }

  /** Returns one comment inspection step against the logical table cell. */
  public PlannedInspection comments() {
    return new PlannedInspection(selector, Queries.comments());
  }

  /** Returns one effective-value assertion step against the logical table cell. */
  public PlannedAssertion valueEquals(Values.ExpectedValue expectedValue) {
    return new PlannedAssertion(selector, Checks.cellValue(expectedValue));
  }

  /** Returns one rendered display-value assertion step against the logical table cell. */
  public PlannedAssertion displayValueEquals(String displayValue) {
    return new PlannedAssertion(selector, Checks.displayValue(displayValue));
  }

  /** Returns one formula-text assertion step against the logical table cell. */
  public PlannedAssertion formulaEquals(String formula) {
    return new PlannedAssertion(selector, Checks.formulaText(formula));
  }
}
