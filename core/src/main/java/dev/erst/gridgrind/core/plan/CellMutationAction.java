package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.drawing.CommentInput;
import dev.erst.gridgrind.core.model.formula.ArrayFormulaInput;
import dev.erst.gridgrind.core.model.style.CellStylePatchInput;
import dev.erst.gridgrind.core.model.value.CellGridInput;
import dev.erst.gridgrind.core.model.value.CellInput;
import dev.erst.gridgrind.core.model.value.CellRowInput;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import java.util.Objects;

/** Mutation family for cells, ranges, hyperlinks, comments, and cell styles. */
public sealed interface CellMutationAction extends MutationAction {
  /** Sets a single cell to the given value. */
  @Variant(
      id = "SET_CELL",
      summary = "Write one typed value to a single cell.",
      targetSelectors = {CellSelector.ByAddress.class, TableCellSelector.ByColumnName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.CELL_MUTATION)
  record SetCell(CellInput value) implements CellMutationAction {
    public SetCell {
      Objects.requireNonNull(value, "value must not be null");
    }
  }

  /** Sets a rectangular region of cells from a row-major grid of values. */
  @Variant(
      id = "SET_RANGE",
      summary = "Write a rectangular grid of typed or compact homogeneous values.",
      targetSelectors = {RangeSelector.ByRange.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.CELL_MUTATION)
  record SetRange(CellGridInput rows) implements CellMutationAction {
    public SetRange {
      Objects.requireNonNull(rows, "rows must not be null");
    }
  }

  /** Clears all cell values and styles within the specified range. */
  @Variant(
      id = "CLEAR_RANGE",
      summary = "Clear all cell values and styles within the selected range.",
      targetSelectors = {RangeSelector.ByRange.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.CELL_MUTATION)
  record ClearRange() implements CellMutationAction {
    public ClearRange {}
  }

  /** Creates or replaces one dedicated array-formula group over the addressed range. */
  @Variant(
      id = "SET_ARRAY_FORMULA",
      summary = "Author one contiguous single-cell or multi-cell array-formula group.",
      targetSelectors = {RangeSelector.ByRange.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.CELL_MUTATION)
  record SetArrayFormula(ArrayFormulaInput formula) implements CellMutationAction {
    public SetArrayFormula {
      Objects.requireNonNull(formula, "formula must not be null");
    }
  }

  /** Removes the array-formula group containing the addressed cell and clears the group. */
  @Variant(
      id = "CLEAR_ARRAY_FORMULA",
      summary = "Remove the stored array-formula group targeted by any member cell.",
      targetSelectors = {CellSelector.ByAddress.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.CELL_MUTATION)
  record ClearArrayFormula() implements CellMutationAction {
    public ClearArrayFormula {}
  }

  /** Replaces the hyperlink attached to a single cell. */
  @Variant(
      id = "SET_HYPERLINK",
      summary = "Replace the hyperlink attached to a single cell.",
      targetSelectors = {CellSelector.ByAddress.class, TableCellSelector.ByColumnName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.CELL_MUTATION)
  record SetHyperlink(HyperlinkTarget target) implements CellMutationAction {
    public SetHyperlink {
      Objects.requireNonNull(target, "target must not be null");
    }
  }

  /** Removes any hyperlink attached to a single existing cell. */
  @Variant(
      id = "CLEAR_HYPERLINK",
      summary = "Remove any hyperlink attached to a single existing cell.",
      targetSelectors = {CellSelector.ByAddress.class, TableCellSelector.ByColumnName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.CELL_MUTATION)
  record ClearHyperlink() implements CellMutationAction {
    public ClearHyperlink {}
  }

  /** Replaces the plain-text comment attached to a single cell. */
  @Variant(
      id = "SET_COMMENT",
      summary = "Replace the plain-text comment attached to a single cell.",
      targetSelectors = {CellSelector.ByAddress.class, TableCellSelector.ByColumnName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.CELL_MUTATION)
  record SetComment(CommentInput comment) implements CellMutationAction {
    public SetComment {
      Objects.requireNonNull(comment, "comment must not be null");
    }
  }

  /** Removes any comment attached to a single existing cell. */
  @Variant(
      id = "CLEAR_COMMENT",
      summary = "Remove any comment attached to a single existing cell.",
      targetSelectors = {CellSelector.ByAddress.class, TableCellSelector.ByColumnName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.CELL_MUTATION)
  record ClearComment() implements CellMutationAction {
    public ClearComment {}
  }

  /** Applies a style patch to every cell in the specified range. */
  @Variant(
      id = "APPLY_STYLE",
      summary = "Apply a style patch to every cell in the selected range.",
      targetSelectors = {RangeSelector.ByRange.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.CELL_MUTATION)
  record ApplyStyle(CellStylePatchInput style) implements CellMutationAction {
    public ApplyStyle {
      Objects.requireNonNull(style, "style must not be null");
    }
  }

  /** Appends a new row of values after the last occupied row on the sheet. */
  @Variant(
      id = "APPEND_ROW",
      summary =
          "Append a new row of typed or compact homogeneous values after the last occupied row.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.CELL_MUTATION)
  record AppendRow(CellRowInput values) implements CellMutationAction {
    public AppendRow {
      Objects.requireNonNull(values, "values must not be null");
    }
  }
}
