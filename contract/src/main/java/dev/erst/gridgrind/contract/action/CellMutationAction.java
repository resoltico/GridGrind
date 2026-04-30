package dev.erst.gridgrind.contract.action;

import dev.erst.gridgrind.contract.dto.ArrayFormulaInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellStyleInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Mutation family for cells, ranges, hyperlinks, comments, and cell styles. */
public sealed interface CellMutationAction extends MutationAction {
  /** Sets a single cell to the given value. */
  record SetCell(CellInput value) implements CellMutationAction {
    public SetCell {
      Objects.requireNonNull(value, "value must not be null");
    }
  }

  /** Sets a rectangular region of cells from a row-major grid of values. */
  record SetRange(List<List<CellInput>> rows) implements CellMutationAction {
    public SetRange {
      rows = MutationAction.Validation.copyRows(rows);
      MutationAction.Validation.requireRectangularRows(rows);
      rows = MutationAction.Validation.freezeRows(rows);
    }
  }

  /** Clears all cell values and styles within the specified range. */
  record ClearRange() implements CellMutationAction {
    public ClearRange {}
  }

  /** Creates or replaces one dedicated array-formula group over the addressed range. */
  record SetArrayFormula(ArrayFormulaInput formula) implements CellMutationAction {
    public SetArrayFormula {
      Objects.requireNonNull(formula, "formula must not be null");
    }
  }

  /** Removes the array-formula group containing the addressed cell and clears the group. */
  record ClearArrayFormula() implements CellMutationAction {
    public ClearArrayFormula {}
  }

  /** Replaces the hyperlink attached to a single cell. */
  record SetHyperlink(HyperlinkTarget target) implements CellMutationAction {
    public SetHyperlink {
      Objects.requireNonNull(target, "target must not be null");
    }
  }

  /** Removes any hyperlink attached to a single existing cell. */
  record ClearHyperlink() implements CellMutationAction {
    public ClearHyperlink {}
  }

  /** Replaces the plain-text comment attached to a single cell. */
  record SetComment(CommentInput comment) implements CellMutationAction {
    public SetComment {
      Objects.requireNonNull(comment, "comment must not be null");
    }
  }

  /** Removes any comment attached to a single existing cell. */
  record ClearComment() implements CellMutationAction {
    public ClearComment {}
  }

  /** Applies a style patch to every cell in the specified range. */
  record ApplyStyle(CellStyleInput style) implements CellMutationAction {
    public ApplyStyle {
      Objects.requireNonNull(style, "style must not be null");
    }
  }

  /** Appends a new row of values after the last occupied row on the sheet. */
  record AppendRow(List<CellInput> values) implements CellMutationAction {
    public AppendRow {
      Objects.requireNonNull(values, "values must not be null");
      values = new ArrayList<>(values);
      if (values.isEmpty()) {
        throw new IllegalArgumentException("values must not be empty");
      }
      for (CellInput item : values) {
        Objects.requireNonNull(item, "values must not contain nulls");
      }
      values = List.copyOf(values);
    }
  }
}
