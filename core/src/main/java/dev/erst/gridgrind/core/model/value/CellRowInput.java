package dev.erst.gridgrind.core.model.value;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.source.TextSourceInput;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Row-oriented authored cell input with compact homogeneous encodings for append operations. */
@Union(discriminator = "type", catalog = true, displayOrder = 17)
public sealed interface CellRowInput
    permits dev.erst.gridgrind.core.model.value.CellRowInput.Typed,
        dev.erst.gridgrind.core.model.value.CellRowInput.TextValues,
        dev.erst.gridgrind.core.model.value.CellRowInput.NumberValues,
        dev.erst.gridgrind.core.model.value.CellRowInput.BooleanValues,
        dev.erst.gridgrind.core.model.value.CellRowInput.ErrorValues,
        dev.erst.gridgrind.core.model.value.CellRowInput.DateValues,
        dev.erst.gridgrind.core.model.value.CellRowInput.DateTimeValues,
        dev.erst.gridgrind.core.model.value.CellRowInput.FormulaValues {

  /** Returns this authored row in canonical per-cell form. */
  List<CellInput> toCellInputs();

  @Variant(id = "TYPED", summary = "Append one row expressed as a per-cell typed value list.")
  record Typed(List<CellInput> cells) implements CellRowInput {
    public Typed {
      cells = copyTyped(cells, "cells");
    }

    @Override
    public List<CellInput> toCellInputs() {
      return cells;
    }
  }

  @Variant(id = "TEXT", summary = "Append one row using a compact homogeneous string list.")
  record TextValues(List<String> cells) implements CellRowInput {
    public TextValues {
      cells = copyScalars(cells, "cells");
    }

    @Override
    public List<CellInput> toCellInputs() {
      return cells.stream()
          .map(TextSourceInput::inline)
          .map(CellInput.Text::new)
          .map(CellInput.class::cast)
          .toList();
    }
  }

  @Variant(id = "NUMBER", summary = "Append one row using a compact homogeneous numeric list.")
  record NumberValues(List<Double> cells) implements CellRowInput {
    public NumberValues {
      cells = copyScalars(cells, "cells");
      cells.forEach(value -> CellInput.Validation.requireFinite(value, "cells element"));
    }

    @Override
    public List<CellInput> toCellInputs() {
      return cells.stream().map(CellInput.NumberValue::new).map(CellInput.class::cast).toList();
    }
  }

  @Variant(id = "BOOLEAN", summary = "Append one row using a compact homogeneous boolean list.")
  record BooleanValues(List<Boolean> cells) implements CellRowInput {
    public BooleanValues {
      cells = copyScalars(cells, "cells");
    }

    @Override
    public List<CellInput> toCellInputs() {
      return cells.stream().map(CellInput.BooleanValue::new).map(CellInput.class::cast).toList();
    }
  }

  @Variant(
      id = "ERROR",
      summary = "Append one row using a compact homogeneous stored Excel error literal list.")
  record ErrorValues(List<String> cells) implements CellRowInput {
    public ErrorValues {
      cells = copyScalars(cells, "cells");
      cells.forEach(
          value -> CellErrorLiteralValidation.requireStoredErrorLiteral(value, "cells element"));
    }

    @Override
    public List<CellInput> toCellInputs() {
      return cells.stream().map(CellInput.ErrorValue::new).map(CellInput.class::cast).toList();
    }
  }

  @Variant(id = "DATE", summary = "Append one row using a compact homogeneous date list.")
  record DateValues(List<LocalDate> cells) implements CellRowInput {
    public DateValues {
      cells = copyScalars(cells, "cells");
    }

    @Override
    public List<CellInput> toCellInputs() {
      return cells.stream().map(CellInput.Date::new).map(CellInput.class::cast).toList();
    }
  }

  @Variant(id = "DATE_TIME", summary = "Append one row using a compact homogeneous date-time list.")
  record DateTimeValues(List<LocalDateTime> cells) implements CellRowInput {
    public DateTimeValues {
      cells = copyScalars(cells, "cells");
    }

    @Override
    public List<CellInput> toCellInputs() {
      return cells.stream().map(CellInput.DateTime::new).map(CellInput.class::cast).toList();
    }
  }

  @Variant(
      id = "FORMULA",
      summary = "Append one row using a compact homogeneous formula text list.")
  record FormulaValues(List<String> cells) implements CellRowInput {
    public FormulaValues {
      cells =
          copyScalars(cells, "cells").stream()
              .map(value -> CellInput.Validation.normalizeInlineFormula(value, "cells element"))
              .toList();
    }

    @Override
    public List<CellInput> toCellInputs() {
      return cells.stream()
          .map(TextSourceInput::inline)
          .map(CellInput.Formula::new)
          .map(CellInput.class::cast)
          .toList();
    }
  }

  private static List<CellInput> copyTyped(List<CellInput> cells, String fieldName) {
    Objects.requireNonNull(cells, fieldName + " must not be null");
    List<CellInput> copy = new ArrayList<>(cells.size());
    if (cells.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    for (CellInput value : cells) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }

  private static <T> List<T> copyScalars(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    if (values.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }
}
