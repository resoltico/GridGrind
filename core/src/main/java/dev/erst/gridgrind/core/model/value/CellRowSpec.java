package dev.erst.gridgrind.core.model.value;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Row-oriented authored cell input with compact homogeneous encodings for append operations. */
@Union(discriminator = "type", catalog = true, displayOrder = 17)
public sealed interface CellRowSpec
    permits dev.erst.gridgrind.core.model.value.CellRowSpec.Typed,
        dev.erst.gridgrind.core.model.value.CellRowSpec.TextValues,
        dev.erst.gridgrind.core.model.value.CellRowSpec.NumberValues,
        dev.erst.gridgrind.core.model.value.CellRowSpec.BooleanValues,
        dev.erst.gridgrind.core.model.value.CellRowSpec.ErrorValues,
        dev.erst.gridgrind.core.model.value.CellRowSpec.DateValues,
        dev.erst.gridgrind.core.model.value.CellRowSpec.DateTimeValues,
        dev.erst.gridgrind.core.model.value.CellRowSpec.FormulaValues {

  /** Returns this authored row in canonical per-cell form. */
  List<CellValueSpec> toCellInputs();

  @Variant(id = "TYPED", summary = "Append one row expressed as a per-cell typed value list.")
  record Typed(List<CellValueSpec> cells) implements CellRowSpec {
    public Typed {
      cells = copyTyped(cells, "cells");
    }

    @Override
    public List<CellValueSpec> toCellInputs() {
      return cells;
    }
  }

  @Variant(id = "TEXT", summary = "Append one row using a compact homogeneous string list.")
  record TextValues(List<String> cells) implements CellRowSpec {
    public TextValues {
      cells = copyScalars(cells, "cells");
    }

    @Override
    public List<CellValueSpec> toCellInputs() {
      return cells.stream()
          .map(TextSourceSpec::inline)
          .map(CellValueSpec.Text::new)
          .map(CellValueSpec.class::cast)
          .toList();
    }
  }

  @Variant(id = "NUMBER", summary = "Append one row using a compact homogeneous numeric list.")
  record NumberValues(List<Double> cells) implements CellRowSpec {
    public NumberValues {
      cells = copyScalars(cells, "cells");
      cells.forEach(value -> CellValueSpec.Validation.requireFinite(value, "cells element"));
    }

    @Override
    public List<CellValueSpec> toCellInputs() {
      return cells.stream()
          .map(CellValueSpec.NumberValue::new)
          .map(CellValueSpec.class::cast)
          .toList();
    }
  }

  @Variant(id = "BOOLEAN", summary = "Append one row using a compact homogeneous boolean list.")
  record BooleanValues(List<Boolean> cells) implements CellRowSpec {
    public BooleanValues {
      cells = copyScalars(cells, "cells");
    }

    @Override
    public List<CellValueSpec> toCellInputs() {
      return cells.stream()
          .map(CellValueSpec.BooleanValue::new)
          .map(CellValueSpec.class::cast)
          .toList();
    }
  }

  @Variant(
      id = "ERROR",
      summary = "Append one row using a compact homogeneous stored Excel error literal list.")
  record ErrorValues(List<String> cells) implements CellRowSpec {
    public ErrorValues {
      cells = copyScalars(cells, "cells");
      cells.forEach(
          value -> CellErrorLiteralValidation.requireStoredErrorLiteral(value, "cells element"));
    }

    @Override
    public List<CellValueSpec> toCellInputs() {
      return cells.stream()
          .map(CellValueSpec.ErrorValue::new)
          .map(CellValueSpec.class::cast)
          .toList();
    }
  }

  @Variant(id = "DATE", summary = "Append one row using a compact homogeneous date list.")
  record DateValues(List<LocalDate> cells) implements CellRowSpec {
    public DateValues {
      cells = copyScalars(cells, "cells");
    }

    @Override
    public List<CellValueSpec> toCellInputs() {
      return cells.stream().map(CellValueSpec.Date::new).map(CellValueSpec.class::cast).toList();
    }
  }

  @Variant(id = "DATE_TIME", summary = "Append one row using a compact homogeneous date-time list.")
  record DateTimeValues(List<LocalDateTime> cells) implements CellRowSpec {
    public DateTimeValues {
      cells = copyScalars(cells, "cells");
    }

    @Override
    public List<CellValueSpec> toCellInputs() {
      return cells.stream()
          .map(CellValueSpec.DateTime::new)
          .map(CellValueSpec.class::cast)
          .toList();
    }
  }

  @Variant(
      id = "FORMULA",
      summary = "Append one row using a compact homogeneous formula text list.")
  record FormulaValues(List<String> cells) implements CellRowSpec {
    public FormulaValues {
      cells =
          copyScalars(cells, "cells").stream()
              .map(value -> CellValueSpec.Validation.normalizeInlineFormula(value, "cells element"))
              .toList();
    }

    @Override
    public List<CellValueSpec> toCellInputs() {
      return cells.stream()
          .map(TextSourceSpec::inline)
          .map(CellValueSpec.Formula::new)
          .map(CellValueSpec.class::cast)
          .toList();
    }
  }

  private static List<CellValueSpec> copyTyped(List<CellValueSpec> cells, String fieldName) {
    Objects.requireNonNull(cells, fieldName + " must not be null");
    List<CellValueSpec> copy = new ArrayList<>(cells.size());
    if (cells.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    for (CellValueSpec value : cells) {
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
