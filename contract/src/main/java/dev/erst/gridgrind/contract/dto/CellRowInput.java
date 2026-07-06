package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Row-oriented authored cell input with compact homogeneous encodings for append operations. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CellRowInput.Typed.class, name = "TYPED"),
  @JsonSubTypes.Type(value = CellRowInput.TextValues.class, name = "TEXT"),
  @JsonSubTypes.Type(value = CellRowInput.NumberValues.class, name = "NUMBER"),
  @JsonSubTypes.Type(value = CellRowInput.BooleanValues.class, name = "BOOLEAN"),
  @JsonSubTypes.Type(value = CellRowInput.ErrorValues.class, name = "ERROR"),
  @JsonSubTypes.Type(value = CellRowInput.DateValues.class, name = "DATE"),
  @JsonSubTypes.Type(value = CellRowInput.DateTimeValues.class, name = "DATE_TIME"),
  @JsonSubTypes.Type(value = CellRowInput.FormulaValues.class, name = "FORMULA")
})
public sealed interface CellRowInput
    permits CellRowInput.Typed,
        CellRowInput.TextValues,
        CellRowInput.NumberValues,
        CellRowInput.BooleanValues,
        CellRowInput.ErrorValues,
        CellRowInput.DateValues,
        CellRowInput.DateTimeValues,
        CellRowInput.FormulaValues {

  /** Returns this authored row in canonical per-cell form. */
  List<CellInput> toCellInputs();

  record Typed(List<CellInput> cells) implements CellRowInput {
    public Typed {
      cells = copyTyped(cells, "cells");
    }

    @Override
    public List<CellInput> toCellInputs() {
      return cells;
    }
  }

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

  record BooleanValues(List<Boolean> cells) implements CellRowInput {
    public BooleanValues {
      cells = copyScalars(cells, "cells");
    }

    @Override
    public List<CellInput> toCellInputs() {
      return cells.stream().map(CellInput.BooleanValue::new).map(CellInput.class::cast).toList();
    }
  }

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

  record DateValues(List<LocalDate> cells) implements CellRowInput {
    public DateValues {
      cells = copyScalars(cells, "cells");
    }

    @Override
    public List<CellInput> toCellInputs() {
      return cells.stream().map(CellInput.Date::new).map(CellInput.class::cast).toList();
    }
  }

  record DateTimeValues(List<LocalDateTime> cells) implements CellRowInput {
    public DateTimeValues {
      cells = copyScalars(cells, "cells");
    }

    @Override
    public List<CellInput> toCellInputs() {
      return cells.stream().map(CellInput.DateTime::new).map(CellInput.class::cast).toList();
    }
  }

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
