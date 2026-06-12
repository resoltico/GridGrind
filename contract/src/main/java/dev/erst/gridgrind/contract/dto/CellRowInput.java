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

  record Typed(List<CellInput> values) implements CellRowInput {
    public Typed {
      values = copyTyped(values);
    }

    @Override
    public List<CellInput> toCellInputs() {
      return values;
    }
  }

  record TextValues(List<String> values) implements CellRowInput {
    public TextValues {
      values = copyScalars(values, "values");
    }

    @Override
    public List<CellInput> toCellInputs() {
      return values.stream()
          .map(TextSourceInput::inline)
          .map(CellInput.Text::new)
          .map(CellInput.class::cast)
          .toList();
    }
  }

  record NumberValues(List<Double> values) implements CellRowInput {
    public NumberValues {
      values = copyScalars(values, "values");
      values.forEach(value -> CellInput.Validation.requireFinite(value, "values element"));
    }

    @Override
    public List<CellInput> toCellInputs() {
      return values.stream().map(CellInput.NumberValue::new).map(CellInput.class::cast).toList();
    }
  }

  record BooleanValues(List<Boolean> values) implements CellRowInput {
    public BooleanValues {
      values = copyScalars(values, "values");
    }

    @Override
    public List<CellInput> toCellInputs() {
      return values.stream().map(CellInput.BooleanValue::new).map(CellInput.class::cast).toList();
    }
  }

  record ErrorValues(List<String> values) implements CellRowInput {
    public ErrorValues {
      values = copyScalars(values, "values");
      values.forEach(value -> CellInput.Validation.requireErrorLiteral(value, "values element"));
    }

    @Override
    public List<CellInput> toCellInputs() {
      return values.stream().map(CellInput.ErrorValue::new).map(CellInput.class::cast).toList();
    }
  }

  record DateValues(List<LocalDate> values) implements CellRowInput {
    public DateValues {
      values = copyScalars(values, "values");
    }

    @Override
    public List<CellInput> toCellInputs() {
      return values.stream().map(CellInput.Date::new).map(CellInput.class::cast).toList();
    }
  }

  record DateTimeValues(List<LocalDateTime> values) implements CellRowInput {
    public DateTimeValues {
      values = copyScalars(values, "values");
    }

    @Override
    public List<CellInput> toCellInputs() {
      return values.stream().map(CellInput.DateTime::new).map(CellInput.class::cast).toList();
    }
  }

  record FormulaValues(List<String> values) implements CellRowInput {
    public FormulaValues {
      values =
          copyScalars(values, "values").stream()
              .map(value -> CellInput.Validation.normalizeInlineFormula(value, "values element"))
              .toList();
    }

    @Override
    public List<CellInput> toCellInputs() {
      return values.stream()
          .map(TextSourceInput::inline)
          .map(CellInput.Formula::new)
          .map(CellInput.class::cast)
          .toList();
    }
  }

  private static List<CellInput> copyTyped(List<CellInput> values) {
    Objects.requireNonNull(values, "values must not be null");
    List<CellInput> copy = new ArrayList<>(values.size());
    if (values.isEmpty()) {
      throw new IllegalArgumentException("values must not be empty");
    }
    for (CellInput value : values) {
      copy.add(Objects.requireNonNull(value, "values must not contain nulls"));
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
