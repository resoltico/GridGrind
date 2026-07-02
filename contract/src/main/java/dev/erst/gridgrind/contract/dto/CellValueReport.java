package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Effective value carried by one formula-cell readback. */
@JsonTypeInfo(
    use = JsonTypeInfo.Id.NAME,
    include = JsonTypeInfo.As.EXISTING_PROPERTY,
    property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CellValueReport.BlankValue.class, name = "BLANK"),
  @JsonSubTypes.Type(value = CellValueReport.TextValue.class, name = "TEXT"),
  @JsonSubTypes.Type(value = CellValueReport.NumberValue.class, name = "NUMBER"),
  @JsonSubTypes.Type(value = CellValueReport.BooleanValue.class, name = "BOOLEAN"),
  @JsonSubTypes.Type(value = CellValueReport.ErrorValue.class, name = "ERROR")
})
public sealed interface CellValueReport {
  /** Canonical published value type. */
  String type();

  /** Blank effective value. */
  record BlankValue() implements CellValueReport {
    @Override
    @JsonProperty
    public String type() {
      return "BLANK";
    }
  }

  /** Text effective value. */
  record TextValue(
      String textValue,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<List<RichTextRunReport>> runs)
      implements CellValueReport {
    public TextValue {
      Objects.requireNonNull(textValue, "textValue must not be null");
      runs = copyRichTextRuns(runs, "runs");
      if (runs.isPresent() && !textValue.equals(concatenateRuns(runs.orElseThrow()))) {
        throw new IllegalArgumentException("runs text must concatenate to the textValue");
      }
    }

    @Override
    @JsonProperty
    public String type() {
      return "TEXT";
    }
  }

  /** Numeric effective value. */
  record NumberValue(
      Double numberValue,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CellTemporalReport> temporal)
      implements CellValueReport {
    public NumberValue {
      Objects.requireNonNull(numberValue, "numberValue must not be null");
      temporal = Objects.requireNonNullElseGet(temporal, Optional::empty);
    }

    @Override
    @JsonProperty
    public String type() {
      return "NUMBER";
    }
  }

  /** Boolean effective value. */
  record BooleanValue(Boolean booleanValue) implements CellValueReport {
    public BooleanValue {
      Objects.requireNonNull(booleanValue, "booleanValue must not be null");
    }

    @Override
    @JsonProperty
    public String type() {
      return "BOOLEAN";
    }
  }

  /** Error effective value. */
  record ErrorValue(String errorValue) implements CellValueReport {
    public ErrorValue {
      Objects.requireNonNull(errorValue, "errorValue must not be null");
      if (errorValue.isBlank()) {
        throw new IllegalArgumentException("errorValue must not be blank");
      }
    }

    @Override
    @JsonProperty
    public String type() {
      return "ERROR";
    }
  }

  private static Optional<List<RichTextRunReport>> copyRichTextRuns(
      Optional<List<RichTextRunReport>> values, String fieldName) {
    Optional<List<RichTextRunReport>> normalized =
        Objects.requireNonNullElseGet(values, Optional::empty);
    if (normalized.isEmpty()) {
      return Optional.empty();
    }
    List<RichTextRunReport> copy = copyValues(normalized.orElseThrow(), fieldName);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    return Optional.of(copy);
  }

  private static String concatenateRuns(List<RichTextRunReport> runs) {
    StringBuilder builder = new StringBuilder();
    for (RichTextRunReport run : runs) {
      builder.append(run.text());
    }
    return builder.toString();
  }

  private static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }
}
