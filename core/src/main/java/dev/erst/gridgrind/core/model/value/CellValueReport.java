package dev.erst.gridgrind.core.model.value;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Effective value carried by one formula-cell readback. */
@Union(discriminator = "type", catalog = true, displayOrder = 2)
public sealed interface CellValueReport {
  /** Canonical published value type. */
  String type();

  /** Blank effective value. */
  @Variant(id = "BLANK", summary = "Effective formula evaluation is blank.")
  record BlankValue() implements CellValueReport {
    @Override
    public String type() {
      return "BLANK";
    }
  }

  /** Text effective value. */
  @Variant(
      id = "TEXT",
      summary = "Effective formula evaluation is text; RICH_TEXT_RUNS projects optional runs.")
  record TextValue(
      String textValue,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"RICH_TEXT_RUNS"})
          Optional<List<RichTextRunReport>> runs)
      implements CellValueReport {
    public TextValue {
      Objects.requireNonNull(textValue, "textValue must not be null");
      runs = copyRichTextRuns(runs, "runs");
      if (runs.isPresent() && !textValue.equals(concatenateRuns(runs.orElseThrow()))) {
        throw new IllegalArgumentException("runs text must concatenate to the textValue");
      }
    }

    @Override
    public String type() {
      return "TEXT";
    }
  }

  /** Numeric effective value. */
  @Variant(
      id = "NUMBER",
      summary =
          "Effective formula evaluation is numeric; TEMPORAL projects derived date, time, or date-time semantics when formatting marks the numeric value as date-like.")
  record NumberValue(
      Double numberValue,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"TEMPORAL"})
          Optional<CellTemporalReport> temporal)
      implements CellValueReport {
    public NumberValue {
      Objects.requireNonNull(numberValue, "numberValue must not be null");
      temporal = Objects.requireNonNullElseGet(temporal, Optional::empty);
    }

    @Override
    public String type() {
      return "NUMBER";
    }
  }

  /** Boolean effective value. */
  @Variant(id = "BOOLEAN", summary = "Effective formula evaluation is boolean.")
  record BooleanValue(Boolean booleanValue) implements CellValueReport {
    public BooleanValue {
      Objects.requireNonNull(booleanValue, "booleanValue must not be null");
    }

    @Override
    public String type() {
      return "BOOLEAN";
    }
  }

  /** Effective reported error literal, including GridGrind-owned evaluation-only states. */
  @Variant(
      id = "ERROR",
      summary =
          "Effective formula evaluation is one GridGrind-owned error literal, including canonical Excel tokens plus evaluator-specific tokens such as #CIRCULAR_REF!.")
  record ErrorValue(String errorValue) implements CellValueReport {
    public ErrorValue {
      errorValue = CellErrorLiteralValidation.requireReportedErrorLiteral(errorValue, "errorValue");
    }

    @Override
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
