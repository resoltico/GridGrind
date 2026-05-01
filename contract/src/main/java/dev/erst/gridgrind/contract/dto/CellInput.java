package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;

/** JSON-friendly typed cell input used at the agent protocol boundary. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CellInput.Blank.class, name = "BLANK"),
  @JsonSubTypes.Type(value = CellInput.Text.class, name = "TEXT"),
  @JsonSubTypes.Type(value = CellInput.RichText.class, name = "RICH_TEXT"),
  @JsonSubTypes.Type(value = CellInput.Numeric.class, name = "NUMBER"),
  @JsonSubTypes.Type(value = CellInput.BooleanValue.class, name = "BOOLEAN"),
  @JsonSubTypes.Type(value = CellInput.Date.class, name = "DATE"),
  @JsonSubTypes.Type(value = CellInput.DateTime.class, name = "DATE_TIME"),
  @JsonSubTypes.Type(value = CellInput.Formula.class, name = "FORMULA")
})
public sealed interface CellInput {

  /** Blank (empty) cell input that clears the target cell. */
  record Blank() implements CellInput {}

  /** Plain string cell input. */
  record Text(TextSourceInput source) implements CellInput {
    public Text {
      source = Validation.required(source, "source");
    }
  }

  /** Structured rich-text string cell input. */
  record RichText(List<RichTextRunInput> runs) implements CellInput {
    public RichText {
      Validation.required(runs, "runs");
      runs = List.copyOf(runs);
      if (runs.isEmpty()) {
        throw new IllegalArgumentException("runs must not be empty");
      }
      for (RichTextRunInput run : runs) {
        Validation.required(run, "runs element");
      }
    }
  }

  /** Numeric cell input stored as a double. */
  record Numeric(double number) implements CellInput {
    public Numeric {
      Validation.requireFinite(number, "number");
    }
  }

  /** Boolean cell input. */
  record BooleanValue(boolean bool) implements CellInput {}

  /** Excel formula cell input. A leading {@code =} sign is stripped automatically if present. */
  record Formula(TextSourceInput source) implements CellInput {
    public Formula {
      source = Validation.required(source, "source");
      if (source instanceof TextSourceInput.Inline inline) {
        String formula = inline.text();
        if (formula.startsWith("=")) {
          formula = formula.substring(1);
        }
        if (formula.isBlank()) {
          throw new IllegalArgumentException("source must not be blank after stripping leading =");
        }
        source = new TextSourceInput.Inline(formula);
      }
    }
  }

  /** ISO-8601 date cell input formatted as a date value in Excel (e.g. {@code "2024-03-15"}). */
  record Date(LocalDate date) implements CellInput {
    public Date {
      Validation.required(date, "date");
    }
  }

  /**
   * ISO-8601 date-time cell input formatted as a date-time value in Excel (e.g. {@code
   * "2024-03-15T09:30:00"}).
   */
  record DateTime(LocalDateTime dateTime) implements CellInput {
    public DateTime {
      Validation.required(dateTime, "dateTime");
    }
  }

  /** Null-checking helpers for CellInput compact constructors. */
  final class Validation {
    private Validation() {}

    static <T> T required(T value, String fieldName) {
      if (value == null) {
        throw new IllegalArgumentException(fieldName + " must not be null");
      }
      return value;
    }

    static void requireFinite(double value, String fieldName) {
      if (!Double.isFinite(value)) {
        throw new IllegalArgumentException(fieldName + " must be finite");
      }
    }
  }
}
