package dev.erst.gridgrind.protocol.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
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
  record Text(String text) implements CellInput {
    public Text {
      Validation.required(text, "text");
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
  record Numeric(Double number) implements CellInput {
    public Numeric {
      Validation.required(number, "number");
    }
  }

  /** Boolean cell input. */
  record BooleanValue(Boolean bool) implements CellInput {
    public BooleanValue {
      Validation.required(bool, "bool");
    }
  }

  /** Excel formula cell input. A leading {@code =} sign is stripped automatically if present. */
  record Formula(String formula) implements CellInput {
    public Formula {
      Validation.required(formula, "formula");
      if (formula.startsWith("=")) {
        formula = formula.substring(1);
      }
      if (formula.isBlank()) {
        throw new IllegalArgumentException("formula must not be blank after stripping leading =");
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
  }
}
