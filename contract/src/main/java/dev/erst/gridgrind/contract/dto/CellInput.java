package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.source.OoxmlTextValidation;
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
  @JsonSubTypes.Type(value = CellInput.NumberValue.class, name = "NUMBER"),
  @JsonSubTypes.Type(value = CellInput.BooleanValue.class, name = "BOOLEAN"),
  @JsonSubTypes.Type(value = CellInput.ErrorValue.class, name = "ERROR"),
  @JsonSubTypes.Type(value = CellInput.Date.class, name = "DATE"),
  @JsonSubTypes.Type(value = CellInput.DateTime.class, name = "DATE_TIME"),
  @JsonSubTypes.Type(value = CellInput.Formula.class, name = "FORMULA"),
  @JsonSubTypes.Type(value = CellInput.RawFormula.class, name = "RAW_FORMULA")
})
public sealed interface CellInput
    permits CellInput.Blank,
        CellInput.Text,
        CellInput.RichText,
        CellInput.NumberValue,
        CellInput.BooleanValue,
        CellInput.ErrorValue,
        CellInput.Date,
        CellInput.DateTime,
        CellInput.Formula,
        CellInput.RawFormula {

  /** Blank (empty) cell input that clears the target cell. */
  record Blank() implements CellInput {}

  /** Source-backed plain string cell input. */
  record Text(TextSourceInput source) implements CellInput {
    public Text {
      source = Validation.requireNonBlankTextSource(source, "source");
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
  record NumberValue(double number) implements CellInput {
    public NumberValue {
      Validation.requireFinite(number, "number");
    }
  }

  /** Boolean cell input. */
  record BooleanValue(boolean bool) implements CellInput {}

  /** Stored Excel error cell input such as {@code #REF!}, {@code #DIV/0!}, or {@code #N/A}. */
  record ErrorValue(String error) implements CellInput {
    public ErrorValue {
      error = Validation.requireNonBlank(error, "error");
      error = CellErrorLiteralValidation.requireStoredErrorLiteral(error, "error");
    }
  }

  /** Parseable Excel formula cell input loaded from one text source. */
  record Formula(TextSourceInput source) implements CellInput {
    public Formula {
      source = Validation.normalizeFormulaSource(source, "source");
    }
  }

  /** Opaque OOXML formula-body input that bypasses POI's normal formula parser on write. */
  record RawFormula(TextSourceInput source) implements CellInput {
    public RawFormula {
      source = Validation.normalizeRawFormulaSource(source, "source");
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

    static TextSourceInput requireNonBlankTextSource(TextSourceInput source, String fieldName) {
      required(source, fieldName);
      if (source instanceof TextSourceInput.Inline inline) {
        return TextSourceInput.inline(
            OoxmlTextValidation.requireXml10CharacterData(
                requireNonBlank(inline.text(), fieldName + ".text"), fieldName + ".text"));
      }
      return source;
    }

    static String requireNonBlank(String value, String fieldName) {
      required(value, fieldName);
      if (value.isBlank()) {
        throw new IllegalArgumentException(fieldName + " must not be blank");
      }
      return value;
    }

    static void requireFinite(double value, String fieldName) {
      if (!Double.isFinite(value)) {
        throw new IllegalArgumentException(fieldName + " must be finite");
      }
    }

    static String normalizeInlineFormula(String value, String fieldName) {
      String normalized = FormulaTextValidation.requireNormalFormulaBody(value, fieldName);
      FormulaInputSecurity.rejectDde(normalized); // LIM-023, LIM-031
      return normalized;
    }

    static TextSourceInput normalizeFormulaSource(TextSourceInput source, String fieldName) {
      required(source, fieldName);
      if (source instanceof TextSourceInput.Inline inline) {
        return TextSourceInput.inline(normalizeInlineFormula(inline.text(), fieldName + ".text"));
      }
      return source;
    }

    static TextSourceInput normalizeRawFormulaSource(TextSourceInput source, String fieldName) {
      required(source, fieldName);
      if (source instanceof TextSourceInput.Inline inline) {
        String value =
            FormulaTextValidation.requireRawFormulaBody(inline.text(), fieldName + ".text");
        return TextSourceInput.inline(value);
      }
      return source;
    }
  }
}
