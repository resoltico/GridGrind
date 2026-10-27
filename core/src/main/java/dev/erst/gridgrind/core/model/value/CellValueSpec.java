package dev.erst.gridgrind.core.model.value;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.formula.FormulaInputSecurity;
import dev.erst.gridgrind.core.model.formula.FormulaTextValidation;
import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.source.OoxmlTextValidation;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;

/** JSON-friendly typed cell input used at the agent protocol boundary. */
@Union(discriminator = "type", catalog = true, displayOrder = 16)
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_DDE_FORMULA_INJECTION,
  dev.erst.gridgrind.core.limits.Limit.ID_WEBSERVICE_FORMULA_REJECTION
})
public sealed interface CellValueSpec
    permits dev.erst.gridgrind.core.model.value.CellValueSpec.Blank,
        dev.erst.gridgrind.core.model.value.CellValueSpec.Text,
        dev.erst.gridgrind.core.model.value.CellValueSpec.RichText,
        dev.erst.gridgrind.core.model.value.CellValueSpec.NumberValue,
        dev.erst.gridgrind.core.model.value.CellValueSpec.BooleanValue,
        dev.erst.gridgrind.core.model.value.CellValueSpec.ErrorValue,
        dev.erst.gridgrind.core.model.value.CellValueSpec.Date,
        dev.erst.gridgrind.core.model.value.CellValueSpec.DateTime,
        dev.erst.gridgrind.core.model.value.CellValueSpec.Formula,
        dev.erst.gridgrind.core.model.value.CellValueSpec.RawFormula {

  /** Blank (empty) cell input that clears the target cell. */
  @Variant(id = "BLANK", summary = "Write an empty cell.")
  record Blank() implements CellValueSpec {}

  /** Source-backed plain string cell input. */
  @Variant(
      id = "TEXT",
      summary =
          "Write one string cell value from an inline, file-backed, or standard-input text source. Blank resolved text is rejected; use BLANK for empty cells.")
  record Text(TextSourceSpec value) implements CellValueSpec {
    public Text {
      value = Validation.requireNonBlankTextSource(value, "value");
    }
  }

  /** Structured rich-text string cell input. */
  @Variant(
      id = "RICH_TEXT",
      summary =
          "Write a structured string cell value with an ordered rich-text run list. Run text concatenates to the stored plain string value and each run may override font attributes independently.")
  record RichText(List<RichTextRunSpec> runs) implements CellValueSpec {
    public RichText {
      Validation.required(runs, "runs");
      runs = List.copyOf(runs);
      if (runs.isEmpty()) {
        throw new IllegalArgumentException("runs must not be empty");
      }
      for (RichTextRunSpec run : runs) {
        Validation.required(run, "runs element");
      }
    }
  }

  /** Numeric cell input stored as a double. */
  @Variant(id = "NUMBER", summary = "Write one numeric cell value.")
  record NumberValue(double value) implements CellValueSpec {
    public NumberValue {
      Validation.requireFinite(value, "value");
    }
  }

  /** Boolean cell input. */
  @Variant(id = "BOOLEAN", summary = "Write one boolean cell value.")
  record BooleanValue(boolean value) implements CellValueSpec {}

  /** Stored Excel error cell input such as {@code #REF!}, {@code #DIV/0!}, or {@code #N/A}. */
  @Variant(
      id = "ERROR",
      summary =
          "Write one exact stored Excel error literal such as #REF!, #DIV/0!, or #N/A. Evaluation-only states such as #CIRCULAR_REF! are reported on readback but cannot be authored as stored cell values.")
  record ErrorValue(String value) implements CellValueSpec {
    public ErrorValue {
      value = CellErrorLiteralValidation.requireStoredErrorLiteral(value, "value");
    }
  }

  /** Parseable Excel formula cell input loaded from one text source. */
  @Variant(
      id = "FORMULA",
      summary =
          "Write one Excel formula from an inline, file-backed, or standard-input text source. Formula text is the OOXML <f> body: it must not begin with = and is parsed by Apache POI.")
  record Formula(TextSourceSpec value) implements CellValueSpec {
    public Formula {
      value = Validation.normalizeFormulaSource(value, "value");
    }
  }

  /** Opaque OOXML formula-body input that bypasses POI's normal formula parser on write. */
  @Variant(
      id = "RAW_FORMULA",
      summary =
          "Write opaque OOXML <f> character data from one text source without POI formula parsing. It must be nonempty, must not begin with =, and rejects XML 1.0-invalid characters as INVALID_FORMULA_TEXT; XML-looking formula string literals are preserved.")
  record RawFormula(TextSourceSpec value) implements CellValueSpec {
    public RawFormula {
      value = Validation.normalizeRawFormulaSource(value, "value");
    }
  }

  /** ISO-8601 date cell input formatted as a date value in Excel (e.g. {@code "2024-03-15"}). */
  @Variant(
      id = "DATE",
      summary =
          "Write an ISO-8601 date value. Stored as an Excel serial number; readback keeps type=NUMBER, and the TEMPORAL facet can expose kind=DATE.")
  record Date(LocalDate value) implements CellValueSpec, ResolvedCellWrite {
    public Date {
      Validation.required(value, "value");
    }
  }

  /**
   * ISO-8601 date-time cell input formatted as a date-time value in Excel (e.g. {@code
   * "2024-03-15T09:30:00"}).
   */
  @Variant(
      id = "DATE_TIME",
      summary =
          "Write an ISO-8601 date-time value. Stored as an Excel serial number; readback keeps type=NUMBER, and the TEMPORAL facet can expose kind=DATE_TIME.")
  record DateTime(LocalDateTime value) implements CellValueSpec, ResolvedCellWrite {
    public DateTime {
      Validation.required(value, "value");
    }
  }

  /** Null-checking helpers for CellValueSpec compact constructors. */
  final class Validation {
    private Validation() {}

    static <T> T required(T value, String fieldName) {
      if (value == null) {
        throw new IllegalArgumentException(fieldName + " must not be null");
      }
      return value;
    }

    static TextSourceSpec requireNonBlankTextSource(TextSourceSpec source, String fieldName) {
      required(source, fieldName);
      if (source instanceof TextSourceSpec.Inline inline) {
        return TextSourceSpec.inline(
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
      FormulaInputSecurity.rejectDde(normalized);
      return normalized;
    }

    static TextSourceSpec normalizeFormulaSource(TextSourceSpec source, String fieldName) {
      required(source, fieldName);
      if (source instanceof TextSourceSpec.Inline inline) {
        return TextSourceSpec.inline(normalizeInlineFormula(inline.text(), fieldName + ".text"));
      }
      return source;
    }

    static TextSourceSpec normalizeRawFormulaSource(TextSourceSpec source, String fieldName) {
      required(source, fieldName);
      if (source instanceof TextSourceSpec.Inline inline) {
        String value =
            FormulaTextValidation.requireRawFormulaBody(inline.text(), fieldName + ".text");
        return TextSourceSpec.inline(value);
      }
      return source;
    }
  }
}
