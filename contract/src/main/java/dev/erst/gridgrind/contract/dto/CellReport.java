package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Structured factual view of one requested or previewed cell. */
@JsonTypeInfo(
    use = JsonTypeInfo.Id.NAME,
    include = JsonTypeInfo.As.EXISTING_PROPERTY,
    property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CellReport.BlankReport.class, name = "BLANK"),
  @JsonSubTypes.Type(value = CellReport.TextReport.class, name = "TEXT"),
  @JsonSubTypes.Type(value = CellReport.NumberReport.class, name = "NUMBER"),
  @JsonSubTypes.Type(value = CellReport.BooleanReport.class, name = "BOOLEAN"),
  @JsonSubTypes.Type(value = CellReport.ErrorReport.class, name = "ERROR"),
  @JsonSubTypes.Type(value = CellReport.FormulaReport.class, name = "FORMULA")
})
public sealed interface CellReport {
  /** Cell address in A1 notation. */
  String address();

  /** Canonical published cell type. */
  String type();

  /** Formatted display string as shown in Excel when the FORMAT facet is projected. */
  Optional<String> displayValue();

  /** Style snapshot captured for this cell when the STYLE facet is projected. */
  Optional<CellStyleReport> style();

  /** Hyperlink metadata when the HYPERLINK facet is projected and metadata exists. */
  Optional<HyperlinkTarget> hyperlink();

  /** Comment metadata when the COMMENT facet is projected and metadata exists. */
  Optional<CommentReport> comment();

  /** Cell report for a cell with no value or formula. */
  record BlankReport(
      String address,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> displayValue,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CellStyleReport> style,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<HyperlinkTarget> hyperlink,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CommentReport> comment)
      implements CellReport {
    public BlankReport {
      Objects.requireNonNull(address, "address must not be null");
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
      displayValue = normalizeOptional(displayValue, "displayValue");
      style = normalizeOptional(style, "style");
      hyperlink = normalizeOptional(hyperlink, "hyperlink");
      comment = normalizeOptional(comment, "comment");
    }

    @Override
    @JsonProperty
    public String type() {
      return "BLANK";
    }
  }

  /** Cell report for a cell containing a plain text value. */
  record TextReport(
      String address,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> displayValue,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CellStyleReport> style,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<HyperlinkTarget> hyperlink,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CommentReport> comment,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> textValue,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<List<RichTextRunReport>> runs)
      implements CellReport {
    public TextReport {
      Objects.requireNonNull(address, "address must not be null");
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
      displayValue = normalizeOptional(displayValue, "displayValue");
      style = normalizeOptional(style, "style");
      hyperlink = normalizeOptional(hyperlink, "hyperlink");
      comment = normalizeOptional(comment, "comment");
      textValue = normalizeOptional(textValue, "textValue");
      runs = copyRichTextRuns(runs, "runs");
      if (runs.isPresent()) {
        if (textValue.isEmpty()) {
          throw new IllegalArgumentException("runs require textValue to be present");
        }
        if (!textValue.orElseThrow().equals(concatenateRuns(runs.orElseThrow()))) {
          throw new IllegalArgumentException("runs text must concatenate to the textValue");
        }
      }
    }

    @Override
    @JsonProperty
    public String type() {
      return "TEXT";
    }
  }

  /** Cell report for a cell containing a numeric value. */
  record NumberReport(
      String address,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> displayValue,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CellStyleReport> style,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<HyperlinkTarget> hyperlink,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CommentReport> comment,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Double> numberValue,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CellTemporalReport> temporal)
      implements CellReport {
    public NumberReport {
      Objects.requireNonNull(address, "address must not be null");
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
      displayValue = normalizeOptional(displayValue, "displayValue");
      style = normalizeOptional(style, "style");
      hyperlink = normalizeOptional(hyperlink, "hyperlink");
      comment = normalizeOptional(comment, "comment");
      numberValue = normalizeOptional(numberValue, "numberValue");
      temporal = normalizeOptional(temporal, "temporal");
    }

    @Override
    @JsonProperty
    public String type() {
      return "NUMBER";
    }
  }

  /** Cell report for a cell containing a boolean value. */
  record BooleanReport(
      String address,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> displayValue,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CellStyleReport> style,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<HyperlinkTarget> hyperlink,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CommentReport> comment,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Boolean> booleanValue)
      implements CellReport {
    public BooleanReport {
      Objects.requireNonNull(address, "address must not be null");
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
      displayValue = normalizeOptional(displayValue, "displayValue");
      style = normalizeOptional(style, "style");
      hyperlink = normalizeOptional(hyperlink, "hyperlink");
      comment = normalizeOptional(comment, "comment");
      booleanValue = normalizeOptional(booleanValue, "booleanValue");
    }

    @Override
    @JsonProperty
    public String type() {
      return "BOOLEAN";
    }
  }

  /** Cell report for a cell in a reported error state, including evaluation-only states. */
  record ErrorReport(
      String address,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> displayValue,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CellStyleReport> style,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<HyperlinkTarget> hyperlink,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CommentReport> comment,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> errorValue)
      implements CellReport {
    public ErrorReport {
      Objects.requireNonNull(address, "address must not be null");
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
      displayValue = normalizeOptional(displayValue, "displayValue");
      style = normalizeOptional(style, "style");
      hyperlink = normalizeOptional(hyperlink, "hyperlink");
      comment = normalizeOptional(comment, "comment");
      errorValue =
          normalizeOptional(errorValue, "errorValue")
              .map(
                  value ->
                      CellErrorLiteralValidation.requireReportedErrorLiteral(value, "errorValue"));
    }

    @Override
    @JsonProperty
    public String type() {
      return "ERROR";
    }
  }

  /** Cell report for a cell containing a formula. */
  record FormulaReport(
      String address,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> displayValue,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CellStyleReport> style,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<HyperlinkTarget> hyperlink,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CommentReport> comment,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> formula,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CellValueReport> evaluation)
      implements CellReport {
    public FormulaReport {
      Objects.requireNonNull(address, "address must not be null");
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
      displayValue = normalizeOptional(displayValue, "displayValue");
      style = normalizeOptional(style, "style");
      hyperlink = normalizeOptional(hyperlink, "hyperlink");
      comment = normalizeOptional(comment, "comment");
      formula = normalizeOptional(formula, "formula");
      evaluation = normalizeOptional(evaluation, "evaluation");
    }

    @Override
    @JsonProperty
    public String type() {
      return "FORMULA";
    }
  }

  private static <T> Optional<T> normalizeOptional(Optional<T> value, String fieldName) {
    Optional<T> normalized = Objects.requireNonNullElseGet(value, Optional::empty);
    normalized.ifPresent(v -> Objects.requireNonNull(v, fieldName + " must not contain nulls"));
    return normalized;
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
