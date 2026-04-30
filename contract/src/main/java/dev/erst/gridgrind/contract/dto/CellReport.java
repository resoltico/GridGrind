package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Structured view of one requested or previewed cell. */
@JsonTypeInfo(
    use = JsonTypeInfo.Id.NAME,
    include = JsonTypeInfo.As.EXISTING_PROPERTY,
    property = "effectiveType")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CellReport.BlankReport.class, name = "BLANK"),
  @JsonSubTypes.Type(value = CellReport.TextReport.class, name = "STRING"),
  @JsonSubTypes.Type(value = CellReport.NumberReport.class, name = "NUMBER"),
  @JsonSubTypes.Type(value = CellReport.BooleanReport.class, name = "BOOLEAN"),
  @JsonSubTypes.Type(value = CellReport.ErrorReport.class, name = "ERROR"),
  @JsonSubTypes.Type(value = CellReport.FormulaReport.class, name = "FORMULA")
})
public sealed interface CellReport {
  /** Cell address in A1 notation. */
  String address();

  /** POI cell type as declared before evaluation. */
  String declaredType();

  /** Effective cell type after formula evaluation. */
  String effectiveType();

  /** Formatted display string as shown in Excel. */
  String displayValue();

  /** Style snapshot captured for this cell. */
  GridGrindWorkbookSurfaceReports.CellStyleReport style();

  /** Hyperlink metadata when the workbook stores a hyperlink for this cell. */
  Optional<HyperlinkTarget> hyperlink();

  /** Comment metadata when the workbook stores a comment for this cell. */
  Optional<GridGrindWorkbookSurfaceReports.CommentReport> comment();

  /** CellReport for a cell with no value or formula. */
  record BlankReport(
      String address,
      String declaredType,
      String displayValue,
      GridGrindWorkbookSurfaceReports.CellStyleReport style,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<HyperlinkTarget> hyperlink,
      @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<GridGrindWorkbookSurfaceReports.CommentReport> comment)
      implements CellReport {
    public BlankReport {
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(declaredType, "declaredType must not be null");
      Objects.requireNonNull(displayValue, "displayValue must not be null");
      Objects.requireNonNull(style, "style must not be null");
      hyperlink = Objects.requireNonNullElseGet(hyperlink, Optional::empty);
      comment = Objects.requireNonNullElseGet(comment, Optional::empty);
    }

    @Override
    @JsonProperty
    public String effectiveType() {
      return "BLANK";
    }
  }

  /** CellReport for a cell containing a plain string value. */
  record TextReport(
      String address,
      String declaredType,
      String displayValue,
      GridGrindWorkbookSurfaceReports.CellStyleReport style,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<HyperlinkTarget> hyperlink,
      @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<GridGrindWorkbookSurfaceReports.CommentReport> comment,
      String stringValue,
      Optional<List<RichTextRunReport>> richText)
      implements CellReport {
    public TextReport {
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(declaredType, "declaredType must not be null");
      Objects.requireNonNull(displayValue, "displayValue must not be null");
      Objects.requireNonNull(style, "style must not be null");
      Objects.requireNonNull(stringValue, "stringValue must not be null");
      hyperlink = Objects.requireNonNullElseGet(hyperlink, Optional::empty);
      comment = Objects.requireNonNullElseGet(comment, Optional::empty);
      richText = copyRichTextRuns(richText, "richText");
      if (richText.isPresent()
          && !stringValue.equals(concatenateRichTextRuns(richText.orElseThrow()))) {
        throw new IllegalArgumentException("richText run text must concatenate to the stringValue");
      }
    }

    @Override
    @JsonProperty
    public String effectiveType() {
      return "STRING";
    }
  }

  /** CellReport for a cell containing a numeric value. */
  record NumberReport(
      String address,
      String declaredType,
      String displayValue,
      GridGrindWorkbookSurfaceReports.CellStyleReport style,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<HyperlinkTarget> hyperlink,
      @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<GridGrindWorkbookSurfaceReports.CommentReport> comment,
      Double numberValue)
      implements CellReport {
    public NumberReport {
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(declaredType, "declaredType must not be null");
      Objects.requireNonNull(displayValue, "displayValue must not be null");
      Objects.requireNonNull(style, "style must not be null");
      hyperlink = Objects.requireNonNullElseGet(hyperlink, Optional::empty);
      comment = Objects.requireNonNullElseGet(comment, Optional::empty);
    }

    @Override
    @JsonProperty
    public String effectiveType() {
      return "NUMBER";
    }
  }

  /** CellReport for a cell containing a boolean value. */
  record BooleanReport(
      String address,
      String declaredType,
      String displayValue,
      GridGrindWorkbookSurfaceReports.CellStyleReport style,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<HyperlinkTarget> hyperlink,
      @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<GridGrindWorkbookSurfaceReports.CommentReport> comment,
      Boolean booleanValue)
      implements CellReport {
    public BooleanReport {
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(declaredType, "declaredType must not be null");
      Objects.requireNonNull(displayValue, "displayValue must not be null");
      Objects.requireNonNull(style, "style must not be null");
      hyperlink = Objects.requireNonNullElseGet(hyperlink, Optional::empty);
      comment = Objects.requireNonNullElseGet(comment, Optional::empty);
    }

    @Override
    @JsonProperty
    public String effectiveType() {
      return "BOOLEAN";
    }
  }

  /** CellReport for a cell in an error state (e.g., #DIV/0!, #REF!). */
  record ErrorReport(
      String address,
      String declaredType,
      String displayValue,
      GridGrindWorkbookSurfaceReports.CellStyleReport style,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<HyperlinkTarget> hyperlink,
      @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<GridGrindWorkbookSurfaceReports.CommentReport> comment,
      String errorValue)
      implements CellReport {
    public ErrorReport {
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(declaredType, "declaredType must not be null");
      Objects.requireNonNull(displayValue, "displayValue must not be null");
      Objects.requireNonNull(style, "style must not be null");
      hyperlink = Objects.requireNonNullElseGet(hyperlink, Optional::empty);
      comment = Objects.requireNonNullElseGet(comment, Optional::empty);
    }

    @Override
    @JsonProperty
    public String effectiveType() {
      return "ERROR";
    }
  }

  /** CellReport for a cell containing a formula, with its evaluated result nested inside. */
  record FormulaReport(
      String address,
      String declaredType,
      String displayValue,
      GridGrindWorkbookSurfaceReports.CellStyleReport style,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<HyperlinkTarget> hyperlink,
      @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<GridGrindWorkbookSurfaceReports.CommentReport> comment,
      String formula,
      CellReport evaluation)
      implements CellReport {
    public FormulaReport {
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(declaredType, "declaredType must not be null");
      Objects.requireNonNull(displayValue, "displayValue must not be null");
      Objects.requireNonNull(style, "style must not be null");
      Objects.requireNonNull(formula, "formula must not be null");
      hyperlink = Objects.requireNonNullElseGet(hyperlink, Optional::empty);
      comment = Objects.requireNonNullElseGet(comment, Optional::empty);
    }

    @Override
    @JsonProperty
    public String effectiveType() {
      return "FORMULA";
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

  private static String concatenateRichTextRuns(List<RichTextRunReport> runs) {
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
