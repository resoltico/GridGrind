package dev.erst.gridgrind.core.model.value;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Structured factual view of one requested or previewed cell. */
@Union(discriminator = "type", catalog = true, displayOrder = 1)
public sealed interface CellReport {
  /** Cell address in A1 notation. */
  CellAddress address();

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
  @Variant(
      id = "BLANK",
      summary =
          "Factual blank cell report. Only address and type are always present; FORMAT, STYLE, HYPERLINK, and COMMENT project the remaining fields.")
  record BlankReport(
      CellAddress address,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"FORMAT"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<String> displayValue,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"STYLE"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<CellStyleReport> style,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"HYPERLINK"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<HyperlinkTarget> hyperlink,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"COMMENT"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<CommentReport> comment)
      implements CellReport {
    public BlankReport {
      Objects.requireNonNull(address, "address must not be null");
      displayValue = normalizeOptional(displayValue, "displayValue");
      style = normalizeOptional(style, "style");
      hyperlink = normalizeOptional(hyperlink, "hyperlink");
      comment = normalizeOptional(comment, "comment");
    }

    @Override
    public String type() {
      return "BLANK";
    }
  }

  /** Cell report for a cell containing a plain text value. */
  @Variant(
      id = "TEXT",
      summary =
          "Factual text cell report. VALUE projects textValue, RICH_TEXT_RUNS projects runs, and the remaining fields are facet-gated.")
  record TextReport(
      CellAddress address,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"FORMAT"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<String> displayValue,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"STYLE"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<CellStyleReport> style,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"HYPERLINK"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<HyperlinkTarget> hyperlink,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"COMMENT"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<CommentReport> comment,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"VALUE"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<String> textValue,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"RICH_TEXT_RUNS"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<List<RichTextRunReport>> runs)
      implements CellReport {
    public TextReport {
      Objects.requireNonNull(address, "address must not be null");
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
    public String type() {
      return "TEXT";
    }
  }

  /** Cell report for a cell containing a numeric value. */
  @Variant(
      id = "NUMBER",
      summary =
          "Factual numeric cell report. VALUE projects numberValue; TEMPORAL adds date, time, or date-time semantics when a numeric Excel value is format-derived rather than a plain number.")
  record NumberReport(
      CellAddress address,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"FORMAT"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<String> displayValue,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"STYLE"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<CellStyleReport> style,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"HYPERLINK"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<HyperlinkTarget> hyperlink,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"COMMENT"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<CommentReport> comment,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"VALUE"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<Double> numberValue,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"TEMPORAL"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<CellTemporalReport> temporal)
      implements CellReport {
    public NumberReport {
      Objects.requireNonNull(address, "address must not be null");
      displayValue = normalizeOptional(displayValue, "displayValue");
      style = normalizeOptional(style, "style");
      hyperlink = normalizeOptional(hyperlink, "hyperlink");
      comment = normalizeOptional(comment, "comment");
      numberValue = normalizeOptional(numberValue, "numberValue");
      temporal = normalizeOptional(temporal, "temporal");
    }

    @Override
    public String type() {
      return "NUMBER";
    }
  }

  /** Cell report for a cell containing a boolean value. */
  @Variant(
      id = "BOOLEAN",
      summary =
          "Factual boolean cell report. VALUE projects booleanValue and the remaining fields are facet-gated.")
  record BooleanReport(
      CellAddress address,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"FORMAT"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<String> displayValue,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"STYLE"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<CellStyleReport> style,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"HYPERLINK"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<HyperlinkTarget> hyperlink,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"COMMENT"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<CommentReport> comment,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"VALUE"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<Boolean> booleanValue)
      implements CellReport {
    public BooleanReport {
      Objects.requireNonNull(address, "address must not be null");
      displayValue = normalizeOptional(displayValue, "displayValue");
      style = normalizeOptional(style, "style");
      hyperlink = normalizeOptional(hyperlink, "hyperlink");
      comment = normalizeOptional(comment, "comment");
      booleanValue = normalizeOptional(booleanValue, "booleanValue");
    }

    @Override
    public String type() {
      return "BOOLEAN";
    }
  }

  /** Cell report for a cell in a reported error state, including evaluation-only states. */
  @Variant(
      id = "ERROR",
      summary =
          "Factual error cell report. VALUE projects errorValue and the remaining fields are facet-gated.")
  record ErrorReport(
      CellAddress address,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"FORMAT"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<String> displayValue,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"STYLE"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<CellStyleReport> style,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"HYPERLINK"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<HyperlinkTarget> hyperlink,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"COMMENT"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<CommentReport> comment,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"VALUE"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<String> errorValue)
      implements CellReport {
    public ErrorReport {
      Objects.requireNonNull(address, "address must not be null");
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
    public String type() {
      return "ERROR";
    }
  }

  /** Cell report for a cell containing a formula. */
  @Variant(
      id = "FORMULA",
      summary =
          "Factual formula cell report with separate projected evaluation. FORMULA projects formula text, VALUE projects evaluation, and the remaining fields are facet-gated.")
  record FormulaReport(
      CellAddress address,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"FORMAT"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<String> displayValue,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"STYLE"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<CellStyleReport> style,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"HYPERLINK"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<HyperlinkTarget> hyperlink,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"COMMENT"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<CommentReport> comment,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"FORMULA"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<String> formula,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"VALUE"})
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<CellValueReport> evaluation)
      implements CellReport {
    public FormulaReport {
      Objects.requireNonNull(address, "address must not be null");
      displayValue = normalizeOptional(displayValue, "displayValue");
      style = normalizeOptional(style, "style");
      hyperlink = normalizeOptional(hyperlink, "hyperlink");
      comment = normalizeOptional(comment, "comment");
      formula = normalizeOptional(formula, "formula");
      evaluation = normalizeOptional(evaluation, "evaluation");
    }

    @Override
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
