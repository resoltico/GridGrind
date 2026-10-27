package dev.erst.gridgrind.core.model.value;

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
  default String type() {
    return dev.erst.gridgrind.core.catalog.ProtocolTypeMetadataSupport.requiredTypeId(getClass());
  }

  /** Effective scalar value when the VALUE facet is projected. */
  Optional<? extends CellValue> value();

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
          "Factual blank cell report. Only address and type are always present; VALUE projects the blank scalar, and FORMAT, STYLE, HYPERLINK, and COMMENT project their metadata.")
  record BlankReport(
      CellAddress address,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"FORMAT"}) Optional<String> displayValue,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"STYLE"}) Optional<CellStyleReport> style,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"HYPERLINK"})
          Optional<HyperlinkTarget> hyperlink,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"COMMENT"}) Optional<CommentReport> comment,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"VALUE"}) Optional<CellValue.Blank> value)
      implements CellReport {
    public BlankReport {
      Objects.requireNonNull(address, "address must not be null");
      displayValue = normalizeOptional(displayValue, "displayValue");
      style = normalizeOptional(style, "style");
      hyperlink = normalizeOptional(hyperlink, "hyperlink");
      comment = normalizeOptional(comment, "comment");
      value = normalizeOptional(value, "value");
    }
  }

  /** Cell report for a cell containing a plain text value. */
  @Variant(
      id = "TEXT",
      summary =
          "Factual text cell report. VALUE projects value, RICH_TEXT_RUNS projects runs, and the remaining fields are facet-gated.")
  record TextReport(
      CellAddress address,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"FORMAT"}) Optional<String> displayValue,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"STYLE"}) Optional<CellStyleReport> style,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"HYPERLINK"})
          Optional<HyperlinkTarget> hyperlink,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"COMMENT"}) Optional<CommentReport> comment,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"VALUE"}) Optional<CellValue.Text> value,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"RICH_TEXT_RUNS"})
          Optional<List<RichTextRunReport>> runs)
      implements CellReport {
    public TextReport {
      Objects.requireNonNull(address, "address must not be null");
      displayValue = normalizeOptional(displayValue, "displayValue");
      style = normalizeOptional(style, "style");
      hyperlink = normalizeOptional(hyperlink, "hyperlink");
      comment = normalizeOptional(comment, "comment");
      value = normalizeOptional(value, "value");
      runs = copyRichTextRuns(runs, "runs");
      if (runs.isPresent()) {
        if (value.isEmpty()) {
          throw new IllegalArgumentException("runs require value to be present");
        }
        if (!value.orElseThrow().value().equals(concatenateRuns(runs.orElseThrow()))) {
          throw new IllegalArgumentException("runs text must concatenate to the value");
        }
      }
    }
  }

  /** Cell report for a cell containing a numeric value. */
  @Variant(
      id = "NUMBER",
      summary =
          "Factual numeric cell report. VALUE projects value; TEMPORAL adds date, time, or date-time semantics when a numeric Excel value is format-derived rather than a plain number.")
  record NumberReport(
      CellAddress address,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"FORMAT"}) Optional<String> displayValue,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"STYLE"}) Optional<CellStyleReport> style,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"HYPERLINK"})
          Optional<HyperlinkTarget> hyperlink,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"COMMENT"}) Optional<CommentReport> comment,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"VALUE"})
          Optional<CellValue.NumberValue> value,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"TEMPORAL"})
          Optional<CellTemporalReport> temporal)
      implements CellReport {
    public NumberReport {
      Objects.requireNonNull(address, "address must not be null");
      displayValue = normalizeOptional(displayValue, "displayValue");
      style = normalizeOptional(style, "style");
      hyperlink = normalizeOptional(hyperlink, "hyperlink");
      comment = normalizeOptional(comment, "comment");
      value = normalizeOptional(value, "value");
      temporal = normalizeOptional(temporal, "temporal");
    }
  }

  /** Cell report for a cell containing a boolean value. */
  @Variant(
      id = "BOOLEAN",
      summary =
          "Factual boolean cell report. VALUE projects value and the remaining fields are facet-gated.")
  record BooleanReport(
      CellAddress address,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"FORMAT"}) Optional<String> displayValue,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"STYLE"}) Optional<CellStyleReport> style,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"HYPERLINK"})
          Optional<HyperlinkTarget> hyperlink,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"COMMENT"}) Optional<CommentReport> comment,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"VALUE"})
          Optional<CellValue.BooleanValue> value)
      implements CellReport {
    public BooleanReport {
      Objects.requireNonNull(address, "address must not be null");
      displayValue = normalizeOptional(displayValue, "displayValue");
      style = normalizeOptional(style, "style");
      hyperlink = normalizeOptional(hyperlink, "hyperlink");
      comment = normalizeOptional(comment, "comment");
      value = normalizeOptional(value, "value");
    }
  }

  /** Cell report for a cell in a reported error state, including evaluation-only states. */
  @Variant(
      id = "ERROR",
      summary =
          "Factual error cell report. VALUE projects value and the remaining fields are facet-gated.")
  record ErrorReport(
      CellAddress address,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"FORMAT"}) Optional<String> displayValue,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"STYLE"}) Optional<CellStyleReport> style,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"HYPERLINK"})
          Optional<HyperlinkTarget> hyperlink,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"COMMENT"}) Optional<CommentReport> comment,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"VALUE"}) Optional<CellValue.ErrorValue> value)
      implements CellReport {
    public ErrorReport {
      Objects.requireNonNull(address, "address must not be null");
      displayValue = normalizeOptional(displayValue, "displayValue");
      style = normalizeOptional(style, "style");
      hyperlink = normalizeOptional(hyperlink, "hyperlink");
      comment = normalizeOptional(comment, "comment");
      value = normalizeOptional(value, "value");
    }
  }

  /** Cell report for a cell containing a formula. */
  @Variant(
      id = "FORMULA",
      summary =
          "Factual formula cell report with separate projected scalar value. FORMULA projects formula text, VALUE projects value, and the remaining fields are facet-gated.")
  record FormulaReport(
      CellAddress address,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"FORMAT"}) Optional<String> displayValue,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"STYLE"}) Optional<CellStyleReport> style,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"HYPERLINK"})
          Optional<HyperlinkTarget> hyperlink,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"COMMENT"}) Optional<CommentReport> comment,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"FORMULA"}) Optional<String> formula,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"VALUE"}) Optional<CellValue> value,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"RICH_TEXT_RUNS"})
          Optional<List<RichTextRunReport>> runs,
      @dev.erst.gridgrind.core.metadata.ProjectedBy({"TEMPORAL"})
          Optional<CellTemporalReport> temporal)
      implements CellReport {
    public FormulaReport {
      Objects.requireNonNull(address, "address must not be null");
      displayValue = normalizeOptional(displayValue, "displayValue");
      style = normalizeOptional(style, "style");
      hyperlink = normalizeOptional(hyperlink, "hyperlink");
      comment = normalizeOptional(comment, "comment");
      formula = normalizeOptional(formula, "formula");
      value = normalizeOptional(value, "value");
      runs = copyRichTextRuns(runs, "runs");
      temporal = normalizeOptional(temporal, "temporal");
      if (runs.isPresent()) {
        if (!(value.orElse(null) instanceof CellValue.Text text)) {
          throw new IllegalArgumentException("runs require a text value to be present");
        }
        if (!text.value().equals(concatenateRuns(runs.orElseThrow()))) {
          throw new IllegalArgumentException("runs text must concatenate to the value");
        }
      }
      if (temporal.isPresent() && !(value.orElse(null) instanceof CellValue.NumberValue)) {
        throw new IllegalArgumentException("temporal requires a numeric value to be present");
      }
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
