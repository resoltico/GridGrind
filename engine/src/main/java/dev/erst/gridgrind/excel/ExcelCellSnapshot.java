package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelReportedCellErrorLiteral;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Immutable snapshot of a cell after formatting and, when needed, formula evaluation. */
public sealed interface ExcelCellSnapshot {
  /** A1-style address of the cell, such as {@code B4}. */
  String address();

  /** Canonical factual cell type: BLANK, TEXT, NUMBER, BOOLEAN, ERROR, or FORMULA. */
  String type();

  /** Formatted display string as Excel would render it in the cell. */
  String displayValue();

  /** Formatting style applied to the cell at the time of the snapshot. */
  ExcelCellStyleSnapshot style();

  /** Hyperlink and comment metadata captured for the cell at snapshot time. */
  ExcelCellMetadataSnapshot metadata();

  record BlankSnapshot(
      String address,
      String displayValue,
      ExcelCellStyleSnapshot style,
      ExcelCellMetadataSnapshot metadata)
      implements ExcelCellSnapshot {
    @Override
    public String type() {
      return "BLANK";
    }
  }

  record TextSnapshot(
      String address,
      String displayValue,
      ExcelCellStyleSnapshot style,
      ExcelCellMetadataSnapshot metadata,
      String textValue,
      @Nullable ExcelRichTextSnapshot richText)
      implements ExcelCellSnapshot {
    public TextSnapshot {
      java.util.Objects.requireNonNull(address, "address must not be null");
      java.util.Objects.requireNonNull(displayValue, "displayValue must not be null");
      java.util.Objects.requireNonNull(style, "style must not be null");
      java.util.Objects.requireNonNull(metadata, "metadata must not be null");
      java.util.Objects.requireNonNull(textValue, "textValue must not be null");
      if (richText != null && !textValue.equals(richText.plainText())) {
        throw new IllegalArgumentException("richText run text must concatenate to the textValue");
      }
    }

    @Override
    public String type() {
      return "TEXT";
    }
  }

  record NumberSnapshot(
      String address,
      String displayValue,
      ExcelCellStyleSnapshot style,
      ExcelCellMetadataSnapshot metadata,
      Double numberValue)
      implements ExcelCellSnapshot {
    @Override
    public String type() {
      return "NUMBER";
    }
  }

  record BooleanSnapshot(
      String address,
      String displayValue,
      ExcelCellStyleSnapshot style,
      ExcelCellMetadataSnapshot metadata,
      Boolean booleanValue)
      implements ExcelCellSnapshot {
    @Override
    public String type() {
      return "BOOLEAN";
    }
  }

  record ErrorSnapshot(
      String address,
      String displayValue,
      ExcelCellStyleSnapshot style,
      ExcelCellMetadataSnapshot metadata,
      String errorValue)
      implements ExcelCellSnapshot {
    public ErrorSnapshot {
      java.util.Objects.requireNonNull(address, "address must not be null");
      java.util.Objects.requireNonNull(displayValue, "displayValue must not be null");
      java.util.Objects.requireNonNull(style, "style must not be null");
      java.util.Objects.requireNonNull(metadata, "metadata must not be null");
      java.util.Objects.requireNonNull(errorValue, "errorValue must not be null");
      errorValue = ExcelReportedCellErrorLiteral.fromWireValue(errorValue).wireValue();
    }

    @Override
    public String type() {
      return "ERROR";
    }
  }

  record FormulaSnapshot(
      String address,
      String displayValue,
      ExcelCellStyleSnapshot style,
      ExcelCellMetadataSnapshot metadata,
      String formula,
      Optional<ExcelCellSnapshot> evaluation)
      implements ExcelCellSnapshot {
    public FormulaSnapshot {
      java.util.Objects.requireNonNull(address, "address must not be null");
      java.util.Objects.requireNonNull(displayValue, "displayValue must not be null");
      java.util.Objects.requireNonNull(style, "style must not be null");
      java.util.Objects.requireNonNull(metadata, "metadata must not be null");
      java.util.Objects.requireNonNull(formula, "formula must not be null");
      evaluation = java.util.Objects.requireNonNullElseGet(evaluation, Optional::empty);
      if (evaluation.isPresent() && evaluation.orElseThrow() instanceof FormulaSnapshot) {
        throw new IllegalArgumentException("formula evaluation must not itself be FORMULA");
      }
    }

    @Override
    public String type() {
      return "FORMULA";
    }
  }
}
