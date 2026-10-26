package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.protection.CellProtectionPatch;
import dev.erst.gridgrind.core.model.style.CellAlignmentPatch;
import dev.erst.gridgrind.core.model.style.CellBorderPatch;
import dev.erst.gridgrind.core.model.style.CellFillSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import java.util.Objects;
import java.util.Optional;

/** Agent-facing style patch that can be applied to one cell or a rectangular range. */
public record ExcelCellStyle(
    Optional<String> numberFormat,
    Optional<CellAlignmentPatch> alignment,
    Optional<ExcelCellFont> font,
    Optional<CellFillSpec> fill,
    Optional<CellBorderPatch> border,
    Optional<CellProtectionPatch> protection) {
  public ExcelCellStyle {
    Objects.requireNonNull(numberFormat, "numberFormat must not be null");
    Objects.requireNonNull(alignment, "alignment must not be null");
    Objects.requireNonNull(font, "font must not be null");
    Objects.requireNonNull(fill, "fill must not be null");
    Objects.requireNonNull(border, "border must not be null");
    Objects.requireNonNull(protection, "protection must not be null");
    numberFormat.ifPresent(
        value -> {
          if (value.isBlank()) {
            throw new IllegalArgumentException("numberFormat must not be blank");
          }
        });
    if (numberFormat.isEmpty()
        && alignment.isEmpty()
        && font.isEmpty()
        && fill.isEmpty()
        && border.isEmpty()
        && protection.isEmpty()) {
      throw new IllegalArgumentException("style must set at least one attribute");
    }
  }

  /** Creates a style patch that only changes the number format. */
  public static ExcelCellStyle numberFormat(String numberFormat) {
    return new ExcelCellStyle(
        Optional.of(numberFormat),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }

  /** Creates a style patch that only changes font emphasis. */
  public static ExcelCellStyle emphasis(Boolean bold, Boolean italic) {
    return new ExcelCellStyle(
        Optional.empty(),
        Optional.empty(),
        Optional.of(
            new ExcelCellFont(
                Optional.ofNullable(bold),
                Optional.ofNullable(italic),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty())),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }

  /** Creates a style patch that only changes horizontal and vertical alignment. */
  public static ExcelCellStyle alignment(
      ExcelHorizontalAlignment horizontalAlignment, ExcelVerticalAlignment verticalAlignment) {
    if (horizontalAlignment == null && verticalAlignment == null) {
      throw new IllegalArgumentException("alignment must set at least one attribute");
    }
    return new ExcelCellStyle(
        Optional.empty(),
        Optional.of(
            new CellAlignmentPatch(
                Optional.empty(),
                Optional.ofNullable(horizontalAlignment),
                Optional.ofNullable(verticalAlignment),
                Optional.empty(),
                Optional.empty())),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }
}
