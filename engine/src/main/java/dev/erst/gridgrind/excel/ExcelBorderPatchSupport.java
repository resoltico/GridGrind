package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Consumer;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Merges border patches onto a POI cell style while enforcing valid side/color combinations. */
final class ExcelBorderPatchSupport {
  private final XSSFWorkbook workbook;

  ExcelBorderPatchSupport(XSSFWorkbook workbook) {
    this.workbook = workbook;
  }

  void applyBorderPatch(XSSFCellStyle cellStyle, ExcelBorder border) {
    applyBorderSidePatch(
        mergedBorderSide(border.all(), border.top()),
        cellStyle::setBorderTop,
        cellStyle::setTopBorderColor);
    applyBorderSidePatch(
        mergedBorderSide(border.all(), border.right()),
        cellStyle::setBorderRight,
        cellStyle::setRightBorderColor);
    applyBorderSidePatch(
        mergedBorderSide(border.all(), border.bottom()),
        cellStyle::setBorderBottom,
        cellStyle::setBottomBorderColor);
    applyBorderSidePatch(
        mergedBorderSide(border.all(), border.left()),
        cellStyle::setBorderLeft,
        cellStyle::setLeftBorderColor);
  }

  private Optional<ExcelBorderSide> mergedBorderSide(
      Optional<ExcelBorderSide> defaultSide, Optional<ExcelBorderSide> explicitSide) {
    return effectiveBorderSide(
        mergedBorderStyle(defaultSide, explicitSide), mergedBorderColor(defaultSide, explicitSide));
  }

  private Optional<ExcelBorderStyle> mergedBorderStyle(
      Optional<ExcelBorderSide> defaultSide, Optional<ExcelBorderSide> explicitSide) {
    Objects.requireNonNull(defaultSide, "defaultSide must not be null");
    Objects.requireNonNull(explicitSide, "explicitSide must not be null");
    Optional<ExcelBorderStyle> explicitStyle = explicitSide.flatMap(ExcelBorderSide::style);
    return explicitStyle.isPresent() ? explicitStyle : defaultSide.flatMap(ExcelBorderSide::style);
  }

  private Optional<ExcelColor> mergedBorderColor(
      Optional<ExcelBorderSide> defaultSide, Optional<ExcelBorderSide> explicitSide) {
    Objects.requireNonNull(defaultSide, "defaultSide must not be null");
    Objects.requireNonNull(explicitSide, "explicitSide must not be null");
    if (explicitSide.isPresent() && explicitSide.orElseThrow().color().isPresent()) {
      return explicitSide.orElseThrow().color();
    }
    return defaultSide.isEmpty() || defaultSide.orElseThrow().color().isEmpty()
        ? Optional.empty()
        : defaultSide.orElseThrow().color();
  }

  private Optional<ExcelBorderSide> effectiveBorderSide(
      Optional<ExcelBorderStyle> style, Optional<ExcelColor> color) {
    Objects.requireNonNull(style, "style must not be null");
    Objects.requireNonNull(color, "color must not be null");
    if (style.isEmpty() && color.isEmpty()) {
      return Optional.empty();
    }
    if (color.isPresent() && (style.isEmpty() || style.orElseThrow() == ExcelBorderStyle.NONE)) {
      throw new IllegalArgumentException("border side color requires an effective border style");
    }
    return Optional.of(new ExcelBorderSide(style, color));
  }

  private void applyBorderSidePatch(
      Optional<ExcelBorderSide> sidePatch,
      Consumer<BorderStyle> styleSetter,
      Consumer<XSSFColor> colorSetter) {
    Objects.requireNonNull(sidePatch, "sidePatch must not be null");
    if (sidePatch.isEmpty()) {
      return;
    }
    ExcelBorderSide resolved = sidePatch.orElseThrow();
    styleSetter.accept(ExcelCellStylePoiBridge.toPoi(resolved.style().orElseThrow()));
    if (resolved.style().orElseThrow() == ExcelBorderStyle.NONE) {
      return;
    }
    if (resolved.color().isPresent()) {
      colorSetter.accept(ExcelColorSupport.toXssfColor(workbook, resolved.color().orElseThrow()));
    }
  }
}
