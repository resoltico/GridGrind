package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.CellBorderPatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
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

  void applyBorderPatch(XSSFCellStyle cellStyle, CellBorderPatch border) {
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

  private Optional<BorderSidePatch> mergedBorderSide(
      Optional<BorderSidePatch> defaultSide, Optional<BorderSidePatch> explicitSide) {
    return effectiveBorderSide(
        mergedBorderStyle(defaultSide, explicitSide), mergedBorderColor(defaultSide, explicitSide));
  }

  private Optional<ExcelBorderStyle> mergedBorderStyle(
      Optional<BorderSidePatch> defaultSide, Optional<BorderSidePatch> explicitSide) {
    Objects.requireNonNull(defaultSide, "defaultSide must not be null");
    Objects.requireNonNull(explicitSide, "explicitSide must not be null");
    Optional<ExcelBorderStyle> explicitStyle = explicitSide.flatMap(BorderSidePatch::style);
    return explicitStyle.isPresent() ? explicitStyle : defaultSide.flatMap(BorderSidePatch::style);
  }

  private Optional<ColorSpec> mergedBorderColor(
      Optional<BorderSidePatch> defaultSide, Optional<BorderSidePatch> explicitSide) {
    Objects.requireNonNull(defaultSide, "defaultSide must not be null");
    Objects.requireNonNull(explicitSide, "explicitSide must not be null");
    if (explicitSide.isPresent() && explicitSide.orElseThrow().color().isPresent()) {
      return explicitSide.orElseThrow().color();
    }
    return defaultSide.isEmpty() || defaultSide.orElseThrow().color().isEmpty()
        ? Optional.empty()
        : defaultSide.orElseThrow().color();
  }

  private Optional<BorderSidePatch> effectiveBorderSide(
      Optional<ExcelBorderStyle> style, Optional<ColorSpec> color) {
    Objects.requireNonNull(style, "style must not be null");
    Objects.requireNonNull(color, "color must not be null");
    if (style.isEmpty() && color.isEmpty()) {
      return Optional.empty();
    }
    if (color.isPresent() && (style.isEmpty() || style.orElseThrow() == ExcelBorderStyle.NONE)) {
      throw new IllegalArgumentException("border side color requires an effective border style");
    }
    return Optional.of(new BorderSidePatch(style, color));
  }

  private void applyBorderSidePatch(
      Optional<BorderSidePatch> sidePatch,
      Consumer<BorderStyle> styleSetter,
      Consumer<XSSFColor> colorSetter) {
    Objects.requireNonNull(sidePatch, "sidePatch must not be null");
    if (sidePatch.isEmpty()) {
      return;
    }
    BorderSidePatch resolved = sidePatch.orElseThrow();
    styleSetter.accept(ExcelCellStylePoiBridge.toPoi(resolved.style().orElseThrow()));
    if (resolved.style().orElseThrow() == ExcelBorderStyle.NONE) {
      return;
    }
    if (resolved.color().isPresent()) {
      colorSetter.accept(ExcelColorSupport.toXssfColor(workbook, resolved.color().orElseThrow()));
    }
  }
}
