package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.CellAlignmentReport;
import dev.erst.gridgrind.contract.dto.CellBorderReport;
import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellColorReport;
import dev.erst.gridgrind.contract.dto.CellFillReport;
import dev.erst.gridgrind.contract.dto.CellFontReport;
import dev.erst.gridgrind.contract.dto.CellGradientFillReport;
import dev.erst.gridgrind.contract.dto.CellGradientStopReport;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.CellStyleReport;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import org.jspecify.annotations.Nullable;

/** Owns invariant checks for cell style, color, border, gradient, and font payloads. */
final class WorkbookInvariantCellStyleChecks {
  private WorkbookInvariantCellStyleChecks() {}

  static void requireCellStyleShape(CellStyleReport style) {
    WorkbookInvariantChecks.require(style != null, "style must not be null");
    WorkbookInvariantChecks.require(style.numberFormat() != null, "numberFormat must not be null");
    requireCellAlignmentShape(style.alignment());
    requireCellFontShape(style.font());
    requireCellFillShape(style.fill());
    requireCellBorderShape(style.border());
    requireCellProtectionShape(style.protection());
  }

  static void requireCellAlignmentShape(CellAlignmentReport alignment) {
    WorkbookInvariantChecks.require(alignment != null, "alignment must not be null");
    WorkbookInvariantChecks.require(
        alignment.horizontalAlignment() != null, "horizontalAlignment must not be null");
    WorkbookInvariantChecks.require(
        alignment.verticalAlignment() != null, "verticalAlignment must not be null");
    WorkbookInvariantChecks.require(
        alignment.textRotation() >= 0 && alignment.textRotation() <= 180,
        "textRotation must be between 0 and 180 inclusive");
    WorkbookInvariantChecks.require(
        alignment.indentation() >= 0 && alignment.indentation() <= 250,
        "indentation must be between 0 and 250 inclusive");
  }

  static void requireCellFontShape(CellFontReport font) {
    WorkbookInvariantChecks.require(font != null, "font must not be null");
    WorkbookInvariantChecks.require(font.fontName() != null, "fontName must not be null");
    WorkbookInvariantChecks.require(!font.fontName().isBlank(), "fontName must not be blank");
    requireFontHeightShape(font.fontHeight());
    if (font.fontColor() != null) {
      requireCellColorShape(font.fontColor(), "fontColor");
    }
  }

  static void requireCellFillShape(CellFillReport fill) {
    WorkbookInvariantChecks.require(fill != null, "fill must not be null");
    switch (fill) {
      case CellFillReport.PatternOnly pattern -> {
        WorkbookInvariantChecks.require(pattern.pattern() != null, "fill pattern must not be null");
      }
      case CellFillReport.PatternForeground pattern -> {
        WorkbookInvariantChecks.require(pattern.pattern() != null, "fill pattern must not be null");
        requireCellColorShape(pattern.foregroundColor(), "fill foregroundColor");
      }
      case CellFillReport.PatternBackground pattern -> {
        WorkbookInvariantChecks.require(pattern.pattern() != null, "fill pattern must not be null");
        WorkbookInvariantChecks.require(
            pattern.pattern() != ExcelFillPattern.SOLID,
            "SOLID fills must not carry backgroundColor");
        requireCellColorShape(pattern.backgroundColor(), "fill backgroundColor");
      }
      case CellFillReport.PatternForegroundBackground pattern -> {
        WorkbookInvariantChecks.require(pattern.pattern() != null, "fill pattern must not be null");
        WorkbookInvariantChecks.require(
            pattern.pattern() != ExcelFillPattern.SOLID,
            "SOLID fills must not carry backgroundColor");
        requireCellColorShape(pattern.foregroundColor(), "fill foregroundColor");
        requireCellColorShape(pattern.backgroundColor(), "fill backgroundColor");
      }
      case CellFillReport.Gradient gradient -> requireCellGradientFillShape(gradient.gradient());
    }
  }

  static void requireCellBorderShape(CellBorderReport border) {
    WorkbookInvariantChecks.require(border != null, "border must not be null");
    requireCellBorderSideShape(border.top(), "top");
    requireCellBorderSideShape(border.right(), "right");
    requireCellBorderSideShape(border.bottom(), "bottom");
    requireCellBorderSideShape(border.left(), "left");
  }

  static void requireCellBorderSideShape(CellBorderSideReport side, String label) {
    WorkbookInvariantChecks.require(side != null, label + " border side must not be null");
    WorkbookInvariantChecks.require(side.style() != null, label + " border style must not be null");
    if (side.color() != null) {
      requireCellColorShape(side.color(), label + " border color");
    }
  }

  static void requireCellGradientFillShape(CellGradientFillReport gradient) {
    WorkbookInvariantChecks.require(gradient != null, "gradient fill must not be null");
    WorkbookInvariantChecks.require(
        gradient.stops() != null, "gradient fill stops must not be null");
    WorkbookInvariantChecks.require(
        !gradient.stops().isEmpty(), "gradient fill stops must not be empty");
    switch (gradient) {
      case CellGradientFillReport.Linear linear -> {
        if (linear.degree() != null) {
          WorkbookInvariantChecks.require(
              Double.isFinite(linear.degree()), "linear gradient degree must be finite");
        }
      }
      case CellGradientFillReport.Path path -> {
        requireFiniteOrNull(path.left(), "path gradient left");
        requireFiniteOrNull(path.right(), "path gradient right");
        requireFiniteOrNull(path.top(), "path gradient top");
        requireFiniteOrNull(path.bottom(), "path gradient bottom");
      }
    }
    for (CellGradientStopReport stop : gradient.stops()) {
      WorkbookInvariantChecks.require(stop != null, "gradient fill stop must not be null");
      WorkbookInvariantChecks.require(
          Double.isFinite(stop.position()) && stop.position() >= 0.0d && stop.position() <= 1.0d,
          "gradient fill stop position must be between 0.0 and 1.0");
      requireCellColorShape(stop.color(), "gradient fill stop color");
    }
  }

  static void requireCellColorShape(CellColorReport color, String label) {
    WorkbookInvariantChecks.require(color != null, label + " must not be null");
    switch (color) {
      case CellColorReport.Rgb rgb ->
          WorkbookInvariantChecks.requireNonBlank(rgb.rgb(), label + " rgb");
      case CellColorReport.Theme theme ->
          WorkbookInvariantChecks.require(
              theme.theme() >= 0, label + " theme must not be negative");
      case CellColorReport.Indexed indexed ->
          WorkbookInvariantChecks.require(
              indexed.indexed() >= 0, label + " indexed must not be negative");
    }
    color.tint().ifPresent(tint -> requireFiniteOrNull(tint, label + " tint"));
  }

  static void requireCellProtectionShape(CellProtectionReport protection) {
    WorkbookInvariantChecks.require(protection != null, "protection must not be null");
  }

  static void requireFontHeightShape(FontHeightReport fontHeight) {
    WorkbookInvariantChecks.require(fontHeight != null, "fontHeight must not be null");
    ExcelFontHeight expected = new ExcelFontHeight(fontHeight.twips());
    WorkbookInvariantChecks.require(
        expected.points().compareTo(fontHeight.points()) == 0,
        "fontHeight points must match twips");
  }

  private static void requireFiniteOrNull(@Nullable Double value, String label) {
    if (value != null) {
      WorkbookInvariantChecks.require(Double.isFinite(value), label + " must be finite");
    }
  }
}
