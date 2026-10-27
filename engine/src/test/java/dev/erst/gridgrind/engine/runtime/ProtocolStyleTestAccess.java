package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellGradientFillReport;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;

/** Test-only helpers for inspecting sealed protocol style variants. */
final class ProtocolStyleTestAccess {
  private ProtocolStyleTestAccess() {}

  static ExcelFillPattern fillPattern(CellFillReport fill) {
    return switch (fill) {
      case CellFillReport.PatternOnly pattern -> pattern.pattern();
      case CellFillReport.PatternForeground pattern -> pattern.pattern();
      case CellFillReport.PatternBackground pattern -> pattern.pattern();
      case CellFillReport.PatternForegroundBackground pattern -> pattern.pattern();
      case CellFillReport.Gradient ignored -> ExcelFillPattern.NONE;
    };
  }

  static CellColorReport fillForegroundColor(CellFillReport fill) {
    return switch (fill) {
      case CellFillReport.PatternForeground pattern -> pattern.foregroundColor();
      case CellFillReport.PatternForegroundBackground pattern -> pattern.foregroundColor();
      default -> null;
    };
  }

  static CellColorReport fillBackgroundColor(CellFillReport fill) {
    return switch (fill) {
      case CellFillReport.PatternBackground pattern -> pattern.backgroundColor();
      case CellFillReport.PatternForegroundBackground pattern -> pattern.backgroundColor();
      default -> null;
    };
  }

  static CellGradientFillReport fillGradient(CellFillReport fill) {
    if (fill instanceof CellFillReport.Gradient gradient) {
      return gradient.gradient();
    }
    return null;
  }
}
