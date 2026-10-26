package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellGradientFillReport;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;

/** Test-only helpers for inspecting sealed engine style variants. */
final class ExcelStyleTestAccess {
  private ExcelStyleTestAccess() {}

  static dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle borderStyle(
      dev.erst.gridgrind.core.model.style.CellBorderSideReport side) {
    return switch (side) {
      case dev.erst.gridgrind.core.model.style.CellBorderSideReport.None ignored ->
          dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle.NONE;
      case dev.erst.gridgrind.core.model.style.CellBorderSideReport.DefaultColor border ->
          border.style();
      case dev.erst.gridgrind.core.model.style.CellBorderSideReport.Colored border ->
          border.style();
    };
  }

  static dev.erst.gridgrind.core.model.style.CellColorReport borderColor(
      dev.erst.gridgrind.core.model.style.CellBorderSideReport side) {
    return side instanceof dev.erst.gridgrind.core.model.style.CellBorderSideReport.Colored colored
        ? colored.color()
        : null;
  }

  static ExcelFillPattern fillPattern(ExcelCellFill fill) {
    return switch (fill) {
      case ExcelCellFill.PatternOnly pattern -> pattern.pattern();
      case ExcelCellFill.PatternForeground pattern -> pattern.pattern();
      case ExcelCellFill.PatternBackground pattern -> pattern.pattern();
      case ExcelCellFill.PatternForegroundBackground pattern -> pattern.pattern();
      case ExcelCellFill.Gradient ignored -> ExcelFillPattern.NONE;
    };
  }

  static ExcelFillPattern fillPattern(CellFillReport fill) {
    return switch (fill) {
      case CellFillReport.PatternOnly pattern -> pattern.pattern();
      case CellFillReport.PatternForeground pattern -> pattern.pattern();
      case CellFillReport.PatternBackground pattern -> pattern.pattern();
      case CellFillReport.PatternForegroundBackground pattern -> pattern.pattern();
      case CellFillReport.Gradient ignored -> ExcelFillPattern.NONE;
    };
  }

  static ExcelColor fillForegroundColor(ExcelCellFill fill) {
    return switch (fill) {
      case ExcelCellFill.PatternForeground pattern -> pattern.foregroundColor();
      case ExcelCellFill.PatternForegroundBackground pattern -> pattern.foregroundColor();
      default -> null;
    };
  }

  static CellColorReport fillForegroundColor(CellFillReport fill) {
    return switch (fill) {
      case CellFillReport.PatternForeground pattern -> pattern.foregroundColor();
      case CellFillReport.PatternForegroundBackground pattern -> pattern.foregroundColor();
      default -> null;
    };
  }

  static ExcelColor fillBackgroundColor(ExcelCellFill fill) {
    return switch (fill) {
      case ExcelCellFill.PatternBackground pattern -> pattern.backgroundColor();
      case ExcelCellFill.PatternForegroundBackground pattern -> pattern.backgroundColor();
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

  static ExcelGradientFill fillGradient(ExcelCellFill fill) {
    if (fill instanceof ExcelCellFill.Gradient gradient) {
      return gradient.gradient();
    }
    return null;
  }

  static CellGradientFillReport fillGradient(CellFillReport fill) {
    if (fill instanceof CellFillReport.Gradient gradient) {
      return gradient.gradient();
    }
    return null;
  }

  static String gradientType(ExcelGradientFill gradient) {
    return switch (gradient) {
      case ExcelGradientFill.Linear ignored -> "LINEAR";
      case ExcelGradientFill.Path ignored -> "PATH";
    };
  }

  static String gradientType(CellGradientFillReport gradient) {
    return switch (gradient) {
      case CellGradientFillReport.Linear ignored -> "LINEAR";
      case CellGradientFillReport.Path ignored -> "PATH";
    };
  }

  static Double gradientDegree(ExcelGradientFill gradient) {
    return switch (gradient) {
      case ExcelGradientFill.Linear linear -> linear.degree().orElse(null);
      case ExcelGradientFill.Path ignored -> null;
    };
  }

  static Double gradientDegree(CellGradientFillReport gradient) {
    return switch (gradient) {
      case CellGradientFillReport.Linear linear -> linear.degree();
      case CellGradientFillReport.Path ignored -> null;
    };
  }

  static Double gradientLeft(ExcelGradientFill gradient) {
    return switch (gradient) {
      case ExcelGradientFill.Path path -> path.left().orElse(null);
      case ExcelGradientFill.Linear ignored -> null;
    };
  }

  static Double gradientLeft(CellGradientFillReport gradient) {
    return switch (gradient) {
      case CellGradientFillReport.Path path -> path.left();
      case CellGradientFillReport.Linear ignored -> null;
    };
  }

  static Double gradientRight(ExcelGradientFill gradient) {
    return switch (gradient) {
      case ExcelGradientFill.Path path -> path.right().orElse(null);
      case ExcelGradientFill.Linear ignored -> null;
    };
  }

  static Double gradientRight(CellGradientFillReport gradient) {
    return switch (gradient) {
      case CellGradientFillReport.Path path -> path.right();
      case CellGradientFillReport.Linear ignored -> null;
    };
  }

  static Double gradientTop(ExcelGradientFill gradient) {
    return switch (gradient) {
      case ExcelGradientFill.Path path -> path.top().orElse(null);
      case ExcelGradientFill.Linear ignored -> null;
    };
  }

  static Double gradientTop(CellGradientFillReport gradient) {
    return switch (gradient) {
      case CellGradientFillReport.Path path -> path.top();
      case CellGradientFillReport.Linear ignored -> null;
    };
  }

  static Double gradientBottom(ExcelGradientFill gradient) {
    return switch (gradient) {
      case ExcelGradientFill.Path path -> path.bottom().orElse(null);
      case ExcelGradientFill.Linear ignored -> null;
    };
  }

  static Double gradientBottom(CellGradientFillReport gradient) {
    return switch (gradient) {
      case CellGradientFillReport.Path path -> path.bottom();
      case CellGradientFillReport.Linear ignored -> null;
    };
  }
}
