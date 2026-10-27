package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFillSpec;
import dev.erst.gridgrind.core.model.style.CellGradientFillReport;
import dev.erst.gridgrind.core.model.style.CellGradientFillSpec;
import dev.erst.gridgrind.core.model.style.ColorSpec;
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

  static ExcelFillPattern fillPattern(CellFillSpec fill) {
    return switch (fill) {
      case CellFillSpec.PatternOnly pattern -> pattern.pattern();
      case CellFillSpec.PatternForeground pattern -> pattern.pattern();
      case CellFillSpec.PatternBackground pattern -> pattern.pattern();
      case CellFillSpec.PatternForegroundBackground pattern -> pattern.pattern();
      case CellFillSpec.Gradient ignored -> ExcelFillPattern.NONE;
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

  static ColorSpec fillForegroundColor(CellFillSpec fill) {
    return switch (fill) {
      case CellFillSpec.PatternForeground pattern -> pattern.foregroundColor();
      case CellFillSpec.PatternForegroundBackground pattern -> pattern.foregroundColor();
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

  static ColorSpec fillBackgroundColor(CellFillSpec fill) {
    return switch (fill) {
      case CellFillSpec.PatternBackground pattern -> pattern.backgroundColor();
      case CellFillSpec.PatternForegroundBackground pattern -> pattern.backgroundColor();
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

  static CellGradientFillSpec fillGradient(CellFillSpec fill) {
    if (fill instanceof CellFillSpec.Gradient gradient) {
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

  static String gradientType(CellGradientFillSpec gradient) {
    return switch (gradient) {
      case CellGradientFillSpec.Linear ignored -> "LINEAR";
      case CellGradientFillSpec.Path ignored -> "PATH";
    };
  }

  static String gradientType(CellGradientFillReport gradient) {
    return switch (gradient) {
      case CellGradientFillReport.Linear ignored -> "LINEAR";
      case CellGradientFillReport.Path ignored -> "PATH";
    };
  }

  static Double gradientDegree(CellGradientFillSpec gradient) {
    return switch (gradient) {
      case CellGradientFillSpec.Linear linear -> linear.degree().orElse(null);
      case CellGradientFillSpec.Path ignored -> null;
    };
  }

  static Double gradientDegree(CellGradientFillReport gradient) {
    return switch (gradient) {
      case CellGradientFillReport.Linear linear -> linear.degree().orElse(null);
      case CellGradientFillReport.Path ignored -> null;
    };
  }

  static Double gradientLeft(CellGradientFillSpec gradient) {
    return switch (gradient) {
      case CellGradientFillSpec.Path path -> path.left().orElse(null);
      case CellGradientFillSpec.Linear ignored -> null;
    };
  }

  static Double gradientLeft(CellGradientFillReport gradient) {
    return switch (gradient) {
      case CellGradientFillReport.Path path -> path.left().orElse(null);
      case CellGradientFillReport.Linear ignored -> null;
    };
  }

  static Double gradientRight(CellGradientFillSpec gradient) {
    return switch (gradient) {
      case CellGradientFillSpec.Path path -> path.right().orElse(null);
      case CellGradientFillSpec.Linear ignored -> null;
    };
  }

  static Double gradientRight(CellGradientFillReport gradient) {
    return switch (gradient) {
      case CellGradientFillReport.Path path -> path.right().orElse(null);
      case CellGradientFillReport.Linear ignored -> null;
    };
  }

  static Double gradientTop(CellGradientFillSpec gradient) {
    return switch (gradient) {
      case CellGradientFillSpec.Path path -> path.top().orElse(null);
      case CellGradientFillSpec.Linear ignored -> null;
    };
  }

  static Double gradientTop(CellGradientFillReport gradient) {
    return switch (gradient) {
      case CellGradientFillReport.Path path -> path.top().orElse(null);
      case CellGradientFillReport.Linear ignored -> null;
    };
  }

  static Double gradientBottom(CellGradientFillSpec gradient) {
    return switch (gradient) {
      case CellGradientFillSpec.Path path -> path.bottom().orElse(null);
      case CellGradientFillSpec.Linear ignored -> null;
    };
  }

  static Double gradientBottom(CellGradientFillReport gradient) {
    return switch (gradient) {
      case CellGradientFillReport.Path path -> path.bottom().orElse(null);
      case CellGradientFillReport.Linear ignored -> null;
    };
  }
}
