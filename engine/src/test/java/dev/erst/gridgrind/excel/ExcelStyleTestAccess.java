package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;

/** Test-only helpers for inspecting sealed engine style variants. */
final class ExcelStyleTestAccess {
  private ExcelStyleTestAccess() {}

  static ExcelFillPattern fillPattern(ExcelCellFill fill) {
    return switch (fill) {
      case ExcelCellFill.PatternOnly pattern -> pattern.pattern();
      case ExcelCellFill.PatternForeground pattern -> pattern.pattern();
      case ExcelCellFill.PatternBackground pattern -> pattern.pattern();
      case ExcelCellFill.PatternForegroundBackground pattern -> pattern.pattern();
      case ExcelCellFill.Gradient ignored -> ExcelFillPattern.NONE;
    };
  }

  static ExcelFillPattern fillPattern(ExcelCellFillSnapshot fill) {
    return switch (fill) {
      case ExcelCellFillSnapshot.PatternOnly pattern -> pattern.pattern();
      case ExcelCellFillSnapshot.PatternForeground pattern -> pattern.pattern();
      case ExcelCellFillSnapshot.PatternBackground pattern -> pattern.pattern();
      case ExcelCellFillSnapshot.PatternForegroundBackground pattern -> pattern.pattern();
      case ExcelCellFillSnapshot.Gradient ignored -> ExcelFillPattern.NONE;
    };
  }

  static ExcelColor fillForegroundColor(ExcelCellFill fill) {
    return switch (fill) {
      case ExcelCellFill.PatternForeground pattern -> pattern.foregroundColor();
      case ExcelCellFill.PatternForegroundBackground pattern -> pattern.foregroundColor();
      default -> null;
    };
  }

  static ExcelColorSnapshot fillForegroundColor(ExcelCellFillSnapshot fill) {
    return switch (fill) {
      case ExcelCellFillSnapshot.PatternForeground pattern -> pattern.foregroundColor();
      case ExcelCellFillSnapshot.PatternForegroundBackground pattern -> pattern.foregroundColor();
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

  static ExcelColorSnapshot fillBackgroundColor(ExcelCellFillSnapshot fill) {
    return switch (fill) {
      case ExcelCellFillSnapshot.PatternBackground pattern -> pattern.backgroundColor();
      case ExcelCellFillSnapshot.PatternForegroundBackground pattern -> pattern.backgroundColor();
      default -> null;
    };
  }

  static ExcelGradientFill fillGradient(ExcelCellFill fill) {
    if (fill instanceof ExcelCellFill.Gradient gradient) {
      return gradient.gradient();
    }
    return null;
  }

  static ExcelGradientFillSnapshot fillGradient(ExcelCellFillSnapshot fill) {
    if (fill instanceof ExcelCellFillSnapshot.Gradient gradient) {
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

  static String gradientType(ExcelGradientFillSnapshot gradient) {
    return switch (gradient) {
      case ExcelGradientFillSnapshot.Linear ignored -> "LINEAR";
      case ExcelGradientFillSnapshot.Path ignored -> "PATH";
    };
  }

  static Double gradientDegree(ExcelGradientFill gradient) {
    return switch (gradient) {
      case ExcelGradientFill.Linear linear -> linear.degree();
      case ExcelGradientFill.Path ignored -> null;
    };
  }

  static Double gradientDegree(ExcelGradientFillSnapshot gradient) {
    return switch (gradient) {
      case ExcelGradientFillSnapshot.Linear linear -> linear.degree();
      case ExcelGradientFillSnapshot.Path ignored -> null;
    };
  }

  static Double gradientLeft(ExcelGradientFill gradient) {
    return switch (gradient) {
      case ExcelGradientFill.Path path -> path.left();
      case ExcelGradientFill.Linear ignored -> null;
    };
  }

  static Double gradientLeft(ExcelGradientFillSnapshot gradient) {
    return switch (gradient) {
      case ExcelGradientFillSnapshot.Path path -> path.left();
      case ExcelGradientFillSnapshot.Linear ignored -> null;
    };
  }

  static Double gradientRight(ExcelGradientFill gradient) {
    return switch (gradient) {
      case ExcelGradientFill.Path path -> path.right();
      case ExcelGradientFill.Linear ignored -> null;
    };
  }

  static Double gradientRight(ExcelGradientFillSnapshot gradient) {
    return switch (gradient) {
      case ExcelGradientFillSnapshot.Path path -> path.right();
      case ExcelGradientFillSnapshot.Linear ignored -> null;
    };
  }

  static Double gradientTop(ExcelGradientFill gradient) {
    return switch (gradient) {
      case ExcelGradientFill.Path path -> path.top();
      case ExcelGradientFill.Linear ignored -> null;
    };
  }

  static Double gradientTop(ExcelGradientFillSnapshot gradient) {
    return switch (gradient) {
      case ExcelGradientFillSnapshot.Path path -> path.top();
      case ExcelGradientFillSnapshot.Linear ignored -> null;
    };
  }

  static Double gradientBottom(ExcelGradientFill gradient) {
    return switch (gradient) {
      case ExcelGradientFill.Path path -> path.bottom();
      case ExcelGradientFill.Linear ignored -> null;
    };
  }

  static Double gradientBottom(ExcelGradientFillSnapshot gradient) {
    return switch (gradient) {
      case ExcelGradientFillSnapshot.Path path -> path.bottom();
      case ExcelGradientFillSnapshot.Linear ignored -> null;
    };
  }
}
