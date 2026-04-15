package dev.erst.gridgrind.excel;

import java.util.Locale;

/** Shared normalization and geometry rules for authored gradient fills. */
public final class ExcelGradientFillGeometry {
  private ExcelGradientFillGeometry() {}

  /** Normalizes one authored gradient type to GridGrind's canonical upper-case form. */
  public static String normalizeType(String type) {
    String normalizedType = type == null ? "LINEAR" : type.strip().toUpperCase(Locale.ROOT);
    if (normalizedType.isBlank()) {
      throw new IllegalArgumentException("type must not be blank");
    }
    if (!"LINEAR".equals(normalizedType) && !"PATH".equals(normalizedType)) {
      throw new IllegalArgumentException("type must be LINEAR or PATH");
    }
    return normalizedType;
  }

  /** Returns the effective factual gradient type after considering path-offset geometry. */
  public static String effectiveType(
      String explicitType, Double left, Double right, Double top, Double bottom) {
    if (explicitType != null) {
      return normalizeType(explicitType);
    }
    return hasPathOffsets(left, right, top, bottom) ? "PATH" : "LINEAR";
  }

  /** Rejects authored geometry that mixes the linear and path gradient models. */
  public static void requireCompatibleGeometry(
      String type, Double degree, Double left, Double right, Double top, Double bottom) {
    if ("LINEAR".equals(type) && hasPathOffsets(left, right, top, bottom)) {
      throw new IllegalArgumentException(
          "LINEAR gradients do not accept left, right, top, or bottom offsets");
    }
    if ("PATH".equals(type) && degree != null) {
      throw new IllegalArgumentException("PATH gradients do not accept degree");
    }
  }

  private static boolean hasPathOffsets(Double left, Double right, Double top, Double bottom) {
    return left != null || right != null || top != null || bottom != null;
  }
}
