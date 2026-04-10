package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Normalizes XMLBeans sqref payloads into stable trimmed string ranges. */
final class ExcelSqrefSupport {
  private ExcelSqrefSupport() {}

  static List<String> normalizedSqref(List<?> sqref) {
    Objects.requireNonNull(sqref, "sqref must not be null");

    List<String> normalized = new ArrayList<>(sqref.size());
    for (Object value : sqref) {
      String text = Objects.toString(value, "").trim();
      if (!text.isEmpty()) {
        normalized.add(text);
      }
    }
    return List.copyOf(normalized);
  }
}
