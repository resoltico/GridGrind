package dev.erst.gridgrind.cli.render;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Objects;

/** Validation helpers for CLI-owned discovery metadata. */
final class HelpValidation {
  private HelpValidation() {}

  static List<String> copyStrings(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<String> copy = new java.util.ArrayList<>(values.size());
    for (String value : values) {
      copy.add(TextValidation.requireNonBlank(value, fieldName));
    }
    return List.copyOf(copy);
  }
}
