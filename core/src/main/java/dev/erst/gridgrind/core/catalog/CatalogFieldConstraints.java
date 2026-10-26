package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.metadata.FieldRule;
import java.util.Arrays;
import java.util.List;

/** Converts colocated scalar rules into the public constraint vocabulary. */
final class CatalogFieldConstraints {
  private CatalogFieldConstraints() {}

  static List<FieldConstraint> forComponent(Class<?> owner, String fieldName) {
    if (!owner.isRecord()) {
      return List.of();
    }
    return Arrays.stream(owner.getRecordComponents())
        .filter(component -> component.getName().equals(fieldName))
        .findFirst()
        .stream()
        .flatMap(component -> Arrays.stream(component.getAnnotationsByType(FieldRule.class)))
        .map(CatalogFieldConstraints::constraint)
        .toList();
  }

  private static FieldConstraint constraint(FieldRule rule) {
    return switch (rule.kind()) {
      case NON_BLANK -> new FieldConstraint.NonBlank();
      case STRING_PATTERN -> new FieldConstraint.StringPattern(rule.text());
      case LENGTH_RANGE ->
          new FieldConstraint.LengthRange((int) rule.minimum(), (int) rule.maximum());
      case NUMBER_RANGE -> new FieldConstraint.NumberRange(rule.minimum(), rule.maximum());
      case INTEGRAL -> new FieldConstraint.Integral();
      case PATH_SUFFIX -> new FieldConstraint.PathSuffix(rule.text());
    };
  }
}
