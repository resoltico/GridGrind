package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Factual criterion family loaded from one autofilter filter column. */
public sealed interface ExcelAutofilterFilterCriterionSnapshot
    permits ExcelAutofilterFilterCriterionSnapshot.Values,
        ExcelAutofilterFilterCriterionSnapshot.Custom,
        ExcelAutofilterFilterCriterionSnapshot.Dynamic,
        ExcelAutofilterFilterCriterionSnapshot.Top10,
        ExcelAutofilterFilterCriterionSnapshot.Color,
        ExcelAutofilterFilterCriterionSnapshot.Icon {

  /** One explicit filter-values list, optionally including blanks. */
  record Values(List<String> values, boolean includeBlank)
      implements ExcelAutofilterFilterCriterionSnapshot {
    public Values {
      values = List.copyOf(Objects.requireNonNull(values, "values must not be null"));
      for (String value : values) {
        Objects.requireNonNull(value, "values must not contain null values");
      }
    }
  }

  /** One ordered custom-filter predicate set. */
  record Custom(boolean and, List<CustomCondition> conditions)
      implements ExcelAutofilterFilterCriterionSnapshot {
    public Custom {
      conditions = List.copyOf(Objects.requireNonNull(conditions, "conditions must not be null"));
      if (conditions.isEmpty()) {
        throw new IllegalArgumentException("conditions must not be empty");
      }
      for (CustomCondition condition : conditions) {
        Objects.requireNonNull(condition, "conditions must not contain null values");
      }
    }
  }

  /** One dynamic-filter rule. */
  record Dynamic(String type, Double value, Double maxValue)
      implements ExcelAutofilterFilterCriterionSnapshot {
    public Dynamic {
      Objects.requireNonNull(type, "type must not be null");
      if (type.isBlank()) {
        throw new IllegalArgumentException("type must not be blank");
      }
      requireFiniteOrNull(value, "value");
      requireFiniteOrNull(maxValue, "maxValue");
    }
  }

  /** One top-10 filter rule. */
  record Top10(boolean top, boolean percent, double value, Double filterValue)
      implements ExcelAutofilterFilterCriterionSnapshot {
    public Top10 {
      if (!Double.isFinite(value) || value < 0.0d) {
        throw new IllegalArgumentException("value must be finite and non-negative");
      }
      requireFiniteOrNull(filterValue, "filterValue");
    }
  }

  /** One color-based filter rule. */
  record Color(boolean cellColor, ExcelColorSnapshot color)
      implements ExcelAutofilterFilterCriterionSnapshot {
    public Color {}
  }

  /** One icon-based filter rule. */
  record Icon(String iconSet, int iconId) implements ExcelAutofilterFilterCriterionSnapshot {
    public Icon {
      Objects.requireNonNull(iconSet, "iconSet must not be null");
      if (iconSet.isBlank()) {
        throw new IllegalArgumentException("iconSet must not be blank");
      }
      if (iconId < 0) {
        throw new IllegalArgumentException("iconId must not be negative");
      }
    }
  }

  /** One custom-filter comparison inside a custom-filter criterion family. */
  record CustomCondition(String operator, String value) {
    public CustomCondition {
      Objects.requireNonNull(operator, "operator must not be null");
      Objects.requireNonNull(value, "value must not be null");
      if (operator.isBlank()) {
        throw new IllegalArgumentException("operator must not be blank");
      }
      if (value.isBlank()) {
        throw new IllegalArgumentException("value must not be blank");
      }
    }
  }

  private static void requireFiniteOrNull(Double value, String fieldName) {
    if (value != null && !Double.isFinite(value)) {
      throw new IllegalArgumentException(fieldName + " must be finite when provided");
    }
  }
}
