package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Authored autofilter criterion for a mutable workbook filter column. */
public sealed interface ExcelAutofilterFilterCriterion
    permits ExcelAutofilterFilterCriterion.Values,
        ExcelAutofilterFilterCriterion.Custom,
        ExcelAutofilterFilterCriterion.Dynamic,
        ExcelAutofilterFilterCriterion.Top10,
        ExcelAutofilterFilterCriterion.Color,
        ExcelAutofilterFilterCriterion.Icon {

  record Values(List<String> values, boolean includeBlank)
      implements ExcelAutofilterFilterCriterion {
    public Values {
      values = copyStrings(values, "values");
    }
  }

  record Custom(boolean and, List<CustomCondition> conditions)
      implements ExcelAutofilterFilterCriterion {
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

  record Dynamic(String type, Double value, Double maxValue)
      implements ExcelAutofilterFilterCriterion {
    public Dynamic {
      Objects.requireNonNull(type, "type must not be null");
      if (type.isBlank()) {
        throw new IllegalArgumentException("type must not be blank");
      }
      value = finiteOrNull(value, "value");
      maxValue = finiteOrNull(maxValue, "maxValue");
    }
  }

  record Top10(int value, boolean top, boolean percent) implements ExcelAutofilterFilterCriterion {
    public Top10 {
      if (value <= 0) {
        throw new IllegalArgumentException("value must be greater than 0");
      }
    }
  }

  record Color(boolean cellColor, ExcelColor color) implements ExcelAutofilterFilterCriterion {
    public Color {
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  record Icon(String iconSet, int iconId) implements ExcelAutofilterFilterCriterion {
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

  private static List<String> copyStrings(List<String> values, String fieldName) {
    List<String> copy =
        List.copyOf(Objects.requireNonNull(values, fieldName + " must not be null"));
    for (String value : copy) {
      Objects.requireNonNull(value, fieldName + " must not contain null values");
    }
    return copy;
  }

  private static Double finiteOrNull(Double value, String fieldName) {
    if (value == null) {
      return null;
    }
    if (!Double.isFinite(value)) {
      throw new IllegalArgumentException(fieldName + " must be finite");
    }
    return value;
  }
}
