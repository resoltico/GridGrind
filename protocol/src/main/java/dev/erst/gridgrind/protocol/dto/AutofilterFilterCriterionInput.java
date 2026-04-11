package dev.erst.gridgrind.protocol.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.List;
import java.util.Objects;

/** Authored autofilter criterion for one filter column. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = AutofilterFilterCriterionInput.Values.class, name = "VALUES"),
  @JsonSubTypes.Type(value = AutofilterFilterCriterionInput.Custom.class, name = "CUSTOM"),
  @JsonSubTypes.Type(value = AutofilterFilterCriterionInput.Dynamic.class, name = "DYNAMIC"),
  @JsonSubTypes.Type(value = AutofilterFilterCriterionInput.Top10.class, name = "TOP10"),
  @JsonSubTypes.Type(value = AutofilterFilterCriterionInput.Color.class, name = "COLOR"),
  @JsonSubTypes.Type(value = AutofilterFilterCriterionInput.Icon.class, name = "ICON")
})
public sealed interface AutofilterFilterCriterionInput
    permits AutofilterFilterCriterionInput.Values,
        AutofilterFilterCriterionInput.Custom,
        AutofilterFilterCriterionInput.Dynamic,
        AutofilterFilterCriterionInput.Top10,
        AutofilterFilterCriterionInput.Color,
        AutofilterFilterCriterionInput.Icon {

  /** Explicit retained filter values plus whether blanks are included. */
  record Values(List<String> values, boolean includeBlank)
      implements AutofilterFilterCriterionInput {
    public Values {
      values = copyStrings(values, "values");
    }
  }

  /** Custom comparator-based autofilter criterion. */
  record Custom(boolean and, List<CustomConditionInput> conditions)
      implements AutofilterFilterCriterionInput {
    public Custom {
      conditions = List.copyOf(Objects.requireNonNull(conditions, "conditions must not be null"));
      if (conditions.isEmpty()) {
        throw new IllegalArgumentException("conditions must not be empty");
      }
      for (CustomConditionInput condition : conditions) {
        Objects.requireNonNull(condition, "conditions must not contain null values");
      }
    }
  }

  /** Dynamic-date or moving-window autofilter criterion. */
  record Dynamic(String type, Double value, Double maxValue)
      implements AutofilterFilterCriterionInput {
    public Dynamic {
      Objects.requireNonNull(type, "type must not be null");
      if (type.isBlank()) {
        throw new IllegalArgumentException("type must not be blank");
      }
      value = finiteOrNull(value, "value");
      maxValue = finiteOrNull(maxValue, "maxValue");
    }
  }

  /** Top or bottom N or percent autofilter criterion. */
  record Top10(int value, boolean top, boolean percent) implements AutofilterFilterCriterionInput {
    public Top10 {
      if (value <= 0) {
        throw new IllegalArgumentException("value must be greater than 0");
      }
    }
  }

  /** Color-based autofilter criterion. */
  record Color(boolean cellColor, ColorInput color) implements AutofilterFilterCriterionInput {
    public Color {
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  /** Icon-based autofilter criterion. */
  record Icon(String iconSet, int iconId) implements AutofilterFilterCriterionInput {
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

  /** One custom comparator inside a custom autofilter criterion. */
  record CustomConditionInput(String operator, String value) {
    public CustomConditionInput {
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
