package dev.erst.gridgrind.core.model.table;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.style.ColorInput;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Authored autofilter criterion for one filter column. */
@Union(discriminator = "type", catalog = true, displayOrder = 63)
public sealed interface AutofilterFilterCriterionInput
    permits dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionInput.Values,
        dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionInput.Custom,
        dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionInput.Dynamic,
        dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionInput.Top10,
        dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionInput.Color,
        dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionInput.Icon {

  /** Explicit retained filter values plus whether blanks are included. */
  @Variant(
      id = "VALUES",
      summary = "Retain rows whose cell values match one or more explicit values.")
  record Values(List<String> values, boolean includeBlank)
      implements AutofilterFilterCriterionInput {
    public Values {
      values = copyStrings(values, "values");
    }
  }

  /** Custom comparator-based autofilter criterion. */
  @Variant(
      id = "CUSTOM",
      summary = "Retain rows that satisfy one or two comparator-based custom conditions.")
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
  @Variant(
      id = "DYNAMIC",
      summary = "Retain rows using one dynamic-date or moving-window autofilter rule.")
  record Dynamic(
      String type,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Double> value,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Double> maxValue)
      implements AutofilterFilterCriterionInput {
    public Dynamic {
      Objects.requireNonNull(type, "type must not be null");
      if (type.isBlank()) {
        throw new IllegalArgumentException("type must not be blank");
      }
      Objects.requireNonNull(value, "value must not be null");
      Objects.requireNonNull(maxValue, "maxValue must not be null");
      value = value.map(entry -> finite(entry, "value"));
      maxValue = maxValue.map(entry -> finite(entry, "maxValue"));
    }
  }

  /** Top or bottom N or percent autofilter criterion. */
  @Variant(id = "TOP10", summary = "Retain top or bottom N or percent values.")
  record Top10(int value, boolean top, boolean percent) implements AutofilterFilterCriterionInput {
    public Top10 {
      if (value <= 0) {
        throw new IllegalArgumentException("value must be greater than 0");
      }
    }
  }

  /** Color-based autofilter criterion. */
  @Variant(id = "COLOR", summary = "Retain rows matching one cell color or font color criterion.")
  record Color(boolean cellColor, ColorInput color) implements AutofilterFilterCriterionInput {
    public Color {
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  /** Icon-based autofilter criterion. */
  @Variant(id = "ICON", summary = "Retain rows matching one icon-set member.")
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
  @CatalogType(
      id = "AutofilterCustomConditionInput",
      summary = "One comparator-value pair nested inside a custom autofilter criterion.",
      displayOrder = 55)
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

  private static Double finite(Double value, String fieldName) {
    if (!Double.isFinite(value)) {
      throw new IllegalArgumentException(fieldName + " must be finite");
    }
    return value;
  }
}
