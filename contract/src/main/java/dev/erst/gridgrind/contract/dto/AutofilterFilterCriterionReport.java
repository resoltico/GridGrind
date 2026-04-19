package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Factual criterion family loaded from one autofilter filter column. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = AutofilterFilterCriterionReport.Values.class, name = "VALUES"),
  @JsonSubTypes.Type(value = AutofilterFilterCriterionReport.Custom.class, name = "CUSTOM"),
  @JsonSubTypes.Type(value = AutofilterFilterCriterionReport.Dynamic.class, name = "DYNAMIC"),
  @JsonSubTypes.Type(value = AutofilterFilterCriterionReport.Top10.class, name = "TOP10"),
  @JsonSubTypes.Type(value = AutofilterFilterCriterionReport.Color.class, name = "COLOR"),
  @JsonSubTypes.Type(value = AutofilterFilterCriterionReport.Icon.class, name = "ICON")
})
public sealed interface AutofilterFilterCriterionReport
    permits AutofilterFilterCriterionReport.Values,
        AutofilterFilterCriterionReport.Custom,
        AutofilterFilterCriterionReport.Dynamic,
        AutofilterFilterCriterionReport.Top10,
        AutofilterFilterCriterionReport.Color,
        AutofilterFilterCriterionReport.Icon {

  /** One explicit filter-values list, optionally including blanks. */
  record Values(List<String> values, boolean includeBlank)
      implements AutofilterFilterCriterionReport {
    public Values {
      values = copyValues(values, "values");
    }
  }

  /** One ordered custom-filter predicate set. */
  record Custom(boolean and, List<CustomConditionReport> conditions)
      implements AutofilterFilterCriterionReport {
    public Custom {
      conditions = copyValues(conditions, "conditions");
      if (conditions.isEmpty()) {
        throw new IllegalArgumentException("conditions must not be empty");
      }
    }
  }

  /** One dynamic-filter rule. */
  record Dynamic(String type, Double value, Double maxValue)
      implements AutofilterFilterCriterionReport {
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
      implements AutofilterFilterCriterionReport {
    public Top10 {
      if (!Double.isFinite(value) || value < 0.0d) {
        throw new IllegalArgumentException("value must be finite and non-negative");
      }
      requireFiniteOrNull(filterValue, "filterValue");
    }
  }

  /** One color-based filter rule. */
  record Color(boolean cellColor, CellColorReport color)
      implements AutofilterFilterCriterionReport {
    public Color {}
  }

  /** One icon-based filter rule. */
  record Icon(String iconSet, int iconId) implements AutofilterFilterCriterionReport {
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
  record CustomConditionReport(String operator, String value) {
    public CustomConditionReport {
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

  private static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain null values"));
    }
    return List.copyOf(copy);
  }
}
