package dev.erst.gridgrind.core.model.table;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Factual criterion family loaded from one autofilter filter column. */
@Union(discriminator = "type")
public sealed interface AutofilterFilterCriterionReport
    permits AutofilterFilterCriterionReport.Values,
        AutofilterFilterCriterionReport.Custom,
        AutofilterFilterCriterionReport.Dynamic,
        AutofilterFilterCriterionReport.Top10,
        AutofilterFilterCriterionReport.Color,
        AutofilterFilterCriterionReport.Icon {

  /** One explicit filter-values list, optionally including blanks. */
  @Variant(id = "VALUES", summary = "One explicit filter-values list, optionally including blanks.")
  record Values(List<String> values, boolean includeBlank)
      implements AutofilterFilterCriterionReport {
    public Values {
      values = copyValues(values, "values");
    }
  }

  /** One ordered custom-filter predicate set. */
  @Variant(id = "CUSTOM", summary = "One ordered custom-filter predicate set.")
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
  @Variant(id = "DYNAMIC", summary = "One dynamic-filter rule.")
  record Dynamic(String dynamicType, Optional<Double> value, Optional<Double> maxValue)
      implements AutofilterFilterCriterionReport {
    public Dynamic {
      dynamicType = TextValidation.requireNonBlank(dynamicType, "dynamicType");
      requireFinite(value, "value");
      requireFinite(maxValue, "maxValue");
    }
  }

  /** One top-10 filter rule. */
  @Variant(id = "TOP10", summary = "One top-10 filter rule.")
  record Top10(boolean top, boolean percent, double value, Optional<Double> filterValue)
      implements AutofilterFilterCriterionReport {
    public Top10 {
      if (!Double.isFinite(value) || value < 0.0d) {
        throw new IllegalArgumentException("value must be finite and non-negative");
      }
      requireFinite(filterValue, "filterValue");
    }
  }

  /** One color-based filter rule. */
  @Variant(id = "COLOR", summary = "One color-based filter rule.")
  record Color(boolean cellColor, Optional<CellColorReport> color)
      implements AutofilterFilterCriterionReport {
    public Color {
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  /** One icon-based filter rule. */
  @Variant(id = "ICON", summary = "One icon-based filter rule.")
  record Icon(String iconSet, int iconId) implements AutofilterFilterCriterionReport {
    public Icon {
      iconSet = TextValidation.requireNonBlank(iconSet, "iconSet");
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
      TextValidation.requireNonBlank(operator, "operator");
      TextValidation.requireNonBlank(value, "value");
    }
  }

  private static void requireFinite(Optional<Double> value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.filter(number -> !Double.isFinite(number)).isPresent()) {
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
