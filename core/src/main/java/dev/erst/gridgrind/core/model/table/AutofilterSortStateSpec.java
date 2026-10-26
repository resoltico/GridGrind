package dev.erst.gridgrind.core.model.table;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.vocabulary.ExcelAutofilterSortMethod;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Authored sort-state payload nested under sheet autofilter authoring. */
@CatalogType(
    id = "AutofilterSortStateSpec",
    summary = "Authored autofilter sort-state payload with one or more ordered sort conditions.",
    displayOrder = 57)
public record AutofilterSortStateSpec(
    RangeRef range,
    @ProtocolField(optional = true) @WireDefault("false") boolean caseSensitive,
    @ProtocolField(optional = true) @WireDefault("false") boolean columnSort,
    Optional<ExcelAutofilterSortMethod> sortMethod,
    List<AutofilterSortConditionSpec> conditions) {
  /** Creates a sort-state payload with no explicit sort-method override. */
  public static AutofilterSortStateSpec withoutSortMethod(
      RangeRef range,
      boolean caseSensitive,
      boolean columnSort,
      List<AutofilterSortConditionSpec> conditions) {
    return new AutofilterSortStateSpec(
        range, caseSensitive, columnSort, Optional.empty(), conditions);
  }

  public AutofilterSortStateSpec {
    Objects.requireNonNull(range, "range must not be null");
    Optional<ExcelAutofilterSortMethod> normalizedSortMethod =
        Objects.requireNonNull(sortMethod, "sortMethod must not be null");
    sortMethod = normalizedSortMethod;
    conditions = copyConditions(conditions);
    if (conditions.isEmpty()) {
      throw new IllegalArgumentException("conditions must not be empty");
    }
  }

  private static List<AutofilterSortConditionSpec> copyConditions(
      List<AutofilterSortConditionSpec> conditions) {
    Objects.requireNonNull(conditions, "conditions must not be null");
    List<AutofilterSortConditionSpec> copiedConditions = new ArrayList<>(conditions.size());
    for (AutofilterSortConditionSpec condition : conditions) {
      copiedConditions.add(
          Objects.requireNonNull(condition, "conditions must not contain null values"));
    }
    return List.copyOf(copiedConditions);
  }
}
