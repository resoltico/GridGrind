package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Selects which named ranges a read operation should inspect or analyze. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = NamedRangeSelection.All.class, name = "ALL"),
  @JsonSubTypes.Type(value = NamedRangeSelection.Selected.class, name = "SELECTED")
})
public sealed interface NamedRangeSelection
    permits NamedRangeSelection.All, NamedRangeSelection.Selected {

  /** Selects every user-facing named range in the workbook. */
  record All() implements NamedRangeSelection {}

  /** Selects only the named ranges matched by the provided exact selectors. */
  record Selected(List<NamedRangeSelector> selectors) implements NamedRangeSelection {
    public Selected {
      selectors = copySelectors(selectors);
    }
  }

  private static List<NamedRangeSelector> copySelectors(List<NamedRangeSelector> selectors) {
    Objects.requireNonNull(selectors, "selectors must not be null");
    List<NamedRangeSelector> copy = List.copyOf(selectors);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException("selectors must not be empty");
    }
    Set<NamedRangeSelector> unique = new LinkedHashSet<>();
    for (NamedRangeSelector selector : copy) {
      Objects.requireNonNull(selector, "selectors must not contain nulls");
      if (!unique.add(selector)) {
        throw new IllegalArgumentException("selectors must not contain duplicates");
      }
    }
    return List.copyOf(copy);
  }
}
