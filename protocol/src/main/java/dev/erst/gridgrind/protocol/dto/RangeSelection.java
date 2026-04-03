package dev.erst.gridgrind.protocol.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Selects which A1-style ranges on one sheet a document-intelligence read should inspect. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = RangeSelection.All.class, name = "ALL"),
  @JsonSubTypes.Type(value = RangeSelection.Selected.class, name = "SELECTED")
})
public sealed interface RangeSelection permits RangeSelection.All, RangeSelection.Selected {

  /** Selects every matching structure on the sheet. */
  record All() implements RangeSelection {}

  /** Selects only structures whose ranges intersect the provided A1-style ranges. */
  record Selected(List<String> ranges) implements RangeSelection {
    public Selected {
      ranges = copyRanges(ranges);
    }
  }

  private static List<String> copyRanges(List<String> ranges) {
    Objects.requireNonNull(ranges, "ranges must not be null");
    List<String> copy = List.copyOf(ranges);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException("ranges must not be empty");
    }
    Set<String> unique = new LinkedHashSet<>();
    for (String range : copy) {
      requireNonBlank(range, "ranges");
      if (!unique.add(range)) {
        throw new IllegalArgumentException("ranges must not contain duplicates");
      }
    }
    return List.copyOf(copy);
  }

  private static void requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not contain nulls");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not contain blank values");
    }
  }
}
