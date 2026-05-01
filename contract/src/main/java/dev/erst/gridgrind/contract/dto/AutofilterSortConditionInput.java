package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** One authored autofilter sort condition expressed as an explicit variant. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = AutofilterSortConditionInput.Value.class, name = "VALUE"),
  @JsonSubTypes.Type(value = AutofilterSortConditionInput.CellColor.class, name = "CELL_COLOR"),
  @JsonSubTypes.Type(value = AutofilterSortConditionInput.FontColor.class, name = "FONT_COLOR"),
  @JsonSubTypes.Type(value = AutofilterSortConditionInput.Icon.class, name = "ICON")
})
public sealed interface AutofilterSortConditionInput
    permits AutofilterSortConditionInput.Value,
        AutofilterSortConditionInput.CellColor,
        AutofilterSortConditionInput.FontColor,
        AutofilterSortConditionInput.Icon {
  /** Worksheet-local range covered by this sort condition. */
  String range();

  /** Whether the condition sorts descending instead of ascending. */
  boolean descending();

  /** Sorts by the ordinary cell value with no auxiliary discriminator payload. */
  record Value(String range, boolean descending) implements AutofilterSortConditionInput {
    public Value {
      range = requireRange(range);
    }
  }

  /** Sorts by the cell fill color referenced by one explicit RGB/theme color. */
  record CellColor(String range, boolean descending, ColorInput color)
      implements AutofilterSortConditionInput {
    public CellColor {
      range = requireRange(range);
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  /** Sorts by the rendered font color referenced by one explicit RGB/theme color. */
  record FontColor(String range, boolean descending, ColorInput color)
      implements AutofilterSortConditionInput {
    public FontColor {
      range = requireRange(range);
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  /** Sorts by the icon identifier inside the active icon-set definition. */
  record Icon(String range, boolean descending, int iconId)
      implements AutofilterSortConditionInput {
    public Icon {
      range = requireRange(range);
      if (iconId < 0) {
        throw new IllegalArgumentException("iconId must not be negative");
      }
    }
  }

  private static String requireRange(String range) {
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    return range;
  }
}
