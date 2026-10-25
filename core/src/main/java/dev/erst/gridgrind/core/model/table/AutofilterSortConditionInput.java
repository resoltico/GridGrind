package dev.erst.gridgrind.core.model.table;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.core.model.style.ColorInput;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
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
  RangeRef range();

  /** Whether the condition sorts descending instead of ascending. */
  boolean descending();

  /** Sorts by the ordinary cell value with no auxiliary discriminator payload. */
  record Value(RangeRef range, boolean descending) implements AutofilterSortConditionInput {
    public Value {
      java.util.Objects.requireNonNull(range, "range must not be null");
    }
  }

  /** Sorts by the cell fill color referenced by one explicit RGB/theme color. */
  record CellColor(RangeRef range, boolean descending, ColorInput color)
      implements AutofilterSortConditionInput {
    public CellColor {
      java.util.Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  /** Sorts by the rendered font color referenced by one explicit RGB/theme color. */
  record FontColor(RangeRef range, boolean descending, ColorInput color)
      implements AutofilterSortConditionInput {
    public FontColor {
      java.util.Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  /** Sorts by the icon identifier inside the active icon-set definition. */
  record Icon(RangeRef range, boolean descending, int iconId)
      implements AutofilterSortConditionInput {
    public Icon {
      java.util.Objects.requireNonNull(range, "range must not be null");
      if (iconId < 0) {
        throw new IllegalArgumentException("iconId must not be negative");
      }
    }
  }
}
