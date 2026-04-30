package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** One factual autofilter sort condition reported from a workbook. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = AutofilterSortConditionReport.Value.class, name = "VALUE"),
  @JsonSubTypes.Type(value = AutofilterSortConditionReport.CellColor.class, name = "CELL_COLOR"),
  @JsonSubTypes.Type(value = AutofilterSortConditionReport.FontColor.class, name = "FONT_COLOR"),
  @JsonSubTypes.Type(value = AutofilterSortConditionReport.Icon.class, name = "ICON")
})
public sealed interface AutofilterSortConditionReport
    permits AutofilterSortConditionReport.Value,
        AutofilterSortConditionReport.CellColor,
        AutofilterSortConditionReport.FontColor,
        AutofilterSortConditionReport.Icon {
  /** Worksheet-local range covered by this sort condition. */
  String range();

  /** Whether the condition sorts descending instead of ascending. */
  boolean descending();

  /** Reports one ordinary value-based sort condition. */
  record Value(String range, boolean descending) implements AutofilterSortConditionReport {
    public Value {
      range = requireRange(range);
    }
  }

  /** Reports one cell-fill-color sort condition. */
  record CellColor(String range, boolean descending, CellColorReport color)
      implements AutofilterSortConditionReport {
    public CellColor {
      range = requireRange(range);
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  /** Reports one font-color sort condition. */
  record FontColor(String range, boolean descending, CellColorReport color)
      implements AutofilterSortConditionReport {
    public FontColor {
      range = requireRange(range);
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  /** Reports one icon-based sort condition. */
  record Icon(String range, boolean descending, int iconId)
      implements AutofilterSortConditionReport {
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
