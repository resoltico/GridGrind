package dev.erst.gridgrind.core.model.table;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
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
  RangeLabel range();

  /** Whether the condition sorts descending instead of ascending. */
  boolean descending();

  /** Reports one ordinary value-based sort condition. */
  record Value(RangeLabel range, boolean descending) implements AutofilterSortConditionReport {
    public Value {
      java.util.Objects.requireNonNull(range, "range must not be null");
    }
  }

  /** Reports one cell-fill-color sort condition. */
  record CellColor(RangeLabel range, boolean descending, CellColorReport color)
      implements AutofilterSortConditionReport {
    public CellColor {
      java.util.Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  /** Reports one font-color sort condition. */
  record FontColor(RangeLabel range, boolean descending, CellColorReport color)
      implements AutofilterSortConditionReport {
    public FontColor {
      java.util.Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  /** Reports one icon-based sort condition. */
  record Icon(RangeLabel range, boolean descending, int iconId)
      implements AutofilterSortConditionReport {
    public Icon {
      java.util.Objects.requireNonNull(range, "range must not be null");
      if (iconId < 0) {
        throw new IllegalArgumentException("iconId must not be negative");
      }
    }
  }
}
