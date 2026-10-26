package dev.erst.gridgrind.core.model.table;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import java.util.Objects;

/** One authored autofilter sort condition expressed as an explicit variant. */
@Union(discriminator = "type", catalog = true, displayOrder = 22)
public sealed interface AutofilterSortConditionSpec
    permits dev.erst.gridgrind.core.model.table.AutofilterSortConditionSpec.Value,
        dev.erst.gridgrind.core.model.table.AutofilterSortConditionSpec.CellColor,
        dev.erst.gridgrind.core.model.table.AutofilterSortConditionSpec.FontColor,
        dev.erst.gridgrind.core.model.table.AutofilterSortConditionSpec.Icon {
  /** Worksheet-local range covered by this sort condition. */
  RangeRef range();

  /** Whether the condition sorts descending instead of ascending. */
  boolean descending();

  /** Sorts by the ordinary cell value with no auxiliary discriminator payload. */
  @Variant(
      id = "VALUE",
      summary = "Sort by the ordinary cell value with no auxiliary discriminator payload.")
  record Value(RangeRef range, boolean descending) implements AutofilterSortConditionSpec {
    public Value {
      java.util.Objects.requireNonNull(range, "range must not be null");
    }
  }

  /** Sorts by the cell fill color referenced by one explicit RGB/theme color. */
  @Variant(
      id = "CELL_COLOR",
      summary = "Sort by the cell fill color referenced by one explicit color.")
  record CellColor(RangeRef range, boolean descending, ColorSpec color)
      implements AutofilterSortConditionSpec {
    public CellColor {
      java.util.Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  /** Sorts by the rendered font color referenced by one explicit RGB/theme color. */
  @Variant(
      id = "FONT_COLOR",
      summary = "Sort by the rendered font color referenced by one explicit color.")
  record FontColor(RangeRef range, boolean descending, ColorSpec color)
      implements AutofilterSortConditionSpec {
    public FontColor {
      java.util.Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  /** Sorts by the icon identifier inside the active icon-set definition. */
  @Variant(id = "ICON", summary = "Sort by the icon id inside the active icon-set definition.")
  record Icon(RangeRef range, boolean descending, int iconId)
      implements AutofilterSortConditionSpec {
    public Icon {
      java.util.Objects.requireNonNull(range, "range must not be null");
      if (iconId < 0) {
        throw new IllegalArgumentException("iconId must not be negative");
      }
    }
  }
}
