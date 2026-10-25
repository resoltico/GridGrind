package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.foundation.ExcelPaneRegion;
import java.util.Objects;

/** Pane state captured from a sheet layout read. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = PaneReport.None.class, name = "NONE"),
  @JsonSubTypes.Type(value = PaneReport.Frozen.class, name = "FROZEN"),
  @JsonSubTypes.Type(value = PaneReport.Split.class, name = "SPLIT")
})
public sealed interface PaneReport permits PaneReport.None, PaneReport.Frozen, PaneReport.Split {
  /** Sheet has no active pane split or freeze state. */
  record None() implements PaneReport {}

  /** Sheet is frozen at the provided split and visible-origin coordinates. */
  record Frozen(int splitColumn, int splitRow, int leftmostColumn, int topRow)
      implements PaneReport {}

  /** Sheet uses split panes with explicit split offsets, visible origin, and active pane. */
  record Split(
      int xSplitPosition,
      int ySplitPosition,
      int leftmostColumn,
      int topRow,
      ExcelPaneRegion activePane)
      implements PaneReport {
    public Split {
      Objects.requireNonNull(activePane, "activePane must not be null");
    }
  }
}
