package dev.erst.gridgrind.core.model.layout;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.ExcelPaneRegion;
import java.util.Objects;

/** Pane state captured from a sheet layout read. */
@Union(discriminator = "type")
public sealed interface PaneReport permits PaneReport.None, PaneReport.Frozen, PaneReport.Split {
  /** Sheet has no active pane split or freeze state. */
  @Variant(id = "NONE", summary = "Sheet has no active pane split or freeze state.")
  record None() implements PaneReport {}

  /** Sheet is frozen at the provided split and visible-origin coordinates. */
  @Variant(
      id = "FROZEN",
      summary = "Sheet is frozen at the provided split and visible-origin coordinates.")
  record Frozen(int splitColumn, int splitRow, int leftmostColumn, int topRow)
      implements PaneReport {}

  /** Sheet uses split panes with explicit split offsets, visible origin, and active pane. */
  @Variant(
      id = "SPLIT",
      summary =
          "Sheet uses split panes with explicit split offsets, visible origin, and active pane.")
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
