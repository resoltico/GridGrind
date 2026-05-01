package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import java.util.Objects;

/** Sheet-scoped fluent target. */
public final class SheetTarget {
  private final SheetSelector.ByName selector;

  SheetTarget(SheetSelector.ByName selector) {
    this.selector = Objects.requireNonNull(selector, "selector must not be null");
  }

  SheetSelector.ByName selector() {
    return selector;
  }

  /** Returns one sheet-creation mutation step. */
  public PlannedMutation ensureExists() {
    return new PlannedMutation(selector, new WorkbookMutationAction.EnsureSheet());
  }

  /** Returns one sheet-rename mutation step. */
  public PlannedMutation renameTo(String newSheetName) {
    return new PlannedMutation(selector, new WorkbookMutationAction.RenameSheet(newSheetName));
  }

  /** Returns one sheet-delete mutation step. */
  public PlannedMutation delete() {
    return new PlannedMutation(selector, new WorkbookMutationAction.DeleteSheet());
  }

  /** Returns one sheet-zoom mutation step. */
  public PlannedMutation setZoom(int zoomPercent) {
    return new PlannedMutation(selector, new WorkbookMutationAction.SetSheetZoom(zoomPercent));
  }

  /** Returns one print-layout clearing mutation step. */
  public PlannedMutation clearPrintLayout() {
    return new PlannedMutation(selector, new WorkbookMutationAction.ClearPrintLayout());
  }

  /** Returns one sheet-summary inspection step. */
  public PlannedInspection summary() {
    return new PlannedInspection(selector, Queries.sheetSummary());
  }

  /** Returns one sheet-layout inspection step. */
  public PlannedInspection layout() {
    return new PlannedInspection(selector, Queries.sheetLayout());
  }

  /** Returns one print-layout inspection step. */
  public PlannedInspection printLayout() {
    return new PlannedInspection(selector, Queries.printLayout());
  }

  /** Returns one merged-regions inspection step. */
  public PlannedInspection mergedRegions() {
    return new PlannedInspection(selector, Queries.mergedRegions());
  }

  /** Returns one autofilter inspection step. */
  public PlannedInspection autofilters() {
    return new PlannedInspection(selector, Queries.autofilters());
  }

  /** Returns one chart inventory inspection step for this sheet. */
  public PlannedInspection charts() {
    return new PlannedInspection(selector, Queries.charts());
  }

  /** Returns one drawing-object inventory inspection step for this sheet. */
  public PlannedInspection drawingObjects() {
    return new PlannedInspection(
        new DrawingObjectSelector.AllOnSheet(selector.name()), Queries.drawingObjects());
  }

  /** Returns one formula-surface inspection step. */
  public PlannedInspection formulaSurface() {
    return new PlannedInspection(selector, Queries.formulaSurface());
  }

  /** Returns one formula-health analysis step. */
  public PlannedInspection formulaHealth() {
    return new PlannedInspection(selector, Queries.formulaHealth());
  }
}
