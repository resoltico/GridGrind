package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.SheetName;

/** Selects one or more sheet-local drawing objects. */
@Union(discriminator = "type", catalog = true, displayOrder = 37)
public sealed interface DrawingObjectSelector extends Selector
    permits dev.erst.gridgrind.core.plan.DrawingObjectSelector.AllOnSheet,
        dev.erst.gridgrind.core.plan.DrawingObjectSelector.ByName {

  /** Selects every drawing object on one sheet. */
  @Variant(
      id = "DRAWING_OBJECT_ALL_ON_SHEET",
      summary = "Select every drawing object on one sheet.")
  record AllOnSheet(SheetName sheetName) implements DrawingObjectSelector {
    public AllOnSheet {
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects one drawing object by sheet-local object name. */
  @Variant(
      id = "DRAWING_OBJECT_BY_NAME",
      summary = "Select one drawing object by exact sheet-local object name.")
  record ByName(SheetName sheetName, String objectName) implements DrawingObjectSelector {
    public ByName {
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      objectName = SelectorValueValidation.requireNonBlank(objectName, "objectName");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }
}
