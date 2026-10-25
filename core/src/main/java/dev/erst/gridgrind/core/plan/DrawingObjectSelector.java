package dev.erst.gridgrind.core.plan;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.core.vocabulary.SheetName;

/** Selects one or more sheet-local drawing objects. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(
      value = DrawingObjectSelector.AllOnSheet.class,
      name = "DRAWING_OBJECT_ALL_ON_SHEET"),
  @JsonSubTypes.Type(value = DrawingObjectSelector.ByName.class, name = "DRAWING_OBJECT_BY_NAME")
})
public sealed interface DrawingObjectSelector extends Selector
    permits DrawingObjectSelector.AllOnSheet, DrawingObjectSelector.ByName {

  /** Selects every drawing object on one sheet. */
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
