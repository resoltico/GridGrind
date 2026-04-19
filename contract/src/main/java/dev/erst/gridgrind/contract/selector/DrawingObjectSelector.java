package dev.erst.gridgrind.contract.selector;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** Selects one or more sheet-local drawing objects. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = DrawingObjectSelector.AllOnSheet.class, name = "ALL_ON_SHEET"),
  @JsonSubTypes.Type(value = DrawingObjectSelector.ByName.class, name = "BY_NAME")
})
public sealed interface DrawingObjectSelector extends Selector
    permits DrawingObjectSelector.AllOnSheet, DrawingObjectSelector.ByName {

  /** Selects every drawing object on one sheet. */
  record AllOnSheet(String sheetName) implements DrawingObjectSelector {
    public AllOnSheet {
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects one drawing object by sheet-local object name. */
  record ByName(String sheetName, String objectName) implements DrawingObjectSelector {
    public ByName {
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
      objectName = SelectorSupport.requireNonBlank(objectName, "objectName");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }
}
