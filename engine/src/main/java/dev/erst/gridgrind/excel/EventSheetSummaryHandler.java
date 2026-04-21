package dev.erst.gridgrind.excel;

import org.apache.poi.ss.util.CellReference;
import org.xml.sax.Attributes;
import org.xml.sax.helpers.DefaultHandler;

/** SAX handler that extracts the factual sheet-summary surface supported by EVENT_READ. */
final class EventSheetSummaryHandler extends DefaultHandler {
  private int physicalRowCount;
  private int lastRowIndex = -1;
  private int lastColumnIndex = -1;
  private int nextRowIndex;
  private boolean selected;
  private WorkbookReadResult.SheetProtection protection =
      new WorkbookReadResult.SheetProtection.Unprotected();

  @Override
  public void startElement(String uri, String localName, String qName, Attributes attributes) {
    String element = localName == null || localName.isEmpty() ? qName : localName;
    switch (element) {
      case "sheetView" -> selected = selected || booleanAttribute(attributes, "tabSelected");
      case "sheetProtection" -> protection = sheetProtection(attributes);
      case "row" -> handleRow(attributes);
      case "c" -> handleCell(attributes);
      case "col" -> handleColumn(attributes);
      default -> {}
    }
  }

  private void handleRow(Attributes attributes) {
    physicalRowCount++;
    String rowRef = attributes.getValue("r");
    int rowIndex = rowRef == null ? nextRowIndex : Integer.parseInt(rowRef) - 1;
    lastRowIndex = Math.max(lastRowIndex, rowIndex);
    nextRowIndex = rowIndex + 1;
  }

  private void handleCell(Attributes attributes) {
    String cellRef = attributes.getValue("r");
    if (cellRef == null || cellRef.isBlank()) {
      return;
    }
    lastColumnIndex = Math.max(lastColumnIndex, new CellReference(cellRef).getCol());
  }

  private void handleColumn(Attributes attributes) {
    String max = attributes.getValue("max");
    if (max == null || max.isBlank()) {
      return;
    }
    lastColumnIndex = Math.max(lastColumnIndex, Integer.parseInt(max) - 1);
  }

  private WorkbookReadResult.SheetProtection sheetProtection(Attributes attributes) {
    if (!booleanAttribute(attributes, "sheet")) {
      return new WorkbookReadResult.SheetProtection.Unprotected();
    }
    return new WorkbookReadResult.SheetProtection.Protected(
        new ExcelSheetProtectionSettings(
            booleanAttribute(attributes, "autoFilter"),
            booleanAttribute(attributes, "deleteColumns"),
            booleanAttribute(attributes, "deleteRows"),
            booleanAttribute(attributes, "formatCells"),
            booleanAttribute(attributes, "formatColumns"),
            booleanAttribute(attributes, "formatRows"),
            booleanAttribute(attributes, "insertColumns"),
            booleanAttribute(attributes, "insertHyperlinks"),
            booleanAttribute(attributes, "insertRows"),
            booleanAttribute(attributes, "objects"),
            booleanAttribute(attributes, "pivotTables"),
            booleanAttribute(attributes, "scenarios"),
            booleanAttribute(attributes, "selectLockedCells"),
            booleanAttribute(attributes, "selectUnlockedCells"),
            booleanAttribute(attributes, "sort")));
  }

  private static boolean booleanAttribute(Attributes attributes, String name) {
    String value = attributes.getValue(name);
    return "1".equals(value) || "true".equalsIgnoreCase(value);
  }

  EventSheetSummary summary() {
    return new EventSheetSummary(
        selected, protection, physicalRowCount, lastRowIndex, lastColumnIndex);
  }
}
