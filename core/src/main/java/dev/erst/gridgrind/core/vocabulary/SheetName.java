package dev.erst.gridgrind.core.vocabulary;

/** An Excel worksheet name, preserving its authored spelling. */
public record SheetName(String value) implements WireString {
  public SheetName { // LIM-003
    ExcelSheetNames.requireValid(value, "sheetName");
  }

  @Override
  public String toString() {
    return value;
  }
}
