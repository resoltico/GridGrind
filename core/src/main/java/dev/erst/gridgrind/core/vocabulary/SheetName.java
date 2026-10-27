package dev.erst.gridgrind.core.vocabulary;

/** An Excel worksheet name, preserving its authored spelling. */
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_SHEET_NAME_CONTRACT
})
public record SheetName(String value) implements WireString {
  public SheetName {
    ExcelSheetNames.requireValid(value, "sheetName");
  }

  @Override
  public String toString() {
    return value;
  }
}
