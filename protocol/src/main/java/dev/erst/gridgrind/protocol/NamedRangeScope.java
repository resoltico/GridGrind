package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import java.util.Objects;

/** Protocol-facing scope of a named range, either workbook-wide or local to one sheet. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = NamedRangeScope.Workbook.class, name = "WORKBOOK"),
  @JsonSubTypes.Type(value = NamedRangeScope.Sheet.class, name = "SHEET")
})
public sealed interface NamedRangeScope permits NamedRangeScope.Workbook, NamedRangeScope.Sheet {

  /** Converts this protocol scope into the workbook-core representation. */
  ExcelNamedRangeScope toExcelNamedRangeScope();

  /** Workbook-wide named-range scope. */
  record Workbook() implements NamedRangeScope {
    @Override
    public ExcelNamedRangeScope toExcelNamedRangeScope() {
      return new ExcelNamedRangeScope.WorkbookScope();
    }
  }

  /** Sheet-local named-range scope bound to one sheet name. */
  record Sheet(String sheetName) implements NamedRangeScope {
    public Sheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }

    @Override
    public ExcelNamedRangeScope toExcelNamedRangeScope() {
      return new ExcelNamedRangeScope.SheetScope(sheetName);
    }
  }
}
