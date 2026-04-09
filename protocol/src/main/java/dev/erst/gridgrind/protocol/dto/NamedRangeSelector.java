package dev.erst.gridgrind.protocol.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelSheetNames;
import java.util.Objects;

/** Exact selector used to request one or more named ranges during workbook analysis. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = NamedRangeSelector.ByName.class, name = "BY_NAME"),
  @JsonSubTypes.Type(value = NamedRangeSelector.WorkbookScope.class, name = "WORKBOOK_SCOPE"),
  @JsonSubTypes.Type(value = NamedRangeSelector.SheetScope.class, name = "SHEET_SCOPE")
})
public sealed interface NamedRangeSelector
    permits NamedRangeSelector.ByName,
        NamedRangeSelector.WorkbookScope,
        NamedRangeSelector.SheetScope {

  /** Matches named ranges by name across all scopes. */
  record ByName(String name) implements NamedRangeSelector {
    public ByName {
      name = requireNonBlank(name, "name");
    }
  }

  /** Matches the workbook-scoped named range with the given name. */
  record WorkbookScope(String name) implements NamedRangeSelector {
    public WorkbookScope {
      name = requireNonBlank(name, "name");
    }
  }

  /** Matches the sheet-scoped named range with the given name on the given sheet. */
  record SheetScope(String name, String sheetName) implements NamedRangeSelector {
    public SheetScope {
      name = requireNonBlank(name, "name");
      sheetName = requireNonBlank(sheetName, "sheetName");
      ExcelSheetNames.requireValid(sheetName, "sheetName");
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
