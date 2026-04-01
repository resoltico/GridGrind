package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import java.util.Arrays;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for workbook-read operation invariants. */
class WorkbookReadOperationTest {
  @Test
  void getWindowAcceptsWindowAtTheMaximumCellLimit() {
    // 500 * 500 = 250,000 == MAX_WINDOW_CELLS — must not throw
    assertDoesNotThrow(() -> new WorkbookReadOperation.GetWindow("w", "Sheet1", "A1", 500, 500));
  }

  @Test
  void getWindowRejectsWindowExceedingMaximumCellLimit() {
    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class,
            () -> new WorkbookReadOperation.GetWindow("w", "Sheet1", "A1", 501, 500));

    assertTrue(
        exception.getMessage().contains("rowCount * columnCount must not exceed"),
        "message should state the limit");
    assertTrue(
        exception.getMessage().contains("250500"), "message should include the actual cell count");
  }

  @Test
  void getSheetSchemaAcceptsWindowAtTheMaximumCellLimit() {
    assertDoesNotThrow(
        () -> new WorkbookReadOperation.GetSheetSchema("s", "Sheet1", "A1", 500, 500));
  }

  @Test
  void getSheetSchemaRejectsWindowExceedingMaximumCellLimit() {
    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class,
            () -> new WorkbookReadOperation.GetSheetSchema("s", "Sheet1", "A1", 1000, 1000));

    assertTrue(exception.getMessage().contains("rowCount * columnCount must not exceed"));
    assertTrue(exception.getMessage().contains("1000000"));
  }

  @Test
  void getCellsRejectsBlankAddresses() {
    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class,
            () -> new WorkbookReadOperation.GetCells("cells", "Budget", Arrays.asList("A1", " ")));

    assertEquals("addresses must not be blank", exception.getMessage());
  }

  @Test
  void getCellsRejectsDuplicateAddresses() {
    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class,
            () -> new WorkbookReadOperation.GetCells("cells", "Budget", Arrays.asList("A1", "A1")));

    assertEquals("addresses must not contain duplicates", exception.getMessage());
  }

  @Test
  void getCellsCopiesAddresses() {
    List<String> addresses = Arrays.asList("A1", "B4");
    WorkbookReadOperation.GetCells operation =
        new WorkbookReadOperation.GetCells("cells", "Budget", addresses);

    assertEquals(List.of("A1", "B4"), operation.addresses());
  }

  @Test
  void getDataValidationsRequiresSheetNameAndSelection() {
    WorkbookReadOperation.GetDataValidations operation =
        new WorkbookReadOperation.GetDataValidations(
            "validations", "Budget", new RangeSelection.Selected(List.of("A1:B2")));

    assertEquals("Budget", operation.sheetName());
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadOperation.GetDataValidations("validations", "Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadOperation.GetDataValidations(
                "validations", " ", new RangeSelection.All()));
  }

  @Test
  void analyzeDataValidationHealthRequiresSelection() {
    WorkbookReadOperation.AnalyzeDataValidationHealth operation =
        new WorkbookReadOperation.AnalyzeDataValidationHealth(
            "health", new SheetSelection.Selected(List.of("Budget")));

    assertEquals("health", operation.requestId());
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadOperation.AnalyzeDataValidationHealth("health", null));
  }

  @Test
  void getAutofiltersAndAnalyzeAutofilterHealthRequireSheetSelectionState() {
    WorkbookReadOperation.GetAutofilters getAutofilters =
        new WorkbookReadOperation.GetAutofilters("filters", "Budget");
    WorkbookReadOperation.AnalyzeAutofilterHealth analyzeAutofilterHealth =
        new WorkbookReadOperation.AnalyzeAutofilterHealth(
            "autofilter-health", new SheetSelection.All());

    assertEquals("Budget", getAutofilters.sheetName());
    assertEquals("autofilter-health", analyzeAutofilterHealth.requestId());
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadOperation.GetAutofilters("filters", " "));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadOperation.AnalyzeAutofilterHealth("autofilter-health", null));
  }

  @Test
  void getTablesAndAnalyzeTableHealthRequireTableSelection() {
    WorkbookReadOperation.GetTables getTables =
        new WorkbookReadOperation.GetTables("tables", new TableSelection.All());
    WorkbookReadOperation.AnalyzeTableHealth analyzeTableHealth =
        new WorkbookReadOperation.AnalyzeTableHealth(
            "table-health", new TableSelection.ByNames(List.of("BudgetTable")));

    assertEquals("tables", getTables.requestId());
    assertEquals(
        List.of("BudgetTable"), ((TableSelection.ByNames) analyzeTableHealth.selection()).names());
    assertThrows(
        NullPointerException.class, () -> new WorkbookReadOperation.GetTables("tables", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadOperation.AnalyzeTableHealth("table-health", null));
  }
}
