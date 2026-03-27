package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import java.util.Arrays;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for workbook-read operation invariants. */
class WorkbookReadOperationTest {
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
}
