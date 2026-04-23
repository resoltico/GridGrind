package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.excel.foundation.ExcelAddressLists;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for canonical exact-cell address-list validation. */
class ExcelAddressListsTest {
  @Test
  void copyNonEmptyDistinctAddressesRejectsMalformedAddressesWithIndexedContext() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelAddressLists.copyNonEmptyDistinctAddresses(List.of("A1:B2")));

    assertEquals(
        "addresses[0] address must be a single-cell A1-style address", failure.getMessage());
  }

  @Test
  void copyNonEmptyDistinctAddressesRejectsOutOfBoundsAddressesWithIndexedContext() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelAddressLists.copyNonEmptyDistinctAddresses(List.of("XFE1")));

    assertEquals(
        "addresses[0] addresses must stay within Excel .xlsx bounds", failure.getMessage());
  }

  @Test
  void copyNonEmptyDistinctAddressesRejectsOutOfBoundsRowNumbersWithIndexedContext() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelAddressLists.copyNonEmptyDistinctAddresses(List.of("A1048577")));

    assertEquals(
        "addresses[0] addresses must stay within Excel .xlsx bounds", failure.getMessage());
  }
}
