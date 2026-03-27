package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for ExcelCellSelection constructor invariants and defensive copies. */
class ExcelCellSelectionTest {
  @Test
  void createsAllUsedCellsSelection() {
    ExcelCellSelection.AllUsedCells selection = new ExcelCellSelection.AllUsedCells();

    assertInstanceOf(ExcelCellSelection.AllUsedCells.class, selection);
  }

  @Test
  void selectedCopiesAddressesAndPreservesOrder() {
    List<String> addresses = new ArrayList<>(List.of("A1", "B2"));

    ExcelCellSelection.Selected selection = new ExcelCellSelection.Selected(addresses);
    addresses.clear();

    assertEquals(List.of("A1", "B2"), selection.addresses());
  }

  @Test
  void selectedRejectsInvalidAddressCollections() {
    assertThrows(NullPointerException.class, () -> new ExcelCellSelection.Selected(null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelCellSelection.Selected(List.of()));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelCellSelection.Selected(List.of("A1", "A1")));
    assertThrows(
        NullPointerException.class, () -> new ExcelCellSelection.Selected(List.of("A1", null)));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelCellSelection.Selected(List.of(" ")));
  }
}
