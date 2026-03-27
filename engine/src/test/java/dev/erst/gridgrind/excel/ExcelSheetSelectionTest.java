package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for ExcelSheetSelection constructor invariants and defensive copies. */
class ExcelSheetSelectionTest {
  @Test
  void createsAllSelection() {
    ExcelSheetSelection.All selection = new ExcelSheetSelection.All();

    assertInstanceOf(ExcelSheetSelection.All.class, selection);
  }

  @Test
  void selectedCopiesSheetNamesAndPreservesOrder() {
    List<String> sheetNames = new ArrayList<>(List.of("Budget", "Summary"));

    ExcelSheetSelection.Selected selection = new ExcelSheetSelection.Selected(sheetNames);
    sheetNames.clear();

    assertEquals(List.of("Budget", "Summary"), selection.sheetNames());
  }

  @Test
  void selectedRejectsInvalidSheetNameCollections() {
    assertThrows(NullPointerException.class, () -> new ExcelSheetSelection.Selected(null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelSheetSelection.Selected(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelSheetSelection.Selected(List.of("Budget", "Budget")));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelSheetSelection.Selected(List.of("Budget", null)));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelSheetSelection.Selected(List.of(" ")));
  }
}
