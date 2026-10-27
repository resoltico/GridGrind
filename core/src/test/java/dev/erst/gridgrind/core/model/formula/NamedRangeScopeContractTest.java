package dev.erst.gridgrind.core.model.formula;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import org.junit.jupiter.api.Test;

/** Observed scope labels and authored scope identifiers have different contracts. */
class NamedRangeScopeContractTest {
  @Test
  void observedScopePreservesStoredLabels() {
    assertEquals(
        "stored[]",
        new NamedRangeScopeReport.Sheet(new SheetLabel("stored[]")).sheetName().value());
    assertThrows(NullPointerException.class, () -> new NamedRangeScopeReport.Sheet(null));
  }

  @Test
  void authoredScopeRetainsWorksheetNameValidation() {
    assertThrows(NullPointerException.class, () -> new NamedRangeScope.Sheet(null));
    assertThrows(
        IllegalArgumentException.class, () -> new NamedRangeScope.Sheet(new SheetName(" ")));
    assertThrows(
        IllegalArgumentException.class, () -> new NamedRangeScope.Sheet(new SheetName("stored[]")));
  }
}
