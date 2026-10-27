package dev.erst.gridgrind.core.model.protection;

import static org.junit.jupiter.api.Assertions.*;

import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Ensures materialized credentials cannot escape through ordinary debug rendering. */
class ResolvedWorkbookProtectionTest {
  @Test
  void debugRenderingRetainsLockStateAndRedactsCredentials() {
    String workbookSecret = "workbook-sensitive-sentinel";
    String revisionsSecret = "revisions-sensitive-sentinel";
    var value =
        new ResolvedWorkbookProtection(
            true, false, true, Optional.of(workbookSecret), Optional.of(revisionsSecret));
    String debug = value.toString();
    assertFalse(debug.contains(workbookSecret));
    assertFalse(debug.contains(revisionsSecret));
    assertTrue(debug.contains("structureLocked=true"));
    assertTrue(debug.contains("revisionsLocked=true"));
    assertEquals(Optional.of(workbookSecret), value.workbookPassword());
    assertEquals(Optional.of(revisionsSecret), value.revisionsPassword());
  }
}
