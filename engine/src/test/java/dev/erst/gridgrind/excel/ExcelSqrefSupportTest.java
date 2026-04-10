package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for XMLBeans sqref normalization. */
class ExcelSqrefSupportTest {
  @Test
  void normalizesListBackedSqrefPayloads() {
    assertEquals(
        List.of("A1:A3", "B2"),
        ExcelSqrefSupport.normalizedSqref(List.of("  A1:A3  ", "", " ", "B2")));
  }
}
