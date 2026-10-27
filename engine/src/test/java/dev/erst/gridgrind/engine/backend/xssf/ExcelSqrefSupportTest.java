package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.engine.poi.SqrefRanges;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for XMLBeans sqref normalization. */
class ExcelSqrefSupportTest {
  @Test
  void normalizesListBackedSqrefPayloads() {
    assertEquals(
        List.of("A1:A3", "B2"), SqrefRanges.normalizedSqref(List.of("  A1:A3  ", "", " ", "B2")));
  }
}
