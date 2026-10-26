package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.value.CellReport;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for ExcelPreviewRow compact constructor validation and defensive copy behaviour. */
class ExcelPreviewRowTest {
  @Test
  void coercesNullCellsToEmptyList() {
    ExcelPreviewRow row = new ExcelPreviewRow(3, null);
    assertEquals(3, row.rowIndex());
    assertEquals(List.of(), row.cells());
  }

  @Test
  void defensivelyCopiesMutableCellsList() {
    java.util.List<CellReport> mutable = new java.util.ArrayList<>();
    ExcelPreviewRow row = new ExcelPreviewRow(0, mutable);
    assertNotSame(mutable, row.cells());
    assertThrows(UnsupportedOperationException.class, () -> row.cells().add(null));
  }
}
