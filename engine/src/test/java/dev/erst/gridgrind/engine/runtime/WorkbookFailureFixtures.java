package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.engine.backend.xssf.CellNotFoundException;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import java.io.IOException;
import java.io.UncheckedIOException;

/** Actual native failures used to verify runtime diagnostic translation. */
final class WorkbookFailureFixtures {
  private WorkbookFailureFixtures() {}

  static CellNotFoundException missingCell(CellAddress address) {
    try (var workbook = XssfWorkbooks.create()) {
      var cells = workbook.getOrCreateSheet("Fixture").cells();
      return assertThrows(CellNotFoundException.class, () -> cells.text(address));
    } catch (IOException failure) {
      throw new UncheckedIOException(failure);
    }
  }
}
