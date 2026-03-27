package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for ExcelCellMetadataSnapshot factory helpers and defaults. */
class ExcelCellMetadataSnapshotTest {
  @Test
  void createsEmptyAndOptionalMetadataSnapshots() {
    assertEquals(
        ExcelCellMetadataSnapshot.empty(),
        new ExcelCellMetadataSnapshot(Optional.empty(), Optional.empty()));

    ExcelCellMetadataSnapshot snapshot =
        ExcelCellMetadataSnapshot.of(
            new ExcelHyperlink.Document("Budget!B4"),
            new ExcelComment("Review", "GridGrind", false));

    assertEquals(Optional.of(new ExcelHyperlink.Document("Budget!B4")), snapshot.hyperlink());
    assertEquals(Optional.of(new ExcelComment("Review", "GridGrind", false)), snapshot.comment());
  }

  @Test
  @SuppressWarnings("NullOptional")
  void coalescesNullOptionalsToEmptyOptionals() {
    ExcelCellMetadataSnapshot snapshot = new ExcelCellMetadataSnapshot(null, null);

    assertEquals(Optional.empty(), snapshot.hyperlink());
    assertEquals(Optional.empty(), snapshot.comment());
  }
}
