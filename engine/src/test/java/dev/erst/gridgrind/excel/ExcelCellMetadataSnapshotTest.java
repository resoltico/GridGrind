package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
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
            Optional.of(new HyperlinkTarget.Document("Budget!B4")),
            Optional.of(
                new CommentReport(
                    "Review", "GridGrind", false, Optional.empty(), Optional.empty())));

    assertEquals(Optional.of(new HyperlinkTarget.Document("Budget!B4")), snapshot.hyperlink());
    assertEquals(
        Optional.of(
            new CommentReport("Review", "GridGrind", false, Optional.empty(), Optional.empty())),
        snapshot.comment());
  }

  @Test
  @SuppressWarnings("NullOptional")
  void rejectsNullOptionals() {
    assertThrows(
        NullPointerException.class, () -> new ExcelCellMetadataSnapshot(null, Optional.empty()));
    assertThrows(
        NullPointerException.class, () -> new ExcelCellMetadataSnapshot(Optional.empty(), null));
  }
}
