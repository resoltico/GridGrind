package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for table-related value objects and constructor invariants. */
class ExcelTableValueTypesTest {
  @Test
  void tableDefinitionAndSnapshotValidateAndCopyCollections() {
    List<String> columnNames = new ArrayList<>(List.of("Owner", "Task"));

    ExcelTableDefinition definition =
        new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None());
    ExcelTableSnapshot snapshot =
        new ExcelTableSnapshot(
            "Queue", "Ops", "A1:B3", 1, 0, columnNames, new ExcelTableStyleSnapshot.None(), true);
    columnNames.clear();

    assertEquals("Queue", definition.name());
    assertEquals(List.of("Owner", "Task"), snapshot.columnNames());

    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelTableDefinition("Queue", " ", "A1:B2", false, new ExcelTableStyle.None()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelTableDefinition(
                "Queue",
                "12345678901234567890123456789012",
                "A1:B2",
                false,
                new ExcelTableStyle.None()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelTableDefinition("Queue", "Ops", " ", false, new ExcelTableStyle.None()));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelTableDefinition("Queue", "Ops", null, false, new ExcelTableStyle.None()));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelTableDefinition("Queue", "Ops", "A1:B2", false, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelTableSnapshot(
                "Queue",
                "Ops",
                "A1:B3",
                -1,
                0,
                List.of("Owner"),
                new ExcelTableStyleSnapshot.None(),
                false));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelTableSnapshot(
                "Queue",
                "Ops",
                "A1:B3",
                1,
                0,
                List.of("Owner", null),
                new ExcelTableStyleSnapshot.None(),
                false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelTableSnapshot(
                "Queue",
                "Ops",
                "A1:B3",
                1,
                -1,
                List.of("Owner"),
                new ExcelTableStyleSnapshot.None(),
                false));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelTableSnapshot("Queue", "Ops", "A1:B3", 1, 0, List.of("Owner"), null, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelTableSnapshot(
                "Queue",
                " ",
                "A1:B3",
                1,
                0,
                List.of("Owner"),
                new ExcelTableStyleSnapshot.None(),
                false));
  }

  @Test
  void tableStyleAndSelectionValidateNamesAndDuplicates() {
    ExcelTableStyle.Named style =
        new ExcelTableStyle.Named("TableStyleMedium2", false, false, true, false);
    ExcelTableSelection.ByNames selection = new ExcelTableSelection.ByNames(List.of("Queue"));

    assertEquals("TableStyleMedium2", style.name());
    assertEquals(List.of("Queue"), selection.names());

    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelTableStyle.Named(" ", false, false, true, false));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelTableStyle.Named(null, false, false, true, false));
    assertThrows(IllegalArgumentException.class, () -> new ExcelTableSelection.ByNames(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelTableSelection.ByNames(List.of("Queue", "queue")));
    assertThrows(
        NullPointerException.class, () -> new ExcelTableSelection.ByNames(List.of("Queue", null)));
  }
}
