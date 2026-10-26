package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for table-related value objects and constructor invariants. */
class ExcelTableValueTypesTest {
  @Test
  void tableDefinitionAndSnapshotValidateAndCopyCollections() {
    List<String> columnNames = new ArrayList<>(List.of("Owner", "Task"));

    ResolvedTable definition =
        new ResolvedTable(
            new TableName("Queue"),
            new SheetName("Ops"),
            new RangeRef("A1:B3"),
            false,
            new TableStyleSpec.None());
    TableEntryReport snapshot =
        new TableEntryReport(
            new dev.erst.gridgrind.core.vocabulary.TableLabel("Queue"),
            new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
            new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"),
            1,
            0,
            columnNames,
            new TableStyleReport.None(),
            true);
    columnNames.clear();

    assertEquals(new TableName("Queue"), definition.name());
    assertEquals(List.of("Owner", "Task"), snapshot.structure().columnNames());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedTable(
                new TableName("Queue"),
                new SheetName(" "),
                new RangeRef("A1:B2"),
                false,
                new TableStyleSpec.None()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedTable(
                new TableName("Queue"),
                new SheetName("12345678901234567890123456789012"),
                new RangeRef("A1:B2"),
                false,
                new TableStyleSpec.None()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedTable(
                new TableName("Queue"),
                new SheetName("Ops"),
                new RangeRef(" "),
                false,
                new TableStyleSpec.None()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ResolvedTable(
                new TableName("Queue"),
                new SheetName("Ops"),
                new RangeRef(null),
                false,
                new TableStyleSpec.None()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ResolvedTable(
                new TableName("Queue"), new SheetName("Ops"), new RangeRef("A1:B2"), false, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new TableEntryReport(
                new dev.erst.gridgrind.core.vocabulary.TableLabel("Queue"),
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"),
                -1,
                0,
                List.of("Owner"),
                new TableStyleReport.None(),
                false));
    assertThrows(
        NullPointerException.class,
        () ->
            new TableEntryReport(
                new dev.erst.gridgrind.core.vocabulary.TableLabel("Queue"),
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"),
                1,
                0,
                List.of("Owner", null),
                new TableStyleReport.None(),
                false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new TableEntryReport(
                new dev.erst.gridgrind.core.vocabulary.TableLabel("Queue"),
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"),
                1,
                -1,
                List.of("Owner"),
                new TableStyleReport.None(),
                false));
    assertThrows(
        NullPointerException.class,
        () ->
            new TableEntryReport(
                new dev.erst.gridgrind.core.vocabulary.TableLabel("Queue"),
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"),
                1,
                0,
                List.of("Owner"),
                null,
                false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new TableEntryReport(
                new dev.erst.gridgrind.core.vocabulary.TableLabel("Queue"),
                new dev.erst.gridgrind.core.vocabulary.SheetLabel(" "),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"),
                1,
                0,
                List.of("Owner"),
                new TableStyleReport.None(),
                false));
  }

  @Test
  void tableStyleAndSelectionValidateNamesAndDuplicates() {
    TableStyleSpec.Named style =
        new TableStyleSpec.Named("TableStyleMedium2", false, false, true, false);
    TableSelector.ByNames selection =
        new TableSelector.ByNames(
            List.of("Queue").stream()
                .map(dev.erst.gridgrind.core.vocabulary.TableName::new)
                .toList());

    assertEquals("TableStyleMedium2", style.name());
    assertEquals(
        List.of("Queue"),
        selection.names().stream()
            .map(dev.erst.gridgrind.core.vocabulary.TableName::value)
            .toList());

    assertThrows(
        IllegalArgumentException.class,
        () -> new TableStyleSpec.Named(" ", false, false, true, false));
    assertThrows(
        NullPointerException.class,
        () -> new TableStyleSpec.Named(null, false, false, true, false));
    assertThrows(IllegalArgumentException.class, () -> new TableSelector.ByNames(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new TableSelector.ByNames(
                List.of("Queue", "queue").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.TableName::new)
                    .toList()));
    assertThrows(
        NullPointerException.class,
        () ->
            new TableSelector.ByNames(
                List.of("Queue", null).stream()
                    .map(dev.erst.gridgrind.core.vocabulary.TableName::new)
                    .toList()));
  }
}
