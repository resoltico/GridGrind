package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for factual table report invariants and defensive copies. */
class TableEntryReportTest {
  @Test
  void validatesAndCopiesColumnMetadata() {
    List<String> columnNames = new ArrayList<>(List.of("Item", "Amount"));

    TableEntryReport report =
        new TableEntryReport(
            "BudgetTable", "Budget", "A1:B4", 1, 1, columnNames, new TableStyleReport.None(), true);
    columnNames.clear();

    assertEquals(List.of("Item", "Amount"), report.columnNames());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new TableEntryReport(
                "BudgetTable",
                " ",
                "A1:B4",
                1,
                1,
                List.of("Item"),
                new TableStyleReport.None(),
                true));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new TableEntryReport(
                "BudgetTable",
                "Budget",
                "A1:B4",
                -1,
                1,
                List.of("Item"),
                new TableStyleReport.None(),
                true));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new TableEntryReport(
                "BudgetTable",
                "Budget",
                "A1:B4",
                1,
                -1,
                List.of("Item"),
                new TableStyleReport.None(),
                true));
    assertThrows(
        NullPointerException.class,
        () ->
            new TableEntryReport(
                "BudgetTable", "Budget", "A1:B4", 1, 1, List.of("Item"), null, true));
  }
}
