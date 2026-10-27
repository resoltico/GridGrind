package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Canonical table-name selection invariants. */
class TableSelectorTest {

  @Test
  void validatesByNamesSelection() {
    assertThrows(IllegalArgumentException.class, () -> new TableSelector.ByNames(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new TableSelector.ByNames(
                List.of(new TableName("BudgetTable"), new TableName("budgettable"))));
    assertThrows(
        NullPointerException.class,
        () -> new TableSelector.ByNames(List.of(new TableName("BudgetTable"), null)));
  }
}
