package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Verifies executable operand ownership and selector contracts independently of POI. */
class ResolvedStepTest {
  @Test
  void sharesSourceFreeActionsAndRejectsMismatchedTargets() {
    var action = new CellMutationAction.ClearHyperlink();
    var target = new CellSelector.ByAddress(new SheetName("Orders"), new CellAddress("B2"));
    var step = new ResolvedStep(target, action);
    assertSame(action, step.action());
    assertSame(target, step.target());
    assertSame(CellMutationAction.ClearHyperlink.class, action.authoredType());
    assertThrows(
        IllegalArgumentException.class,
        () -> new ResolvedStep(new SheetSelector.ByName(new SheetName("Orders")), action));
  }

  @Test
  void resolvedGridsOwnBothCollectionLevelsAndRequireRectangularContent() {
    List<ResolvedCellWrite> row = new ArrayList<>(List.of(new CellValue.Text("original")));
    List<List<ResolvedCellWrite>> rows = new ArrayList<>(List.of(row));
    var action = new ResolvedMutationAction.SetRange(rows);
    row.set(0, new CellValue.Text("changed"));
    rows.clear();
    assertEquals(List.of(List.of(new CellValue.Text("original"))), action.rows());
    assertThrows(
        IllegalArgumentException.class, () -> new ResolvedMutationAction.SetRange(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ResolvedMutationAction.SetRange(List.of(List.of())));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedMutationAction.SetRange(
                List.of(List.of(new CellValue.Blank()), List.of())));
  }

  @Test
  void materializedProtectionNeverPrintsCredentials() {
    var action =
        new ResolvedMutationAction.SetSheetProtection(
            new SheetProtectionSettings(
                false, false, false, false, false, false, false, false, false, false, false, false,
                false, false, false),
            Optional.of("credential-sentinel"));
    assertFalse(action.toString().contains("credential-sentinel"));
    assertSame(WorkbookMutationAction.SetSheetProtection.class, action.authoredType());
  }
}
