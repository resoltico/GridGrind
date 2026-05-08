package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.*;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.step.WorkbookStepTargeting;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct coverage for catalog target-surface helper extraction. */
class TypeEntryTargetingSupportTest {
  @Test
  void resolvesOptionalTargetSurfaceAcrossSupportedRecordFamilies() {
    Optional<WorkbookStepTargeting.TargetSurface> mutationSurface =
        TypeEntryTargetingSupport.optionalTargetSurfaceFor(
            WorkbookMutationAction.EnsureSheet.class);
    Optional<WorkbookStepTargeting.TargetSurface> assertionSurface =
        TypeEntryTargetingSupport.optionalTargetSurfaceFor(PresenceAssertion.TablePresent.class);
    Optional<WorkbookStepTargeting.TargetSurface> querySurface =
        TypeEntryTargetingSupport.optionalTargetSurfaceFor(
            WorkbookIntrospectionQuery.GetWorkbookSummary.class);

    assertTrue(mutationSurface.isPresent());
    assertTrue(assertionSurface.isPresent());
    assertTrue(querySurface.isPresent());
    assertFalse(TypeEntryTargetingSupport.targetSelectorEntries(mutationSurface).isEmpty());
    assertFalse(TypeEntryTargetingSupport.targetSelectorEntries(assertionSurface).isEmpty());
    assertFalse(TypeEntryTargetingSupport.targetSelectorEntries(querySurface).isEmpty());
  }

  @Test
  void omitsTargetSurfaceForUntargetedRecords() {
    assertEquals(
        Optional.empty(),
        TypeEntryTargetingSupport.optionalTargetSurfaceFor(
            WorkbookPlan.WorkbookPersistence.None.class));
    assertEquals(List.of(), TypeEntryTargetingSupport.targetSelectorEntries(Optional.empty()));
  }
}
