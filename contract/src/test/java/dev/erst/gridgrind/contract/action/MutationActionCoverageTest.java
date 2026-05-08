package dev.erst.gridgrind.contract.action;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

/** Edge-path coverage for mutation-action selector metadata lookup. */
class MutationActionCoverageTest {
  @Test
  void rejectsUnmappedAndMetadataFreeMutationActionTypes() {
    @SuppressWarnings("unchecked")
    Class<? extends MutationAction> nonRecordActionType =
        (Class<? extends MutationAction>) MutationAction.class;

    IllegalArgumentException nonRecordFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> MutationAction.allowedTargetTypesForType(nonRecordActionType));
    assertEquals(
        "No target-type mapping configured for action class dev.erst.gridgrind.contract.action.MutationAction",
        nonRecordFailure.getMessage());

    @SuppressWarnings("unchecked")
    Class<? extends MutationAction> missingMetadataActionType =
        (Class<? extends MutationAction>) (Class<?>) MissingMetadataActionRecord.class;

    IllegalArgumentException missingMetadataFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> MutationAction.allowedTargetTypesForType(missingMetadataActionType));
    assertEquals(
        "No target-type mapping configured for action class "
            + MissingMetadataActionRecord.class.getName(),
        missingMetadataFailure.getMessage());
  }

  private record MissingMetadataActionRecord() {}
}
