package dev.erst.gridgrind.contract.query;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Edge-path coverage for inspection-query selector metadata lookup. */
class InspectionQueryCoverageTest {
  @Test
  void discoversTargetsFromTheCanonicalOperationContract() {
    assertEquals(
        List.of(WorkbookSelector.class),
        List.of(
            InspectionQuery.allowedTargetTypes(
                new WorkbookIntrospectionQuery.GetWorkbookSummary())));
    assertEquals(
        List.of(WorkbookSelector.class),
        List.of(
            InspectionQuery.allowedTargetTypesForType(
                WorkbookIntrospectionQuery.GetWorkbookSummary.class)));
  }

  @Test
  void rejectsUnmappedAndMetadataFreeInspectionQueryTypes() {
    @SuppressWarnings("unchecked")
    Class<? extends InspectionQuery> nonRecordQueryType =
        (Class<? extends InspectionQuery>) InspectionQuery.class;

    IllegalArgumentException nonRecordFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> InspectionQuery.allowedTargetTypesForType(nonRecordQueryType));
    assertEquals(
        "No target-type mapping configured for query class dev.erst.gridgrind.contract.query.InspectionQuery",
        nonRecordFailure.getMessage());

    @SuppressWarnings("unchecked")
    Class<? extends InspectionQuery> missingMetadataQueryType =
        (Class<? extends InspectionQuery>) (Class<?>) MissingMetadataQueryRecord.class;

    IllegalArgumentException missingMetadataFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> InspectionQuery.allowedTargetTypesForType(missingMetadataQueryType));
    assertEquals(
        "No target-type mapping configured for query class "
            + MissingMetadataQueryRecord.class.getName(),
        missingMetadataFailure.getMessage());
  }

  private record MissingMetadataQueryRecord() {}
}
