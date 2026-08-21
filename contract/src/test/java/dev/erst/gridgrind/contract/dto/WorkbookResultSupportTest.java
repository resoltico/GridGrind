package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct coverage for package-private support behind the public GridGrind response DTOs. */
class WorkbookResultSupportTest {
  @Test
  void copyOptionalValuesRejectsPresentEmptyListsAndCopiesPresentValues() {
    assertEquals(
        ExecutionJournal.Status.SUCCEEDED,
        WorkbookResult.syntheticSuccessJournal().outcome().status());
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST,
        assertInstanceOf(
                ExecutionJournal.Outcome.Failed.class,
                WorkbookResult.syntheticFailureJournal(GridGrindProblemCode.INVALID_REQUEST)
                    .outcome())
            .problemCode());
    assertEquals(
        ExecutionJournal.Status.FAILED,
        WorkbookResult.syntheticJournal(
                ExecutionJournal.Status.FAILED, Optional.of(GridGrindProblemCode.INVALID_REQUEST))
            .outcome()
            .status());
    assertEquals(
        Optional.of(List.of("Budget")),
        WorkbookResultSupport.copyOptionalValues(Optional.of(List.of("Budget")), "sheetNames"));
    assertEquals(
        "sheetNames must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    WorkbookResultSupport.copyOptionalValues(Optional.of(List.of()), "sheetNames"))
            .getMessage());
  }
}
