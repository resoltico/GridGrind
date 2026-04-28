package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct coverage for package-private support behind the public GridGrind response DTOs. */
class GridGrindResponseSupportTest {
  @Test
  void copyOptionalValuesRejectsPresentEmptyListsAndCopiesPresentValues() {
    assertEquals(
        ExecutionJournal.Status.SUCCEEDED,
        GridGrindResponse.syntheticSuccessJournal().outcome().status());
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST,
        GridGrindResponse.syntheticFailureJournal(GridGrindProblemCode.INVALID_REQUEST)
            .outcome()
            .failureCode()
            .orElseThrow());
    assertEquals(
        ExecutionJournal.Status.FAILED,
        GridGrindResponse.syntheticJournal(
                ExecutionJournal.Status.FAILED, Optional.of(GridGrindProblemCode.INVALID_REQUEST))
            .outcome()
            .status());
    assertEquals(
        Optional.of(List.of("Budget")),
        GridGrindResponseSupport.copyOptionalValues(Optional.of(List.of("Budget")), "sheetNames"));
    assertEquals(
        "sheetNames must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    GridGrindResponseSupport.copyOptionalValues(
                        Optional.of(List.of()), "sheetNames"))
            .getMessage());
  }
}
