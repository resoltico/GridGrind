package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import dev.erst.gridgrind.contract.dto.WorkbookResults;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Locks overwrite-persistence response invariants that depend on request/source pairing. */
class WorkbookInvariantOverwritePersistenceTest {
  @Test
  void acceptsInvalidOverwriteFailuresThatOmitUnavailableSourcePaths() {
    WorkbookPlan request =
        ProtocolStepSupport.request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.Overwrite(),
            List.of(),
            List.of());
    WorkbookResult.Failure response =
        WorkbookResults.failure(
            GridGrindProtocolVersion.V2,
            new WorkbookResultPersistence.PersistenceOutcome.Overwritten(
                Optional.empty(), new WorkbookResultPersistence.WriteResult.NotWritten()),
            GridGrindProblemDetail.Problem.of(
                GridGrindProblemCode.INVALID_REQUEST,
                "OVERWRITE persistence requires an EXISTING source; a NEW workbook has no source"
                    + " file to overwrite",
                new ProblemContext.ValidateRequest(
                    ProblemContextRequestSurfaces.RequestShape.known("NEW", "OVERWRITE"))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
    assertDoesNotThrow(
        () -> WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response));
  }
}
