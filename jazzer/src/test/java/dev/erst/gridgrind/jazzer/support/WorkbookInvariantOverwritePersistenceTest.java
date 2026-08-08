package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponsePersistence;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
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
    GridGrindResponse.Failure response =
        GridGrindResponses.failure(
            GridGrindProtocolVersion.V2,
            new GridGrindResponsePersistence.PersistenceOutcome.Overwritten(
                Optional.empty(), new GridGrindResponsePersistence.WriteResult.NotWritten()),
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
