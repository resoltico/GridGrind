package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.result.WorkbookResultPersistence;
import dev.erst.gridgrind.core.result.WorkbookResults;
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
            new WorkbookPlan.WorkbookPersistence.Overwrite(
                dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec.none()),
            List.of(),
            List.of());
    WorkbookResult.Failure response =
        WorkbookResults.failure(
            ProtocolVersion.CURRENT,
            new WorkbookResultPersistence.PersistenceOutcome.Overwritten(
                Optional.empty(), new WorkbookResultPersistence.PublicationOutcome.NotAttempted()),
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
