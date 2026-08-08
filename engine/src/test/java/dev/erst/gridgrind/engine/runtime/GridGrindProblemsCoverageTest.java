package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.assertion.AssertionFailure;
import dev.erst.gridgrind.contract.assertion.CellAssertion;
import dev.erst.gridgrind.contract.dto.CellScalarValue;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.MessageInvariant;
import dev.erst.gridgrind.contract.selector.CellSelector;
import java.nio.file.AccessDeniedException;
import java.nio.file.FileAlreadyExistsException;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Covers payload-location enrichment branches in {@link GridGrindProblems}. */
class GridGrindProblemsCoverageTest {
  @Test
  void enrichContextMapsPayloadMetadataToEachJsonLocationShape() {
    ProblemContext.ReadRequest readContext =
        new ProblemContext.ReadRequest(
            ProblemContextRequestSurfaces.RequestInput.requestFile("/tmp/request.json"),
            ProblemContextRequestSurfaces.JsonLocation.unavailable());

    ProblemContext.ReadRequest pathOnlyContext =
        (ProblemContext.ReadRequest)
            GridGrindProblems.enrichContext(
                readContext,
                badRequest(Optional.of("steps[0].target"), Optional.of(11), Optional.empty()));
    ProblemContext.ReadRequest pathOnlyFromMissingLineContext =
        (ProblemContext.ReadRequest)
            GridGrindProblems.enrichContext(
                readContext,
                badRequest(Optional.of("steps[0].target"), Optional.empty(), Optional.of(7)));
    ProblemContext.ReadRequest lineColumnContext =
        (ProblemContext.ReadRequest)
            GridGrindProblems.enrichContext(
                readContext, badRequest(Optional.empty(), Optional.of(11), Optional.of(7)));
    ProblemContext.ReadRequest locatedContext =
        (ProblemContext.ReadRequest)
            GridGrindProblems.enrichContext(
                readContext,
                badRequest(Optional.of("steps[0].target"), Optional.of(11), Optional.of(7)));
    ProblemContext.ReadRequest unavailableContext =
        (ProblemContext.ReadRequest)
            GridGrindProblems.enrichContext(
                readContext, badRequest(Optional.empty(), Optional.of(11), Optional.empty()));
    ProblemContext.ReadRequest unavailableFromMissingLineContext =
        (ProblemContext.ReadRequest)
            GridGrindProblems.enrichContext(
                readContext, badRequest(Optional.empty(), Optional.empty(), Optional.of(7)));

    assertEquals(Optional.of("steps[0].target"), pathOnlyContext.jsonPath());
    assertEquals(Optional.empty(), pathOnlyContext.jsonLine());
    assertEquals(Optional.empty(), pathOnlyContext.jsonColumn());

    assertEquals(Optional.of("steps[0].target"), pathOnlyFromMissingLineContext.jsonPath());
    assertEquals(Optional.empty(), pathOnlyFromMissingLineContext.jsonLine());
    assertEquals(Optional.empty(), pathOnlyFromMissingLineContext.jsonColumn());

    assertEquals(Optional.empty(), lineColumnContext.jsonPath());
    assertEquals(Optional.of(11), lineColumnContext.jsonLine());
    assertEquals(Optional.of(7), lineColumnContext.jsonColumn());

    assertEquals(Optional.of("steps[0].target"), locatedContext.jsonPath());
    assertEquals(Optional.of(11), locatedContext.jsonLine());
    assertEquals(Optional.of(7), locatedContext.jsonColumn());

    assertEquals(Optional.empty(), unavailableContext.jsonPath());
    assertEquals(Optional.empty(), unavailableContext.jsonLine());
    assertEquals(Optional.empty(), unavailableContext.jsonColumn());

    assertEquals(Optional.empty(), unavailableFromMissingLineContext.jsonPath());
    assertEquals(Optional.empty(), unavailableFromMissingLineContext.jsonLine());
    assertEquals(Optional.empty(), unavailableFromMissingLineContext.jsonColumn());
  }

  @Test
  void fromExceptionUsesCauseSpecificPublicResolutions() {
    ProblemContext.ExecuteStep assertionContext =
        new ProblemContext.ExecuteStep(
            ProblemContextRequestSurfaces.RequestShape.known("NEW", "NONE"),
            new ProblemContextWorkbookSurfaces.StepReference(
                0, "assert-total", "ASSERTION", "EXPECT_CELL_VALUE"),
            ProblemContextWorkbookSurfaces.ProblemLocation.cell("Budget", "B2"));
    ProblemContext.PersistWorkbook saveAsContext =
        new ProblemContext.PersistWorkbook(
            ProblemContextRequestSurfaces.RequestShape.known("EXISTING", "SAVE_AS"),
            ProblemContextWorkbookSurfaces.PersistenceReference.saveAs("/tmp/out.xlsx"));
    ProblemContext.OpenWorkbook openWorkbookContext =
        new ProblemContext.OpenWorkbook(
            ProblemContextRequestSurfaces.RequestShape.known("EXISTING", "NONE"),
            ProblemContextWorkbookSurfaces.WorkbookReference.existingFile("/tmp/source.xlsx"));
    ProblemContext.WriteResponse writeResponseContext =
        new ProblemContext.WriteResponse(
            ProblemContextRequestSurfaces.ResponseOutput.responseFile("/tmp/report.json"));

    GridGrindProblemDetail.Problem assertionFailure =
        GridGrindProblems.fromException(
            new AssertionFailedException(
                "EXPECT_CELL_VALUE mismatched effective values at B2",
                new AssertionFailure(
                    "assert-total",
                    "EXPECT_CELL_VALUE",
                    new CellSelector.ByAddress("Budget", "B2"),
                    new CellAssertion.CellValue(new CellScalarValue.NumberValue(1200.0d)),
                    java.util.List.of())),
            assertionContext);
    GridGrindProblemDetail.Problem saveAsConflict =
        GridGrindProblems.fromException(
            new FileAlreadyExistsException("/tmp/out.xlsx"), saveAsContext);
    GridGrindProblemDetail.Problem openWorkbookDenied =
        GridGrindProblems.fromException(
            new AccessDeniedException("/tmp/source.xlsx"), openWorkbookContext);
    GridGrindProblemDetail.Problem writeResponseDenied =
        GridGrindProblems.fromException(
            new AccessDeniedException("/tmp/report.json"), writeResponseContext);

    assertEquals(GridGrindProblemCode.ASSERTION_FAILED, assertionFailure.code());
    assertEquals(
        "Inspect problem.assertionFailure observations, then adjust the failing assertion or preceding workbook mutations and retry.",
        assertionFailure.resolution());

    assertEquals(GridGrindProblemCode.IO_ERROR, saveAsConflict.code());
    assertEquals(
        "Choose a new SAVE_AS destination path, remove the conflicting file, or set"
            + " SAVE_AS.ifExists=REPLACE before retrying.",
        saveAsConflict.resolution());
    assertEquals(
        "Could not write workbook to /tmp/out.xlsx: already exists; SAVE_AS.ifExists=REJECT"
            + " requires a new destination path. Use ifExists=REPLACE to allow"
            + " create-or-replace.",
        saveAsConflict.message());

    assertEquals(
        "Check the source workbook path, permissions, and file locks before retrying.",
        openWorkbookDenied.resolution());
    assertEquals(
        "Check the --response destination path, parent directory permissions, free disk space, and file locks before retrying.",
        writeResponseDenied.resolution());
  }

  @Test
  void internalProblemsNeverExposeThrowableMessages() {
    ProblemContext.ExecuteRequest context =
        new ProblemContext.ExecuteRequest(ProblemContextRequestSurfaces.RequestShape.unknown());

    GridGrindProblemDetail.Problem classified =
        GridGrindProblems.fromException(new AssertionError("secret runtime detail"), context);
    GridGrindProblemDetail.Problem explicit =
        GridGrindProblems.problem(
            GridGrindProblemCode.INTERNAL_ERROR,
            "another secret runtime detail",
            context,
            new IllegalStateException("third secret runtime detail"));

    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, classified.code());
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR.title(), classified.message());
    assertEquals(
        GridGrindProblemCode.INTERNAL_ERROR.title(), classified.causes().getFirst().message());
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR.title(), explicit.message());
    assertEquals(
        GridGrindProblemCode.INTERNAL_ERROR.title(), explicit.causes().getFirst().message());
  }

  private static InvalidRequestException badRequest(
      Optional<String> jsonPath, Optional<Integer> jsonLine, Optional<Integer> jsonColumn) {
    return new InvalidRequestException(
        new MessageInvariant("bad request", jsonPath),
        jsonPath,
        jsonLine,
        jsonColumn,
        new IllegalArgumentException("bad"));
  }
}
