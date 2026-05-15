package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
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
                new InvalidRequestException(
                    "bad request",
                    Optional.of("steps[0].target"),
                    Optional.of(11),
                    Optional.empty(),
                    new IllegalArgumentException("bad")));
    ProblemContext.ReadRequest pathOnlyFromMissingLineContext =
        (ProblemContext.ReadRequest)
            GridGrindProblems.enrichContext(
                readContext,
                new InvalidRequestException(
                    "bad request",
                    Optional.of("steps[0].target"),
                    Optional.empty(),
                    Optional.of(7),
                    new IllegalArgumentException("bad")));
    ProblemContext.ReadRequest lineColumnContext =
        (ProblemContext.ReadRequest)
            GridGrindProblems.enrichContext(
                readContext,
                new InvalidRequestException(
                    "bad request",
                    Optional.empty(),
                    Optional.of(11),
                    Optional.of(7),
                    new IllegalArgumentException("bad")));
    ProblemContext.ReadRequest locatedContext =
        (ProblemContext.ReadRequest)
            GridGrindProblems.enrichContext(
                readContext,
                new InvalidRequestException(
                    "bad request",
                    Optional.of("steps[0].target"),
                    Optional.of(11),
                    Optional.of(7),
                    new IllegalArgumentException("bad")));
    ProblemContext.ReadRequest unavailableContext =
        (ProblemContext.ReadRequest)
            GridGrindProblems.enrichContext(
                readContext,
                new InvalidRequestException(
                    "bad request",
                    Optional.empty(),
                    Optional.of(11),
                    Optional.empty(),
                    new IllegalArgumentException("bad")));
    ProblemContext.ReadRequest unavailableFromMissingLineContext =
        (ProblemContext.ReadRequest)
            GridGrindProblems.enrichContext(
                readContext,
                new InvalidRequestException(
                    "bad request",
                    Optional.empty(),
                    Optional.empty(),
                    Optional.of(7),
                    new IllegalArgumentException("bad")));

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
}
