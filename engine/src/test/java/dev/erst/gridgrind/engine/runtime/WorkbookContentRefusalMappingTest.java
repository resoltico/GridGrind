package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape;
import dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.WorkbookReference;
import dev.erst.gridgrind.core.problem.WorkbookContentIssue;
import dev.erst.gridgrind.core.problem.WorkbookContentLocation;
import dev.erst.gridgrind.engine.backend.xssf.UnexpectedNativeContentException;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookContentRefusalException;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests mapping and safe context wiring; native readers own their actual refusal rules. */
class WorkbookContentRefusalMappingTest {
  @Test
  void knownRefusalPreservesOriginalSourceAndReliableNativeLocationWithoutCauseText() {
    var location =
        new WorkbookContentLocation(
            Optional.empty(), Optional.empty(), Optional.empty(), Optional.of(1), Optional.of(31));
    var nativeFailure =
        new IllegalArgumentException("synthetic-private-content-not-for-diagnostics");
    var refusal =
        new WorkbookContentRefusalException(
            WorkbookContentIssue.XML_SOURCE_REFUSED, Optional.of(location), nativeFailure);
    var problem =
        GridGrindProblems.fromException(
            refusal,
            new ProblemContext.OpenWorkbook(
                RequestShape.known("EXISTING", "NONE"),
                WorkbookReference.existingFile("original-source.xlsx"),
                Optional.empty()));
    assertEquals(GridGrindProblemCode.INVALID_WORKBOOK_CONTENT, problem.code());
    assertFalse(problem.message().contains("synthetic-private"));
    assertTrue(
        problem.causes().stream()
            .noneMatch(cause -> cause.message().contains("synthetic-private")));
    var context = assertInstanceOf(ProblemContext.OpenWorkbook.class, problem.context());
    assertEquals(Optional.of("original-source.xlsx"), context.sourceWorkbookPath());
    assertEquals(Optional.of(location), context.content());
    assertSame(nativeFailure, refusal.getCause());
  }

  @Test
  void unknownNativeArgumentExceptionRemainsInternalRatherThanAnAuthoredInputFault() {
    var failure = new IllegalArgumentException("synthetic-private-definition-fault");
    var unexpected = new UnexpectedNativeContentException(failure);
    var problem =
        GridGrindProblems.fromException(
            unexpected, new ProblemContext.ExecuteRequest(RequestShape.known("NEW", "NONE")));
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, problem.code());
    assertFalse(problem.message().contains("synthetic-private"));
    assertSame(failure, unexpected.getCause());
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST,
        GridGrindProblems.codeFor(failure),
        "The scoped native wrapper must not change ordinary authored argument validation");
  }
}
