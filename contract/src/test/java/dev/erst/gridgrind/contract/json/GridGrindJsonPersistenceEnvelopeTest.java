package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;

import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import dev.erst.gridgrind.contract.dto.WorkbookResults;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.node.ObjectNode;

/** Locks the top-level-only persistence envelope contract for success and failure responses. */
class GridGrindJsonPersistenceEnvelopeTest {
  @Test
  void successResponsesKeepPersistenceOnlyAtTheTopLevel() throws IOException {
    WorkbookResult persistSuccess =
        WorkbookResults.success(
            GridGrindProtocolVersion.V2,
            new WorkbookResultPersistence.PersistenceOutcome.SavedAs(
                "out/report.xlsx",
                new WorkbookResultPersistence.WriteResult.Written("/work/out/report.xlsx")),
            List.of(),
            List.of(),
            List.of());

    byte[] responseBytes = GridGrindJsonOutput.writeWorkbookResultBytes(persistSuccess);
    String rendered = new String(responseBytes, StandardCharsets.UTF_8);
    ObjectNode responseTree =
        (ObjectNode) GridGrindJsonMapperSupport.JSON_MAPPER.readTree(responseBytes);
    ObjectNode journal = (ObjectNode) responseTree.path("journal");

    assertEquals(1, rendered.split("\"persistence\":", -1).length - 1);
    assertEquals("SAVE_AS", responseTree.path("persistence").path("type").asText());
    assertEquals("WRITTEN", responseTree.path("persistence").path("write").path("status").asText());
    assertEquals(
        "/work/out/report.xlsx",
        responseTree.path("persistence").path("write").path("executionPath").asText());
    assertFalse(journal.has("persistence"));
  }

  @Test
  void persistFailureContextsDoNotDuplicateTheTopLevelPersistenceBlock() throws IOException {
    WorkbookResult persistFailure =
        WorkbookResults.failure(
            GridGrindProtocolVersion.V2,
            new WorkbookResultPersistence.PersistenceOutcome.SavedAs(
                "out/report.xlsx", new WorkbookResultPersistence.WriteResult.NotWritten()),
            new GridGrindProblemDetail.Problem(
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.IO_ERROR,
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.IO_ERROR.category(),
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.IO_ERROR.recovery(),
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.IO_ERROR.title(),
                "write failed",
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.IO_ERROR.resolution(),
                new dev.erst.gridgrind.contract.dto.ProblemContext.PersistWorkbook(
                    dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape
                        .known("NEW", "SAVE_AS"),
                    dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces
                        .PersistenceReference.saveAs("/work/out/report.xlsx")),
                List.of()));

    byte[] responseBytes = GridGrindJsonOutput.writeWorkbookResultBytes(persistFailure);
    String rendered = new String(responseBytes, StandardCharsets.UTF_8);
    ObjectNode responseTree =
        (ObjectNode) GridGrindJsonMapperSupport.JSON_MAPPER.readTree(responseBytes);
    ObjectNode context = (ObjectNode) responseTree.path("problem").path("context");

    assertEquals(1, rendered.split("\"persistence\":", -1).length - 1);
    assertEquals("SAVE_AS", responseTree.path("persistence").path("type").asText());
    assertEquals(
        "NOT_WRITTEN", responseTree.path("persistence").path("write").path("status").asText());
    assertFalse(context.has("persistence"));
    assertEquals("/work/out/report.xlsx", context.path("persistencePath").asText());
    assertFalse(context.has("sourceWorkbookPath"));
  }

  @Test
  void overwriteFailuresWithoutASourcePathStillKeepTheTopLevelPersistenceEnvelope()
      throws IOException {
    WorkbookResult persistFailure =
        WorkbookResults.failure(
            GridGrindProtocolVersion.V2,
            new WorkbookResultPersistence.PersistenceOutcome.Overwritten(
                Optional.empty(), new WorkbookResultPersistence.WriteResult.NotWritten()),
            new GridGrindProblemDetail.Problem(
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_REQUEST,
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_REQUEST.category(),
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_REQUEST.recovery(),
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_REQUEST.title(),
                "OVERWRITE persistence requires an EXISTING source",
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_REQUEST.resolution(),
                new dev.erst.gridgrind.contract.dto.ProblemContext.ValidateRequest(
                    dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape
                        .known("NEW", "OVERWRITE")),
                List.of()));

    byte[] responseBytes = GridGrindJsonOutput.writeWorkbookResultBytes(persistFailure);
    String rendered = new String(responseBytes, StandardCharsets.UTF_8);
    ObjectNode responseTree =
        (ObjectNode) GridGrindJsonMapperSupport.JSON_MAPPER.readTree(responseBytes);

    assertEquals(1, rendered.split("\"persistence\":", -1).length - 1);
    assertEquals("OVERWRITE", responseTree.path("persistence").path("type").asText());
    assertEquals(
        "NOT_WRITTEN", responseTree.path("persistence").path("write").path("status").asText());
    assertFalse(responseTree.path("persistence").has("sourcePath"));
    assertEquals(persistFailure, GridGrindJson.readWorkbookResult(responseBytes));
  }
}
