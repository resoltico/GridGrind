package dev.erst.gridgrind.cli.discovery;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.JsonNode;

/** Additional round-trip and stream-ownership coverage for the CLI discovery JSON codec. */
class GridGrindCliJsonCoverageTest {
  @Test
  void discoveryRoundTripsDoNotCloseCallerOwnedStreams() throws IOException {
    TaskCatalog taskCatalog = GridGrindTaskCatalog.catalog();
    RecipeCatalog recipeCatalog = GridGrindRecipeCatalog.catalog();
    RecipeCatalogDetail recipeCatalogDetail =
        GridGrindRecipeCatalog.lookupFor("BUDGET").orElseThrow();
    RecipeCatalogDetail taskRecipeCatalogDetail =
        GridGrindRecipeCatalog.lookupFor("DASHBOARD").orElseThrow();
    RecipeKeywordMatchReport taskKeywordMatchReport = sampleRecipeKeywordMatchReport();
    ShippedExampleCatalog exampleCatalog = GridGrindShippedExamples.catalog();
    ProtocolCatalogSearchReport protocolCatalogSearchReport = sampleProtocolCatalogSearchReport();
    CommandError commandError = sampleCommandError();

    assertEquals(
        taskCatalog,
        GridGrindCliJson.readBytes(GridGrindCliJson.writeBytes(taskCatalog), TaskCatalog.class));
    assertEquals(
        recipeCatalog,
        GridGrindCliJson.readBytes(
            GridGrindCliJson.writeBytes(recipeCatalog), RecipeCatalog.class));
    assertEquals(
        recipeCatalogDetail,
        GridGrindCliJson.readBytes(
            GridGrindCliJson.writeBytes(recipeCatalogDetail), RecipeCatalogDetail.class));
    assertEquals(
        taskRecipeCatalogDetail,
        GridGrindCliJson.readBytes(
            GridGrindCliJson.writeBytes(taskRecipeCatalogDetail), RecipeCatalogDetail.class));
    assertEquals(
        taskKeywordMatchReport,
        GridGrindCliJson.readBytes(
            GridGrindCliJson.writeBytes(taskKeywordMatchReport), RecipeKeywordMatchReport.class));
    assertEquals(
        exampleCatalog,
        GridGrindCliJson.readBytes(
            GridGrindCliJson.writeBytes(exampleCatalog), ShippedExampleCatalog.class));
    assertEquals(
        protocolCatalogSearchReport,
        ProtocolCatalogCliJson.readProtocolCatalogSearchReport(
            ProtocolCatalogCliJson.writeProtocolCatalogSearchReportBytes(
                protocolCatalogSearchReport)));
    assertEquals(
        commandError,
        GridGrindCliJson.readBytes(
            GridGrindCliJson.writeBytes(commandError), CommandError.class));

    try (TrackingInputStream taskCatalogStream =
            new TrackingInputStream(GridGrindCliJson.writeBytes(taskCatalog));
        TrackingInputStream recipeCatalogStream =
            new TrackingInputStream(GridGrindCliJson.writeBytes(recipeCatalog));
        TrackingInputStream taskKeywordMatchReportStream =
            new TrackingInputStream(GridGrindCliJson.writeBytes(taskKeywordMatchReport));
        TrackingInputStream exampleCatalogStream =
            new TrackingInputStream(GridGrindCliJson.writeBytes(exampleCatalog));
        TrackingInputStream protocolCatalogSearchReportStream =
            new TrackingInputStream(
                ProtocolCatalogCliJson.writeProtocolCatalogSearchReportBytes(
                    protocolCatalogSearchReport));
        TrackingInputStream commandErrorStream =
            new TrackingInputStream(GridGrindCliJson.writeBytes(commandError))) {
      assertEquals(taskCatalog, GridGrindCliJsonStreams.readTaskCatalog(taskCatalogStream));
      assertEquals(recipeCatalog, GridGrindCliJsonStreams.readRecipeCatalog(recipeCatalogStream));
      assertEquals(
          taskKeywordMatchReport,
          GridGrindCliJsonStreams.readRecipeKeywordMatchReport(taskKeywordMatchReportStream));
      assertEquals(
          exampleCatalog, GridGrindCliJsonStreams.readShippedExampleCatalog(exampleCatalogStream));
      assertEquals(
          protocolCatalogSearchReport,
          GridGrindCliJsonStreams.readProtocolCatalogSearchReport(
              protocolCatalogSearchReportStream));
      assertEquals(commandError, GridGrindCliJsonStreams.readCommandError(commandErrorStream));
      assertFalse(taskCatalogStream.closed);
      assertFalse(recipeCatalogStream.closed);
      assertFalse(taskKeywordMatchReportStream.closed);
      assertFalse(exampleCatalogStream.closed);
      assertFalse(protocolCatalogSearchReportStream.closed);
      assertFalse(commandErrorStream.closed);
    }
  }

  @Test
  void discoveryWritersAndReadersRejectNullArguments() {
    TaskEntry task = GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow();
    RecipeKeywordMatchReport taskKeywordMatchReport = sampleRecipeKeywordMatchReport();
    ShippedExampleCatalog exampleCatalog = GridGrindShippedExamples.catalog();
    ProtocolCatalogSearchReport protocolCatalogSearchReport = sampleProtocolCatalogSearchReport();
    CommandError commandError = sampleCommandError();

    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.readBytes(null, TaskCatalog.class))
            .getMessage());
    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.readBytes(null, RecipeCatalog.class))
            .getMessage());
    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.readBytes(null, RecipeCatalogDetail.class))
            .getMessage());
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class, () -> GridGrindCliJsonStreams.readTaskCatalog(null))
            .getMessage());
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class, () -> GridGrindCliJsonStreams.readRecipeCatalog(null))
            .getMessage());
    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.readBytes(null, RecipeKeywordMatchReport.class))
            .getMessage());
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJsonStreams.readRecipeKeywordMatchReport(null))
            .getMessage());
    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.readBytes(null, ShippedExampleCatalog.class))
            .getMessage());
    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class,
                () -> ProtocolCatalogCliJson.readProtocolCatalogSearchReport((byte[]) null))
            .getMessage());
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJsonStreams.readShippedExampleCatalog(null))
            .getMessage());
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJsonStreams.readProtocolCatalogSearchReport(null))
            .getMessage());
    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.readBytes(null, CommandError.class))
            .getMessage());
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class, () -> GridGrindCliJsonStreams.readCommandError(null))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(NullPointerException.class, () -> GridGrindCliJson.writeValue(null, task))
            .getMessage());
    assertEquals(
        "value must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.writeValue(new ByteArrayOutputStream(), null))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.writeValue(null, taskKeywordMatchReport))
            .getMessage());
    assertEquals(
        "value must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.writeValue(new ByteArrayOutputStream(), null))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(
                NullPointerException.class, () -> GridGrindCliJson.writeValue(null, exampleCatalog))
            .getMessage());
    assertEquals(
        "value must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.writeValue(new ByteArrayOutputStream(), null))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    ProtocolCatalogCliJson.writeProtocolCatalogSearchReport(
                        null, protocolCatalogSearchReport, false))
            .getMessage());
    assertEquals(
        "value must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    ProtocolCatalogCliJson.writeProtocolCatalogSearchReport(
                        new ByteArrayOutputStream(), null, false))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(
                NullPointerException.class, () -> GridGrindCliJson.writeValue(null, commandError))
            .getMessage());
    assertEquals(
        "value must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.writeValue(new ByteArrayOutputStream(), null))
            .getMessage());
  }

  @Test
  void discoverySerializersOmitExplicitNullProperties() throws IOException {
    String taskCatalogJson =
        new String(
            GridGrindCliJson.writeBytes(GridGrindTaskCatalog.catalog()), StandardCharsets.UTF_8);
    String exampleCatalogJson =
        new String(
            GridGrindCliJson.writeBytes(GridGrindShippedExamples.catalog()),
            StandardCharsets.UTF_8);
    ByteArrayOutputStream taskEntryOutput = new ByteArrayOutputStream();
    GridGrindCliJson.writeValue(
        taskEntryOutput, GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow());
    String taskEntryJson = taskEntryOutput.toString(StandardCharsets.UTF_8);

    assertFalse(taskCatalogJson.contains(": null"));
    assertFalse(exampleCatalogJson.contains(": null"));
    assertFalse(taskEntryJson.contains(": null"));

    ByteArrayOutputStream taskCatalogOutput = new ByteArrayOutputStream();
    GridGrindCliJson.writeValue(taskCatalogOutput, GridGrindTaskCatalog.catalog());
    assertFalse(taskCatalogOutput.toString(StandardCharsets.UTF_8).contains(": null"));

    ByteArrayOutputStream taskKeywordMatchOutput = new ByteArrayOutputStream();
    GridGrindCliJson.writeValue(taskKeywordMatchOutput, sampleRecipeKeywordMatchReport());
    assertFalse(taskKeywordMatchOutput.toString(StandardCharsets.UTF_8).contains(": null"));

    ByteArrayOutputStream exampleCatalogOutput = new ByteArrayOutputStream();
    GridGrindCliJson.writeValue(exampleCatalogOutput, GridGrindShippedExamples.catalog());
    assertFalse(exampleCatalogOutput.toString(StandardCharsets.UTF_8).contains(": null"));

    ByteArrayOutputStream protocolCatalogSearchOutput = new ByteArrayOutputStream();
    ProtocolCatalogCliJson.writeProtocolCatalogSearchReport(
        protocolCatalogSearchOutput, sampleProtocolCatalogSearchReport(), false);
    assertFalse(protocolCatalogSearchOutput.toString(StandardCharsets.UTF_8).contains(": null"));

    ByteArrayOutputStream commandErrorOutput = new ByteArrayOutputStream();
    GridGrindCliJson.writeValue(commandErrorOutput, sampleCommandError());
    assertFalse(commandErrorOutput.toString(StandardCharsets.UTF_8).contains(": null"));

    assertTrue(
        GridGrindCliJsonStreams.readTree("{\"hello\":true}".getBytes(StandardCharsets.UTF_8))
            .path("hello")
            .asBoolean());
  }

  @Test
  void commandErrorJsonKeepsCommandFactsSeparateFromTheSharedProblemCore()
      throws IOException {
    JsonNode parseArgumentsDiagnostic =
        GridGrindCliJsonStreams.readTree(GridGrindCliJson.writeBytes(sampleCommandError()));
    GridGrindProblemDetail.Problem readRequestProblem =
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST_SHAPE,
            "Unknown field 'bogus'",
            new ProblemContext.ReadRequest(
                RequestInput.standardInput(), JsonLocation.located("steps[0].target.type", 7, 13)));
    JsonNode readRequestDiagnostic =
        GridGrindCliJsonStreams.readTree(
            GridGrindCliJson.writeBytes(
                new CommandError(
                    GridGrindProtocolVersion.current(), "execute", List.of(readRequestProblem))));

    assertEquals(
        Set.of("protocolVersion", "command", "problems", "status"),
        fieldNames(parseArgumentsDiagnostic));
    assertFalse(parseArgumentsDiagnostic.has("code"));
    assertFalse(parseArgumentsDiagnostic.has("message"));
    assertFalse(parseArgumentsDiagnostic.has("resolution"));
    assertFalse(parseArgumentsDiagnostic.has("argument"));
    assertFalse(parseArgumentsDiagnostic.has("location"));
    assertFalse(parseArgumentsDiagnostic.has("jsonPath"));
    assertFalse(parseArgumentsDiagnostic.has("exitCode"));
    assertFalse(parseArgumentsDiagnostic.has("suggestions"));
    assertFalse(parseArgumentsDiagnostic.has("transport"));
    assertEquals(
        "NAMED",
        parseArgumentsDiagnostic
            .path("problems")
            .path(0)
            .path("context")
            .path("argument")
            .path("type")
            .asText());
    assertEquals(
        "--query",
        parseArgumentsDiagnostic
            .path("problems")
            .path(0)
            .path("context")
            .path("argument")
            .path("argument")
            .asText());
    assertEquals("REJECTED", parseArgumentsDiagnostic.path("status").asText());

    assertEquals(
        Set.of("protocolVersion", "command", "problems", "status"),
        fieldNames(readRequestDiagnostic));
    assertFalse(readRequestDiagnostic.has("location"));
    assertFalse(readRequestDiagnostic.has("jsonPath"));
    assertFalse(readRequestDiagnostic.has("jsonLine"));
    assertFalse(readRequestDiagnostic.has("jsonColumn"));
    assertEquals(
        "STANDARD_INPUT",
        readRequestDiagnostic
            .path("problems")
            .path(0)
            .path("context")
            .path("request")
            .path("type")
            .asText());
    assertEquals(
        "LOCATED",
        readRequestDiagnostic
            .path("problems")
            .path(0)
            .path("context")
            .path("json")
            .path("type")
            .asText());
    assertEquals(
        "steps[0].target.type",
        readRequestDiagnostic
            .path("problems")
            .path(0)
            .path("context")
            .path("json")
            .path("jsonPath")
            .asText());
    assertEquals(
        7,
        readRequestDiagnostic
            .path("problems")
            .path(0)
            .path("context")
            .path("json")
            .path("jsonLine")
            .asInt());
    assertEquals(
        13,
        readRequestDiagnostic
            .path("problems")
            .path(0)
            .path("context")
            .path("json")
            .path("jsonColumn")
            .asInt());
    assertEquals("REJECTED", readRequestDiagnostic.path("status").asText());
  }

  private static RecipeKeywordMatchReport sampleRecipeKeywordMatchReport() {
    TaskEntry task = GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow();
    return new RecipeKeywordMatchReport(
        GridGrindProtocolVersion.current(),
        "Create a monthly sales dashboard with charts",
        List.of("monthly", "sales", "dashboard", "chart"),
        List.of("monthly", "sales"),
        List.of("dashboard", "charts", "summary"),
        List.of(
            new RecipeKeywordMatchReport.Candidate(
                task.id(),
                RecipeView.TASK_STARTER,
                task.narrative().summary(),
                42,
                List.of("dashboard", "chart"),
                List.of("summary", "discovery term"))));
  }

  private static CommandError sampleCommandError() {
    return new CommandError(
        GridGrindProtocolVersion.current(),
        "print-recipe-keyword-match",
        List.of(
            GridGrindProblemDetail.Problem.of(
                GridGrindProblemCode.INVALID_ARGUMENTS,
                "message",
                new ProblemContext.ParseArguments(CliArgument.named("--query")))));
  }

  private static ProtocolCatalogSearchReport sampleProtocolCatalogSearchReport() {
    return new ProtocolCatalogSearchReport(
        GridGrindProtocolVersion.current(),
        "chart title",
        List.of(
            new ProtocolCatalogSearchHit(
                "mutationActionTypes",
                "SET_CHART",
                "mutationActionTypes:SET_CHART",
                "ENTRY",
                "Create or mutate one supported simple chart on one sheet.",
                List.of("SET_CHART"),
                List.of("chartInputType:ChartInput"))));
  }

  private static Set<String> fieldNames(JsonNode node) {
    Set<String> names = new LinkedHashSet<>();
    for (var property : node.properties()) {
      names.add(property.getKey());
    }
    return names;
  }

  /** Tracks whether the codec attempts to close one caller-owned input stream. */
  private static final class TrackingInputStream extends InputStream {
    private final ByteArrayInputStream delegate;
    private boolean closed;

    private TrackingInputStream(byte[] bytes) {
      this.delegate = new ByteArrayInputStream(bytes);
    }

    @Override
    public int read() {
      return delegate.read();
    }

    @Override
    public int read(byte[] buffer, int offset, int length) {
      return delegate.read(buffer, offset, length);
    }

    @Override
    public void close() {
      closed = true;
    }
  }
}
