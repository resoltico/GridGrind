package dev.erst.gridgrind.cli.discovery;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Additional round-trip and stream-ownership coverage for the CLI discovery JSON codec. */
class GridGrindCliJsonCoverageTest {
  @Test
  void discoveryRoundTripsDoNotCloseCallerOwnedStreams() throws IOException {
    TaskCatalog taskCatalog = GridGrindTaskCatalog.catalog();
    TaskKeywordMatchReport taskKeywordMatchReport = sampleTaskKeywordMatchReport();
    ShippedExampleCatalog exampleCatalog = GridGrindShippedExamples.catalog();
    ProtocolCatalogSearchReport protocolCatalogSearchReport = sampleProtocolCatalogSearchReport();
    CliFailureReport cliFailureReport = sampleCliFailureReport();

    assertEquals(
        taskCatalog,
        GridGrindCliJson.readBytes(GridGrindCliJson.writeBytes(taskCatalog), TaskCatalog.class));
    assertEquals(
        taskKeywordMatchReport,
        GridGrindCliJson.readBytes(
            GridGrindCliJson.writeBytes(taskKeywordMatchReport), TaskKeywordMatchReport.class));
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
        cliFailureReport,
        GridGrindCliJson.readBytes(
            GridGrindCliJson.writeBytes(cliFailureReport), CliFailureReport.class));

    try (TrackingInputStream taskCatalogStream =
            new TrackingInputStream(GridGrindCliJson.writeBytes(taskCatalog));
        TrackingInputStream taskKeywordMatchReportStream =
            new TrackingInputStream(GridGrindCliJson.writeBytes(taskKeywordMatchReport));
        TrackingInputStream exampleCatalogStream =
            new TrackingInputStream(GridGrindCliJson.writeBytes(exampleCatalog));
        TrackingInputStream protocolCatalogSearchReportStream =
            new TrackingInputStream(
                ProtocolCatalogCliJson.writeProtocolCatalogSearchReportBytes(
                    protocolCatalogSearchReport));
        TrackingInputStream cliFailureReportStream =
            new TrackingInputStream(GridGrindCliJson.writeBytes(cliFailureReport))) {
      assertEquals(taskCatalog, GridGrindCliJsonStreams.readTaskCatalog(taskCatalogStream));
      assertEquals(
          taskKeywordMatchReport,
          GridGrindCliJsonStreams.readTaskKeywordMatchReport(taskKeywordMatchReportStream));
      assertEquals(
          exampleCatalog, GridGrindCliJsonStreams.readShippedExampleCatalog(exampleCatalogStream));
      assertEquals(
          protocolCatalogSearchReport,
          GridGrindCliJsonStreams.readProtocolCatalogSearchReport(
              protocolCatalogSearchReportStream));
      assertEquals(
          cliFailureReport, GridGrindCliJsonStreams.readCliFailureReport(cliFailureReportStream));
      assertFalse(taskCatalogStream.closed);
      assertFalse(taskKeywordMatchReportStream.closed);
      assertFalse(exampleCatalogStream.closed);
      assertFalse(protocolCatalogSearchReportStream.closed);
      assertFalse(cliFailureReportStream.closed);
    }
  }

  @Test
  void discoveryWritersAndReadersRejectNullArguments() {
    TaskEntry task = GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow();
    TaskKeywordMatchReport taskKeywordMatchReport = sampleTaskKeywordMatchReport();
    ShippedExampleCatalog exampleCatalog = GridGrindShippedExamples.catalog();
    ProtocolCatalogSearchReport protocolCatalogSearchReport = sampleProtocolCatalogSearchReport();
    CliFailureReport cliFailureReport = sampleCliFailureReport();

    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.readBytes(null, TaskCatalog.class))
            .getMessage());
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class, () -> GridGrindCliJsonStreams.readTaskCatalog(null))
            .getMessage());
    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.readBytes(null, TaskKeywordMatchReport.class))
            .getMessage());
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJsonStreams.readTaskKeywordMatchReport(null))
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
                () -> GridGrindCliJson.readBytes(null, CliFailureReport.class))
            .getMessage());
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJsonStreams.readCliFailureReport(null))
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
                NullPointerException.class,
                () -> GridGrindCliJson.writeValue(null, cliFailureReport))
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
    GridGrindCliJson.writeValue(taskKeywordMatchOutput, sampleTaskKeywordMatchReport());
    assertFalse(taskKeywordMatchOutput.toString(StandardCharsets.UTF_8).contains(": null"));

    ByteArrayOutputStream exampleCatalogOutput = new ByteArrayOutputStream();
    GridGrindCliJson.writeValue(exampleCatalogOutput, GridGrindShippedExamples.catalog());
    assertFalse(exampleCatalogOutput.toString(StandardCharsets.UTF_8).contains(": null"));

    ByteArrayOutputStream protocolCatalogSearchOutput = new ByteArrayOutputStream();
    ProtocolCatalogCliJson.writeProtocolCatalogSearchReport(
        protocolCatalogSearchOutput, sampleProtocolCatalogSearchReport(), false);
    assertFalse(protocolCatalogSearchOutput.toString(StandardCharsets.UTF_8).contains(": null"));

    ByteArrayOutputStream cliFailureReportOutput = new ByteArrayOutputStream();
    GridGrindCliJson.writeValue(cliFailureReportOutput, sampleCliFailureReport());
    assertFalse(cliFailureReportOutput.toString(StandardCharsets.UTF_8).contains(": null"));

    assertTrue(
        GridGrindCliJsonStreams.readTree("{\"hello\":true}".getBytes(StandardCharsets.UTF_8))
            .path("hello")
            .asBoolean());
  }

  private static TaskKeywordMatchReport sampleTaskKeywordMatchReport() {
    TaskEntry task = GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow();
    return new TaskKeywordMatchReport(
        GridGrindProtocolVersion.current(),
        "Create a monthly sales dashboard with charts",
        List.of("monthly", "sales", "dashboard", "chart"),
        List.of("monthly", "sales"),
        List.of("dashboard", "charts", "summary"),
        List.of(
            new TaskKeywordMatchReport.Candidate(
                task.id(),
                task.narrative().summary(),
                42,
                List.of("dashboard", "chart"),
                List.of("summary", "discovery term"))));
  }

  private static CliFailureReport sampleCliFailureReport() {
    return new CliFailureReport(
        GridGrindProtocolVersion.current(),
        2,
        "print-task-keyword-match",
        "match-query",
        dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_ARGUMENTS,
        "message",
        java.util.Optional.empty(),
        java.util.Optional.of("--query"),
        List.of("gridgrind --print-task-catalog"),
        java.util.Optional.of("Use a fuller query."));
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
