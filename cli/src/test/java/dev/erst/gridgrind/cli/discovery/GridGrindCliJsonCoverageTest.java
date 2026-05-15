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
    CliFailureReport cliFailureReport = sampleCliFailureReport();

    assertEquals(
        taskCatalog,
        GridGrindCliJson.readTaskCatalog(GridGrindCliJson.writeTaskCatalogBytes(taskCatalog)));
    assertEquals(
        taskKeywordMatchReport,
        GridGrindCliJson.readTaskKeywordMatchReport(
            GridGrindCliJson.writeTaskKeywordMatchReportBytes(taskKeywordMatchReport)));
    assertEquals(
        exampleCatalog,
        GridGrindCliJson.readShippedExampleCatalog(
            GridGrindCliJson.writeShippedExampleCatalogBytes(exampleCatalog)));
    assertEquals(
        cliFailureReport,
        GridGrindCliJson.readCliFailureReport(
            GridGrindCliJson.writeCliFailureReportBytes(cliFailureReport)));

    try (TrackingInputStream taskCatalogStream =
            new TrackingInputStream(GridGrindCliJson.writeTaskCatalogBytes(taskCatalog));
        TrackingInputStream taskKeywordMatchReportStream =
            new TrackingInputStream(
                GridGrindCliJson.writeTaskKeywordMatchReportBytes(taskKeywordMatchReport));
        TrackingInputStream exampleCatalogStream =
            new TrackingInputStream(
                GridGrindCliJson.writeShippedExampleCatalogBytes(exampleCatalog));
        TrackingInputStream cliFailureReportStream =
            new TrackingInputStream(
                GridGrindCliJson.writeCliFailureReportBytes(cliFailureReport))) {
      assertEquals(taskCatalog, GridGrindCliJson.readTaskCatalog(taskCatalogStream));
      assertEquals(
          taskKeywordMatchReport,
          GridGrindCliJson.readTaskKeywordMatchReport(taskKeywordMatchReportStream));
      assertEquals(
          exampleCatalog, GridGrindCliJson.readShippedExampleCatalog(exampleCatalogStream));
      assertEquals(cliFailureReport, GridGrindCliJson.readCliFailureReport(cliFailureReportStream));
      assertFalse(taskCatalogStream.closed);
      assertFalse(taskKeywordMatchReportStream.closed);
      assertFalse(exampleCatalogStream.closed);
      assertFalse(cliFailureReportStream.closed);
    }
  }

  @Test
  void discoveryWritersAndReadersRejectNullArguments() {
    TaskEntry task = GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow();
    TaskKeywordMatchReport taskKeywordMatchReport = sampleTaskKeywordMatchReport();
    ShippedExampleCatalog exampleCatalog = GridGrindShippedExamples.catalog();
    CliFailureReport cliFailureReport = sampleCliFailureReport();

    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class, () -> GridGrindCliJson.readTaskCatalog((byte[]) null))
            .getMessage());
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.readTaskCatalog((InputStream) null))
            .getMessage());
    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.readTaskKeywordMatchReport((byte[]) null))
            .getMessage());
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.readTaskKeywordMatchReport((InputStream) null))
            .getMessage());
    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.readShippedExampleCatalog((byte[]) null))
            .getMessage());
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.readShippedExampleCatalog((InputStream) null))
            .getMessage());
    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.readCliFailureReport((byte[]) null))
            .getMessage());
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.readCliFailureReport((InputStream) null))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(NullPointerException.class, () -> GridGrindCliJson.writeTaskEntry(null, task))
            .getMessage());
    assertEquals(
        "value must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.writeTaskCatalog(new ByteArrayOutputStream(), null))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.writeTaskKeywordMatchReport(null, taskKeywordMatchReport))
            .getMessage());
    assertEquals(
        "value must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    GridGrindCliJson.writeTaskKeywordMatchReport(new ByteArrayOutputStream(), null))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.writeShippedExampleCatalog(null, exampleCatalog))
            .getMessage());
    assertEquals(
        "value must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    GridGrindCliJson.writeShippedExampleCatalog(new ByteArrayOutputStream(), null))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.writeCliFailureReport(null, cliFailureReport))
            .getMessage());
    assertEquals(
        "value must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.writeCliFailureReport(new ByteArrayOutputStream(), null))
            .getMessage());
  }

  @Test
  void discoverySerializersOmitExplicitNullProperties() throws IOException {
    String taskCatalogJson =
        new String(
            GridGrindCliJson.writeTaskCatalogBytes(GridGrindTaskCatalog.catalog()),
            StandardCharsets.UTF_8);
    String exampleCatalogJson =
        new String(
            GridGrindCliJson.writeShippedExampleCatalogBytes(GridGrindShippedExamples.catalog()),
            StandardCharsets.UTF_8);
    ByteArrayOutputStream taskEntryOutput = new ByteArrayOutputStream();
    GridGrindCliJson.writeTaskEntry(
        taskEntryOutput, GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow());
    String taskEntryJson = taskEntryOutput.toString(StandardCharsets.UTF_8);

    assertFalse(taskCatalogJson.contains(": null"));
    assertFalse(exampleCatalogJson.contains(": null"));
    assertFalse(taskEntryJson.contains(": null"));

    ByteArrayOutputStream taskCatalogOutput = new ByteArrayOutputStream();
    GridGrindCliJson.writeTaskCatalog(taskCatalogOutput, GridGrindTaskCatalog.catalog());
    assertFalse(taskCatalogOutput.toString(StandardCharsets.UTF_8).contains(": null"));

    ByteArrayOutputStream taskKeywordMatchOutput = new ByteArrayOutputStream();
    GridGrindCliJson.writeTaskKeywordMatchReport(
        taskKeywordMatchOutput, sampleTaskKeywordMatchReport());
    assertFalse(taskKeywordMatchOutput.toString(StandardCharsets.UTF_8).contains(": null"));

    ByteArrayOutputStream exampleCatalogOutput = new ByteArrayOutputStream();
    GridGrindCliJson.writeShippedExampleCatalog(
        exampleCatalogOutput, GridGrindShippedExamples.catalog());
    assertFalse(exampleCatalogOutput.toString(StandardCharsets.UTF_8).contains(": null"));

    ByteArrayOutputStream cliFailureReportOutput = new ByteArrayOutputStream();
    GridGrindCliJson.writeCliFailureReport(cliFailureReportOutput, sampleCliFailureReport());
    assertFalse(cliFailureReportOutput.toString(StandardCharsets.UTF_8).contains(": null"));

    assertTrue(
        GridGrindCliJson.readTree("{\"hello\":true}".getBytes(StandardCharsets.UTF_8))
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
        dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_ARGUMENTS,
        "message",
        CliFailureLocation.unavailable(),
        java.util.Optional.of("--query"),
        List.of("gridgrind --print-task-catalog"),
        java.util.Optional.of("Use a fuller query."));
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
