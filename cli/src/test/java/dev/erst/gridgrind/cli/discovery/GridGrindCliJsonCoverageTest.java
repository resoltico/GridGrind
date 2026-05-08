package dev.erst.gridgrind.cli.discovery;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
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
    TaskPlanTemplate taskPlanTemplate = sampleTaskPlanTemplate();
    TaskKeywordMatchReport taskKeywordMatchReport = sampleTaskKeywordMatchReport();
    ShippedExampleCatalog exampleCatalog = GridGrindShippedExamples.catalog();

    assertEquals(
        taskCatalog,
        GridGrindCliJson.readTaskCatalog(GridGrindCliJson.writeTaskCatalogBytes(taskCatalog)));
    assertEquals(
        taskPlanTemplate,
        GridGrindCliJson.readTaskPlanTemplate(
            GridGrindCliJson.writeTaskPlanTemplateBytes(taskPlanTemplate)));
    assertEquals(
        taskKeywordMatchReport,
        GridGrindCliJson.readTaskKeywordMatchReport(
            GridGrindCliJson.writeTaskKeywordMatchReportBytes(taskKeywordMatchReport)));
    assertEquals(
        exampleCatalog,
        GridGrindCliJson.readShippedExampleCatalog(
            GridGrindCliJson.writeShippedExampleCatalogBytes(exampleCatalog)));

    try (TrackingInputStream taskCatalogStream =
            new TrackingInputStream(GridGrindCliJson.writeTaskCatalogBytes(taskCatalog));
        TrackingInputStream taskPlanStream =
            new TrackingInputStream(GridGrindCliJson.writeTaskPlanTemplateBytes(taskPlanTemplate));
        TrackingInputStream taskKeywordMatchReportStream =
            new TrackingInputStream(
                GridGrindCliJson.writeTaskKeywordMatchReportBytes(taskKeywordMatchReport));
        TrackingInputStream exampleCatalogStream =
            new TrackingInputStream(
                GridGrindCliJson.writeShippedExampleCatalogBytes(exampleCatalog))) {
      assertEquals(taskCatalog, GridGrindCliJson.readTaskCatalog(taskCatalogStream));
      assertEquals(taskPlanTemplate, GridGrindCliJson.readTaskPlanTemplate(taskPlanStream));
      assertEquals(
          taskKeywordMatchReport,
          GridGrindCliJson.readTaskKeywordMatchReport(taskKeywordMatchReportStream));
      assertEquals(
          exampleCatalog, GridGrindCliJson.readShippedExampleCatalog(exampleCatalogStream));
      assertFalse(taskCatalogStream.closed);
      assertFalse(taskPlanStream.closed);
      assertFalse(taskKeywordMatchReportStream.closed);
      assertFalse(exampleCatalogStream.closed);
    }
  }

  @Test
  void discoveryWritersAndReadersRejectNullArguments() {
    TaskEntry task = GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow();
    TaskPlanTemplate taskPlanTemplate = sampleTaskPlanTemplate();
    TaskKeywordMatchReport taskKeywordMatchReport = sampleTaskKeywordMatchReport();
    ShippedExampleCatalog exampleCatalog = GridGrindShippedExamples.catalog();

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
                () -> GridGrindCliJson.readTaskPlanTemplate((byte[]) null))
            .getMessage());
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.readTaskPlanTemplate((InputStream) null))
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
                () -> GridGrindCliJson.writeTaskPlanTemplate(null, taskPlanTemplate))
            .getMessage());
    assertEquals(
        "value must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindCliJson.writeTaskPlanTemplate(new ByteArrayOutputStream(), null))
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

    ByteArrayOutputStream taskPlanOutput = new ByteArrayOutputStream();
    GridGrindCliJson.writeTaskPlanTemplate(taskPlanOutput, sampleTaskPlanTemplate());
    assertFalse(taskPlanOutput.toString(StandardCharsets.UTF_8).contains(": null"));

    ByteArrayOutputStream taskKeywordMatchOutput = new ByteArrayOutputStream();
    GridGrindCliJson.writeTaskKeywordMatchReport(
        taskKeywordMatchOutput, sampleTaskKeywordMatchReport());
    assertFalse(taskKeywordMatchOutput.toString(StandardCharsets.UTF_8).contains(": null"));

    ByteArrayOutputStream exampleCatalogOutput = new ByteArrayOutputStream();
    GridGrindCliJson.writeShippedExampleCatalog(
        exampleCatalogOutput, GridGrindShippedExamples.catalog());
    assertFalse(exampleCatalogOutput.toString(StandardCharsets.UTF_8).contains(": null"));
  }

  private static TaskPlanTemplate sampleTaskPlanTemplate() {
    TaskEntry task = GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow();
    return new TaskPlanTemplate(
        GridGrindProtocolVersion.current(),
        task,
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs("sample-dashboard.xlsx"),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of()),
        List.of("Replace the sample output workbook path before execution."));
  }

  private static TaskKeywordMatchReport sampleTaskKeywordMatchReport() {
    TaskEntry task = GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow();
    TaskPlanTemplate starterTemplate = sampleTaskPlanTemplate();
    return new TaskKeywordMatchReport(
        GridGrindProtocolVersion.current(),
        "Create a monthly sales dashboard with charts",
        List.of("monthly", "sales", "dashboard", "chart"),
        List.of("monthly", "sales"),
        List.of("dashboard", "charts", "summary"),
        List.of(
            new TaskKeywordMatchReport.Candidate(
                task,
                42,
                List.of("dashboard", "chart"),
                List.of("Matched summary \"dashboard\" via dashboard and chart."),
                starterTemplate)));
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
