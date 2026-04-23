package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.catalog.CliSurface;
import dev.erst.gridgrind.contract.catalog.GridGrindContractText;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.catalog.ShippedExampleEntry;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import org.junit.jupiter.api.Test;

/** Help, version, and documentation integration tests for GridGrindCli. */
class GridGrindCliHelpTextTest extends GridGrindCliTestSupport {
  @Test
  void versionFrom_returnsUnknown_whenImplementationVersionIsAbsent() {
    assertEquals("unknown", GridGrindCli.versionFrom(null));
  }

  @Test
  void versionFrom_returnsVersion_whenImplementationVersionIsPresent() {
    assertEquals("0.4.1", GridGrindCli.versionFrom("0.4.1"));
  }

  @Test
  void versionFlagPrintsVersionLineToStdoutAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"--version"}, new ByteArrayInputStream(new byte[0]), stdout);

    assertEquals(0, exitCode);
    // Version reads from the JAR manifest Implementation-Version attribute.
    // When running from the test classpath (no JAR), the attribute is absent and "unknown" is used.
    // The description comes from the processed gridgrind.properties resource on the test classpath.
    String output = stdout.toString(StandardCharsets.UTF_8);
    assertTrue(output.startsWith("GridGrind unknown\n"), "must start with GridGrind unknown");
    assertTrue(output.endsWith("\n"), "must end with newline");
    assertTrue(output.lines().count() >= 2, "must have at least two lines");
  }

  @Test
  void licenseFlagPrintsLicenseTextToStdoutAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"--license"}, new ByteArrayInputStream(new byte[0]), stdout);

    assertEquals(0, exitCode);
    String license = stdout.toString(StandardCharsets.UTF_8);
    // The license text is absent from the test classpath; expect the fallback notice.
    assertFalse(license.isBlank());
  }

  @Test
  void licenseText_containsMitLicense_whenResourcePresent() {
    String mit = "MIT License\n\nCopyright (c) 2026 Ervins Strauhmanis\n";
    InputStream own = new ByteArrayInputStream(mit.getBytes(StandardCharsets.UTF_8));

    String result = GridGrindCli.licenseText(own, null, null, null);

    assertTrue(result.contains("MIT License"));
    assertTrue(result.contains("Ervins Strauhmanis"));
    assertFalse(
        result.contains("Third-party notices and licenses:"), "no third-party section when absent");
  }

  @Test
  void licenseText_containsThirdPartySection_whenDependencyLicensesPresent() {
    InputStream own = new ByteArrayInputStream("MIT License\n".getBytes(StandardCharsets.UTF_8));
    InputStream notice = new ByteArrayInputStream("NOTICE info\n".getBytes(StandardCharsets.UTF_8));
    InputStream apache =
        new ByteArrayInputStream("Apache License\n".getBytes(StandardCharsets.UTF_8));
    InputStream bsd = new ByteArrayInputStream("BSD License\n".getBytes(StandardCharsets.UTF_8));

    String result = GridGrindCli.licenseText(own, notice, apache, bsd);

    assertTrue(result.contains("MIT License"));
    assertTrue(result.contains("Third-party notices and licenses:"));
    assertTrue(result.contains("NOTICE info"));
    assertTrue(result.contains("Apache License"));
    assertTrue(result.contains("BSD License"));
  }

  @Test
  void licenseText_returnsFallback_whenAllResourcesAbsent() {
    String result = GridGrindCli.licenseText(null, null, null, null);

    assertFalse(result.isBlank());
    assertTrue(result.contains("not available"));
  }

  @Test
  void licenseText_thirdPartyOnly_whenOwnAbsent() {
    InputStream apache =
        new ByteArrayInputStream("Apache License\n".getBytes(StandardCharsets.UTF_8));

    String result = GridGrindCli.licenseText(null, null, apache, null);

    assertTrue(result.contains("Apache License"));
    assertFalse(result.contains("---"), "no separator when own license is absent");
  }

  @Test
  void licenseText_ensuresTrailingNewline_whenContentLacksIt() {
    InputStream own = new ByteArrayInputStream("MIT License".getBytes(StandardCharsets.UTF_8));

    String result = GridGrindCli.licenseText(own, null, null, null);

    assertTrue(result.endsWith("\n"), "must end with newline even when source text does not");
  }

  @Test
  void licenseText_skipsUnreadableStream() {
    // Pass the broken stream directly to avoid a PMD CloseResource warning;
    // append() closes it via try-with-resources even on IOException.
    String result =
        GridGrindCli.licenseText(
            new InputStream() {
              @Override
              public int read() throws IOException {
                throw new IOException("simulated read failure");
              }

              @Override
              public int read(byte[] buf, int off, int len) throws IOException {
                throw new IOException("simulated read failure");
              }
            },
            null,
            null,
            null);

    // The broken stream is skipped; all streams absent triggers the fallback.
    assertTrue(result.contains("not available"));
  }

  @Test
  void licenseText_containsNotice_whenNoticePresent() {
    InputStream own = new ByteArrayInputStream("MIT License\n".getBytes(StandardCharsets.UTF_8));
    InputStream notice =
        new ByteArrayInputStream("NOTICE content\n".getBytes(StandardCharsets.UTF_8));

    String result = GridGrindCli.licenseText(own, notice, null, null);

    assertTrue(result.contains("Third-party notices and licenses:"));
    assertTrue(result.contains("NOTICE content"));
  }

  @Test
  void productHeader_formatsVersionAndDescription() {
    assertEquals(
        "GridGrind 1.0.0\nA description", GridGrindCli.productHeader("1.0.0", "A description"));
  }

  @Test
  void helpFlagsPrintUsageAndReturnExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int longExitCode =
        new GridGrindCli()
            .run(new String[] {"--help"}, new ByteArrayInputStream(new byte[0]), stdout);

    assertEquals(0, longExitCode);
    String help = stdout.toString(StandardCharsets.UTF_8);
    assertTrue(help.contains("GridGrind"));
    assertTrue(help.contains("Usage:"));
    assertTrue(help.contains("Minimal Valid Request:"));
    assertTrue(help.contains("--request <path>"));
    assertTrue(help.contains("--print-request-template"));
    assertTrue(help.contains("--print-task-catalog"));
    assertTrue(help.contains("--print-task-plan <id>"));
    assertTrue(help.contains("--print-goal-plan <goal>"));
    assertTrue(help.contains("--doctor-request"));
    assertTrue(help.contains("--print-protocol-catalog"));
    assertTrue(help.contains("--print-example <id>"));
    assertTrue(help.contains("--help, -h"));
    assertTrue(help.contains("blob/main/docs/QUICK_REFERENCE.md"));
    assertTrue(help.contains("Coordinate Systems:"));
    assertTrue(
        help.contains(
            GridGrindProtocolCatalog.catalog()
                .cliSurface()
                .fileWorkflow()
                .entries()
                .getLast()
                .value()));
    assertTrue(
        help.contains("type group accepted by polymorphic fields"),
        "help must explain what the protocol catalog publishes");
    assertTrue(
        help.contains("high-level office-work recipes"),
        "help must explain what the task catalog publishes");
    assertTrue(
        help.contains("cellInputTypes:FORMULA"),
        "help must explain how to qualify duplicate catalog ids");

    ByteArrayOutputStream shortStdout = new ByteArrayOutputStream();
    int shortExitCode =
        new GridGrindCli()
            .run(new String[] {"-h"}, new ByteArrayInputStream(new byte[0]), shortStdout);
    assertEquals(0, shortExitCode);
    assertEquals(help, shortStdout.toString(StandardCharsets.UTF_8));
  }

  @Test
  void helpTextUsesVersionedDocumentationRoutesWhenVersionKnown() {
    String help = GridGrindCli.helpText("0.9.0");

    assertTrue(help.contains("GridGrind 0.9.0"));
    assertTrue(help.contains("ghcr.io/resoltico/gridgrind:0.9.0"));
    assertTrue(help.contains("blob/v0.9.0/docs/QUICK_REFERENCE.md"));
    assertTrue(help.contains("blob/v0.9.0/docs/OPERATIONS.md"));
    assertTrue(help.contains("blob/v0.9.0/docs/ERRORS.md"));
  }

  @Test
  void helpTextContainsProductDescription() {
    String help = GridGrindCli.helpText("1.0.0");

    // The description is either the default fallback or the value from the properties resource.
    // Either way the version line and description line must both be present.
    assertTrue(help.contains("GridGrind 1.0.0"), "help must contain the version line");
    // The description line appears immediately after the version line.
    int versionLineEnd = help.indexOf("GridGrind 1.0.0") + "GridGrind 1.0.0".length();
    String afterVersion = help.substring(versionLineEnd).stripLeading();
    assertFalse(
        afterVersion.startsWith("Usage:"),
        "A description line must appear between the version and Usage:");
  }

  @Test
  void helpTextDockerExampleUsesMountedWorkingDirectoryPaths() {
    String help = GridGrindCli.helpText("1.0.0");

    assertTrue(help.contains("-w /workdir"), "Docker example must show -w /workdir");
    assertTrue(
        help.contains("--request request.json"),
        "Docker example must use request paths relative to the mounted workdir");
    assertTrue(
        help.contains("--response response.json"),
        "Docker example must use response paths relative to the mounted workdir");
  }

  @Test
  void helpTextExplainsFileWorkflow() {
    String help = GridGrindCli.helpText("1.0.0");
    CliSurface cliSurface = GridGrindProtocolCatalog.catalog().cliSurface();

    assertTrue(help.contains("File Workflow:"));
    for (CliSurface.DefinitionEntry entry : cliSurface.fileWorkflow().entries()) {
      assertTrue(
          help.contains(entry.label() + ":"),
          () -> "help must include file workflow label: " + entry.label());
      assertTrue(
          help.contains(entry.value()),
          () -> "help must include file workflow value: " + entry.value());
    }
  }

  @Test
  void helpTextIncludesCoordinateSystemsTable() {
    String help = GridGrindCli.helpText("1.0.0");
    CliSurface cliSurface = GridGrindProtocolCatalog.catalog().cliSurface();

    assertTrue(help.contains("Coordinate Systems:"));
    for (CliSurface.CoordinateSystemEntry entry : cliSurface.coordinateSystems().entries()) {
      assertTrue(
          help.contains(entry.pattern()),
          () -> "help must include coordinate pattern " + entry.pattern());
      assertTrue(
          help.contains(entry.convention()),
          () -> "help must include coordinate convention " + entry.convention());
    }
  }

  @Test
  void helpTextListsBuiltInGeneratedExamples() {
    String help = GridGrindCli.helpText("1.0.0");
    CliSurface cliSurface = GridGrindProtocolCatalog.catalog().cliSurface();

    assertTrue(help.contains("Built-in generated examples:"));
    assertTrue(help.contains(cliSurface.discovery().printOneExampleCommand()));
    assertTrue(help.contains(GridGrindContractText.workbookFindingsDiscoverySummary()));
    assertTrue(help.contains(GridGrindContractText.stepKindSummary()));
    for (ShippedExampleEntry example : GridGrindProtocolCatalog.catalog().shippedExamples()) {
      assertTrue(help.contains(example.id()), () -> "help must include example id " + example.id());
      assertTrue(
          help.contains("examples/" + example.fileName()),
          () -> "help must include example file " + example.fileName());
    }
  }

  @Test
  void helpTextDocumentsFormulaAuthoringBoundaries() {
    String help = GridGrindCli.helpText("1.0.0");

    assertTrue(help.contains("Formula authoring:"), "help must include formula-authoring label");
    assertTrue(help.contains(GridGrindContractText.formulaAuthoringLimitSummary()));
    assertTrue(
        help.contains("Loaded formula support:"), "help must include loaded-formula-support label");
    assertTrue(help.contains(GridGrindContractText.loadedFormulaSupportSummary()));
  }

  @Test
  void helpTextIncludesStructuralEditLimitNotes() {
    String help = GridGrindCli.helpText("1.0.0");
    for (CliSurface.DefinitionEntry entry :
        GridGrindProtocolCatalog.catalog().cliSurface().limits().entries()) {
      if ("Row structural edits".equals(entry.label())
          || "Column structural edits".equals(entry.label())
          || "Chart mutations".equals(entry.label())
          || "Chart title formulas".equals(entry.label())
          || "Drawing validation".equals(entry.label())) {
        assertTrue(
            help.contains(entry.label() + ":"),
            () -> "help must include limit label: " + entry.label());
        assertTrue(
            help.contains(entry.value()), () -> "help must include limit value: " + entry.value());
      }
    }
  }

  @Test
  void descriptionFrom_returnsFallback_whenResourceAbsent() {
    // Object.class lives in the bootstrap classloader which has no gridgrind.properties.
    assertEquals("GridGrind", GridGrindCli.descriptionFrom(Object.class));
  }

  @Test
  void descriptionFrom_returnsFallback_whenStreamIsNull() {
    assertEquals("GridGrind", GridGrindCli.descriptionFrom((InputStream) null));
  }

  @Test
  void descriptionFrom_returnsDescription_fromInputStream() {
    InputStream stream =
        new ByteArrayInputStream("description=Custom Description".getBytes(StandardCharsets.UTF_8));
    assertEquals("Custom Description", GridGrindCli.descriptionFrom(stream));
  }

  @Test
  void descriptionFrom_returnsFallback_whenDescriptionIsBlank() {
    InputStream stream =
        new ByteArrayInputStream("description=   ".getBytes(StandardCharsets.UTF_8));
    assertEquals("GridGrind", GridGrindCli.descriptionFrom(stream));
  }

  @Test
  void descriptionFrom_returnsFallback_whenStreamThrowsOnRead() throws IOException {
    try (InputStream broken =
        new InputStream() {
          @Override
          public int read() throws IOException {
            throw new IOException("simulated read failure");
          }

          @Override
          public int read(byte[] bytes, int offset, int length) throws IOException {
            throw new IOException("simulated read failure");
          }
        }) {
      assertEquals("GridGrind", GridGrindCli.descriptionFrom(broken));
    }
  }

  @Test
  void requestTemplateTextRendersUtf8TemplateBytes() {
    assertEquals(
        "{\"protocolVersion\":\"V1\"}",
        GridGrindCli.requestTemplateText(
            () -> "{\"protocolVersion\":\"V1\"}".getBytes(StandardCharsets.UTF_8)));
  }

  @Test
  void requestTemplateTextWrapsTemplateSerializationFailures() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindCli.requestTemplateText(
                    () -> {
                      throw new IOException("synthetic failure");
                    }));

    assertEquals("Failed to render the built-in request template", failure.getMessage());
    assertEquals("synthetic failure", failure.getCause().getMessage());
  }

  @Test
  void printRequestTemplateFlagPrintsValidRequestAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-request-template"},
                new ByteArrayInputStream("ignored".getBytes(StandardCharsets.UTF_8)),
                stdout);

    WorkbookPlan request = GridGrindJson.readRequest(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(GridGrindProtocolCatalog.requestTemplate(), request);
  }

  @Test
  void helpTextMentionsZeroSheetsForNewWorkbook() {
    String help = GridGrindCli.helpText("1.0.0");

    assertTrue(
        help.contains("zero sheets"),
        "help must mention that a NEW workbook starts with zero sheets");
    assertTrue(help.contains("ENSURE_SHEET"), "help must mention ENSURE_SHEET as the remedy");
  }

  @Test
  void helpTextListsKeyLimitsUpfront() {
    String help = GridGrindCli.helpText("1.0.0");

    assertTrue(help.contains(".xlsx"), "help must state the .xlsx-only file format limit");
    assertTrue(help.contains("31"), "help must state the 31-character sheet name limit");
    assertTrue(
        help.contains("250,000"), "help must state the GET_WINDOW / GET_SHEET_SCHEMA cell limit");
    assertTrue(help.contains("255"), "help must state the column width limit");
    assertTrue(help.contains("409"), "help must state the row height limit");
    assertTrue(
        help.contains("AREA, AREA_3D, BAR, BAR_3D")
            && help.contains("SURFACE_3D")
            && help.contains("DOUGHNUT"),
        "help must state the authored chart boundary");
    assertTrue(
        help.contains("NUMBER"),
        "help must note that DATE/DATE_TIME inputs are stored as NUMBER on read-back");
  }

  @Test
  void helpTextMentionsOptionalRequestFields() {
    String help = GridGrindCli.helpText("1.0.0");

    assertTrue(
        help.contains("protocolVersion"), "help must mention that protocolVersion is optional");
    assertTrue(
        help.contains("persistence is optional"), "help must mention that persistence is optional");
    assertTrue(
        help.contains("execution is optional"), "help must mention that execution is optional");
    assertTrue(
        help.contains("formulaEnvironment is optional"),
        "help must mention that formulaEnvironment is optional");
    assertTrue(help.contains("steps is optional"), "help must mention that steps is optional");
    assertTrue(
        help.contains("ASSERTION steps for first-class verification"),
        "help must mention assertion steps");
    assertTrue(
        help.contains("do not send step.type"),
        "help must explain that step kind is inferred instead of authored as step.type");
    assertTrue(
        help.contains("mutations, assertions, and inspections may be interleaved"),
        "help must describe the ordered step model");
    assertTrue(help.contains("EVENT_READ mode"), "help must describe EVENT_READ mode limits");
    assertTrue(
        help.contains("STREAMING_WRITE mode"), "help must describe STREAMING_WRITE mode limits");
    assertTrue(
        help.contains(GridGrindContractText.eventReadInspectionQueryTypePhrase()),
        "help must describe the canonical EVENT_READ read surface");
    assertTrue(
        help.contains(GridGrindContractText.streamingWriteMutationActionTypePhrase()),
        "help must expose the canonical streaming-write operation name");
    assertFalse(
        help.contains("FORCE_FORMULA_RECALC_ON_OPEN"),
        "help must not expose the rejected legacy shorthand");
  }
}
