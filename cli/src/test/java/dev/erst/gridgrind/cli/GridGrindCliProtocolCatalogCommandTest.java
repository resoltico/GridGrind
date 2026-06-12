package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import dev.erst.gridgrind.cli.discovery.ProtocolCatalogCliJson;
import dev.erst.gridgrind.cli.discovery.ProtocolCatalogIndexReport;
import dev.erst.gridgrind.cli.discovery.ProtocolCatalogSearchReport;
import dev.erst.gridgrind.contract.catalog.Catalog;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Protocol-catalog command integration tests for {@link GridGrindCli}. */
class GridGrindCliProtocolCatalogCommandTest extends GridGrindCliTestSupport {
  @Test
  void printProtocolCatalogFlagPrintsCurrentCatalogAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog"},
                new ByteArrayInputStream("ignored".getBytes(StandardCharsets.UTF_8)),
                stdout);

    ProtocolCatalogIndexReport indexReport =
        ProtocolCatalogCliJson.readProtocolCatalogIndexReport(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(
        GridGrindProtocolCatalog.catalog().protocolVersion(), indexReport.protocolVersion());
    assertEquals("type", indexReport.discriminatorField());
    assertEquals("WorkbookPlan", indexReport.requestTypeId());
    assertFalse(indexReport.topLevelGroups().isEmpty());
    assertFalse(indexReport.lookupNamespaces().isEmpty());
    assertTrue(
        indexReport.lookupNamespaces().stream()
            .anyMatch(namespace -> "<topLevelGroup>:<id>".equals(namespace.shape())));
    assertFalse(
        stdout.toString(StandardCharsets.UTF_8).contains(": null"),
        "compact protocol catalog index output must omit explicit null placeholders");
  }

  @Test
  void printProtocolCatalogCanWriteItsPayloadToAnExplicitResponsePath() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-protocol-catalog-", ".json");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--response", responsePath.toString()},
                InputStream.nullInputStream(),
                stdout);

    ProtocolCatalogIndexReport indexReport =
        ProtocolCatalogCliJson.readProtocolCatalogIndexReport(Files.readAllBytes(responsePath));

    assertEquals(0, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertEquals(
        GridGrindProtocolCatalog.catalog().protocolVersion(), indexReport.protocolVersion());
    assertFalse(indexReport.topLevelGroups().isEmpty());
  }

  @Test
  void printProtocolCatalogFullFlagPrintsCurrentCatalogAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--full"},
                new ByteArrayInputStream("ignored".getBytes(StandardCharsets.UTF_8)),
                stdout);

    Catalog catalog = GridGrindJson.readProtocolCatalog(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(GridGrindProtocolCatalog.catalog(), catalog);
    assertFalse(
        stdout.toString(StandardCharsets.UTF_8).contains(": null"),
        "full protocol catalog output must omit explicit null placeholders");
  }

  @Test
  void printProtocolCatalogWithUnexpectedTrailingArgReturnsError() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--version"},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    assertEquals(2, exitCode);
    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertTrue(failure.message().contains("--version"));
  }

  @Test
  void printProtocolCatalogWithLookupFilterReturnsMatchingEntry() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--lookup", "SET_CELL"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    String output = stdout.toString(StandardCharsets.UTF_8).trim();
    assertTrue(output.contains("\"SET_CELL\""), "output must contain the entry id");
    assertTrue(
        output.contains("\"fields\""), "filtered catalog output must contain field descriptors");
    assertTrue(
        output.contains("\"targetSelectors\""),
        "filtered catalog output must expose allowed target selector families");
    assertTrue(
        output.contains("\"CellSelector\""),
        "filtered catalog output must identify the target selector family");
    assertFalse(
        output.contains(": null"),
        "filtered catalog entry output must omit explicit null placeholders");
  }

  @Test
  void printProtocolCatalogRejectsBlankOperationAndSearchValuesWithStructuredFailures()
      throws IOException {
    ByteArrayOutputStream blankOperationStdout = new ByteArrayOutputStream();
    int blankOperationExitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--lookup", ""},
                InputStream.nullInputStream(),
                blankOperationStdout,
                blankOperationStdout);
    CliFailureReport blankOperationFailure = cliFailure(blankOperationStdout.toByteArray());

    assertEquals(2, blankOperationExitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, blankOperationFailure.code());
    assertEquals(java.util.Optional.of("--lookup"), blankOperationFailure.argument());
    assertTrue(
        blankOperationFailure.message().contains("protocol catalog lookup id must not be blank"));

    ByteArrayOutputStream blankSearchStdout = new ByteArrayOutputStream();
    int blankSearchExitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--search", ""},
                InputStream.nullInputStream(),
                blankSearchStdout,
                blankSearchStdout);
    CliFailureReport blankSearchFailure = cliFailure(blankSearchStdout.toByteArray());

    assertEquals(2, blankSearchExitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, blankSearchFailure.code());
    assertEquals(java.util.Optional.of("--search"), blankSearchFailure.argument());
    assertTrue(blankSearchFailure.message().contains("search query must not be blank"));
  }

  @Test
  void printProtocolCatalogWithQualifiedLookupFilterReturnsMatchingNestedEntry()
      throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--lookup", "cellInputTypes:FORMULA"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    String output = stdout.toString(StandardCharsets.UTF_8).trim();
    assertTrue(output.contains("\"FORMULA\""), "output must contain the entry id");
    assertTrue(output.contains("\"source\""), "qualified lookup must expose the source field");
    assertFalse(
        output.contains("\"refersToFormula\""),
        "qualified lookup must not silently return the named-range report variant");
  }

  @Test
  void printProtocolCatalogWithNestedGroupFilterReturnsMatchingNestedGroup() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--lookup", "nestedTypes:cellInputTypes"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    String output = stdout.toString(StandardCharsets.UTF_8).trim();
    assertTrue(output.contains("\"group\" : \"cellInputTypes\""));
    assertTrue(output.contains("\"discriminatorField\" : \"type\""));
    assertTrue(output.contains("\"TEXT\""));
  }

  @Test
  void printProtocolCatalogWithPlainGroupFilterReturnsMatchingPlainGroup() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--lookup", "chartInputType"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    String output = stdout.toString(StandardCharsets.UTF_8).trim();
    assertTrue(output.contains("\"group\" : \"chartInputType\""));
    assertTrue(output.contains("\"ChartInput\""));
    assertTrue(output.contains("\"plots\""));
  }

  @Test
  void printProtocolCatalogWithAmbiguousLookupReturnsErrorAndCandidates() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--lookup", "FORMULA"},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    assertEquals(2, exitCode);
    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertTrue(failure.message().contains("Ambiguous lookup id: FORMULA"));
    assertTrue(failure.message().contains("cellInputTypes:FORMULA"));
    assertTrue(failure.message().contains("namedRangeReportTypes:FORMULA"));
  }

  @Test
  void printProtocolCatalogWithSheetLayoutFilterMentionsPresentation() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--lookup", "GET_SHEET_LAYOUT"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    String output = stdout.toString(StandardCharsets.UTF_8).trim();
    assertTrue(output.contains("\"GET_SHEET_LAYOUT\""), "output must contain the entry id");
    assertTrue(output.contains("presentation"), "summary must mention layout.presentation");
    assertTrue(output.contains("outlineLevel"), "summary must mention row/column outline state");
  }

  @Test
  void printProtocolCatalogWithUnknownLookupReturnsError() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--lookup", "BOGUS_XYZ"},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    assertEquals(2, exitCode);
    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertTrue(failure.message().contains("BOGUS_XYZ"));
    assertTrue(failure.message().contains("--print-protocol-catalog --search \"sheet layout\""));
    assertTrue(failure.message().contains("--print-protocol-catalog"));
  }

  @Test
  void printProtocolCatalogWritesStructuredFailuresToTheResponsePathWhenConfigured()
      throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-protocol-error-", ".json");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--print-protocol-catalog",
                  "--lookup",
                  "BOGUS_XYZ",
                  "--response",
                  responsePath.toString()
                },
                InputStream.nullInputStream(),
                stdout,
                stderr);

    CliFailureReport failure = cliFailure(Files.readAllBytes(responsePath));

    assertEquals(2, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertTrue(
        stderr
            .toString(StandardCharsets.UTF_8)
            .contains(
                "GridGrind wrote the CLI failure report to " + responsePath.toAbsolutePath()));
    assertTrue(stderr.toString(StandardCharsets.UTF_8).contains("[INVALID_ARGUMENTS:"));
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertTrue(failure.message().contains("BOGUS_XYZ"));
  }

  @Test
  void printProtocolCatalogWithSearchReturnsRankedMatches() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--search", "sheet layout"},
                InputStream.nullInputStream(),
                stdout);

    ProtocolCatalogSearchReport report =
        ProtocolCatalogCliJson.readProtocolCatalogSearchReport(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals("sheet layout", report.query());
    assertFalse(report.matches().isEmpty());
    assertEquals(
        "inspectionQueryTypes:GET_SHEET_LAYOUT",
        report.matches().stream()
            .map(match -> match.qualifiedId())
            .filter("inspectionQueryTypes:GET_SHEET_LAYOUT"::equals)
            .findFirst()
            .orElseThrow());
    assertEquals(
        "ENTRY",
        report.matches().stream()
            .filter(match -> "inspectionQueryTypes:GET_SHEET_LAYOUT".equals(match.qualifiedId()))
            .findFirst()
            .orElseThrow()
            .kind());
    String output = stdout.toString(StandardCharsets.UTF_8);
    assertFalse(output.contains("\"stepTemplate\""));
    assertFalse(output.contains("\"supportingMatches\""));
  }

  @Test
  void printProtocolCatalogRejectsOperationAndSearchTogether() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--print-protocol-catalog", "--lookup", "SET_CELL", "--search", "cell"
                },
                InputStream.nullInputStream(),
                stdout,
                stderr);

    assertEquals(2, exitCode);
    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertTrue(failure.message().contains("--lookup"));
    assertTrue(failure.message().contains("--search"));
  }
}
