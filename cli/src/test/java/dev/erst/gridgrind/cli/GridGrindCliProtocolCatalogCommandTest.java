package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.CliDiagnostic;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.cli.discovery.ProtocolCatalogCliJson;
import dev.erst.gridgrind.cli.discovery.ProtocolCatalogIndexReport;
import dev.erst.gridgrind.cli.discovery.ProtocolCatalogSearchReport;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.catalog.GridGrindRequestSurfaceContractText;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.JsonNode;

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
    assertFalse(indexReport.fieldMetadataKeys().isEmpty());
    assertFalse(indexReport.lookupNamespaces().isEmpty());
    assertTrue(
        indexReport.lookupNamespaces().stream()
            .anyMatch(namespace -> "<topLevelGroup>:<id>".equals(namespace.shape())));
    assertTrue(
        indexReport.fieldMetadataKeys().stream()
            .anyMatch(key -> "projectedByFacets".equals(key.name())));
    assertTrue(
        indexReport.fieldMetadataKeys().stream().anyMatch(key -> "noteRefs".equals(key.name())));
    assertTrue(
        indexReport.fieldMetadataKeys().stream()
            .anyMatch(key -> "enumValueDocs".equals(key.name())));
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
  void printProtocolCatalogFullFlagReturnsScopedLookupGuidance() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--full"},
                new ByteArrayInputStream("ignored".getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    assertEquals(2, exitCode);
    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--full"), parseArgumentsContext(failure).argumentName());
    assertEquals("Unknown argument: --full", failure.problem().message());
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
    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertTrue(failure.problem().message().contains("--version"));
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
  void pathBearingLookupPublishesSharedNotesOnceAndCarriesStableNoteRefs() throws IOException {
    String noteText = GridGrindRequestSurfaceContractText.requestOwnedPathResolutionSummary();
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--lookup", "sourceTypes:EXISTING"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    JsonNode output = GridGrindCliJson.readBytes(stdout.toByteArray(), JsonNode.class);
    String rendered = stdout.toString(StandardCharsets.UTF_8);

    assertEquals("EXISTING", output.path("id").asText());
    assertEquals("requestOwnedPathRule", output.path("noteRefs").get(0).asText());
    assertEquals("requestOwnedPathRule", output.path("notes").get(0).path("id").asText());
    assertEquals(noteText, output.path("notes").get(0).path("text").asText());
    assertEquals(1, occurrenceCount(rendered, noteText));
  }

  @Test
  void appendRowAndSetRangeStepTemplatesPreferTypedCellWrappers() throws IOException {
    JsonNode setRange = protocolCatalogLookupJson("mutationActionTypes:SET_RANGE");
    JsonNode appendRow = protocolCatalogLookupJson("mutationActionTypes:APPEND_ROW");

    assertEquals(
        "TYPED",
        setRange
            .path("stepTemplate")
            .path("template")
            .path("action")
            .path("rows")
            .path("type")
            .asText());
    assertTrue(
        setRange.path("stepTemplate").path("template").path("action").path("rows").has("cells"));
    assertEquals(
        "TYPED",
        appendRow
            .path("stepTemplate")
            .path("template")
            .path("action")
            .path("values")
            .path("type")
            .asText());
    assertTrue(
        appendRow.path("stepTemplate").path("template").path("action").path("values").has("cells"));
  }

  @Test
  void projectionAwareReadCatalogTemplatesStayMinimalAndSparseByDefault() throws IOException {
    JsonNode getCells = protocolCatalogLookupJson("inspectionQueryTypes:GET_CELLS");
    JsonNode getWindow = protocolCatalogLookupJson("inspectionQueryTypes:GET_WINDOW");
    JsonNode getSheetSchema = protocolCatalogLookupJson("inspectionQueryTypes:GET_SHEET_SCHEMA");

    assertFalse(getCells.path("stepTemplate").path("template").path("query").has("projection"));
    assertFalse(getWindow.path("stepTemplate").path("template").path("query").has("projection"));
    assertFalse(getWindow.path("stepTemplate").path("template").path("query").has("includeBlanks"));
    assertFalse(
        getSheetSchema.path("stepTemplate").path("template").path("query").has("projection"));
  }

  @Test
  void ooxmlWriteEncryptionLookupPublishesModeLessStrongOnlyShape() throws IOException {
    JsonNode encryption = protocolCatalogLookupJson("plainTypes:ooxmlEncryptionInputType");
    JsonNode type = encryption.path("type");

    assertFalse(type.path("fields").toString().contains("\"mode\""));
    assertEquals("OPTIONAL", catalogField(type, "cipher").path("requirement").asText());
    assertEquals("OPTIONAL", catalogField(type, "hash").path("requirement").asText());
    assertTrue(type.path("summary").asText().contains("AGILE packages only"));
    assertTrue(type.path("summary").asText().contains("AES_256"));
    assertTrue(type.path("summary").asText().contains("SHA_512"));
  }

  @Test
  void cellReadLookupPublishesFacetGatingMetadata() throws IOException {
    JsonNode cellReports = protocolCatalogLookupJson("nestedTypes:cellReportTypes");
    JsonNode cellValues = protocolCatalogLookupJson("nestedTypes:cellValueReportTypes");
    JsonNode projection = protocolCatalogLookupJson("plainTypes:cellReadProjectionType");
    JsonNode number = catalogType(cellReports, "NUMBER");
    JsonNode formula = catalogType(cellReports, "FORMULA");
    JsonNode text = catalogType(cellReports, "TEXT");
    JsonNode evaluatedText = catalogType(cellValues, "TEXT");
    JsonNode evaluatedNumber = catalogType(cellValues, "NUMBER");
    JsonNode facets = catalogField(projection.path("type"), "facets");

    assertEquals("OPTIONAL", catalogField(number, "displayValue").path("requirement").asText());
    assertEquals(
        "FORMAT", catalogField(number, "displayValue").path("projectedByFacets").get(0).asText());
    assertEquals("OPTIONAL", catalogField(number, "style").path("requirement").asText());
    assertEquals("STYLE", catalogField(number, "style").path("projectedByFacets").get(0).asText());
    assertEquals("OPTIONAL", catalogField(number, "hyperlink").path("requirement").asText());
    assertEquals(
        "HYPERLINK", catalogField(number, "hyperlink").path("projectedByFacets").get(0).asText());
    assertEquals("OPTIONAL", catalogField(number, "comment").path("requirement").asText());
    assertEquals(
        "COMMENT", catalogField(number, "comment").path("projectedByFacets").get(0).asText());
    assertEquals("OPTIONAL", catalogField(number, "numberValue").path("requirement").asText());
    assertEquals(
        "VALUE", catalogField(number, "numberValue").path("projectedByFacets").get(0).asText());
    assertEquals("OPTIONAL", catalogField(number, "temporal").path("requirement").asText());
    assertEquals(
        "TEMPORAL", catalogField(number, "temporal").path("projectedByFacets").get(0).asText());
    assertTrue(number.path("summary").asText().contains("TEMPORAL"));
    assertEquals("OPTIONAL", catalogField(text, "runs").path("requirement").asText());
    assertEquals(
        "RICH_TEXT_RUNS", catalogField(text, "runs").path("projectedByFacets").get(0).asText());
    assertEquals("OPTIONAL", catalogField(formula, "formula").path("requirement").asText());
    assertEquals(
        "FORMULA", catalogField(formula, "formula").path("projectedByFacets").get(0).asText());
    assertEquals("OPTIONAL", catalogField(formula, "evaluation").path("requirement").asText());
    assertEquals(
        "VALUE", catalogField(formula, "evaluation").path("projectedByFacets").get(0).asText());
    assertFalse(catalogField(evaluatedText, "textValue").has("projectedByFacets"));
    assertEquals("OPTIONAL", catalogField(evaluatedText, "runs").path("requirement").asText());
    assertEquals(
        "RICH_TEXT_RUNS",
        catalogField(evaluatedText, "runs").path("projectedByFacets").get(0).asText());
    assertEquals(
        "OPTIONAL", catalogField(evaluatedNumber, "temporal").path("requirement").asText());
    assertEquals(
        "TEMPORAL",
        catalogField(evaluatedNumber, "temporal").path("projectedByFacets").get(0).asText());
    assertEquals("FORMAT", facets.path("enumValueDocs").get(2).path("value").asText());
    assertTrue(
        facets.path("enumValueDocs").get(2).path("summary").asText().contains("displayValue"));
    assertEquals("VALUE", facets.path("enumValueDocs").get(0).path("value").asText());
    assertTrue(
        facets
            .path("enumValueDocs")
            .get(0)
            .path("summary")
            .asText()
            .contains("formula evaluation"));
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
    CliDiagnostic blankOperationFailure = cliDiagnostic(blankOperationStdout.toByteArray());

    assertEquals(2, blankOperationExitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, blankOperationFailure.problem().code());
    assertEquals(
        java.util.Optional.of("--lookup"),
        parseArgumentsContext(blankOperationFailure).argumentName());
    assertTrue(
        blankOperationFailure
            .problem()
            .message()
            .contains("protocol catalog lookup id must not be blank"));

    ByteArrayOutputStream blankSearchStdout = new ByteArrayOutputStream();
    int blankSearchExitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--search", ""},
                InputStream.nullInputStream(),
                blankSearchStdout,
                blankSearchStdout);
    CliDiagnostic blankSearchFailure = cliDiagnostic(blankSearchStdout.toByteArray());

    assertEquals(2, blankSearchExitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, blankSearchFailure.problem().code());
    assertEquals(
        java.util.Optional.of("--search"),
        parseArgumentsContext(blankSearchFailure).argumentName());
    assertTrue(blankSearchFailure.problem().message().contains("search query must not be blank"));
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

  private static JsonNode protocolCatalogLookupJson(String lookupId) throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--lookup", lookupId},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    return GridGrindCliJson.readBytes(stdout.toByteArray(), JsonNode.class);
  }

  private static JsonNode catalogType(JsonNode group, String id) {
    for (JsonNode type : group.path("types")) {
      if (id.equals(type.path("id").asText())) {
        return type;
      }
    }
    throw new IllegalArgumentException("Missing catalog type " + id);
  }

  private static JsonNode catalogField(JsonNode type, String name) {
    for (JsonNode field : type.path("fields")) {
      if (name.equals(field.path("name").asText())) {
        return field;
      }
    }
    throw new IllegalArgumentException("Missing catalog field " + name);
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
    JsonNode output = GridGrindCliJson.readBytes(stdout.toByteArray(), JsonNode.class);
    assertEquals("cellInputTypes", output.path("group").asText());
    assertEquals("type", output.path("discriminatorField").asText());
    assertTrue(stdout.toString(StandardCharsets.UTF_8).contains("\"TEXT\""));
  }

  @Test
  void printProtocolCatalogWithTopLevelGroupFilterReturnsMatchingTopLevelGroup()
      throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--lookup", "mutationActionTypes"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    JsonNode output = GridGrindCliJson.readBytes(stdout.toByteArray(), JsonNode.class);
    assertEquals("mutationActionTypes", output.path("group").asText());
    assertTrue(output.path("types").isArray());
    assertTrue(stdout.toString(StandardCharsets.UTF_8).contains("\"SET_CELL\""));
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
    JsonNode output = GridGrindCliJson.readBytes(stdout.toByteArray(), JsonNode.class);
    assertEquals("chartInputType", output.path("group").asText());
    String rendered = stdout.toString(StandardCharsets.UTF_8);
    assertTrue(rendered.contains("\"ChartInput\""));
    assertTrue(rendered.contains("\"plots\""));
  }

  @Test
  void printProtocolCatalogPrettyFlagIndentsJson() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--pretty"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    String output = stdout.toString(StandardCharsets.UTF_8);
    assertTrue(output.startsWith("{\n"));
    assertTrue(output.contains("\n  \"protocolVersion\" : "));
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
    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertTrue(failure.problem().message().contains("Ambiguous lookup id: FORMULA"));
    assertTrue(failure.problem().message().contains("cellInputTypes:FORMULA"));
    assertTrue(failure.problem().message().contains("namedRangeReportTypes:FORMULA"));
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
    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertTrue(failure.problem().message().contains("BOGUS_XYZ"));
    assertTrue(
        failure.problem().message().contains("--print-protocol-catalog --search \"BOGUS_XYZ\""));
    assertTrue(failure.problem().message().contains("--print-protocol-catalog"));
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

    CliDiagnostic failure = cliDiagnostic(Files.readAllBytes(responsePath));
    CliDiagnostic stderrDiagnostic = cliDiagnosticOnStderr(stderr);

    assertEquals(2, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertEquals(failure, stderrDiagnostic);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertTrue(failure.problem().message().contains("BOGUS_XYZ"));
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
    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertTrue(failure.problem().message().contains("--lookup"));
    assertTrue(failure.problem().message().contains("--search"));
  }

  private static int occurrenceCount(String text, String needle) {
    return text.split(java.util.regex.Pattern.quote(needle), -1).length - 1;
  }
}
