package dev.erst.gridgrind.json.codec;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.PresenceAssertion;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;
import tools.jackson.core.JacksonException;

/** Verifies the authority document binds directly to the canonical core model. */
class GrantCodecTest {
  private static final Path ROOT = Path.of("/work/gridgrind").toAbsolutePath().normalize();
  private static final String WORKBOOK = "{\"type\":\"WORKBOOK_WIDE\"}";
  private static final String NONE = "{\"type\":\"NONE\"}";
  private static final String MINIMUM = "{\"type\":\"MINIMUM_ONLY\"}";

  @Test
  void acceptanceCannotBindAnInspectionAsATerminalAssertion() {
    String policy =
        """
        {"type":"REQUIRE_ALL","requirements":[{"type":"TERMINAL_ASSERTIONS","assertions":[{"stepId":"wrong-kind","target":{"type":"WORKBOOK_CURRENT"},"query":{"type":"GET_WORKBOOK_SUMMARY"}}]}]}
        """;
    assertThrows(JacksonException.class, () -> read(grant("[]", WORKBOOK, NONE, policy)));
  }

  @Test
  void acceptanceStepsRoundTripWithoutAnIntermediateHostModel() throws IOException {
    var assertion =
        new AssertionStep(
            new StepId("verify-sheet"),
            new SheetSelector.ByName(new SheetName("Budget")),
            new PresenceAssertion.SheetPresent());
    var inspection =
        new InspectionStep(
            new StepId("capture-summary"),
            new WorkbookSelector.Current(),
            new WorkbookIntrospectionQuery.GetWorkbookSummary());
    var policy =
        new AcceptancePolicy.RequireAll(
            List.of(
                new AcceptancePolicy.Requirement.TerminalAssertions(List.of(assertion)),
                new AcceptancePolicy.Requirement.PreserveInspectionFacts(List.of(inspection))));
    var authority =
        new ExecutionGrant(
            List.of(),
            List.of(),
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new ExecutionGrant.PublicationAuthority.None(),
            List.of(),
            policy);
    assertEquals(
        authority,
        GridGrindJson.readGrant(GridGrindJsonOutput.writeGrantBytes(authority, false), ROOT));
  }

  @Test
  void resolvesPathsAgainstTheExplicitAuthorityRootAndWritesFilesystemStrings() throws IOException {
    String document =
        grant(
            "[{\"type\":\"FILE\",\"path\":\"input.xlsx\"}]",
            WORKBOOK,
            "{\"type\":\"SAVE_AS\",\"path\":\"out/../out/report.xlsx\",\"ifExists\":\"REJECT\"}",
            MINIMUM);
    var authority = read(document);
    var readable =
        assertInstanceOf(
            ExecutionGrant.ReadAuthority.File.class, authority.readableResources().getFirst());
    var publication =
        assertInstanceOf(
            ExecutionGrant.PublicationAuthority.SaveAs.class, authority.publicationAuthority());
    assertEquals(ROOT.resolve("input.xlsx"), readable.path());
    assertEquals(ROOT.resolve("out/report.xlsx"), publication.path());
    assertEquals(WorkbookPlan.WorkbookPersistence.IfExists.REJECT, publication.ifExists());
    var serialized =
        GridGrindJsonMapperSupport.JSON_MAPPER.readTree(
            GridGrindJsonOutput.writeGrantBytes(authority, false));
    assertEquals(
        ROOT.resolve("input.xlsx").toString(),
        serialized.path("readableResources").get(0).path("path").asString());
    assertEquals(
        authority,
        GridGrindJson.readGrant(
            GridGrindJsonOutput.writeGrantBytes(authority, false), ROOT.resolve("other")));
  }

  @Test
  void relativePathsDoNotSilentlyUseTheProcessWorkingDirectory() throws IOException {
    String document =
        grant("[{\"type\":\"FILE\",\"path\":\"input.xlsx\"}]", WORKBOOK, NONE, MINIMUM);
    assertThrows(
        JacksonException.class,
        () -> GridGrindJsonMapperSupport.JSON_MAPPER.readValue(document, ExecutionGrant.class));
    var otherRoot = ROOT.resolve("other");
    var authority = GridGrindJson.readGrant(document.getBytes(StandardCharsets.UTF_8), otherRoot);
    assertEquals(
        otherRoot.resolve("input.xlsx"),
        assertInstanceOf(
                ExecutionGrant.ReadAuthority.File.class, authority.readableResources().getFirst())
            .path());
  }

  @Test
  void rejectsUnknownFieldsWithAnOtherwiseValidAuthority() {
    String document = grant("[]", WORKBOOK, NONE, MINIMUM);
    String unknown =
        document.substring(0, document.lastIndexOf('}')) + ",\"authority\":\"agent-authored\"}";
    assertThrows(JacksonException.class, () -> read(unknown));
  }

  @Test
  void bindsOpaquePartPreservationWithoutTranslation() throws IOException {
    var authority =
        read(
            grant(
                "[]",
                WORKBOOK,
                NONE,
                "{\"type\":\"REQUIRE_ALL\",\"requirements\":[{\"type\":\"PRESERVE_OPAQUE_OOXML_PARTS\",\"partNames\":[\"/customXml/item1.xml\"]}]}"));
    var requirement =
        assertInstanceOf(
            AcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts.class,
            assertInstanceOf(AcceptancePolicy.RequireAll.class, authority.acceptancePolicy())
                .requirements()
                .getFirst());
    assertEquals(List.of("/customXml/item1.xml"), requirement.partNames());
  }

  @Test
  void bindsSelectedSheetsAndRequiredCalculation() throws IOException {
    var authority =
        read(
            grant(
                "[]",
                "{\"type\":\"SELECTED_SHEETS\",\"sheetNames\":[\"Summary\",\"Ledger\"]}",
                NONE,
                "{\"type\":\"REQUIRE_ALL\",\"requirements\":[{\"type\":\"REQUIRE_CALCULATION\"}]}"));
    assertEquals(
        List.of("Summary", "Ledger"),
        assertInstanceOf(
                ExecutionGrant.WorkbookTargetAuthority.SelectedSheets.class,
                authority.targetAuthority())
            .sheetNames());
    assertInstanceOf(
        AcceptancePolicy.Requirement.RequireCalculation.class,
        assertInstanceOf(AcceptancePolicy.RequireAll.class, authority.acceptancePolicy())
            .requirements()
            .getFirst());
  }

  @Test
  void bindsStdinAndOverwriteSourceAuthorities() throws IOException {
    var authority =
        read(
            grant(
                "[{\"type\":\"STANDARD_INPUT\"}]",
                WORKBOOK,
                "{\"type\":\"OVERWRITE_SOURCE\"}",
                MINIMUM));
    assertEquals(List.of(new ExecutionGrant.ReadAuthority.Stdin()), authority.readableResources());
    assertEquals(
        new ExecutionGrant.PublicationAuthority.OverwriteSource(),
        authority.publicationAuthority());
  }

  @ParameterizedTest
  @ValueSource(strings = {"\" \"", "7", "null", "\"\\u0000\""})
  void rejectsInvalidPathLiterals(String literal) {
    String document =
        grant("[{\"type\":\"FILE\",\"path\":" + literal + "}]", WORKBOOK, NONE, MINIMUM);
    assertThrows(JacksonException.class, () -> read(document));
  }

  @Test
  void rejectsBlankSelectedSheetNames() {
    assertThrows(
        JacksonException.class,
        () ->
            read(
                grant(
                    "[]", "{\"type\":\"SELECTED_SHEETS\",\"sheetNames\":[\"\"]}", NONE, MINIMUM)));
  }

  @Test
  void failClosedAuthorityRoundTrips() throws IOException {
    var authority = ExecutionGrant.denied();
    assertEquals(
        authority,
        GridGrindJson.readGrant(GridGrindJsonOutput.writeGrantBytes(authority, false), ROOT));
  }

  private static ExecutionGrant read(String document) throws IOException {
    return GridGrindJson.readGrant(document.getBytes(StandardCharsets.UTF_8), ROOT);
  }

  private static String grant(
      String resources, String target, String publication, String acceptance) {
    return """
        {"readableResources":%s,"operationIds":[],"targetAuthority":%s,"publicationAuthority":%s,"allowedSecretReferences":[],"acceptancePolicy":%s}
        """
        .formatted(resources, target, publication, acceptance);
  }
}
