package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.assertion.CellAssertion;
import dev.erst.gridgrind.contract.dto.CellScalarValue;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import tools.jackson.core.JacksonException;

/** Covers strict CLI host-grant decoding and translation into engine authority. */
class CliExecutionGrantDocumentTest {
  @Test
  void translatesExplicitDocumentPathsAgainstTheCliWorkingDirectory() throws IOException {
    CliExecutionGrantDocument document =
        GridGrindCliJson.readBytes(
            """
            {
              "readableResources": [{ "type": "FILE", "path": "input.xlsx" }],
              "operationIds": ["ENSURE_SHEET"],
              "targetAuthority": { "type": "WORKBOOK_WIDE" },
              "publicationAuthority": {
                "type": "SAVE_AS",
                "path": "out/report.xlsx",
                "ifExists": "REJECT"
              },
              "allowedSecretReferences": [],
              "acceptancePolicy": { "type": "MINIMUM_ONLY" }
            }
            """
                .getBytes(java.nio.charset.StandardCharsets.UTF_8),
            CliExecutionGrantDocument.class);

    GridGrindExecutionGrant.Bounded grant =
        assertInstanceOf(
            GridGrindExecutionGrant.Bounded.class,
            document.toExecutionGrant(Path.of("/work/gridgrind")));
    GridGrindExecutionGrant.ReadAuthority.File readable =
        assertInstanceOf(
            GridGrindExecutionGrant.ReadAuthority.File.class, grant.readableResources().getFirst());
    GridGrindExecutionGrant.PublicationAuthority.SaveAs publication =
        assertInstanceOf(
            GridGrindExecutionGrant.PublicationAuthority.SaveAs.class,
            grant.publicationAuthority());

    assertEquals(Path.of("/work/gridgrind/input.xlsx"), readable.path());
    assertEquals(Path.of("/work/gridgrind/out/report.xlsx"), publication.path());
    assertEquals(WorkbookPlan.WorkbookPersistence.IfExists.REJECT, publication.ifExists());
  }

  @Test
  void rejectsUnknownGrantFields() {
    assertThrows(
        JacksonException.class,
        () ->
            GridGrindCliJson.readBytes(
                """
                {
                  "readableResources": [],
                  "operationIds": [],
                  "targetAuthority": { "type": "WORKBOOK_WIDE" },
                  "publicationAuthority": { "type": "NONE" },
                  "acceptancePolicy": { "type": "MINIMUM_ONLY" },
                  "authority": "agent-authored"
                }
                """
                    .getBytes(java.nio.charset.StandardCharsets.UTF_8),
                CliExecutionGrantDocument.class));
  }

  @Test
  void translatesExplicitOpaquePartPreservation() throws IOException {
    CliExecutionGrantDocument document =
        GridGrindCliJson.readBytes(
            """
            {
              "readableResources": [],
              "operationIds": [],
              "targetAuthority": { "type": "WORKBOOK_WIDE" },
              "publicationAuthority": { "type": "NONE" },
              "allowedSecretReferences": [],
              "acceptancePolicy": {
                "type": "REQUIRE_ALL",
                "requirements": [{
                  "type": "PRESERVE_OPAQUE_OOXML_PARTS",
                  "partNames": ["/customXml/item1.xml"]
                }]
              }
            }
            """
                .getBytes(java.nio.charset.StandardCharsets.UTF_8),
            CliExecutionGrantDocument.class);

    GridGrindExecutionGrant.Bounded grant =
        assertInstanceOf(
            GridGrindExecutionGrant.Bounded.class,
            document.toExecutionGrant(Path.of("/work/gridgrind")));
    GridGrindHostAcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts requirement =
        assertInstanceOf(
            GridGrindHostAcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts.class,
            assertInstanceOf(
                    GridGrindHostAcceptancePolicy.RequireAll.class, grant.hostAcceptancePolicy())
                .requirements()
                .getFirst());

    assertEquals(List.of("/customXml/item1.xml"), requirement.partNames());
  }

  @Test
  void translatesSelectedSheetsAndRequiredCalculation() throws IOException {
    CliExecutionGrantDocument document =
        GridGrindCliJson.readBytes(
            """
            {
              "readableResources": [],
              "operationIds": [],
              "targetAuthority": {
                "type": "SELECTED_SHEETS",
                "sheetNames": ["Summary", "Ledger"]
              },
              "publicationAuthority": { "type": "NONE" },
              "allowedSecretReferences": [],
              "acceptancePolicy": {
                "type": "REQUIRE_ALL",
                "requirements": [{ "type": "REQUIRE_CALCULATION" }]
              }
            }
            """
                .getBytes(java.nio.charset.StandardCharsets.UTF_8),
            CliExecutionGrantDocument.class);

    GridGrindExecutionGrant.Bounded grant =
        assertInstanceOf(
            GridGrindExecutionGrant.Bounded.class,
            document.toExecutionGrant(Path.of("/work/gridgrind")));
    GridGrindExecutionGrant.WorkbookTargetAuthority.SelectedSheets targetAuthority =
        assertInstanceOf(
            GridGrindExecutionGrant.WorkbookTargetAuthority.SelectedSheets.class,
            grant.workbookTargetAuthority());
    assertInstanceOf(
        GridGrindHostAcceptancePolicy.Requirement.RequireCalculation.class,
        assertInstanceOf(
                GridGrindHostAcceptancePolicy.RequireAll.class, grant.hostAcceptancePolicy())
            .requirements()
            .getFirst());

    assertEquals(List.of("Summary", "Ledger"), targetAuthority.sheetNames());
  }

  @Test
  void constructsStandaloneGrantRequirementVariants() {
    CliGrantTargetAuthority.SelectedSheets selectedSheets =
        new CliGrantTargetAuthority.SelectedSheets(List.of("Summary"));
    CliGrantRequiredCalculation requiredCalculation = new CliGrantRequiredCalculation();

    assertEquals(
        List.of("Summary"),
        assertInstanceOf(
                GridGrindExecutionGrant.WorkbookTargetAuthority.SelectedSheets.class,
                selectedSheets.toEngineAuthority())
            .sheetNames());
    assertInstanceOf(
        GridGrindHostAcceptancePolicy.Requirement.RequireCalculation.class,
        requiredCalculation.toEngineRequirement());
  }

  @Test
  void constructsSemanticPreservationRequirement() {
    CliGrantAcceptancePolicy.Requirement.PreserveInspectionFacts preservation =
        new CliGrantAcceptancePolicy.Requirement.PreserveInspectionFacts(
            List.of(
                new InspectionStep(
                    "preserve-summary",
                    new SheetSelector.ByName("Summary"),
                    new SheetIntrospectionQuery.GetSheetSummary())));

    GridGrindHostAcceptancePolicy.Requirement.PreserveInspectionFacts translated =
        assertInstanceOf(
            GridGrindHostAcceptancePolicy.Requirement.PreserveInspectionFacts.class,
            preservation.toEngineRequirement());

    assertEquals(
        List.of("preserve-summary"),
        translated.inspections().stream().map(InspectionStep::stepId).toList());
  }

  @Test
  void constructsTerminalAssertionRequirement() {
    AssertionStep assertion =
        new AssertionStep(
            "assert-owner",
            new CellSelector.ByAddress("Summary", "A1"),
            new CellAssertion.CellValue(new CellScalarValue.Text("Approved")));
    CliGrantAcceptancePolicy.Requirement.TerminalAssertions requirement =
        new CliGrantAcceptancePolicy.Requirement.TerminalAssertions(List.of(assertion));

    GridGrindHostAcceptancePolicy.Requirement.TerminalAssertions translated =
        assertInstanceOf(
            GridGrindHostAcceptancePolicy.Requirement.TerminalAssertions.class,
            requirement.toEngineRequirement());

    assertEquals(
        List.of("assert-owner"),
        translated.assertions().stream().map(AssertionStep::stepId).toList());
  }

  @Test
  void preservesGrantReadPathForTransportDiagnostics() {
    Path grantPath = Path.of("grant.json");
    CliGrantReadException exception =
        new CliGrantReadException(grantPath, new IOException("cannot read grant"));

    assertEquals(grantPath, exception.grantPath());
  }

  @Test
  void translatesEveryPublicationAuthorityVariant() {
    Path workingDirectory = Path.of("/work/gridgrind");

    assertInstanceOf(
        GridGrindExecutionGrant.PublicationAuthority.None.class,
        new CliGrantPublicationAuthority.None().toEngineAuthority(workingDirectory));
    GridGrindExecutionGrant.PublicationAuthority.SaveAs saveAs =
        assertInstanceOf(
            GridGrindExecutionGrant.PublicationAuthority.SaveAs.class,
            new CliGrantPublicationAuthority.SaveAs(
                    "out.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REPLACE)
                .toEngineAuthority(workingDirectory));
    assertEquals(workingDirectory.resolve("out.xlsx"), saveAs.path());
    assertInstanceOf(
        GridGrindExecutionGrant.PublicationAuthority.OverwriteSource.class,
        new CliGrantPublicationAuthority.OverwriteSource().toEngineAuthority(workingDirectory));
  }

  @Test
  void rejectsBlankGrantValues() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CliGrantPublicationAuthority.SaveAs(
                " ", WorkbookPlan.WorkbookPersistence.IfExists.REJECT));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CliGrantTargetAuthority.SelectedSheets(List.of("")));
  }
}
