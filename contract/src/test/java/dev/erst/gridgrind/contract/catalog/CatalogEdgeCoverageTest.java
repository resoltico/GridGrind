package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.catalog.gather.CatalogFieldMetadataSupport;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Function;
import org.junit.jupiter.api.Test;

/** Additional catalog-validation coverage for private and package-private helper seams. */
class CatalogEdgeCoverageTest {
  @Test
  void catalogAndTypeEntriesCoverStrictProtocolVersionAndLookupPaths() {
    TypeEntry requestType = new TypeEntry("REQUEST", "Summary", List.of());
    Catalog catalog =
        new Catalog(
            GridGrindProtocolVersion.current(),
            "type",
            requestType,
            new CliSurface(
                new CliSurface.CliSection("Usage", List.of("gridgrind --help")),
                new CliSurface.CliWorkflowSection(
                    "Workflows",
                    List.of(
                        new CliSurface.WorkflowEntry(
                            "Discover", List.of("gridgrind --print-task-catalog")))),
                new CliSurface.CliSection("Execution", List.of("Executes requests.")),
                new CliSurface.CliDefinitionSection(
                    "Limits", List.of(new CliSurface.DefinitionEntry("One", "Value"))),
                new CliSurface.CliSection("Request", List.of("Request line")),
                new CliSurface.CliDefinitionSection(
                    "Files", List.of(new CliSurface.DefinitionEntry("Input", "Reads input"))),
                new CliSurface.CliTableSection(
                    "Coordinates",
                    "Pattern",
                    "Meaning",
                    List.of(new CliSurface.CoordinateSystemEntry("A1", "Excel"))),
                new CliSurface.CliTemplateSection("Minimal valid request"),
                new CliSurface.CliCommandExample(
                    "Read from stdin", List.of("cat request.json | gridgrind"), null),
                new CliSurface.CliCommandExample(
                    "Run in Docker", List.of("docker run {{CONTAINER_TAG}}"), "Uses the image"),
                new CliSurface.CliDiscoverySection(
                    "Discovery",
                    List.of("List built-in examples"),
                    "Built-in examples",
                    "Print one example",
                    "Protocol catalog note",
                    "gridgrind --print-example WORKBOOK_HEALTH"),
                new CliSurface.CliReferenceSection(
                    "Docs",
                    List.of(new CliSurface.ReferenceEntry("Quick start", "docs/QUICK_START.md"))),
                new CliSurface.CliDefinitionSection(
                    "Flags", List.of(new CliSurface.DefinitionEntry("--help", "Show help"))),
                "msg"),
            List.of(),
            List.of(),
            List.of(),
            List.of(),
            List.of(),
            List.of(),
            List.of(),
            List.of(),
            List.of());

    assertEquals(GridGrindProtocolVersion.current(), catalog.protocolVersion());
    assertTrue(requestType.field("missing").isEmpty());
    assertEquals(
        "protocolVersion must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new Catalog(
                        null,
                        "type",
                        requestType,
                        catalog.cliSurface(),
                        List.of(),
                        List.of(),
                        List.of(),
                        List.of(),
                        List.of(),
                        List.of(),
                        List.of(),
                        List.of(),
                        List.of()))
            .getMessage());
    assertEquals(
        "name must not be null",
        assertThrows(NullPointerException.class, () -> requestType.field(null)).getMessage());
  }

  @Test
  void catalogRecordValidationUsesProductOwnedNullAndBlankMessages() {
    TypeEntry typeEntry = new TypeEntry("REQUEST", "Summary", List.of());

    assertEquals(
        "entries must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () ->
                    CatalogRecordValidation.copyEntries(
                        java.util.Arrays.asList(typeEntry, null), "entries"))
            .getMessage());
    assertEquals(
        "groups must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () ->
                    CatalogRecordValidation.copyGroups(
                        java.util.Arrays.asList((NestedTypeGroup) null), "groups"))
            .getMessage());
    assertEquals(
        "groups must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () ->
                    CatalogRecordValidation.copyPlainGroups(
                        java.util.Arrays.asList((PlainTypeGroup) null), "groups"))
            .getMessage());
    assertEquals(
        "names must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> CatalogRecordValidation.copyStrings(List.of(" "), "names"))
            .getMessage());
    assertEquals(
        "fields must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () ->
                    CatalogRecordValidation.copyFieldEntries(
                        java.util.Arrays.asList((FieldEntry) null), "fields"))
            .getMessage());
  }

  @Test
  void protocolCatalogValidationCoversMissingAnnotationsMismatchesAndDuplicateMaps() {
    assertTrue(
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolCatalog.validateReverseGroupMappings(
                        CatalogFieldMetadataSupport.registeredNestedTypes(), Set.of()))
            .getMessage()
            .startsWith("Field-shape plain group map contains type with no catalog descriptor: "));

    assertTrue(
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolCatalog.validateCoverage(
                        JsonTypeOnlySealedType.class, Map.of(JsonTypeOnlyRecord.class, "ONLY")))
            .getMessage()
            .contains("@JsonSubTypes"));
    assertDoesNotThrow(
        () ->
            GridGrindProtocolCatalog.validateCoverage(
                WrongDiscriminatorAnnotatedSealedType.class,
                Map.of(WrongDiscriminatorRecord.class, "WRONG_FIELD")));
    assertTrue(
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolCatalog.validateCoverage(
                        AnnotatedSealedType.class, Map.of(AnnotatedRecord.class, "WRONG")))
            .getMessage()
            .contains("Catalog id mismatch"));
    assertTrue(
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolCatalog.validateCoverage(AnnotatedSealedType.class, Map.of()))
            .getMessage()
            .contains("Catalog coverage mismatch"));
    assertEquals(
        "Catalog entry %s does not target a record type".formatted(NonRecordSubtype.class),
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolCatalog.validateCoverage(
                        NonRecordAnnotatedSealedType.class,
                        Map.of(NonRecordSubtype.class, "NON_RECORD")))
            .getMessage());

    IllegalStateException duplicateFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalog.toOrderedMap(
                    List.of(new DuplicateFixture("A", "one"), new DuplicateFixture("A", "two")),
                    (Function<DuplicateFixture, String>) DuplicateFixture::id,
                    (Function<DuplicateFixture, String>) DuplicateFixture::value,
                    "fixture"));
    assertEquals(
        "Duplicate fixture detected while building the protocol catalog: one / two",
        duplicateFailure.getMessage());

    assertEquals(
        List.of("stepId", "query"),
        GridGrindProtocolCatalog.requiredFields(InspectionStep.class, List.of("target")));
    assertEquals(
        List.of("stepId", "action"),
        GridGrindProtocolCatalog.requiredFields(MutationStep.class, List.of("target")));
  }

  @Test
  void gridGrindContractTextTypeMapRejectsMissingJsonSubtypes() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> GridGrindContractText.typeNamesByClass(MissingJsonSubtypes.class));
    assertEquals(MissingJsonSubtypes.class + " is missing @JsonSubTypes", failure.getMessage());
  }

  @Test
  void privateCatalogIdGuardAndCliCommandValidationStayCovered() {
    assertEquals(
        "description must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new CliSurface.CliCommandExample("Example", List.of("gridgrind"), " "))
            .getMessage());
    IllegalStateException mismatch =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalog.requireMatchingCatalogId(
                    "BROKEN", "SET_CELL", CellMutationAction.SetCell.class));
    assertTrue(mismatch.getMessage().contains("Catalog type id mismatch"));
  }

  /** Duplicate-id fixture used to cover ordered catalog-map rejection. */
  private record DuplicateFixture(String id, String value) {}

  /** Sealed type missing `@JsonSubTypes` to cover annotation validation. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  private sealed interface JsonTypeOnlySealedType permits JsonTypeOnlyRecord {}

  /** Minimal subtype for missing `@JsonSubTypes` coverage. */
  private record JsonTypeOnlyRecord() implements JsonTypeOnlySealedType {}

  /** Sealed type with a valid subtype id mapping. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({@JsonSubTypes.Type(value = AnnotatedRecord.class, name = "RIGHT")})
  private sealed interface AnnotatedSealedType permits AnnotatedRecord {}

  /** Minimal annotated subtype for catalog-id mismatch coverage. */
  private record AnnotatedRecord() implements AnnotatedSealedType {}

  /** Sealed type using the wrong discriminator field. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
  @JsonSubTypes({@JsonSubTypes.Type(value = WrongDiscriminatorRecord.class, name = "WRONG_FIELD")})
  private sealed interface WrongDiscriminatorAnnotatedSealedType permits WrongDiscriminatorRecord {}

  /** Minimal wrong-discriminator subtype. */
  private record WrongDiscriminatorRecord() implements WrongDiscriminatorAnnotatedSealedType {}

  /** Sealed type whose subtype is not a record. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({@JsonSubTypes.Type(value = NonRecordSubtype.class, name = "NON_RECORD")})
  private sealed interface NonRecordAnnotatedSealedType permits NonRecordSubtype {}

  /** Non-record subtype used to verify coverage rejection. */
  private static final class NonRecordSubtype implements NonRecordAnnotatedSealedType {}

  /** Type with no subtype annotations for contract-text validation coverage. */
  private static final class MissingJsonSubtypes {}
}
