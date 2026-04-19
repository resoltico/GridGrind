package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
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
  void catalogAndTypeEntriesCoverDefaultsAndLookupPaths() {
    TypeEntry requestType = new TypeEntry("REQUEST", "Summary", List.of());
    Catalog catalog =
        new Catalog(
            null,
            "type",
            requestType,
            new CliSurface(List.of(), List.of(), List.of(), List.of(), List.of(), List.of(), "msg"),
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
    assertNull(requestType.field("missing"));
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
