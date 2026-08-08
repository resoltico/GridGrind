package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.action.CellMutationAction;
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
                    GridGrindProtocolCatalogCoverageValidator.validateReverseGroupMappings(
                        CatalogFieldShapeRegistry.registeredNestedTypes(), Set.of()))
            .getMessage()
            .startsWith("Field-shape plain group map contains type with no catalog descriptor: "));

    assertTrue(
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolCatalogCoverageValidator.validateCoverage(
                        JsonTypeOnlySealedType.class, Map.of(JsonTypeOnlyRecord.class, "ONLY")))
            .getMessage()
            .contains("must declare a non-blank @JsonTypeName or @ProtocolTypeMetadata id"));
    assertDoesNotThrow(
        () ->
            GridGrindProtocolCatalogCoverageValidator.validateCoverage(
                WrongDiscriminatorAnnotatedSealedType.class,
                Map.of(WrongDiscriminatorRecord.class, "WRONG_FIELD")));
    assertTrue(
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolCatalogCoverageValidator.validateCoverage(
                        AnnotatedSealedType.class, Map.of(AnnotatedRecord.class, "WRONG")))
            .getMessage()
            .contains("Catalog id mismatch"));
    assertTrue(
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolCatalogCoverageValidator.validateCoverage(
                        AnnotatedSealedType.class, Map.of()))
            .getMessage()
            .contains("Catalog coverage mismatch"));
    assertEquals(
        "Catalog entry %s does not target a record type".formatted(NonRecordSubtype.class),
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolCatalogCoverageValidator.validateCoverage(
                        NonRecordAnnotatedSealedType.class,
                        Map.of(NonRecordSubtype.class, "NON_RECORD")))
            .getMessage());

    IllegalStateException duplicateFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalogCoverageValidator.toOrderedMap(
                    List.of(new DuplicateFixture("A", "one"), new DuplicateFixture("A", "two")),
                    (Function<DuplicateFixture, String>) DuplicateFixture::id,
                    (Function<DuplicateFixture, String>) DuplicateFixture::value,
                    "fixture"));
    assertEquals(
        "Duplicate fixture detected while building the protocol catalog: one / two",
        duplicateFailure.getMessage());

    assertEquals(
        List.of("stepId", "target", "query"),
        CatalogTypeEntryFactory.requiredFields(InspectionStep.class));
    assertEquals(
        List.of("stepId", "target", "action"),
        CatalogTypeEntryFactory.requiredFields(MutationStep.class));
  }

  @Test
  void gridGrindContractTextTypeMapRejectsMissingJsonSubtypes() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> GridGrindContractText.typeNamesByClass(MissingJsonSubtypes.class));
    assertEquals(
        MissingJsonSubtypes.class
            + " must be sealed or declare @JsonSubTypes for discriminator discovery",
        failure.getMessage());
  }

  @Test
  void privateCatalogIdGuardStaysCovered() {
    IllegalStateException mismatch =
        assertThrows(
            IllegalStateException.class,
            () ->
                CatalogTypeEntryFactory.requireMatchingCatalogId(
                    "BROKEN", "SET_CELL", CellMutationAction.SetCell.class));
    assertTrue(mismatch.getMessage().contains("Catalog type id mismatch"));
  }

  @Test
  void descriptorFactoryHelpersRejectInvalidShapesAndHideIgnoredFields() {
    IllegalStateException missingDiscriminator =
        assertThrows(
            IllegalStateException.class,
            () ->
                CatalogTypeEntryFactory.discriminatorFieldFor(
                    MissingDiscriminatorSealedType.class));
    assertEquals(
        "Catalog coverage requires "
            + MissingDiscriminatorSealedType.class.getName()
            + " to declare a non-blank @JsonTypeInfo property",
        missingDiscriminator.getMessage());

    TypeEntry ignoredFieldEntry =
        CatalogTypeEntryFactory.typeEntry(
            IgnoredFieldRecord.class, "IgnoredFieldRecord", "summary");
    assertEquals(
        List.of("visible"), ignoredFieldEntry.fields().stream().map(FieldEntry::name).toList());
  }

  @Test
  void noteAwareCatalogHelpersResolveValidateAndNormalizeSharedRuleMetadata() {
    TypeEntry directNoted =
        new TypeEntry(
            "DIRECT",
            "summary",
            List.of(),
            List.of(),
            java.util.Optional.empty(),
            List.of("sharedRule"),
            java.util.Optional.empty());
    TypeEntry factoryNoted =
        CatalogTypeEntryFactory.typeEntry(
            NoteAwareRecord.class, "NoteAwareRecord", "summary", List.of("sharedRule"));
    CatalogPlainTypeDescriptor plainDescriptor =
        CatalogTypeEntryFactory.plainTypeDescriptorWithNotes(
            "noteAwareGroup",
            NoteAwareRecord.class,
            "NoteAwareRecord",
            "summary",
            List.of("sharedRule"));
    Catalog catalog =
        new Catalog(
            GridGrindProtocolVersion.current(),
            "type",
            new TypeEntry("REQUEST", "Summary", List.of()),
            List.of(directNoted),
            List.of(),
            List.of(),
            List.of(),
            List.of(),
            List.of(),
            List.of(),
            List.of(new PlainTypeGroup("noteAwareGroup", plainDescriptor.typeEntry())),
            List.of(new CatalogNote("sharedRule", "Shared rule text.")));

    assertEquals(List.of("sharedRule"), directNoted.noteRefs());
    assertEquals(List.of("sharedRule"), factoryNoted.noteRefs());
    assertEquals(List.of("sharedRule"), plainDescriptor.typeEntry().noteRefs());
    assertTrue(
        new TypeEntry(
                "EMPTY",
                "summary",
                List.of(),
                List.of(),
                java.util.Optional.empty(),
                null,
                java.util.Optional.empty())
            .noteRefs()
            .isEmpty());
    assertEquals(
        List.of(new CatalogNote("sharedRule", "Shared rule text.")),
        CatalogNoteResolutionSupport.referencedNotes(
            catalog, new TopLevelTypeGroup("sourceTypes", List.of(directNoted))));
    assertTrue(CatalogNoteResolutionSupport.referencedNotes(catalog, new Object()).isEmpty());
    assertEquals(
        "Shared rule text.",
        CatalogNoteResolutionSupport.referencedNoteText(
            catalog, List.of("sharedRule", "sharedRule")));
    assertEquals(
        "Catalog note id missingRule is not published",
        assertThrows(
                IllegalStateException.class,
                () ->
                    CatalogNoteResolutionSupport.referencedNotes(
                        catalog,
                        new TypeEntry(
                            "BROKEN",
                            "summary",
                            List.of(),
                            List.of(),
                            java.util.Optional.empty(),
                            List.of("missingRule"),
                            java.util.Optional.empty())))
            .getMessage());
    assertEquals(
        "Catalog entry BROKEN references unknown note id missingRule",
        assertThrows(
                IllegalStateException.class,
                () ->
                    CatalogNoteResolutionSupport.validateCatalogNoteRefs(
                        List.of(
                            List.of(
                                new TypeEntry(
                                    "BROKEN",
                                    "summary",
                                    List.of(),
                                    List.of(),
                                    java.util.Optional.empty(),
                                    List.of("missingRule"),
                                    java.util.Optional.empty()))),
                        List.of(),
                        List.of(),
                        List.of(new CatalogNote("sharedRule", "Shared rule text."))))
            .getMessage());
  }

  @Test
  void descriptorMapAndTopLevelGroupGuardsRejectBrokenInputs() {
    IllegalStateException duplicateFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                CatalogDescriptorMaps.uniqueMap(
                    List.of(new DuplicateFixture("A", "one"), new DuplicateFixture("A", "two")),
                    DuplicateFixture::id,
                    DuplicateFixture::value,
                    "duplicate key: "));
    assertEquals("duplicate key: A", duplicateFailure.getMessage());

    assertEquals(
        "typeDescriptors must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new CatalogTopLevelTypeDescriptorGroup(
                        "group", AnnotatedSealedType.class, List.of()))
            .getMessage());
  }

  /** Duplicate-id fixture used to cover ordered catalog-map rejection. */
  private record DuplicateFixture(String id, String value) {}

  /** Minimal record used to cover note-aware type-descriptor helpers. */
  private record NoteAwareRecord(String value) {}

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

  /** Sealed type with a blank discriminator property. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "")
  private sealed interface MissingDiscriminatorSealedType permits MissingDiscriminatorRecord {}

  /** Minimal subtype for missing-discriminator coverage. */
  private record MissingDiscriminatorRecord() implements MissingDiscriminatorSealedType {}

  /** Sealed type whose subtype is not a record. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({@JsonSubTypes.Type(value = NonRecordSubtype.class, name = "NON_RECORD")})
  private sealed interface NonRecordAnnotatedSealedType permits NonRecordSubtype {}

  /** Non-record subtype used to verify coverage rejection. */
  private static final class NonRecordSubtype implements NonRecordAnnotatedSealedType {}

  /** Record used to verify that ignored record components do not leak into the catalog surface. */
  private record IgnoredFieldRecord(
      String visible, @CatalogIgnored String hiddenByCatalog, @JsonIgnore String hiddenByJson) {}

  /** Type with no subtype annotations for contract-text validation coverage. */
  private static final class MissingJsonSubtypes {}
}
