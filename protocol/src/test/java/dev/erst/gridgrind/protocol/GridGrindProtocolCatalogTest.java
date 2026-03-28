package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import org.junit.jupiter.api.Test;

/** Tests for the machine-readable protocol catalog and built-in request template. */
class GridGrindProtocolCatalogTest {
  @Test
  void exposesTheCurrentMinimalRequestTemplate() throws IOException {
    GridGrindRequest template = GridGrindProtocolCatalog.requestTemplate();
    GridGrindRequest decoded = GridGrindJson.readRequest(GridGrindJson.writeRequestBytes(template));

    assertEquals(GridGrindProtocolVersion.V1, template.protocolVersion());
    assertInstanceOf(GridGrindRequest.WorkbookSource.New.class, template.source());
    assertInstanceOf(GridGrindRequest.WorkbookPersistence.None.class, template.persistence());
    assertEquals(List.of(), template.operations());
    assertEquals(List.of(), template.reads());
    assertEquals(template, decoded);
  }

  @Test
  void exposesTheCurrentProtocolCatalog() throws IOException {
    GridGrindProtocolCatalog.Catalog catalog = GridGrindProtocolCatalog.catalog();
    GridGrindProtocolCatalog.Catalog decoded =
        GridGrindJson.readProtocolCatalog(GridGrindJson.writeProtocolCatalogBytes(catalog));

    assertEquals(GridGrindProtocolVersion.V1, catalog.protocolVersion());
    assertEquals("type", catalog.discriminatorField());
    assertEquals(List.of("NEW", "EXISTING"), ids(catalog.sourceTypes()));
    assertEquals(List.of("NONE", "OVERWRITE", "SAVE_AS"), ids(catalog.persistenceTypes()));
    assertEquals(23, catalog.operationTypes().size());
    assertEquals(16, catalog.readTypes().size());
    assertEquals(
        List.of(
            "cellInputTypes",
            "hyperlinkTargetTypes",
            "cellSelectionTypes",
            "sheetSelectionTypes",
            "namedRangeSelectionTypes",
            "namedRangeScopeTypes",
            "namedRangeSelectorTypes",
            "fontHeightTypes"),
        catalog.nestedTypes().stream()
            .map(GridGrindProtocolCatalog.NestedTypeGroup::group)
            .toList());
    assertEquals(
        List.of(
            "commentInputType",
            "namedRangeTargetType",
            "cellStyleInputType",
            "cellBorderInputType",
            "cellBorderSideInputType"),
        catalog.plainTypes().stream().map(GridGrindProtocolCatalog.PlainTypeGroup::group).toList());

    GridGrindProtocolCatalog.TypeEntry formulaEntry =
        catalog.nestedTypes().stream()
            .filter(g -> "cellInputTypes".equals(g.group()))
            .findFirst()
            .orElseThrow()
            .types()
            .stream()
            .filter(t -> "FORMULA".equals(t.id()))
            .findFirst()
            .orElseThrow();
    assertTrue(
        formulaEntry.summary().contains("Omit the leading = sign"),
        "FORMULA summary should mention omitting leading = sign");

    GridGrindProtocolCatalog.PlainTypeGroup commentGroup =
        catalog.plainTypes().stream()
            .filter(g -> "commentInputType".equals(g.group()))
            .findFirst()
            .orElseThrow();
    assertTrue(commentGroup.type().requiredFields().contains("text"));
    assertTrue(commentGroup.type().requiredFields().contains("author"));
    assertTrue(commentGroup.type().optionalFields().contains("visible"));

    GridGrindProtocolCatalog.PlainTypeGroup namedRangeGroup =
        catalog.plainTypes().stream()
            .filter(g -> "namedRangeTargetType".equals(g.group()))
            .findFirst()
            .orElseThrow();
    assertTrue(namedRangeGroup.type().requiredFields().contains("sheetName"));
    assertTrue(namedRangeGroup.type().requiredFields().contains("range"));

    GridGrindProtocolCatalog.PlainTypeGroup borderSideGroup =
        catalog.plainTypes().stream()
            .filter(g -> "cellBorderSideInputType".equals(g.group()))
            .findFirst()
            .orElseThrow();
    assertTrue(borderSideGroup.type().requiredFields().contains("style"));

    assertEquals(catalog, decoded);
  }

  @Test
  void requiredFieldsSupportsOptionalFieldFiltering() {
    assertEquals(
        List.of("sheetName"),
        GridGrindProtocolCatalog.requiredFields(
            WorkbookOperation.RenameSheet.class, List.of("newSheetName")));
  }

  @Test
  void requiredFieldsRejectsUnknownOptionalFields() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalog.requiredFields(
                    WorkbookOperation.RenameSheet.class, List.of("missingField")));

    assertEquals(
        "Catalog optional field 'missingField' does not exist on "
            + WorkbookOperation.RenameSheet.class.getName(),
        failure.getMessage());
  }

  @Test
  void validateCoverageRejectsTypesWithoutTypeDiscriminatorMetadata() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () -> GridGrindProtocolCatalog.validateCoverage(String.class, Map.of()));

    assertEquals(
        "Catalog coverage requires java.lang.String to use discriminator field 'type'",
        failure.getMessage());
  }

  @Test
  void validateCoverageRejectsTypesWithWrongDiscriminatorField() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalog.validateCoverage(
                    WrongDiscriminatorTaggedUnion.class,
                    Map.of(WrongDiscriminatorTaggedUnion.Option.class, "OPTION")));

    assertEquals(
        "Catalog coverage requires "
            + WrongDiscriminatorTaggedUnion.class.getName()
            + " to use discriminator field 'type'",
        failure.getMessage());
  }

  @Test
  void validateCoverageRejectsTypesWithoutSubtypesMetadata() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalog.validateCoverage(
                    MissingSubtypesTaggedUnion.class,
                    Map.of(MissingSubtypesTaggedUnion.Option.class, "OPTION")));

    assertEquals(
        "Catalog coverage requires @JsonSubTypes on " + MissingSubtypesTaggedUnion.class.getName(),
        failure.getMessage());
  }

  @Test
  void validateCoverageRejectsDuplicateAnnotationEntries() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalog.validateCoverage(
                    DuplicateAnnotatedTaggedUnion.class,
                    Map.of(DuplicateAnnotatedTaggedUnion.Option.class, "OPTION")));

    assertEquals(
        "Duplicate annotation subtype detected while building the protocol catalog: OPTION /"
            + " ALSO_OPTION",
        failure.getMessage());
  }

  @Test
  void validateCoverageRejectsNonRecordCatalogEntries() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalog.validateCoverage(
                    TaggedUnionForValidation.class, Map.of(String.class, "OPTION")));

    assertEquals(
        "Catalog entry class java.lang.String does not target a record type", failure.getMessage());
  }

  @Test
  void validateCoverageRejectsCoverageMismatches() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalog.validateCoverage(
                    TaggedUnionForValidation.class,
                    Map.of(TaggedUnionForValidation.Option.class, "OPTION")));

    assertTrue(
        failure
            .getMessage()
            .startsWith(
                "Catalog coverage mismatch for "
                    + TaggedUnionForValidation.class.getName()
                    + ": annotated="));
  }

  @Test
  void validateCoverageRejectsIdMismatches() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalog.validateCoverage(
                    TaggedUnionForValidation.class,
                    Map.of(
                        TaggedUnionForValidation.Option.class,
                        "WRONG",
                        TaggedUnionForValidation.Alternative.class,
                        "ALTERNATIVE")));

    assertEquals(
        "Catalog id mismatch for "
            + TaggedUnionForValidation.Option.class.getName()
            + ": annotation=OPTION, catalog=WRONG",
        failure.getMessage());
  }

  @Test
  void catalogDefaultsNullProtocolVersionAndCopiesLists() {
    GridGrindProtocolCatalog.TypeEntry sourceType =
        new GridGrindProtocolCatalog.TypeEntry("NEW", "Create", List.of("type"), List.of("notes"));
    GridGrindProtocolCatalog.NestedTypeGroup nestedTypeGroup =
        new GridGrindProtocolCatalog.NestedTypeGroup("cellInputTypes", List.of(sourceType));
    GridGrindProtocolCatalog.PlainTypeGroup plainTypeGroup =
        new GridGrindProtocolCatalog.PlainTypeGroup("commentInputType", sourceType);

    GridGrindProtocolCatalog.Catalog catalog =
        new GridGrindProtocolCatalog.Catalog(
            null,
            "type",
            List.of(sourceType),
            List.of(sourceType),
            List.of(sourceType),
            List.of(sourceType),
            List.of(nestedTypeGroup),
            List.of(plainTypeGroup));

    assertEquals(GridGrindProtocolVersion.current(), catalog.protocolVersion());
    assertThrows(UnsupportedOperationException.class, () -> catalog.sourceTypes().add(sourceType));
    assertThrows(
        UnsupportedOperationException.class, () -> catalog.plainTypes().add(plainTypeGroup));
  }

  @Test
  void publicCatalogRecordsRejectBlankAndNullValues() {
    assertEquals(
        "group must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindProtocolCatalog.NestedTypeGroup(" ", List.of()))
            .getMessage());
    assertEquals(
        "requiredFields must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindProtocolCatalog.TypeEntry(
                        "ID", "Summary", List.of(" "), List.of()))
            .getMessage());
    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindProtocolCatalog.NestedTypeGroup(
                "cellInputTypes", Arrays.asList((GridGrindProtocolCatalog.TypeEntry) null)));
    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindProtocolCatalog.Catalog(
                GridGrindProtocolVersion.current(),
                "type",
                List.of(),
                List.of(),
                List.of(),
                List.of(),
                Arrays.asList((GridGrindProtocolCatalog.NestedTypeGroup) null),
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindProtocolCatalog.Catalog(
                GridGrindProtocolVersion.current(),
                "type",
                List.of(),
                List.of(),
                List.of(),
                List.of(),
                List.of(),
                Arrays.asList((GridGrindProtocolCatalog.PlainTypeGroup) null)));
    assertEquals(
        "group must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindProtocolCatalog.PlainTypeGroup(
                        " ",
                        new GridGrindProtocolCatalog.TypeEntry(
                            "ID", "Summary", List.of(), List.of())))
            .getMessage());
    assertThrows(
        NullPointerException.class,
        () -> new GridGrindProtocolCatalog.PlainTypeGroup("commentInputType", null));
  }

  private static List<String> ids(List<GridGrindProtocolCatalog.TypeEntry> entries) {
    return entries.stream().map(GridGrindProtocolCatalog.TypeEntry::id).toList();
  }

  /** Tagged union with a missing {@link JsonSubTypes} annotation for validator coverage tests. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, include = JsonTypeInfo.As.PROPERTY, property = "type")
  private sealed interface MissingSubtypesTaggedUnion permits MissingSubtypesTaggedUnion.Option {
    record Option() implements MissingSubtypesTaggedUnion {}
  }

  /** Tagged union that intentionally uses the wrong discriminator field for validator tests. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, include = JsonTypeInfo.As.PROPERTY, property = "mode")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = WrongDiscriminatorTaggedUnion.Option.class, name = "OPTION")
  })
  private sealed interface WrongDiscriminatorTaggedUnion
      permits WrongDiscriminatorTaggedUnion.Option {
    record Option() implements WrongDiscriminatorTaggedUnion {}
  }

  /** Tagged union that intentionally repeats one subtype id to exercise duplicate detection. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, include = JsonTypeInfo.As.PROPERTY, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = DuplicateAnnotatedTaggedUnion.Option.class, name = "OPTION"),
    @JsonSubTypes.Type(value = DuplicateAnnotatedTaggedUnion.Option.class, name = "ALSO_OPTION")
  })
  private sealed interface DuplicateAnnotatedTaggedUnion
      permits DuplicateAnnotatedTaggedUnion.Option {
    record Option() implements DuplicateAnnotatedTaggedUnion {}
  }

  /** Tagged union that mirrors the happy-path coverage shape expected by the catalog validator. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, include = JsonTypeInfo.As.PROPERTY, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = TaggedUnionForValidation.Option.class, name = "OPTION"),
    @JsonSubTypes.Type(value = TaggedUnionForValidation.Alternative.class, name = "ALTERNATIVE")
  })
  private sealed interface TaggedUnionForValidation
      permits TaggedUnionForValidation.Option, TaggedUnionForValidation.Alternative {
    record Option() implements TaggedUnionForValidation {}

    record Alternative() implements TaggedUnionForValidation {}
  }
}
