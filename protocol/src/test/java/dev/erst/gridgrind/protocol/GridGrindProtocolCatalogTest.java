package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.io.IOException;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Objects;
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

    assertCatalogInventory(catalog);
    assertCatalogFieldShapes(catalog);
    assertCatalogSummaries(catalog);
    assertCatalogPolymorphicReferences(catalog);

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
  void fieldShapeBuildsNestedListShapesFromParameterizedTypes() {
    ParameterizedType listOfCellInputs =
        new TestParameterizedType(List.class, List.of(CellInput.class));

    assertEquals(
        new GridGrindProtocolCatalog.FieldShape.ListShape(
            new GridGrindProtocolCatalog.FieldShape.NestedTypeGroupRef("cellInputTypes")),
        GridGrindProtocolCatalog.fieldShape(listOfCellInputs));
  }

  @Test
  void fieldShapeRejectsUnsupportedTypeImplementations() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () -> GridGrindProtocolCatalog.fieldShape(new UnsupportedType("custom-type")));

    assertEquals("Unsupported catalog field type: custom-type", failure.getMessage());
  }

  @Test
  void fieldShapeRejectsListShapesWithoutExactlyOneTypeArgument() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalog.fieldShape(
                    new TestParameterizedType(List.class, List.of())));

    assertEquals(
        "List field must declare exactly one type argument: "
            + new TestParameterizedType(List.class, List.of()),
        failure.getMessage());
  }

  @Test
  void fieldShapeRejectsUnsupportedParameterizedTypes() {
    ParameterizedType unsupportedType =
        new TestParameterizedType(Map.class, List.of(String.class, String.class));

    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () -> GridGrindProtocolCatalog.fieldShape(unsupportedType));

    assertEquals(
        "Unsupported parameterized catalog field type: " + unsupportedType, failure.getMessage());
  }

  @Test
  void fieldShapeRejectsUnsupportedClassTypes() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () -> GridGrindProtocolCatalog.fieldShape(java.util.UUID.class));

    assertEquals("Unsupported catalog field type: java.util.UUID", failure.getMessage());
  }

  @Test
  void isNumericTypeRecognizesNumericAndNonNumericClasses() {
    assertTrue(GridGrindProtocolCatalog.isNumericType(java.math.BigInteger.class));
    assertFalse(GridGrindProtocolCatalog.isNumericType(String.class));
  }

  @Test
  void validateNestedTypeGroupMappingRejectsMismatches() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalog.validateNestedTypeGroupMapping(
                    HyperlinkTarget.class, "wrongGroup"));

    assertEquals(
        "Field-shape nested group mapping mismatch for "
            + HyperlinkTarget.class.getName()
            + ": expected=wrongGroup, mapped=hyperlinkTargetTypes",
        failure.getMessage());
  }

  @Test
  void validatePlainTypeGroupMappingRejectsMismatches() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalog.validatePlainTypeGroupMapping(
                    CommentInput.class, "wrongGroup"));

    assertEquals(
        "Field-shape plain group mapping mismatch for "
            + CommentInput.class.getName()
            + ": expected=wrongGroup, mapped=commentInputType",
        failure.getMessage());
  }

  @Test
  void catalogDefaultsNullProtocolVersionAndCopiesLists() {
    GridGrindProtocolCatalog.TypeEntry sourceType =
        new GridGrindProtocolCatalog.TypeEntry(
            "NEW",
            "Create",
            List.of(
                new GridGrindProtocolCatalog.FieldEntry(
                    "type",
                    GridGrindProtocolCatalog.FieldRequirement.REQUIRED,
                    new GridGrindProtocolCatalog.FieldShape.Scalar(
                        GridGrindProtocolCatalog.ScalarType.STRING),
                    List.of())));
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
        "name must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindProtocolCatalog.FieldEntry(
                        " ",
                        GridGrindProtocolCatalog.FieldRequirement.REQUIRED,
                        new GridGrindProtocolCatalog.FieldShape.Scalar(
                            GridGrindProtocolCatalog.ScalarType.STRING),
                        List.of()))
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
                        " ", new GridGrindProtocolCatalog.TypeEntry("ID", "Summary", List.of())))
            .getMessage());
    assertThrows(
        NullPointerException.class,
        () -> new GridGrindProtocolCatalog.PlainTypeGroup("commentInputType", null));
    assertThrows(
        NullPointerException.class,
        () -> new GridGrindProtocolCatalog.TypeEntry("ID", "Summary", null));
    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindProtocolCatalog.FieldEntry(
                "name",
                null,
                new GridGrindProtocolCatalog.FieldShape.Scalar(
                    GridGrindProtocolCatalog.ScalarType.STRING),
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindProtocolCatalog.FieldEntry(
                "name", GridGrindProtocolCatalog.FieldRequirement.REQUIRED, null, List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindProtocolCatalog.FieldEntry(
                "name",
                GridGrindProtocolCatalog.FieldRequirement.REQUIRED,
                new GridGrindProtocolCatalog.FieldShape.Scalar(
                    GridGrindProtocolCatalog.ScalarType.STRING),
                null));
    IllegalStateException duplicateFieldFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                new GridGrindProtocolCatalog.TypeEntry(
                    "ID",
                    "Summary",
                    List.of(
                        new GridGrindProtocolCatalog.FieldEntry(
                            "name",
                            GridGrindProtocolCatalog.FieldRequirement.REQUIRED,
                            new GridGrindProtocolCatalog.FieldShape.Scalar(
                                GridGrindProtocolCatalog.ScalarType.STRING),
                            List.of()),
                        new GridGrindProtocolCatalog.FieldEntry(
                            "name",
                            GridGrindProtocolCatalog.FieldRequirement.OPTIONAL,
                            new GridGrindProtocolCatalog.FieldShape.Scalar(
                                GridGrindProtocolCatalog.ScalarType.STRING),
                            List.of()))));
    assertTrue(duplicateFieldFailure.getMessage().contains("Duplicate fields"));
  }

  @Test
  void entryForReturnsMatchingEntryByOperationId() {
    assertTrue(GridGrindProtocolCatalog.entryFor("SET_CELL").isPresent());
    assertEquals("SET_CELL", GridGrindProtocolCatalog.entryFor("SET_CELL").get().id());
    assertTrue(GridGrindProtocolCatalog.entryFor("GET_WINDOW").isPresent());
    assertTrue(GridGrindProtocolCatalog.entryFor("NEW").isPresent());
    assertTrue(GridGrindProtocolCatalog.entryFor("SAVE_AS").isPresent());
    assertTrue(GridGrindProtocolCatalog.entryFor("FORMULA").isPresent());
    assertTrue(GridGrindProtocolCatalog.entryFor("UNKNOWN_XYZ").isEmpty());
  }

  @Test
  void duplicateEntryFailureReturnsIllegalStateException() {
    IllegalStateException failure =
        GridGrindProtocolCatalog.duplicateEntryFailure("test label", "LEFT", "RIGHT");
    assertTrue(failure.getMessage().contains("test label"));
    assertTrue(failure.getMessage().contains("LEFT"));
    assertTrue(failure.getMessage().contains("RIGHT"));
  }

  @Test
  void duplicateEntryHandlerThrowsIllegalStateThroughTheReturnedMergeFunction() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalog.duplicateEntryHandler("test label")
                    .apply("LEFT", "RIGHT"));
    assertTrue(failure.getMessage().contains("test label"));
    assertTrue(failure.getMessage().contains("LEFT"));
    assertTrue(failure.getMessage().contains("RIGHT"));
  }

  @Test
  void duplicateEntryMergerThrowsIllegalStateWhenAppliedDirectly() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                new GridGrindProtocolCatalog.DuplicateEntryMerger<>("test label")
                    .apply("LEFT", "RIGHT"));
    assertTrue(failure.getMessage().contains("test label"));
    assertTrue(failure.getMessage().contains("LEFT"));
    assertTrue(failure.getMessage().contains("RIGHT"));
  }

  private static void assertCatalogInventory(GridGrindProtocolCatalog.Catalog catalog) {
    assertEquals(GridGrindProtocolVersion.V1, catalog.protocolVersion());
    assertEquals("type", catalog.discriminatorField());
    assertEquals(List.of("NEW", "EXISTING"), ids(catalog.sourceTypes()));
    assertEquals(List.of("NONE", "OVERWRITE", "SAVE_AS"), ids(catalog.persistenceTypes()));
    assertEquals(25, catalog.operationTypes().size());
    assertEquals(18, catalog.readTypes().size());
    assertEquals(
        List.of(
            "cellInputTypes",
            "hyperlinkTargetTypes",
            "cellSelectionTypes",
            "rangeSelectionTypes",
            "sheetSelectionTypes",
            "namedRangeSelectionTypes",
            "namedRangeScopeTypes",
            "namedRangeSelectorTypes",
            "fontHeightTypes",
            "dataValidationRuleTypes"),
        catalog.nestedTypes().stream()
            .map(GridGrindProtocolCatalog.NestedTypeGroup::group)
            .toList());
    assertEquals(
        List.of(
            "commentInputType",
            "namedRangeTargetType",
            "cellStyleInputType",
            "cellBorderInputType",
            "cellBorderSideInputType",
            "dataValidationInputType",
            "dataValidationPromptInputType",
            "dataValidationErrorAlertInputType"),
        catalog.plainTypes().stream().map(GridGrindProtocolCatalog.PlainTypeGroup::group).toList());
  }

  private static void assertCatalogFieldShapes(GridGrindProtocolCatalog.Catalog catalog) {
    GridGrindProtocolCatalog.PlainTypeGroup commentGroup = plainGroup(catalog, "commentInputType");
    assertEquals(
        GridGrindProtocolCatalog.FieldRequirement.REQUIRED,
        fieldNamed(commentGroup.type(), "text").requirement());
    assertEquals(
        new GridGrindProtocolCatalog.FieldShape.Scalar(GridGrindProtocolCatalog.ScalarType.STRING),
        fieldNamed(commentGroup.type(), "text").shape());
    assertEquals(
        GridGrindProtocolCatalog.FieldRequirement.REQUIRED,
        fieldNamed(commentGroup.type(), "author").requirement());
    assertEquals(
        GridGrindProtocolCatalog.FieldRequirement.OPTIONAL,
        fieldNamed(commentGroup.type(), "visible").requirement());
    assertEquals(
        new GridGrindProtocolCatalog.FieldShape.Scalar(GridGrindProtocolCatalog.ScalarType.BOOLEAN),
        fieldNamed(commentGroup.type(), "visible").shape());

    GridGrindProtocolCatalog.PlainTypeGroup namedRangeGroup =
        plainGroup(catalog, "namedRangeTargetType");
    assertEquals(
        GridGrindProtocolCatalog.FieldRequirement.REQUIRED,
        fieldNamed(namedRangeGroup.type(), "sheetName").requirement());
    assertEquals(
        GridGrindProtocolCatalog.FieldRequirement.REQUIRED,
        fieldNamed(namedRangeGroup.type(), "range").requirement());

    GridGrindProtocolCatalog.FieldEntry borderStyleField =
        fieldNamed(plainGroup(catalog, "cellBorderSideInputType").type(), "style");
    assertEquals(
        GridGrindProtocolCatalog.FieldRequirement.REQUIRED, borderStyleField.requirement());
    assertFalse(
        borderStyleField.enumValues().isEmpty(),
        "cellBorderSideInputType should expose enum values for 'style'");
    assertTrue(
        borderStyleField.enumValues().contains("THIN"),
        "border style enum values should include THIN");
    assertTrue(
        borderStyleField.enumValues().contains("NONE"),
        "border style enum values should include NONE");

    GridGrindProtocolCatalog.PlainTypeGroup styleGroup = plainGroup(catalog, "cellStyleInputType");
    assertTrue(
        fieldNamed(styleGroup.type(), "horizontalAlignment").enumValues().contains("GENERAL"),
        "horizontalAlignment enum values should include GENERAL");
    assertTrue(
        fieldNamed(styleGroup.type(), "verticalAlignment").enumValues().contains("BOTTOM"),
        "verticalAlignment enum values should include BOTTOM");
    assertEquals(
        new GridGrindProtocolCatalog.FieldShape.NestedTypeGroupRef("fontHeightTypes"),
        fieldNamed(styleGroup.type(), "fontHeight").shape());
    assertEquals(
        new GridGrindProtocolCatalog.FieldShape.PlainTypeGroupRef("cellBorderInputType"),
        fieldNamed(styleGroup.type(), "border").shape());

    GridGrindProtocolCatalog.PlainTypeGroup validationGroup =
        plainGroup(catalog, "dataValidationInputType");
    assertEquals(
        new GridGrindProtocolCatalog.FieldShape.NestedTypeGroupRef("dataValidationRuleTypes"),
        fieldNamed(validationGroup.type(), "rule").shape());
    assertEquals(
        GridGrindProtocolCatalog.FieldRequirement.OPTIONAL,
        fieldNamed(validationGroup.type(), "prompt").requirement());
    assertEquals(
        new GridGrindProtocolCatalog.FieldShape.PlainTypeGroupRef("dataValidationPromptInputType"),
        fieldNamed(validationGroup.type(), "prompt").shape());
    assertEquals(
        new GridGrindProtocolCatalog.FieldShape.PlainTypeGroupRef(
            "dataValidationErrorAlertInputType"),
        fieldNamed(validationGroup.type(), "errorAlert").shape());

    GridGrindProtocolCatalog.FieldEntry dataValidationOperatorField =
        fieldNamed(nestedTypeEntry(catalog, "dataValidationRuleTypes", "WHOLE_NUMBER"), "operator");
    assertFalse(
        dataValidationOperatorField.enumValues().isEmpty(),
        "WHOLE_NUMBER.operator should expose comparison-operator enum values");
    assertTrue(
        dataValidationOperatorField.enumValues().contains("BETWEEN"),
        "comparison-operator enum values should include BETWEEN");
  }

  private static void assertCatalogSummaries(GridGrindProtocolCatalog.Catalog catalog) {
    assertTrue(
        nestedTypeEntry(catalog, "cellInputTypes", "FORMULA")
            .summary()
            .contains("Omit the leading = sign"),
        "FORMULA summary should mention omitting leading = sign");
    assertTrue(
        entryNamed(catalog.sourceTypes(), "NEW").summary().contains("zero sheets"),
        "NEW source summary must mention zero sheets");
    assertTrue(
        entryNamed(catalog.sourceTypes(), "NEW").summary().contains("ENSURE_SHEET"),
        "NEW source summary must mention ENSURE_SHEET");
    assertTrue(
        entryNamed(catalog.sourceTypes(), "EXISTING")
            .summary()
            .contains("current execution environment"),
        "EXISTING source summary must explain relative path resolution");
    assertTrue(
        entryNamed(catalog.persistenceTypes(), "OVERWRITE")
            .summary()
            .contains("No path field is accepted"),
        "OVERWRITE summary must explain that the write target comes from source.path");
    assertTrue(
        entryNamed(catalog.persistenceTypes(), "SAVE_AS")
            .summary()
            .contains("current execution environment"),
        "SAVE_AS summary must explain relative path resolution");
    assertTrue(
        entryNamed(catalog.readTypes(), "GET_WINDOW").summary().contains("250000"),
        "GET_WINDOW summary must state the max cell limit");
    assertTrue(
        entryNamed(catalog.operationTypes(), "ENSURE_SHEET").summary().contains("31"),
        "ENSURE_SHEET summary must state the 31-character sheet name limit");
    assertTrue(
        entryNamed(catalog.operationTypes(), "SET_COLUMN_WIDTH").summary().contains("255"),
        "SET_COLUMN_WIDTH summary must state the 255-unit column width limit");
    assertTrue(
        entryNamed(catalog.operationTypes(), "SET_ROW_HEIGHT").summary().contains("1638"),
        "SET_ROW_HEIGHT summary must state the row height limit");
    assertTrue(
        entryNamed(catalog.operationTypes(), "APPEND_ROW").summary().contains("value-bearing"),
        "APPEND_ROW summary must explain value-bearing row semantics");
    assertTrue(
        entryNamed(catalog.operationTypes(), "AUTO_SIZE_COLUMNS")
            .summary()
            .contains("deterministically"),
        "AUTO_SIZE_COLUMNS summary must state deterministic sizing");
    assertTrue(
        entryNamed(catalog.operationTypes(), "SET_DATA_VALIDATION")
            .summary()
            .contains("Overlapping existing rules are normalized"),
        "SET_DATA_VALIDATION summary must describe overlap normalization");
    assertTrue(
        entryNamed(catalog.readTypes(), "GET_DATA_VALIDATIONS")
            .summary()
            .contains("unsupported rules are surfaced explicitly"),
        "GET_DATA_VALIDATIONS summary must explain unsupported-rule reporting");
    assertTrue(
        entryNamed(catalog.readTypes(), "ANALYZE_DATA_VALIDATION_HEALTH")
            .summary()
            .contains("overlapping"),
        "ANALYZE_DATA_VALIDATION_HEALTH summary must mention overlapping rules");
    assertTrue(
        entryNamed(catalog.readTypes(), "ANALYZE_HYPERLINK_HEALTH")
            .summary()
            .contains("persisted path"),
        "ANALYZE_HYPERLINK_HEALTH summary must explain relative FILE resolution");
    assertEquals(
        GridGrindProtocolCatalog.FieldRequirement.REQUIRED,
        fieldNamed(nestedTypeEntry(catalog, "hyperlinkTargetTypes", "FILE"), "path").requirement(),
        "FILE hyperlink targets must expose a required path field");
    assertTrue(
        nestedTypeEntry(catalog, "hyperlinkTargetTypes", "FILE").summary().contains("file: URIs"),
        "FILE hyperlink target summary must mention file: URI normalization");
    assertTrue(
        nestedTypeEntry(catalog, "cellInputTypes", "DATE").summary().contains("NUMBER"),
        "DATE summary must note that read-back declaredType is NUMBER");
  }

  private static void assertCatalogPolymorphicReferences(GridGrindProtocolCatalog.Catalog catalog) {
    assertEquals(
        new GridGrindProtocolCatalog.FieldShape.NestedTypeGroupRef("cellSelectionTypes"),
        fieldNamed(entryNamed(catalog.readTypes(), "GET_HYPERLINKS"), "selection").shape(),
        "GET_HYPERLINKS.selection must point to the cell selection group");
    assertEquals(
        new GridGrindProtocolCatalog.FieldShape.NestedTypeGroupRef("hyperlinkTargetTypes"),
        fieldNamed(entryNamed(catalog.operationTypes(), "SET_HYPERLINK"), "target").shape(),
        "SET_HYPERLINK.target must point to hyperlinkTargetTypes");
    assertEquals(
        new GridGrindProtocolCatalog.FieldShape.ListShape(
            new GridGrindProtocolCatalog.FieldShape.ListShape(
                new GridGrindProtocolCatalog.FieldShape.NestedTypeGroupRef("cellInputTypes"))),
        fieldNamed(entryNamed(catalog.operationTypes(), "SET_RANGE"), "rows").shape(),
        "SET_RANGE.rows must expose the nested cellInputTypes matrix shape");
    assertEquals(
        new GridGrindProtocolCatalog.FieldShape.NestedTypeGroupRef("sheetSelectionTypes"),
        fieldNamed(entryNamed(catalog.readTypes(), "ANALYZE_HYPERLINK_HEALTH"), "selection")
            .shape(),
        "ANALYZE_HYPERLINK_HEALTH.selection must point to sheetSelectionTypes");
    assertEquals(
        new GridGrindProtocolCatalog.FieldShape.PlainTypeGroupRef("dataValidationInputType"),
        fieldNamed(entryNamed(catalog.operationTypes(), "SET_DATA_VALIDATION"), "validation")
            .shape(),
        "SET_DATA_VALIDATION.validation must point to dataValidationInputType");
    assertEquals(
        new GridGrindProtocolCatalog.FieldShape.NestedTypeGroupRef("rangeSelectionTypes"),
        fieldNamed(entryNamed(catalog.operationTypes(), "CLEAR_DATA_VALIDATIONS"), "selection")
            .shape(),
        "CLEAR_DATA_VALIDATIONS.selection must point to rangeSelectionTypes");
    assertEquals(
        new GridGrindProtocolCatalog.FieldShape.NestedTypeGroupRef("rangeSelectionTypes"),
        fieldNamed(entryNamed(catalog.readTypes(), "GET_DATA_VALIDATIONS"), "selection").shape(),
        "GET_DATA_VALIDATIONS.selection must point to rangeSelectionTypes");
    assertEquals(
        new GridGrindProtocolCatalog.FieldShape.NestedTypeGroupRef("sheetSelectionTypes"),
        fieldNamed(entryNamed(catalog.readTypes(), "ANALYZE_DATA_VALIDATION_HEALTH"), "selection")
            .shape(),
        "ANALYZE_DATA_VALIDATION_HEALTH.selection must point to sheetSelectionTypes");
  }

  private static List<String> ids(List<GridGrindProtocolCatalog.TypeEntry> entries) {
    return entries.stream().map(GridGrindProtocolCatalog.TypeEntry::id).toList();
  }

  private static GridGrindProtocolCatalog.TypeEntry entryNamed(
      List<GridGrindProtocolCatalog.TypeEntry> entries, String id) {
    return entries.stream().filter(entry -> id.equals(entry.id())).findFirst().orElseThrow();
  }

  private static GridGrindProtocolCatalog.PlainTypeGroup plainGroup(
      GridGrindProtocolCatalog.Catalog catalog, String group) {
    return catalog.plainTypes().stream()
        .filter(entry -> group.equals(entry.group()))
        .findFirst()
        .orElseThrow();
  }

  private static GridGrindProtocolCatalog.TypeEntry nestedTypeEntry(
      GridGrindProtocolCatalog.Catalog catalog, String group, String id) {
    return catalog.nestedTypes().stream()
        .filter(entry -> group.equals(entry.group()))
        .findFirst()
        .orElseThrow()
        .types()
        .stream()
        .filter(entry -> id.equals(entry.id()))
        .findFirst()
        .orElseThrow();
  }

  private static GridGrindProtocolCatalog.FieldEntry fieldNamed(
      GridGrindProtocolCatalog.TypeEntry entry, String name) {
    GridGrindProtocolCatalog.FieldEntry field = entry.field(name);
    assertNotNull(field, () -> "Expected field '" + name + "' on catalog entry " + entry.id());
    return field;
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

  private record UnsupportedType(String displayName) implements Type {
    @Override
    public String getTypeName() {
      return displayName;
    }

    @Override
    public String toString() {
      return displayName;
    }
  }

  private record TestParameterizedType(Type rawType, List<Type> actualTypeArguments)
      implements ParameterizedType {
    TestParameterizedType {
      Objects.requireNonNull(rawType, "rawType must not be null");
      actualTypeArguments = List.copyOf(actualTypeArguments);
    }

    @Override
    public Type[] getActualTypeArguments() {
      return actualTypeArguments.toArray(Type[]::new);
    }

    @Override
    public Type getRawType() {
      return rawType;
    }

    @Override
    public Type getOwnerType() {
      return null;
    }

    @Override
    public String toString() {
      return rawType.getTypeName()
          + "<"
          + actualTypeArguments.stream()
              .map(Type::getTypeName)
              .collect(java.util.stream.Collectors.joining(", "))
          + ">";
    }
  }
}
