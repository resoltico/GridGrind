package dev.erst.gridgrind.protocol.catalog;

import static org.junit.jupiter.api.Assertions.*;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.protocol.catalog.gather.CatalogFieldMetadataSupport;
import dev.erst.gridgrind.protocol.dto.CellInput;
import dev.erst.gridgrind.protocol.dto.CommentInput;
import dev.erst.gridgrind.protocol.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.protocol.dto.GridGrindRequest;
import dev.erst.gridgrind.protocol.dto.HyperlinkTarget;
import dev.erst.gridgrind.protocol.json.GridGrindJson;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import java.io.IOException;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
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
    Catalog catalog = GridGrindProtocolCatalog.catalog();
    Catalog decoded =
        GridGrindJson.readProtocolCatalog(GridGrindJson.writeProtocolCatalogBytes(catalog));

    assertCatalogInventory(catalog);
    assertCatalogFieldShapes(catalog);
    assertCatalogSummaries(catalog);
    assertCatalogPolymorphicReferences(catalog);

    assertEquals(catalog, decoded);
  }

  @Test
  void requestTemplateAndCatalogEncodeDeterministically() throws IOException {
    assertArrayEquals(
        GridGrindJson.writeRequestBytes(GridGrindProtocolCatalog.requestTemplate()),
        GridGrindJson.writeRequestBytes(GridGrindProtocolCatalog.requestTemplate()));
    assertArrayEquals(
        GridGrindJson.writeProtocolCatalogBytes(GridGrindProtocolCatalog.catalog()),
        GridGrindJson.writeProtocolCatalogBytes(GridGrindProtocolCatalog.catalog()));
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
        new FieldShape.ListShape(new FieldShape.NestedTypeGroupRef("cellInputTypes")),
        CatalogFieldMetadataSupport.fieldShape(listOfCellInputs));
  }

  @Test
  void fieldShapeRejectsUnsupportedTypeImplementations() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () -> CatalogFieldMetadataSupport.fieldShape(new UnsupportedType("custom-type")));

    assertEquals("Unsupported catalog field type: custom-type", failure.getMessage());
  }

  @Test
  void fieldShapeRejectsListShapesWithoutExactlyOneTypeArgument() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                CatalogFieldMetadataSupport.fieldShape(
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
            () -> CatalogFieldMetadataSupport.fieldShape(unsupportedType));

    assertEquals(
        "Unsupported parameterized catalog field type: " + unsupportedType, failure.getMessage());
  }

  @Test
  void fieldShapeRejectsUnsupportedClassTypes() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () -> CatalogFieldMetadataSupport.fieldShape(java.util.UUID.class));

    assertEquals("Unsupported catalog field type: java.util.UUID", failure.getMessage());
  }

  @Test
  void isNumericTypeRecognizesNumericAndNonNumericClasses() {
    assertTrue(CatalogFieldMetadataSupport.isNumericType(java.math.BigInteger.class));
    assertFalse(CatalogFieldMetadataSupport.isNumericType(String.class));
  }

  @Test
  void validateNestedTypeGroupMappingRejectsMismatches() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                CatalogFieldMetadataSupport.validateNestedTypeGroupMapping(
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
                CatalogFieldMetadataSupport.validatePlainTypeGroupMapping(
                    CommentInput.class, "wrongGroup"));

    assertEquals(
        "Field-shape plain group mapping mismatch for "
            + CommentInput.class.getName()
            + ": expected=wrongGroup, mapped=commentInputType",
        failure.getMessage());
  }

  @Test
  void catalogDefaultsNullProtocolVersionAndCopiesLists() {
    TypeEntry sourceType =
        new TypeEntry(
            "NEW",
            "Create",
            List.of(
                new FieldEntry(
                    "type",
                    FieldRequirement.REQUIRED,
                    new FieldShape.Scalar(ScalarType.STRING),
                    List.of())));
    NestedTypeGroup nestedTypeGroup = new NestedTypeGroup("cellInputTypes", List.of(sourceType));
    PlainTypeGroup plainTypeGroup = new PlainTypeGroup("commentInputType", sourceType);

    Catalog catalog =
        new Catalog(
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
        assertThrows(IllegalArgumentException.class, () -> new NestedTypeGroup(" ", List.of()))
            .getMessage());
    assertEquals(
        "name must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new FieldEntry(
                        " ",
                        FieldRequirement.REQUIRED,
                        new FieldShape.Scalar(ScalarType.STRING),
                        List.of()))
            .getMessage());
    assertThrows(
        NullPointerException.class,
        () -> new NestedTypeGroup("cellInputTypes", Arrays.asList((TypeEntry) null)));
    assertThrows(
        NullPointerException.class,
        () ->
            new Catalog(
                GridGrindProtocolVersion.current(),
                "type",
                List.of(),
                List.of(),
                List.of(),
                List.of(),
                Arrays.asList((NestedTypeGroup) null),
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new Catalog(
                GridGrindProtocolVersion.current(),
                "type",
                List.of(),
                List.of(),
                List.of(),
                List.of(),
                List.of(),
                Arrays.asList((PlainTypeGroup) null)));
    assertEquals(
        "group must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new PlainTypeGroup(" ", new TypeEntry("ID", "Summary", List.of())))
            .getMessage());
    assertThrows(NullPointerException.class, () -> new PlainTypeGroup("commentInputType", null));
    assertThrows(NullPointerException.class, () -> new TypeEntry("ID", "Summary", null));
    assertThrows(
        NullPointerException.class,
        () -> new FieldEntry("name", null, new FieldShape.Scalar(ScalarType.STRING), List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new FieldEntry("name", FieldRequirement.REQUIRED, null, List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new FieldEntry(
                "name", FieldRequirement.REQUIRED, new FieldShape.Scalar(ScalarType.STRING), null));
    IllegalStateException duplicateFieldFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                new TypeEntry(
                    "ID",
                    "Summary",
                    List.of(
                        new FieldEntry(
                            "name",
                            FieldRequirement.REQUIRED,
                            new FieldShape.Scalar(ScalarType.STRING),
                            List.of()),
                        new FieldEntry(
                            "name",
                            FieldRequirement.OPTIONAL,
                            new FieldShape.Scalar(ScalarType.STRING),
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
  void fieldShapeGroupRefRejectsBlankGroups() {
    assertEquals(
        "group must not be blank",
        assertThrows(IllegalArgumentException.class, () -> new FieldShape.NestedTypeGroupRef(" "))
            .getMessage());
    assertEquals(
        "group must not be blank",
        assertThrows(IllegalArgumentException.class, () -> new FieldShape.PlainTypeGroupRef(" "))
            .getMessage());
  }

  @Test
  void validateReverseGroupMappingsRejectsNestedTypeWithNoDescriptor() {
    Set<Class<?>> emptyNestedTypes = Set.of();
    Set<Class<?>> allPlainTypes = CatalogFieldMetadataSupport.registeredPlainTypes();

    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalog.validateReverseGroupMappings(
                    emptyNestedTypes, allPlainTypes));

    assertTrue(
        failure
            .getMessage()
            .contains("Field-shape nested group map contains type with no catalog descriptor:"),
        "failure must identify the orphaned nested type");
  }

  @Test
  void validateReverseGroupMappingsRejectsPlainTypeWithNoDescriptor() {
    Set<Class<?>> allNestedTypes = CatalogFieldMetadataSupport.registeredNestedTypes();
    Set<Class<?>> emptyPlainTypes = Set.of();

    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalog.validateReverseGroupMappings(
                    allNestedTypes, emptyPlainTypes));

    assertTrue(
        failure
            .getMessage()
            .contains("Field-shape plain group map contains type with no catalog descriptor:"),
        "failure must identify the orphaned plain type");
  }

  private static void assertCatalogInventory(Catalog catalog) {
    assertEquals(GridGrindProtocolVersion.V1, catalog.protocolVersion());
    assertEquals("type", catalog.discriminatorField());
    assertEquals(List.of("NEW", "EXISTING"), ids(catalog.sourceTypes()));
    assertEquals(List.of("NONE", "OVERWRITE", "SAVE_AS"), ids(catalog.persistenceTypes()));
    assertEquals(52, catalog.operationTypes().size());
    assertEquals(25, catalog.readTypes().size());
    assertEquals(
        List.of(
            "cellInputTypes",
            "hyperlinkTargetTypes",
            "paneTypes",
            "sheetCopyPositionTypes",
            "rowSpanTypes",
            "columnSpanTypes",
            "cellSelectionTypes",
            "rangeSelectionTypes",
            "sheetSelectionTypes",
            "tableSelectionTypes",
            "namedRangeSelectionTypes",
            "namedRangeScopeTypes",
            "namedRangeSelectorTypes",
            "fontHeightTypes",
            "dataValidationRuleTypes",
            "conditionalFormattingRuleTypes",
            "printAreaTypes",
            "printScalingTypes",
            "printTitleRowsTypes",
            "printTitleColumnsTypes",
            "tableStyleTypes"),
        catalog.nestedTypes().stream().map(NestedTypeGroup::group).toList());
    assertEquals(
        List.of(
            "commentInputType",
            "namedRangeTargetType",
            "sheetProtectionSettingsType",
            "cellStyleInputType",
            "cellAlignmentInputType",
            "cellFontInputType",
            "richTextRunInputType",
            "cellFillInputType",
            "cellBorderInputType",
            "cellBorderSideInputType",
            "cellProtectionInputType",
            "dataValidationInputType",
            "dataValidationPromptInputType",
            "dataValidationErrorAlertInputType",
            "conditionalFormattingBlockInputType",
            "headerFooterTextInputType",
            "differentialStyleInputType",
            "differentialBorderInputType",
            "differentialBorderSideInputType",
            "printLayoutInputType",
            "tableInputType"),
        catalog.plainTypes().stream().map(PlainTypeGroup::group).toList());
  }

  private static void assertCatalogFieldShapes(Catalog catalog) {
    assertCatalogCoreFieldShapes(catalog);
    assertCatalogStyleFieldShapes(catalog);
    assertCatalogValidationAndTableFieldShapes(catalog);
    assertCatalogConditionalFormattingFieldShapes(catalog);
    assertCatalogPrintAndPaneFieldShapes(catalog);
  }

  private static void assertCatalogCoreFieldShapes(Catalog catalog) {
    PlainTypeGroup commentGroup = plainGroup(catalog, "commentInputType");
    assertEquals(FieldRequirement.REQUIRED, fieldNamed(commentGroup.type(), "text").requirement());
    assertEquals(
        new FieldShape.Scalar(ScalarType.STRING), fieldNamed(commentGroup.type(), "text").shape());
    assertEquals(
        FieldRequirement.REQUIRED, fieldNamed(commentGroup.type(), "author").requirement());
    assertEquals(
        FieldRequirement.OPTIONAL, fieldNamed(commentGroup.type(), "visible").requirement());
    assertEquals(
        new FieldShape.Scalar(ScalarType.BOOLEAN),
        fieldNamed(commentGroup.type(), "visible").shape());

    PlainTypeGroup namedRangeGroup = plainGroup(catalog, "namedRangeTargetType");
    assertEquals(
        FieldRequirement.REQUIRED, fieldNamed(namedRangeGroup.type(), "sheetName").requirement());
    assertEquals(
        FieldRequirement.REQUIRED, fieldNamed(namedRangeGroup.type(), "range").requirement());

    TypeEntry copySheet = entryNamed(catalog.operationTypes(), "COPY_SHEET");
    assertEquals(FieldRequirement.OPTIONAL, fieldNamed(copySheet, "position").requirement());
    assertEquals(
        new FieldShape.NestedTypeGroupRef("sheetCopyPositionTypes"),
        fieldNamed(copySheet, "position").shape());

    PlainTypeGroup protectionGroup = plainGroup(catalog, "sheetProtectionSettingsType");
    assertTrue(
        fieldNamed(protectionGroup.type(), "autoFilterLocked").shape() instanceof FieldShape.Scalar,
        "sheetProtectionSettingsType boolean fields must be scalar booleans");
    assertEquals(
        new FieldShape.Scalar(ScalarType.BOOLEAN),
        fieldNamed(protectionGroup.type(), "sortLocked").shape());

    assertTrue(
        fieldNamed(entryNamed(catalog.operationTypes(), "SET_SHEET_VISIBILITY"), "visibility")
            .enumValues()
            .contains("VERY_HIDDEN"),
        "SET_SHEET_VISIBILITY.visibility must expose visibility enum values");
    assertEquals(
        new FieldShape.PlainTypeGroupRef("sheetProtectionSettingsType"),
        fieldNamed(entryNamed(catalog.operationTypes(), "SET_SHEET_PROTECTION"), "protection")
            .shape());

    FieldEntry borderStyleField =
        fieldNamed(plainGroup(catalog, "cellBorderSideInputType").type(), "style");
    assertEquals(FieldRequirement.OPTIONAL, borderStyleField.requirement());
    assertFalse(
        borderStyleField.enumValues().isEmpty(),
        "cellBorderSideInputType should expose enum values for 'style'");
    assertTrue(
        borderStyleField.enumValues().contains("THIN"),
        "border style enum values should include THIN");
    assertTrue(
        borderStyleField.enumValues().contains("NONE"),
        "border style enum values should include NONE");
  }

  private static void assertCatalogStyleFieldShapes(Catalog catalog) {
    PlainTypeGroup styleGroup = plainGroup(catalog, "cellStyleInputType");
    assertEquals(
        new FieldShape.PlainTypeGroupRef("cellAlignmentInputType"),
        fieldNamed(styleGroup.type(), "alignment").shape());
    assertEquals(
        new FieldShape.PlainTypeGroupRef("cellFontInputType"),
        fieldNamed(styleGroup.type(), "font").shape());
    assertEquals(
        new FieldShape.PlainTypeGroupRef("cellFillInputType"),
        fieldNamed(styleGroup.type(), "fill").shape());
    assertEquals(
        new FieldShape.PlainTypeGroupRef("cellProtectionInputType"),
        fieldNamed(styleGroup.type(), "protection").shape());
    assertEquals(
        new FieldShape.PlainTypeGroupRef("cellBorderInputType"),
        fieldNamed(styleGroup.type(), "border").shape());

    PlainTypeGroup alignmentGroup = plainGroup(catalog, "cellAlignmentInputType");
    assertTrue(
        fieldNamed(alignmentGroup.type(), "horizontalAlignment").enumValues().contains("GENERAL"),
        "horizontalAlignment enum values should include GENERAL");
    assertTrue(
        fieldNamed(alignmentGroup.type(), "verticalAlignment").enumValues().contains("BOTTOM"),
        "verticalAlignment enum values should include BOTTOM");

    PlainTypeGroup fontGroup = plainGroup(catalog, "cellFontInputType");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("fontHeightTypes"),
        fieldNamed(fontGroup.type(), "fontHeight").shape());

    TypeEntry richTextEntry = nestedTypeEntry(catalog, "cellInputTypes", "RICH_TEXT");
    assertEquals(
        new FieldShape.ListShape(new FieldShape.PlainTypeGroupRef("richTextRunInputType")),
        fieldNamed(richTextEntry, "runs").shape());

    PlainTypeGroup richTextRunGroup = plainGroup(catalog, "richTextRunInputType");
    assertEquals(
        new FieldShape.PlainTypeGroupRef("cellFontInputType"),
        fieldNamed(richTextRunGroup.type(), "font").shape());
    assertEquals(
        FieldRequirement.OPTIONAL, fieldNamed(richTextRunGroup.type(), "font").requirement());

    PlainTypeGroup fillGroup = plainGroup(catalog, "cellFillInputType");
    assertEquals(
        FieldRequirement.OPTIONAL, fieldNamed(fillGroup.type(), "backgroundColor").requirement());
  }

  private static void assertCatalogValidationAndTableFieldShapes(Catalog catalog) {
    PlainTypeGroup validationGroup = plainGroup(catalog, "dataValidationInputType");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("dataValidationRuleTypes"),
        fieldNamed(validationGroup.type(), "rule").shape());
    assertEquals(
        FieldRequirement.OPTIONAL, fieldNamed(validationGroup.type(), "prompt").requirement());
    assertEquals(
        new FieldShape.PlainTypeGroupRef("dataValidationPromptInputType"),
        fieldNamed(validationGroup.type(), "prompt").shape());
    assertEquals(
        new FieldShape.PlainTypeGroupRef("dataValidationErrorAlertInputType"),
        fieldNamed(validationGroup.type(), "errorAlert").shape());

    FieldEntry dataValidationOperatorField =
        fieldNamed(nestedTypeEntry(catalog, "dataValidationRuleTypes", "WHOLE_NUMBER"), "operator");
    assertFalse(
        dataValidationOperatorField.enumValues().isEmpty(),
        "WHOLE_NUMBER.operator should expose comparison-operator enum values");
    assertTrue(
        dataValidationOperatorField.enumValues().contains("BETWEEN"),
        "comparison-operator enum values should include BETWEEN");

    PlainTypeGroup tableGroup = plainGroup(catalog, "tableInputType");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("tableStyleTypes"),
        fieldNamed(tableGroup.type(), "style").shape());
    assertEquals(FieldRequirement.REQUIRED, fieldNamed(tableGroup.type(), "name").requirement());

    FieldEntry tableSelectionNames =
        fieldNamed(nestedTypeEntry(catalog, "tableSelectionTypes", "BY_NAMES"), "names");
    assertEquals(
        new FieldShape.ListShape(new FieldShape.Scalar(ScalarType.STRING)),
        tableSelectionNames.shape());

    FieldEntry tableStyleName =
        fieldNamed(nestedTypeEntry(catalog, "tableStyleTypes", "NAMED"), "name");
    assertEquals(new FieldShape.Scalar(ScalarType.STRING), tableStyleName.shape());
  }

  private static void assertCatalogConditionalFormattingFieldShapes(Catalog catalog) {
    PlainTypeGroup conditionalFormattingBlockGroup =
        plainGroup(catalog, "conditionalFormattingBlockInputType");
    assertEquals(
        new FieldShape.ListShape(new FieldShape.Scalar(ScalarType.STRING)),
        fieldNamed(conditionalFormattingBlockGroup.type(), "ranges").shape());
    assertEquals(
        new FieldShape.ListShape(
            new FieldShape.NestedTypeGroupRef("conditionalFormattingRuleTypes")),
        fieldNamed(conditionalFormattingBlockGroup.type(), "rules").shape());

    PlainTypeGroup differentialStyleGroup = plainGroup(catalog, "differentialStyleInputType");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("fontHeightTypes"),
        fieldNamed(differentialStyleGroup.type(), "fontHeight").shape());
    assertEquals(
        new FieldShape.PlainTypeGroupRef("differentialBorderInputType"),
        fieldNamed(differentialStyleGroup.type(), "border").shape());

    FieldEntry conditionalFormattingOperatorField =
        fieldNamed(
            nestedTypeEntry(catalog, "conditionalFormattingRuleTypes", "CELL_VALUE_RULE"),
            "operator");
    assertTrue(
        conditionalFormattingOperatorField.enumValues().contains("BETWEEN"),
        "conditional-formatting operator enum values should include BETWEEN");
  }

  private static void assertCatalogPrintAndPaneFieldShapes(Catalog catalog) {
    TypeEntry setSheetPane = entryNamed(catalog.operationTypes(), "SET_SHEET_PANE");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("paneTypes"),
        fieldNamed(setSheetPane, "pane").shape(),
        "SET_SHEET_PANE.pane must point to paneTypes");
    assertEquals(
        0,
        nestedTypeEntry(catalog, "paneTypes", "NONE").fields().size(),
        "paneTypes.NONE has no fields");
    assertEquals(
        4,
        nestedTypeEntry(catalog, "paneTypes", "FROZEN").fields().size(),
        "paneTypes.FROZEN has four fields: splitColumn, splitRow, leftmostColumn, topRow");
    assertEquals(
        5,
        nestedTypeEntry(catalog, "paneTypes", "SPLIT").fields().size(),
        "paneTypes.SPLIT has five fields: xSplitPosition, ySplitPosition, leftmostColumn, topRow, activePane");

    PlainTypeGroup printLayoutGroup = plainGroup(catalog, "printLayoutInputType");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("printAreaTypes"),
        fieldNamed(printLayoutGroup.type(), "printArea").shape(),
        "printLayoutInputType.printArea must point to printAreaTypes");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("printScalingTypes"),
        fieldNamed(printLayoutGroup.type(), "scaling").shape(),
        "printLayoutInputType.scaling must point to printScalingTypes");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("printTitleRowsTypes"),
        fieldNamed(printLayoutGroup.type(), "repeatingRows").shape(),
        "printLayoutInputType.repeatingRows must point to printTitleRowsTypes");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("printTitleColumnsTypes"),
        fieldNamed(printLayoutGroup.type(), "repeatingColumns").shape(),
        "printLayoutInputType.repeatingColumns must point to printTitleColumnsTypes");
    assertEquals(
        new FieldShape.PlainTypeGroupRef("headerFooterTextInputType"),
        fieldNamed(printLayoutGroup.type(), "header").shape(),
        "printLayoutInputType.header must point to headerFooterTextInputType");
    assertEquals(
        new FieldShape.PlainTypeGroupRef("headerFooterTextInputType"),
        fieldNamed(printLayoutGroup.type(), "footer").shape(),
        "printLayoutInputType.footer must point to headerFooterTextInputType");
    assertEquals(
        FieldRequirement.OPTIONAL,
        fieldNamed(printLayoutGroup.type(), "printArea").requirement(),
        "printLayoutInputType.printArea must be optional");

    PlainTypeGroup headerFooterGroup = plainGroup(catalog, "headerFooterTextInputType");
    assertEquals(
        new FieldShape.Scalar(ScalarType.STRING),
        fieldNamed(headerFooterGroup.type(), "left").shape(),
        "headerFooterTextInputType.left must be a string scalar");
    assertEquals(
        FieldRequirement.OPTIONAL,
        fieldNamed(headerFooterGroup.type(), "left").requirement(),
        "headerFooterTextInputType.left must be optional");
  }

  private static void assertCatalogSummaries(Catalog catalog) {
    assertTrue(
        nestedTypeEntry(catalog, "cellInputTypes", "FORMULA")
            .summary()
            .contains("leading = sign is accepted and stripped automatically"),
        "FORMULA summary should describe leading = sign handling");
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
        entryNamed(catalog.operationTypes(), "COPY_SHEET")
            .summary()
            .contains("Sheets containing tables"),
        "COPY_SHEET summary must describe explicit non-copyable sheet families");
    assertTrue(
        entryNamed(catalog.operationTypes(), "SET_SELECTED_SHEETS")
            .summary()
            .contains("workbook order"),
        "SET_SELECTED_SHEETS summary must explain workbook-order readback");
    assertTrue(
        entryNamed(catalog.operationTypes(), "DELETE_SHEET")
            .summary()
            .contains("last visible sheet"),
        "DELETE_SHEET summary must explain the last-visible-sheet invariant");
    assertTrue(
        entryNamed(catalog.readTypes(), "GET_WORKBOOK_SUMMARY")
            .summary()
            .contains("workbook.kind=EMPTY"),
        "GET_WORKBOOK_SUMMARY summary must describe the empty-workbook variant");
    assertTrue(
        entryNamed(catalog.readTypes(), "GET_SHEET_SUMMARY").summary().contains("visibility"),
        "GET_SHEET_SUMMARY summary must mention visibility");
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
        entryNamed(catalog.operationTypes(), "SET_CONDITIONAL_FORMATTING")
            .summary()
            .contains("authoritative"),
        "SET_CONDITIONAL_FORMATTING summary must describe authoritative replacement");
    assertTrue(
        entryNamed(catalog.readTypes(), "GET_CONDITIONAL_FORMATTING")
            .summary()
            .contains("unsupported rules"),
        "GET_CONDITIONAL_FORMATTING summary must describe unsupported-rule surfacing");
    assertTrue(
        entryNamed(catalog.readTypes(), "ANALYZE_CONDITIONAL_FORMATTING_HEALTH")
            .summary()
            .contains("priority collisions"),
        "ANALYZE_CONDITIONAL_FORMATTING_HEALTH summary must mention priority collisions");
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
        entryNamed(catalog.operationTypes(), "SET_AUTOFILTER")
            .summary()
            .contains("must not overlap"),
        "SET_AUTOFILTER summary must describe overlap rejection");
    assertTrue(
        entryNamed(catalog.operationTypes(), "SET_TABLE").summary().contains("workbook-global"),
        "SET_TABLE summary must describe workbook-global table names");
    assertTrue(
        entryNamed(catalog.readTypes(), "GET_AUTOFILTERS")
            .summary()
            .contains("sheet- and table-owned"),
        "GET_AUTOFILTERS summary must describe both ownership families");
    assertTrue(
        entryNamed(catalog.readTypes(), "GET_TABLES").summary().contains("workbook-global"),
        "GET_TABLES summary must describe workbook-global selection");
    assertTrue(
        entryNamed(catalog.readTypes(), "ANALYZE_AUTOFILTER_HEALTH")
            .summary()
            .contains("ownership"),
        "ANALYZE_AUTOFILTER_HEALTH summary must mention ownership mismatches");
    assertTrue(
        entryNamed(catalog.readTypes(), "ANALYZE_TABLE_HEALTH").summary().contains("overlaps"),
        "ANALYZE_TABLE_HEALTH summary must mention table overlaps");
    assertTrue(
        entryNamed(catalog.readTypes(), "ANALYZE_HYPERLINK_HEALTH")
            .summary()
            .contains("persisted path"),
        "ANALYZE_HYPERLINK_HEALTH summary must explain relative FILE resolution");
    assertEquals(
        FieldRequirement.REQUIRED,
        fieldNamed(nestedTypeEntry(catalog, "hyperlinkTargetTypes", "FILE"), "path").requirement(),
        "FILE hyperlink targets must expose a required path field");
    assertTrue(
        nestedTypeEntry(catalog, "hyperlinkTargetTypes", "FILE").summary().contains("file: URIs"),
        "FILE hyperlink target summary must mention file: URI normalization");
    assertTrue(
        nestedTypeEntry(catalog, "cellInputTypes", "DATE").summary().contains("NUMBER"),
        "DATE summary must note that read-back declaredType is NUMBER");
    assertTrue(
        nestedTypeEntry(catalog, "cellInputTypes", "RICH_TEXT")
            .summary()
            .contains("ordered rich-text run list"),
        "RICH_TEXT summary must describe ordered run authoring");
    assertTrue(
        entryNamed(catalog.operationTypes(), "SET_SHEET_PANE")
            .summary()
            .contains("NONE, FROZEN, or SPLIT"),
        "SET_SHEET_PANE summary must document the three pane types");
    assertTrue(
        entryNamed(catalog.operationTypes(), "SET_PRINT_LAYOUT").summary().contains("print area"),
        "SET_PRINT_LAYOUT summary must mention print area");
    assertTrue(
        nestedTypeEntry(catalog, "printScalingTypes", "FIT").summary().contains("unconstrained"),
        "FIT scaling summary must explain zero-axis constraint semantics");
  }

  private static void assertCatalogPolymorphicReferences(Catalog catalog) {
    assertEquals(
        new FieldShape.NestedTypeGroupRef("sheetCopyPositionTypes"),
        fieldNamed(entryNamed(catalog.operationTypes(), "COPY_SHEET"), "position").shape(),
        "COPY_SHEET.position must point to sheetCopyPositionTypes");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("cellSelectionTypes"),
        fieldNamed(entryNamed(catalog.readTypes(), "GET_HYPERLINKS"), "selection").shape(),
        "GET_HYPERLINKS.selection must point to the cell selection group");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("hyperlinkTargetTypes"),
        fieldNamed(entryNamed(catalog.operationTypes(), "SET_HYPERLINK"), "target").shape(),
        "SET_HYPERLINK.target must point to hyperlinkTargetTypes");
    assertEquals(
        new FieldShape.ListShape(
            new FieldShape.ListShape(new FieldShape.NestedTypeGroupRef("cellInputTypes"))),
        fieldNamed(entryNamed(catalog.operationTypes(), "SET_RANGE"), "rows").shape(),
        "SET_RANGE.rows must expose the nested cellInputTypes matrix shape");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("sheetSelectionTypes"),
        fieldNamed(entryNamed(catalog.readTypes(), "ANALYZE_HYPERLINK_HEALTH"), "selection")
            .shape(),
        "ANALYZE_HYPERLINK_HEALTH.selection must point to sheetSelectionTypes");
    assertEquals(
        new FieldShape.PlainTypeGroupRef("dataValidationInputType"),
        fieldNamed(entryNamed(catalog.operationTypes(), "SET_DATA_VALIDATION"), "validation")
            .shape(),
        "SET_DATA_VALIDATION.validation must point to dataValidationInputType");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("rangeSelectionTypes"),
        fieldNamed(entryNamed(catalog.operationTypes(), "CLEAR_DATA_VALIDATIONS"), "selection")
            .shape(),
        "CLEAR_DATA_VALIDATIONS.selection must point to rangeSelectionTypes");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("rangeSelectionTypes"),
        fieldNamed(entryNamed(catalog.readTypes(), "GET_DATA_VALIDATIONS"), "selection").shape(),
        "GET_DATA_VALIDATIONS.selection must point to rangeSelectionTypes");
    assertEquals(
        new FieldShape.PlainTypeGroupRef("conditionalFormattingBlockInputType"),
        fieldNamed(
                entryNamed(catalog.operationTypes(), "SET_CONDITIONAL_FORMATTING"),
                "conditionalFormatting")
            .shape(),
        "SET_CONDITIONAL_FORMATTING.conditionalFormatting must point to conditionalFormattingBlockInputType");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("rangeSelectionTypes"),
        fieldNamed(
                entryNamed(catalog.operationTypes(), "CLEAR_CONDITIONAL_FORMATTING"), "selection")
            .shape(),
        "CLEAR_CONDITIONAL_FORMATTING.selection must point to rangeSelectionTypes");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("rangeSelectionTypes"),
        fieldNamed(entryNamed(catalog.readTypes(), "GET_CONDITIONAL_FORMATTING"), "selection")
            .shape(),
        "GET_CONDITIONAL_FORMATTING.selection must point to rangeSelectionTypes");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("sheetSelectionTypes"),
        fieldNamed(
                entryNamed(catalog.readTypes(), "ANALYZE_CONDITIONAL_FORMATTING_HEALTH"),
                "selection")
            .shape(),
        "ANALYZE_CONDITIONAL_FORMATTING_HEALTH.selection must point to sheetSelectionTypes");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("sheetSelectionTypes"),
        fieldNamed(entryNamed(catalog.readTypes(), "ANALYZE_DATA_VALIDATION_HEALTH"), "selection")
            .shape(),
        "ANALYZE_DATA_VALIDATION_HEALTH.selection must point to sheetSelectionTypes");
    assertEquals(
        new FieldShape.PlainTypeGroupRef("tableInputType"),
        fieldNamed(entryNamed(catalog.operationTypes(), "SET_TABLE"), "table").shape(),
        "SET_TABLE.table must point to tableInputType");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("tableSelectionTypes"),
        fieldNamed(entryNamed(catalog.readTypes(), "GET_TABLES"), "selection").shape(),
        "GET_TABLES.selection must point to tableSelectionTypes");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("tableSelectionTypes"),
        fieldNamed(entryNamed(catalog.readTypes(), "ANALYZE_TABLE_HEALTH"), "selection").shape(),
        "ANALYZE_TABLE_HEALTH.selection must point to tableSelectionTypes");
    assertEquals(
        new FieldShape.NestedTypeGroupRef("sheetSelectionTypes"),
        fieldNamed(entryNamed(catalog.readTypes(), "ANALYZE_AUTOFILTER_HEALTH"), "selection")
            .shape(),
        "ANALYZE_AUTOFILTER_HEALTH.selection must point to sheetSelectionTypes");
    assertEquals(
        new FieldShape.PlainTypeGroupRef("printLayoutInputType"),
        fieldNamed(entryNamed(catalog.operationTypes(), "SET_PRINT_LAYOUT"), "printLayout").shape(),
        "SET_PRINT_LAYOUT.printLayout must point to printLayoutInputType");
  }

  private static List<String> ids(List<TypeEntry> entries) {
    return entries.stream().map(TypeEntry::id).toList();
  }

  private static TypeEntry entryNamed(List<TypeEntry> entries, String id) {
    return entries.stream().filter(entry -> id.equals(entry.id())).findFirst().orElseThrow();
  }

  private static PlainTypeGroup plainGroup(Catalog catalog, String group) {
    return catalog.plainTypes().stream()
        .filter(entry -> group.equals(entry.group()))
        .findFirst()
        .orElseThrow();
  }

  private static TypeEntry nestedTypeEntry(Catalog catalog, String group, String id) {
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

  private static FieldEntry fieldNamed(TypeEntry entry, String name) {
    FieldEntry field = entry.field(name);
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
  }
}
