package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.GridGrindJsonOutput;
import dev.erst.gridgrind.contract.step.MutationStep;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import org.junit.jupiter.api.Test;

/** Tests for the machine-readable catalog and built-in request template. */
class GridGrindProtocolCatalogTest {
  @Test
  void exposesMinimalStepBasedRequestTemplate() throws IOException {
    WorkbookPlan template = GridGrindProtocolCatalog.requestTemplate();
    WorkbookPlan decoded =
        GridGrindJson.readRequest(GridGrindJsonOutput.writeRequestBytes(template));
    var templateTree = GridGrindJsonOutput.requestTree(template);

    assertEquals(GridGrindProtocolVersion.V2, template.protocolVersion());
    assertTrue(template.execution().isDefault());
    assertTrue(template.formulaEnvironment().isEmpty());
    assertTrue(template.steps().isEmpty());
    assertFalse(templateTree.has("execution"));
    assertFalse(templateTree.has("formulaEnvironment"));
    assertEquals(
        FieldRequirement.OPTIONAL,
        GridGrindProtocolCatalog.catalog()
            .requestType()
            .field("execution")
            .orElseThrow()
            .requirement());
    assertEquals(
        FieldRequirement.OPTIONAL,
        GridGrindProtocolCatalog.catalog()
            .requestType()
            .field("formulaEnvironment")
            .orElseThrow()
            .requirement());
    assertEquals(template, decoded);
  }

  @Test
  void exposesStepMutationAssertionAndInspectionTypeGroups() throws IOException {
    Catalog catalog = GridGrindProtocolCatalog.catalog();
    Catalog decoded =
        GridGrindJson.readProtocolCatalog(GridGrindJsonOutput.writeProtocolCatalogBytes(catalog));

    assertFalse(catalog.stepTypes().isEmpty());
    assertFalse(catalog.mutationActionTypes().isEmpty());
    assertFalse(catalog.assertionTypes().isEmpty());
    assertFalse(catalog.inspectionQueryTypes().isEmpty());
    assertTrue(GridGrindProtocolCatalog.entryFor("MUTATION").isPresent());
    assertTrue(GridGrindProtocolCatalog.entryFor("ASSERTION").isPresent());
    assertTrue(GridGrindProtocolCatalog.entryFor("SET_CELL").isPresent());
    assertTrue(GridGrindProtocolCatalog.entryFor("EXPECT_CELL_VALUE").isPresent());
    assertTrue(GridGrindProtocolCatalog.entryFor("GET_CELLS").isPresent());
    assertTrue(GridGrindProtocolCatalog.entryFor("cellInputTypes:FORMULA").isPresent());
    assertFalse(GridGrindProtocolCatalog.entryFor("FORMULA").isPresent());
    assertFalse(GridGrindProtocolCatalog.entryFor(":FORMULA").isPresent());
    assertFalse(GridGrindProtocolCatalog.entryFor("cellInputTypes:").isPresent());
    assertTrue(GridGrindProtocolCatalog.lookupValueFor("cellInputTypes").isPresent());
    assertTrue(GridGrindProtocolCatalog.lookupValueFor("nestedTypes:cellInputTypes").isPresent());
    assertTrue(GridGrindProtocolCatalog.lookupValueFor("chartInputType").isPresent());
    assertTrue(GridGrindProtocolCatalog.lookupValueFor("plainTypes:chartInputType").isPresent());
    assertTrue(GridGrindProtocolCatalog.lookupValueFor(":cellInputTypes").isEmpty());
    assertTrue(GridGrindProtocolCatalog.lookupValueFor("nestedTypes:").isEmpty());
    assertTrue(
        GridGrindProtocolCatalog.matchingEntryIds("FORMULA").contains("cellInputTypes:FORMULA"));
    assertTrue(
        GridGrindProtocolCatalog.matchingEntryIds("FORMULA")
            .contains("namedRangeReportTypes:FORMULA"));
    assertTrue(GridGrindProtocolCatalog.matchingEntryIds(":FORMULA").isEmpty());
    assertTrue(GridGrindProtocolCatalog.matchingEntryIds("cellInputTypes:").isEmpty());
    assertEquals(
        List.of("nestedTypes:cellInputTypes"),
        GridGrindProtocolCatalog.matchingLookupIds("cellInputTypes"));
    assertEquals(
        List.of("plainTypes:chartInputType"),
        GridGrindProtocolCatalog.matchingLookupIds("chartInputType"));
    CatalogSearchResult search = GridGrindProtocolCatalog.searchCatalog("sheet layout");
    assertEquals("sheet layout", search.query());
    assertTrue(
        search.matches().stream()
            .anyMatch(
                match -> "inspectionQueryTypes:GET_SHEET_LAYOUT".equals(match.qualifiedId())));
    assertEquals(
        "source",
        GridGrindProtocolCatalog.entryFor("cellInputTypes:FORMULA")
            .orElseThrow()
            .field("source")
            .orElseThrow()
            .name());
    NestedTypeGroup cellInputs =
        (NestedTypeGroup)
            GridGrindProtocolCatalog.lookupValueFor("nestedTypes:cellInputTypes").orElseThrow();
    assertEquals("cellInputTypes", cellInputs.group());
    assertEquals("TEXT", cellInputs.types().get(1).id());
    PlainTypeGroup chartInput =
        (PlainTypeGroup)
            GridGrindProtocolCatalog.lookupValueFor("plainTypes:chartInputType").orElseThrow();
    assertEquals("chartInputType", chartInput.group());
    assertEquals("ChartInput", chartInput.type().id());
    assertEquals(catalog, decoded);
    assertEquals(
        List.of(new TargetSelectorEntry("TableSelector", List.of("TABLE_BY_NAME_ON_SHEET"))),
        GridGrindProtocolCatalog.entryFor("SET_TABLE").orElseThrow().targetSelectors());
    assertEquals(
        Optional.of("Matches the nested analysis query's target selectors."),
        GridGrindProtocolCatalog.entryFor("EXPECT_ANALYSIS_FINDING_PRESENT")
            .orElseThrow()
            .targetSelectorRule());
    TypeEntry present = GridGrindProtocolCatalog.entryFor("EXPECT_TABLE_PRESENT").orElseThrow();
    assertEquals(
        List.of(
            new TargetSelectorEntry(
                "TableSelector",
                List.of("TABLE_ALL", "TABLE_BY_NAME", "TABLE_BY_NAMES", "TABLE_BY_NAME_ON_SHEET"))),
        present.targetSelectors());
    assertEquals(Optional.empty(), present.targetSelectorRule());
  }

  @Test
  void assertionTargetSelectorsNeverReuseOneWireTypeAcrossMultipleFamilies() {
    for (TypeEntry assertionType : GridGrindProtocolCatalog.catalog().assertionTypes()) {
      assertSelectorFamiliesDoNotReuseTypeIds(assertionType);
    }
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static void assertSelectorFamiliesDoNotReuseTypeIds(TypeEntry assertionType) {
    Map<String, String> familyByTypeId = new java.util.TreeMap<>();
    for (TargetSelectorEntry targetSelector : assertionType.targetSelectors()) {
      for (String typeId : targetSelector.typeIds()) {
        String previousFamily = familyByTypeId.putIfAbsent(typeId, targetSelector.family());
        assertEquals(
            null,
            previousFamily,
            () ->
                assertionType.id()
                    + " must not reuse selector type "
                    + typeId
                    + " across "
                    + previousFamily
                    + " and "
                    + targetSelector.family());
      }
    }
  }

  @Test
  void requiredFieldsComeFromRecordMetadataRatherThanCatalogArguments() {
    assertEquals(
        List.of("stepId", "target", "action"),
        CatalogTypeEntryFactory.requiredFields(MutationStep.class));
  }

  @Test
  void requiredFieldsExcludeCatalogAndJsonIgnoredRecordComponents() {
    assertEquals(
        List.of("visible"),
        CatalogTypeEntryFactory.requiredFields(CatalogIgnoredComponentRecord.class));
    assertEquals(
        List.of("visible"),
        CatalogTypeEntryFactory.requiredFields(CatalogIgnoredAccessorRecord.class));
    assertEquals(
        List.of("visible"),
        CatalogTypeEntryFactory.requiredFields(JsonIgnoredComponentRecord.class));
    assertEquals(
        List.of("visible"),
        CatalogTypeEntryFactory.requiredFields(JsonIgnoredAccessorRecord.class));
  }

  @Test
  void requestTemplateAndCatalogEncodeDeterministically() throws IOException {
    assertArrayEquals(
        GridGrindJsonOutput.writeRequestBytes(GridGrindProtocolCatalog.requestTemplate()),
        GridGrindJsonOutput.writeRequestBytes(GridGrindProtocolCatalog.requestTemplate()));
    assertArrayEquals(
        GridGrindJsonOutput.writeProtocolCatalogBytes(GridGrindProtocolCatalog.catalog()),
        GridGrindJsonOutput.writeProtocolCatalogBytes(GridGrindProtocolCatalog.catalog()));
  }

  @Test
  void requestOwnedPathRulePublishesOnceAndRemainsSearchableAcrossPathBearingEntries()
      throws IOException {
    Catalog catalog = GridGrindProtocolCatalog.catalog();
    String noteText = GridGrindRequestSurfaceContractText.requestOwnedPathResolutionSummary();
    String rendered =
        new String(GridGrindJsonOutput.writeProtocolCatalogBytes(catalog), StandardCharsets.UTF_8);

    assertEquals(
        List.of(
            new CatalogNote(GridGrindProtocolCatalogNotes.REQUEST_OWNED_PATH_RULE_ID, noteText)),
        catalog.notes());
    assertEquals(
        List.of(GridGrindProtocolCatalogNotes.REQUEST_OWNED_PATH_RULE_ID),
        GridGrindProtocolCatalog.entryFor("EXISTING").orElseThrow().noteRefs());
    assertEquals(
        List.of(GridGrindProtocolCatalogNotes.REQUEST_OWNED_PATH_RULE_ID),
        GridGrindProtocolCatalog.entryFor("SAVE_AS").orElseThrow().noteRefs());
    assertEquals(
        List.of(GridGrindProtocolCatalogNotes.REQUEST_OWNED_PATH_RULE_ID),
        ((PlainTypeGroup)
                GridGrindProtocolCatalog.lookupValueFor(
                        "plainTypes:formulaExternalWorkbookInputType")
                    .orElseThrow())
            .type()
            .noteRefs());
    assertEquals(1, occurrenceCount(rendered, noteText));

    CatalogSearchResult search = GridGrindProtocolCatalog.searchCatalog("request-owned paths");
    assertTrue(
        search.matches().stream()
            .anyMatch(match -> "sourceTypes:EXISTING".equals(match.qualifiedId())));
    assertTrue(
        search.matches().stream()
            .anyMatch(match -> "persistenceTypes:SAVE_AS".equals(match.qualifiedId())));
    assertTrue(
        search.matches().stream()
            .anyMatch(
                match ->
                    "plainTypes:formulaExternalWorkbookInputType".equals(match.qualifiedId())));
  }

  @Test
  void executionPolicyCatalogEntryPublishesIndependentlyOptionalAxes() {
    PlainTypeGroup executionPolicy =
        (PlainTypeGroup)
            GridGrindProtocolCatalog.lookupValueFor("plainTypes:executionPolicyInputType")
                .orElseThrow();
    TypeEntry entry = executionPolicy.type();

    assertEquals(FieldRequirement.OPTIONAL, entry.field("mode").orElseThrow().requirement());
    assertEquals(FieldRequirement.OPTIONAL, entry.field("journal").orElseThrow().requirement());
    assertEquals(FieldRequirement.OPTIONAL, entry.field("calculation").orElseThrow().requirement());
    assertEquals(
        FieldRequirement.OPTIONAL, entry.field("assertionMode").orElseThrow().requirement());
    assertTrue(entry.summary().contains("omit any nested execution field"));
    assertTrue(entry.summary().contains("FAIL_FAST"));
  }

  @Test
  void catalogPublishesStructuredProgressInsteadOfJournalEvents() {
    NestedTypeGroup progress =
        (NestedTypeGroup)
            GridGrindProtocolCatalog.lookupValueFor("nestedTypes:executionProgressEventTypes")
                .orElseThrow();
    TypeEntry started =
        progress.types().stream()
            .filter(type -> "STARTED".equals(type.id()))
            .findFirst()
            .orElseThrow();
    TypeEntry failed =
        progress.types().stream()
            .filter(type -> "FAILED".equals(type.id()))
            .findFirst()
            .orElseThrow();

    assertEquals("status", progress.discriminatorField());
    assertTrue(started.field("timestamp").isPresent());
    assertTrue(started.field("category").orElseThrow().enumValues().contains("STEP"));
    assertTrue(failed.field("problemCode").orElseThrow().enumValues().contains("IO_ERROR"));
    assertFalse(started.field("problemCode").isPresent());
    assertFalse(failed.field("detail").isPresent());
    assertTrue(
        GridGrindProtocolCatalog.lookupValueFor("plainTypes:executionJournalEventType").isEmpty());
  }

  @Test
  void ooxmlWriteEncryptionCatalogEntryPublishesModeLessStrongOnlyDefaults() {
    PlainTypeGroup encryption =
        (PlainTypeGroup)
            GridGrindProtocolCatalog.lookupValueFor("plainTypes:ooxmlEncryptionInputType")
                .orElseThrow();
    TypeEntry entry = encryption.type();

    assertTrue(entry.field("password").isPresent());
    assertTrue(entry.field("password").orElseThrow().secret());
    assertFalse(entry.field("mode").isPresent());
    assertEquals(FieldRequirement.OPTIONAL, entry.field("cipher").orElseThrow().requirement());
    assertEquals(FieldRequirement.OPTIONAL, entry.field("hash").orElseThrow().requirement());
    assertTrue(entry.summary().contains("AGILE packages only"));
    assertTrue(entry.summary().contains("AES_256"));
    assertTrue(entry.summary().contains("SHA_512"));
  }

  @Test
  void catalogPublishesRecordDeclaredSecretMarkers() {
    PlainTypeGroup signature =
        (PlainTypeGroup)
            GridGrindProtocolCatalog.lookupValueFor("plainTypes:ooxmlSignatureInputType")
                .orElseThrow();

    assertTrue(signature.type().field("keystorePassword").orElseThrow().secret());
    assertTrue(signature.type().field("keyPassword").orElseThrow().secret());
    assertFalse(signature.type().field("pkcs12Path").orElseThrow().secret());
  }

  @Test
  void projectionAwareReadQueriesPublishTheirDefaultedFieldsAsOptional() {
    TypeEntry getCells = GridGrindProtocolCatalog.entryFor("GET_CELLS").orElseThrow();
    TypeEntry getWindow = GridGrindProtocolCatalog.entryFor("GET_WINDOW").orElseThrow();
    TypeEntry getSheetSchema = GridGrindProtocolCatalog.entryFor("GET_SHEET_SCHEMA").orElseThrow();

    assertEquals(
        FieldRequirement.OPTIONAL, getCells.field("projection").orElseThrow().requirement());
    assertTrue(getCells.summary().contains("Omit projection"));

    assertEquals(
        FieldRequirement.OPTIONAL, getWindow.field("projection").orElseThrow().requirement());
    assertEquals(
        FieldRequirement.OPTIONAL, getWindow.field("includeBlanks").orElseThrow().requirement());
    assertEquals(
        Optional.of(false), getWindow.field("includeBlanks").orElseThrow().defaultBoolean());
    assertTrue(getWindow.summary().contains("sparse default"));

    assertEquals(
        FieldRequirement.OPTIONAL, getSheetSchema.field("projection").orElseThrow().requirement());
    assertTrue(getSheetSchema.summary().contains("Omit projection"));
  }

  @Test
  void cellReadCatalogPublishesFacetGatingAsMachineReadableFieldMetadata() {
    NestedTypeGroup group =
        (NestedTypeGroup)
            GridGrindProtocolCatalog.lookupValueFor("nestedTypes:cellReportTypes").orElseThrow();
    NestedTypeGroup valueGroup =
        (NestedTypeGroup)
            GridGrindProtocolCatalog.lookupValueFor("nestedTypes:cellValueReportTypes")
                .orElseThrow();
    NestedTypeGroup inputGroup =
        (NestedTypeGroup)
            GridGrindProtocolCatalog.lookupValueFor("nestedTypes:cellInputTypes").orElseThrow();
    NestedTypeGroup rowInputGroup =
        (NestedTypeGroup)
            GridGrindProtocolCatalog.lookupValueFor("nestedTypes:cellRowInputTypes").orElseThrow();
    NestedTypeGroup gridInputGroup =
        (NestedTypeGroup)
            GridGrindProtocolCatalog.lookupValueFor("nestedTypes:cellGridInputTypes").orElseThrow();
    TypeEntry number =
        group.types().stream().filter(type -> "NUMBER".equals(type.id())).findFirst().orElseThrow();
    TypeEntry formula =
        group.types().stream()
            .filter(type -> "FORMULA".equals(type.id()))
            .findFirst()
            .orElseThrow();
    TypeEntry error =
        group.types().stream().filter(type -> "ERROR".equals(type.id())).findFirst().orElseThrow();
    TypeEntry text =
        group.types().stream().filter(type -> "TEXT".equals(type.id())).findFirst().orElseThrow();
    TypeEntry evaluatedError =
        valueGroup.types().stream()
            .filter(type -> "ERROR".equals(type.id()))
            .findFirst()
            .orElseThrow();
    TypeEntry evaluatedText =
        valueGroup.types().stream()
            .filter(type -> "TEXT".equals(type.id()))
            .findFirst()
            .orElseThrow();
    TypeEntry evaluatedNumber =
        valueGroup.types().stream()
            .filter(type -> "NUMBER".equals(type.id()))
            .findFirst()
            .orElseThrow();
    TypeEntry inputError =
        inputGroup.types().stream()
            .filter(type -> "ERROR".equals(type.id()))
            .findFirst()
            .orElseThrow();
    TypeEntry rowInputError =
        rowInputGroup.types().stream()
            .filter(type -> "ERROR".equals(type.id()))
            .findFirst()
            .orElseThrow();
    TypeEntry gridInputError =
        gridInputGroup.types().stream()
            .filter(type -> "ERROR".equals(type.id()))
            .findFirst()
            .orElseThrow();
    PlainTypeGroup projection =
        (PlainTypeGroup)
            GridGrindProtocolCatalog.lookupValueFor("plainTypes:cellReadProjectionType")
                .orElseThrow();
    List<String> storedErrorLiterals =
        List.of("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A");
    List<String> reportedErrorLiterals =
        List.of(
            "#NULL!",
            "#DIV/0!",
            "#VALUE!",
            "#REF!",
            "#NAME?",
            "#NUM!",
            "#N/A",
            "#CIRCULAR_REF!",
            "#FUNCTION_NOT_IMPLEMENTED!");

    assertEquals(FieldRequirement.REQUIRED, number.field("address").orElseThrow().requirement());
    assertTrue(number.field("address").orElseThrow().projectedByFacets().isEmpty());
    assertEquals(
        FieldRequirement.OPTIONAL, number.field("displayValue").orElseThrow().requirement());
    assertEquals(List.of("FORMAT"), number.field("displayValue").orElseThrow().projectedByFacets());
    assertEquals(FieldRequirement.OPTIONAL, number.field("style").orElseThrow().requirement());
    assertEquals(List.of("STYLE"), number.field("style").orElseThrow().projectedByFacets());
    assertEquals(FieldRequirement.OPTIONAL, number.field("hyperlink").orElseThrow().requirement());
    assertEquals(List.of("HYPERLINK"), number.field("hyperlink").orElseThrow().projectedByFacets());
    assertEquals(FieldRequirement.OPTIONAL, number.field("comment").orElseThrow().requirement());
    assertEquals(List.of("COMMENT"), number.field("comment").orElseThrow().projectedByFacets());
    assertEquals(
        FieldRequirement.OPTIONAL, number.field("numberValue").orElseThrow().requirement());
    assertEquals(List.of("VALUE"), number.field("numberValue").orElseThrow().projectedByFacets());
    assertEquals(FieldRequirement.OPTIONAL, number.field("temporal").orElseThrow().requirement());
    assertEquals(List.of("TEMPORAL"), number.field("temporal").orElseThrow().projectedByFacets());
    assertTrue(number.summary().contains("TEMPORAL"));

    assertEquals(FieldRequirement.OPTIONAL, text.field("textValue").orElseThrow().requirement());
    assertEquals(List.of("VALUE"), text.field("textValue").orElseThrow().projectedByFacets());
    assertEquals(FieldRequirement.OPTIONAL, text.field("runs").orElseThrow().requirement());
    assertEquals(List.of("RICH_TEXT_RUNS"), text.field("runs").orElseThrow().projectedByFacets());
    assertEquals(reportedErrorLiterals, error.field("errorValue").orElseThrow().enumValues());
    assertEquals(FieldRequirement.OPTIONAL, formula.field("formula").orElseThrow().requirement());
    assertEquals(List.of("FORMULA"), formula.field("formula").orElseThrow().projectedByFacets());
    assertEquals(
        FieldRequirement.OPTIONAL, formula.field("evaluation").orElseThrow().requirement());
    assertEquals(List.of("VALUE"), formula.field("evaluation").orElseThrow().projectedByFacets());

    assertEquals(
        FieldRequirement.REQUIRED, evaluatedText.field("textValue").orElseThrow().requirement());
    assertTrue(evaluatedText.field("textValue").orElseThrow().projectedByFacets().isEmpty());
    assertEquals(
        FieldRequirement.OPTIONAL, evaluatedText.field("runs").orElseThrow().requirement());
    assertEquals(
        List.of("RICH_TEXT_RUNS"), evaluatedText.field("runs").orElseThrow().projectedByFacets());
    assertEquals(
        FieldRequirement.REQUIRED,
        evaluatedNumber.field("numberValue").orElseThrow().requirement());
    assertTrue(evaluatedNumber.field("numberValue").orElseThrow().projectedByFacets().isEmpty());
    assertEquals(
        reportedErrorLiterals, evaluatedError.field("errorValue").orElseThrow().enumValues());
    assertEquals(
        FieldRequirement.OPTIONAL, evaluatedNumber.field("temporal").orElseThrow().requirement());
    assertEquals(
        List.of("TEMPORAL"), evaluatedNumber.field("temporal").orElseThrow().projectedByFacets());
    assertEquals(storedErrorLiterals, inputError.field("error").orElseThrow().enumValues());
    assertEquals(storedErrorLiterals, rowInputError.field("cells").orElseThrow().enumValues());
    assertEquals(storedErrorLiterals, gridInputError.field("cells").orElseThrow().enumValues());
    assertEquals(
        "#FUNCTION_NOT_IMPLEMENTED!",
        error.field("errorValue").orElseThrow().enumValueDocs().get(8).value());
    assertEquals(
        List.of(
            new EnumValueDocEntry(
                "VALUE",
                "Project the factual cell value: textValue, numberValue, booleanValue,"
                    + " errorValue, or formula evaluation."),
            new EnumValueDocEntry("STYLE", "Project the style report for each returned cell."),
            new EnumValueDocEntry(
                "FORMAT", "Project displayValue using Excel's formatted display text."),
            new EnumValueDocEntry(
                "HYPERLINK", "Project hyperlink metadata when the cell carries a hyperlink."),
            new EnumValueDocEntry(
                "COMMENT", "Project comment metadata when the cell carries a comment."),
            new EnumValueDocEntry("FORMULA", "Project authored formula text for formula cells."),
            new EnumValueDocEntry(
                "RICH_TEXT_RUNS",
                "Project rich-text runs for text cells and text-valued formula evaluations."),
            new EnumValueDocEntry(
                "TEMPORAL",
                "Project derived date, time, or date-time semantics for date-like numeric"
                    + " values.")),
        projection.type().field("facets").orElseThrow().enumValueDocs());
  }

  @Test
  void validatesCatalogCoverageFailurePaths() {
    IllegalStateException missingNestedDescriptor =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalogCoverageValidator.validateReverseGroupMappings(
                    Set.of(), Set.of()));
    assertTrue(
        missingNestedDescriptor
            .getMessage()
            .contains("Field-shape nested group map contains type with no catalog descriptor"));

    IllegalStateException badWorkbookStepCoverage =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalogCoverageValidator.validateCoverage(
                    dev.erst.gridgrind.contract.step.WorkbookStep.class,
                    Map.of(MutationStep.class, "MUTATION")));
    assertTrue(badWorkbookStepCoverage.getMessage().contains("Catalog coverage mismatch"));

    IllegalStateException nonRecordCoverage =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalogCoverageValidator.validateCoverage(
                    NonAnnotatedSealedType.class, Map.of(NotARecord.class, "BROKEN")));
    assertEquals(
        "Catalog coverage requires "
            + NonAnnotatedSealedType.class.getName()
            + " to declare a non-blank @JsonTypeInfo property",
        nonRecordCoverage.getMessage());

    IllegalStateException blankDiscriminatorCoverage =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalogCoverageValidator.validateCoverage(
                    BlankPropertySealedType.class, Map.of(BlankPropertyRecord.class, "BROKEN")));
    assertEquals(
        "Catalog coverage requires "
            + BlankPropertySealedType.class.getName()
            + " to declare a non-blank @JsonTypeInfo property",
        blankDiscriminatorCoverage.getMessage());
  }

  @Test
  void validatesProjectedFieldMetadataFailurePaths() {
    CatalogProjectedField projectedField = new CatalogProjectedField("displayValue", "FORMAT");

    assertEquals("displayValue", projectedField.name());
    assertEquals(List.of("FORMAT"), projectedField.projectedByFacets());
    assertEquals(
        "projectedByFacets must not be empty",
        assertThrows(
                IllegalArgumentException.class, () -> new CatalogProjectedField("displayValue"))
            .getMessage());
    assertEquals(
        "projectedByFacets require the field requirement to be OPTIONAL",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new FieldEntry(
                        "displayValue",
                        FieldRequirement.REQUIRED,
                        new FieldShape.Scalar(ScalarType.STRING),
                        List.of(),
                        Optional.empty(),
                        List.of(),
                        List.of("FORMAT"),
                        false))
            .getMessage());
    assertEquals(
        "enumValueDocs must document every enumValues entry in published order",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new FieldEntry(
                        "facets",
                        FieldRequirement.REQUIRED,
                        new FieldShape.Scalar(ScalarType.STRING),
                        List.of("VALUE", "FORMAT"),
                        Optional.empty(),
                        List.of(new EnumValueDocEntry("FORMAT", "Only format.")),
                        List.of(),
                        false))
            .getMessage());
    assertEquals(
        "defaultBoolean requires an OPTIONAL BOOLEAN field requirement",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new FieldEntry(
                        "enabled",
                        FieldRequirement.REQUIRED,
                        new FieldShape.Scalar(ScalarType.BOOLEAN),
                        List.of(),
                        Optional.of(false),
                        List.of(),
                        List.of(),
                        false))
            .getMessage());
    assertEquals(
        "defaultBoolean requires an OPTIONAL BOOLEAN field requirement",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new FieldEntry(
                        "label",
                        FieldRequirement.OPTIONAL,
                        new FieldShape.Scalar(ScalarType.STRING),
                        List.of(),
                        Optional.of(false),
                        List.of(),
                        List.of(),
                        false))
            .getMessage());

    IllegalStateException missingProjectedField =
        assertThrows(
            IllegalStateException.class,
            () ->
                CatalogTypeEntryFactory.requiredFields(
                    MutationStep.class, List.of(new CatalogProjectedField("missing", "FORMAT"))));
    assertEquals(
        "Catalog projected field 'missing' does not exist on " + MutationStep.class.getName(),
        missingProjectedField.getMessage());
  }

  private record CatalogIgnoredComponentRecord(String visible, @CatalogIgnored String hidden) {}

  private record CatalogIgnoredAccessorRecord(String visible, String hidden) {
    @SuppressWarnings("UnusedMethod")
    @CatalogIgnored
    @Override
    public String hidden() {
      return hidden;
    }
  }

  private record JsonIgnoredComponentRecord(String visible, @JsonIgnore String hidden) {}

  private record JsonIgnoredAccessorRecord(String visible, String hidden) {
    @SuppressWarnings("UnusedMethod")
    @JsonIgnore
    @Override
    public String hidden() {
      return hidden;
    }
  }

  @Test
  void typeEntriesExposeOptionalFieldLookup() {
    TypeEntry typeEntry =
        new TypeEntry(
            "ASSERTION",
            "summary",
            List.of(
                new FieldEntry(
                    "target",
                    FieldRequirement.REQUIRED,
                    new FieldShape.Scalar(ScalarType.STRING),
                    List.of())));

    assertEquals("target", typeEntry.field("target").orElseThrow().name());
    assertTrue(typeEntry.field("target").orElseThrow().projectedByFacets().isEmpty());
    assertTrue(typeEntry.field("missing").isEmpty());
    assertTrue(typeEntry.targetSelectors().isEmpty());
    assertEquals(Optional.empty(), typeEntry.targetSelectorRule());
    assertEquals(
        "name must not be null",
        assertThrows(NullPointerException.class, () -> typeEntry.field(null)).getMessage());
  }

  @Test
  void typeEntriesCopyTargetSelectorMetadataAndRejectBlankRules() {
    TypeEntry typeEntry =
        new TypeEntry(
            "SET_TABLE",
            "summary",
            List.of(),
            List.of(new TargetSelectorEntry("TableSelector", List.of("TABLE_BY_NAME_ON_SHEET"))),
            Optional.of("Requires the table selector family."));

    assertEquals(
        List.of(new TargetSelectorEntry("TableSelector", List.of("TABLE_BY_NAME_ON_SHEET"))),
        typeEntry.targetSelectors());
    assertEquals(
        Optional.of("Requires the table selector family."), typeEntry.targetSelectorRule());
    assertEquals(
        "targetSelectorRule must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new TypeEntry("BROKEN", "summary", List.of(), List.of(), Optional.of(" ")))
            .getMessage());
  }

  /** Non-annotated sealed type used for discriminator validation. */
  private sealed interface NonAnnotatedSealedType permits NotARecord {}

  /** Non-record subtype used to verify coverage rejection. */
  private static final class NotARecord implements NonAnnotatedSealedType {}

  /** Sealed type with a blank JsonTypeInfo property to cover catalog discriminator validation. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = " ")
  private sealed interface BlankPropertySealedType permits BlankPropertyRecord {}

  /** Record subtype for blank-property discriminator coverage. */
  private record BlankPropertyRecord() implements BlankPropertySealedType {}

  private static int occurrenceCount(String text, String needle) {
    return text.split(java.util.regex.Pattern.quote(needle), -1).length - 1;
  }
}
