package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.step.MutationStep;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Set;
import org.junit.jupiter.api.Test;

/** Tests for the machine-readable catalog and built-in request template. */
class GridGrindProtocolCatalogTest {
  @Test
  void exposesMinimalStepBasedRequestTemplate() throws IOException {
    WorkbookPlan template = GridGrindProtocolCatalog.requestTemplate();
    WorkbookPlan decoded = GridGrindJson.readRequest(GridGrindJson.writeRequestBytes(template));

    assertEquals(GridGrindProtocolVersion.V1, template.protocolVersion());
    assertTrue(template.steps().isEmpty());
    assertEquals(template, decoded);
  }

  @Test
  void exposesStepMutationAssertionAndInspectionTypeGroups() throws IOException {
    Catalog catalog = GridGrindProtocolCatalog.catalog();
    Catalog decoded =
        GridGrindJson.readProtocolCatalog(GridGrindJson.writeProtocolCatalogBytes(catalog));

    assertFalse(catalog.stepTypes().isEmpty());
    assertFalse(catalog.mutationActionTypes().isEmpty());
    assertFalse(catalog.assertionTypes().isEmpty());
    assertFalse(catalog.inspectionQueryTypes().isEmpty());
    assertFalse(catalog.shippedExamples().isEmpty());
    assertFalse(catalog.cliSurface().limits().entries().isEmpty());
    assertEquals("Usage", catalog.cliSurface().usage().label());
    assertEquals(
        "gridgrind --print-example " + GridGrindShippedExamples.examples().getFirst().id(),
        catalog.cliSurface().discovery().printOneExampleCommand());
    assertTrue(GridGrindProtocolCatalog.entryFor("MUTATION").isPresent());
    assertTrue(GridGrindProtocolCatalog.entryFor("ASSERTION").isPresent());
    assertTrue(GridGrindProtocolCatalog.entryFor("SET_CELL").isPresent());
    assertTrue(GridGrindProtocolCatalog.entryFor("EXPECT_CELL_VALUE").isPresent());
    assertTrue(GridGrindProtocolCatalog.entryFor("GET_CELLS").isPresent());
    assertTrue(GridGrindProtocolCatalog.exampleFor("WORKBOOK_HEALTH").isPresent());
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
        List.of(new TargetSelectorEntry("TableSelector", List.of("BY_NAME_ON_SHEET"))),
        GridGrindProtocolCatalog.entryFor("SET_TABLE").orElseThrow().targetSelectors());
    assertEquals(
        "Matches the nested analysis query's target selectors.",
        GridGrindProtocolCatalog.entryFor("EXPECT_ANALYSIS_FINDING_PRESENT")
            .orElseThrow()
            .targetSelectorRule());
    TypeEntry present = GridGrindProtocolCatalog.entryFor("EXPECT_PRESENT").orElseThrow();
    assertEquals(
        List.of(
            new TargetSelectorEntry(
                "NamedRangeSelector",
                List.of("ALL", "ANY_OF", "BY_NAME", "BY_NAMES", "WORKBOOK_SCOPE", "SHEET_SCOPE")),
            new TargetSelectorEntry(
                "TableSelector", List.of("ALL", "BY_NAME", "BY_NAMES", "BY_NAME_ON_SHEET")),
            new TargetSelectorEntry(
                "PivotTableSelector", List.of("ALL", "BY_NAME", "BY_NAMES", "BY_NAME_ON_SHEET")),
            new TargetSelectorEntry("ChartSelector", List.of("ALL_ON_SHEET", "BY_NAME"))),
        present.targetSelectors());
    assertTrue(
        present
            .targetSelectorRule()
            .contains("Shared selector wire types remain family-sensitive"));
  }

  @Test
  void shippedExamplesArePublishedInCatalogOrder() {
    Catalog catalog = GridGrindProtocolCatalog.catalog();

    assertEquals(GridGrindShippedExamples.examples(), GridGrindProtocolCatalog.shippedExamples());
    assertEquals(
        GridGrindShippedExamples.catalogEntries(),
        catalog.shippedExamples(),
        "catalog shippedExamples must mirror the contract-owned generated example registry");
    assertTrue(
        catalog.cliSurface().request().lines().stream()
            .anyMatch(line -> line.contains("STANDARD_INPUT-authored values require --request")));
    assertTrue(
        catalog.cliSurface().request().lines().stream()
            .anyMatch(line -> line.contains("do not send step.type")));
    assertTrue(
        catalog.cliSurface().discovery().lines().stream()
            .anyMatch(line -> line.contains("--doctor-request")));
    assertTrue(
        catalog.cliSurface().discovery().lines().stream()
            .anyMatch(line -> line.contains("--print-task-catalog")));
    assertTrue(
        catalog.cliSurface().discovery().lines().stream()
            .anyMatch(line -> line.contains("--print-task-plan")));
    assertTrue(
        catalog.cliSurface().discovery().lines().stream()
            .anyMatch(line -> line.contains("--print-goal-plan")));
  }

  @Test
  void requiredFieldsFiltersOptionalRecordComponents() {
    assertEquals(
        List.of("stepId", "action"),
        GridGrindProtocolCatalog.requiredFields(MutationStep.class, List.of("target")));
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
  void validatesOptionalFieldDescriptorsAndCoverageFailurePaths() {
    IllegalStateException missingOptionalField =
        assertThrows(
            IllegalStateException.class,
            () -> GridGrindProtocolCatalog.requiredFields(MutationStep.class, List.of("missing")));
    assertEquals(
        "Catalog optional field 'missing' does not exist on " + MutationStep.class.getName(),
        missingOptionalField.getMessage());

    IllegalStateException primitiveOptionalField =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalog.requiredFields(
                    PrimitiveFixture.class, List.of("enabled")));
    assertEquals(
        "Catalog optional field 'enabled' on "
            + PrimitiveFixture.class.getName()
            + " uses primitive component type boolean",
        primitiveOptionalField.getMessage());

    IllegalStateException missingNestedDescriptor =
        assertThrows(
            IllegalStateException.class,
            () -> GridGrindProtocolCatalog.validateReverseGroupMappings(Set.of(), Set.of()));
    assertTrue(
        missingNestedDescriptor
            .getMessage()
            .contains("Field-shape nested group map contains type with no catalog descriptor"));

    IllegalStateException badWorkbookStepCoverage =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalog.validateCoverage(
                    dev.erst.gridgrind.contract.step.WorkbookStep.class,
                    Map.of(MutationStep.class, "MUTATION")));
    assertTrue(badWorkbookStepCoverage.getMessage().contains("Catalog coverage mismatch"));

    IllegalStateException nonRecordCoverage =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolCatalog.validateCoverage(
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
                GridGrindProtocolCatalog.validateCoverage(
                    BlankPropertySealedType.class, Map.of(BlankPropertyRecord.class, "BROKEN")));
    assertEquals(
        "Catalog coverage requires "
            + BlankPropertySealedType.class.getName()
            + " to declare a non-blank @JsonTypeInfo property",
        blankDiscriminatorCoverage.getMessage());
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

    assertEquals("target", typeEntry.field("target").name());
    assertNull(typeEntry.field("missing"));
    assertTrue(typeEntry.targetSelectors().isEmpty());
    assertNull(typeEntry.targetSelectorRule());
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
            List.of(new TargetSelectorEntry("TableSelector", List.of("BY_NAME_ON_SHEET"))),
            "Requires the table selector family.");

    assertEquals(
        List.of(new TargetSelectorEntry("TableSelector", List.of("BY_NAME_ON_SHEET"))),
        typeEntry.targetSelectors());
    assertEquals("Requires the table selector family.", typeEntry.targetSelectorRule());
    assertEquals(
        "targetSelectorRule must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new TypeEntry("BROKEN", "summary", List.of(), List.of(), " "))
            .getMessage());
  }

  /** Primitive record used to verify optional-field validation. */
  private record PrimitiveFixture(boolean enabled) {}

  /** Non-annotated sealed type used for discriminator validation. */
  private sealed interface NonAnnotatedSealedType permits NotARecord {}

  /** Non-record subtype used to verify coverage rejection. */
  private static final class NotARecord implements NonAnnotatedSealedType {}

  /** Sealed type with a blank JsonTypeInfo property to cover catalog discriminator validation. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = " ")
  private sealed interface BlankPropertySealedType permits BlankPropertyRecord {}

  /** Record subtype for blank-property discriminator coverage. */
  private record BlankPropertyRecord() implements BlankPropertySealedType {}
}
