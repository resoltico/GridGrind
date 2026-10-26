package dev.erst.gridgrind.core.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.CellReadFacet;
import dev.erst.gridgrind.core.plan.CellReadProjection;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelChartGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelReportedCellErrorLiteral;
import dev.erst.gridgrind.core.vocabulary.ExcelStoredCellErrorLiteral;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import org.junit.jupiter.api.Test;

/** Direct coverage for field-shape metadata resolution and validation failure paths. */
class CatalogFieldMetadataSupportTest {
  @Test
  void resolvesScalarListUnionAndTopLevelFieldShapes() throws Exception {
    FieldEntry requiredString =
        CatalogFieldMetadataSupport.fieldEntry(
            recordComponent(MetadataFixture.class, "name"), Set.of("notes"));
    FieldEntry optionalList =
        CatalogFieldMetadataSupport.fieldEntry(
            recordComponent(MetadataFixture.class, "notes"), Set.of("notes"));

    assertEquals(FieldRequirement.REQUIRED, requiredString.requirement());
    assertEquals(new FieldShape.Scalar(ScalarType.STRING), requiredString.shape());
    assertEquals(FieldRequirement.OPTIONAL, optionalList.requirement());
    assertEquals(
        new FieldShape.ListShape(new FieldShape.Scalar(ScalarType.STRING)), optionalList.shape());
    assertEquals(
        new FieldShape.TopLevelTypeSetRef("workbookStepTypes"),
        CatalogFieldMetadataSupport.fieldShape(WorkbookStep.class));
    assertEquals(
        List.of("STANDARD", "STACKED", "PERCENT_STACKED"),
        CatalogFieldMetadataSupport.fieldEntry(
                recordComponent(MetadataFixture.class, "grouping"), Set.of("notes"))
            .enumValues());
    assertEquals(
        List.of("STANDARD", "STACKED", "PERCENT_STACKED"),
        CatalogFieldMetadataSupport.fieldEntry(
                recordComponent(MetadataFixture.class, "groupings"), Set.of("notes"))
            .enumValues());
    assertEquals(
        List.of(),
        CatalogFieldMetadataSupport.enumValues(
            recordComponent(MetadataFixture.class, "cells").getGenericType()));
    assertEquals(
        new FieldShape.NestedTypeGroupUnionRef(
            List.of(
                "workbookSelectorTypes",
                "sheetSelectorTypes",
                "cellSelectorTypes",
                "rangeSelectorTypes",
                "rowBandSelectorTypes",
                "columnBandSelectorTypes",
                "tableSelectorTypes",
                "tableRowSelectorTypes",
                "tableCellSelectorTypes",
                "namedRangeSelectorTypes",
                "drawingObjectSelectorTypes",
                "chartSelectorTypes",
                "pivotTableSelectorTypes")),
        CatalogFieldMetadataSupport.fieldShape(dev.erst.gridgrind.core.plan.Selector.class));
    assertEquals(
        new FieldShape.NestedTypeGroupRef("executionModeSpecTypes"),
        CatalogFieldMetadataSupport.fieldShape(ExecutionModeSpec.class));
    assertEquals(
        new FieldShape.NestedTypeGroupRef("textSourceSpecTypes"),
        CatalogFieldMetadataSupport.fieldShape(TextSourceSpec.class));
    assertEquals(
        new FieldShape.NestedTypeGroupRef("binarySourceSpecTypes"),
        CatalogFieldMetadataSupport.fieldShape(BinarySourceSpec.class));
  }

  @Test
  void validatesUnsupportedAndAmbiguousFieldShapeInputs() throws ReflectiveOperationException {
    IllegalStateException unsupportedType =
        assertThrows(
            IllegalStateException.class,
            () -> CatalogFieldMetadataSupport.fieldShape((Type) new UnsupportedType()));
    assertEquals("Unsupported catalog field type: unsupported-type", unsupportedType.getMessage());

    IllegalStateException unsupportedParameterized =
        assertThrows(
            IllegalStateException.class,
            () ->
                CatalogFieldMetadataSupport.fieldShape(
                    (ParameterizedType)
                        recordComponent(MetadataFixture.class, "cells").getGenericType()));
    assertEquals(
        "Unsupported parameterized catalog field type: java.util.Map<java.lang.String, dev.erst.gridgrind.core.model.value.CellValueSpec>",
        unsupportedParameterized.getMessage());

    assertEquals(
        "Unsupported catalog field type: java.lang.Object",
        assertThrows(
                IllegalStateException.class,
                () -> CatalogFieldMetadataSupport.fieldShape(Object.class))
            .getMessage());
    assertEquals(
        "List field must declare exactly one type argument: java.util.List<java.lang.String,java.lang.Integer>",
        assertThrows(
                IllegalStateException.class,
                () -> CatalogFieldMetadataSupport.fieldShape(new InvalidListParameterizedType()))
            .getMessage());
  }

  @Test
  void marksOnlyExplicitSecretProtocolFieldsAsSecret() throws ReflectiveOperationException {
    assertTrue(
        CatalogFieldMetadataSupport.fieldEntry(
                recordComponent(SecretFixture.class, "password"), Set.of())
            .secret());
    assertFalse(
        CatalogFieldMetadataSupport.fieldEntry(
                recordComponent(SecretFixture.class, "label"), Set.of())
            .secret());
  }

  @Test
  void fieldEntryCanPublishProjectedFacetMetadata() throws Exception {
    FieldEntry projected =
        CatalogFieldMetadataSupport.fieldEntry(
            recordComponent(MetadataFixture.class, "name"),
            Set.of("name"),
            Map.of("name", List.of("VALUE", "FORMAT")));

    assertEquals(FieldRequirement.OPTIONAL, projected.requirement());
    assertEquals(List.of("VALUE", "FORMAT"), projected.projectedByFacets());
  }

  @Test
  void fieldEntryCanPublishEnumValueDocsForFacetSelectors() throws Exception {
    FieldEntry facets =
        CatalogFieldMetadataSupport.fieldEntry(
            recordComponent(CellReadProjection.class, "facets"), Set.of());
    FieldEntry errorValue =
        CatalogFieldMetadataSupport.fieldEntry(
            recordComponent(CellValueSpec.ErrorValue.class, "error"), Set.of());
    FieldEntry errorRowValues =
        CatalogFieldMetadataSupport.fieldEntry(
            recordComponent(
                dev.erst.gridgrind.core.model.value.CellRowSpec.ErrorValues.class, "cells"),
            Set.of());
    FieldEntry reportedErrorValue =
        CatalogFieldMetadataSupport.fieldEntry(
            recordComponent(
                dev.erst.gridgrind.core.model.value.CellValueReport.ErrorValue.class, "errorValue"),
            Set.of());

    assertEquals(
        List.of(
            "VALUE",
            "STYLE",
            "FORMAT",
            "HYPERLINK",
            "COMMENT",
            "FORMULA",
            "RICH_TEXT_RUNS",
            "TEMPORAL"),
        facets.enumValues());
    assertEquals("FORMAT", facets.enumValueDocs().get(2).value());
    assertTrue(facets.enumValueDocs().get(2).summary().contains("displayValue"));
    assertEquals("VALUE", facets.enumValueDocs().getFirst().value());
    assertEquals(
        List.of("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"),
        errorValue.enumValues());
    assertEquals(errorValue.enumValues(), errorRowValues.enumValues());
    assertEquals("#NULL!", errorValue.enumValueDocs().getFirst().value());
    assertEquals("#CIRCULAR_REF!", reportedErrorValue.enumValueDocs().get(7).value());
    assertEquals(
        ExcelStoredCellErrorLiteral.orderedWireValues(),
        CatalogFieldMetadataSupport.enumValues(ExcelStoredCellErrorLiteral.class));
    assertEquals(
        ExcelReportedCellErrorLiteral.orderedWireValues(),
        CatalogFieldMetadataSupport.enumValues(ExcelReportedCellErrorLiteral.class));
  }

  @Test
  void enumValueDocumentationRequiresCompleteCoverageOfPublishedTokens() {
    IllegalStateException missingDoc =
        assertThrows(
            IllegalStateException.class,
            () ->
                CatalogEnumValueDocumentationSupport.orderedEnumDocs(
                    CellReadFacet.class, Map.of("VALUE", "Only one token is documented.")));

    assertEquals(
        "Enum value docs must cover every published token for " + CellReadFacet.class.getName(),
        missingDoc.getMessage());
  }

  private record MetadataFixture(
      String name,
      List<String> notes,
      Optional<ExcelChartGrouping> grouping,
      List<ExcelChartGrouping> groupings,
      Map<String, CellValueSpec> cells,
      WorkbookStep step) {}

  /** Type implementation not supported by the catalog field-shape resolver. */
  private static final class UnsupportedType implements Type {
    @Override
    public String getTypeName() {
      return "unsupported-type";
    }

    @Override
    public String toString() {
      return getTypeName();
    }
  }

  private record SecretFixture(
      @dev.erst.gridgrind.core.metadata.ProtocolField(secret = true) String password,
      String label) {}

  /** Invalid synthetic List parameterization used to cover wrong-arity failures. */
  private static final class InvalidListParameterizedType implements ParameterizedType {
    @Override
    public Type[] getActualTypeArguments() {
      return new Type[] {String.class, Integer.class};
    }

    @Override
    public Type getRawType() {
      return java.util.List.class;
    }

    @Override
    public Type getOwnerType() {
      return null;
    }

    @Override
    public String toString() {
      return "java.util.List<java.lang.String,java.lang.Integer>";
    }
  }

  private static java.lang.reflect.RecordComponent recordComponent(
      Class<? extends Record> recordType, String fieldName) {
    return java.util.Arrays.stream(recordType.getRecordComponents())
        .filter(component -> component.getName().equals(fieldName))
        .findFirst()
        .orElseThrow();
  }
}
