package dev.erst.gridgrind.contract.catalog.gather;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.catalog.FieldEntry;
import dev.erst.gridgrind.contract.catalog.FieldRequirement;
import dev.erst.gridgrind.contract.catalog.FieldShape;
import dev.erst.gridgrind.contract.catalog.ScalarType;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
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
        new FieldShape.TopLevelTypeSetRef("stepTypes"),
        CatalogFieldMetadataSupport.fieldShape(WorkbookStep.class));
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
        CatalogFieldMetadataSupport.fieldShape(
            dev.erst.gridgrind.contract.selector.Selector.class));
    assertEquals(
        new FieldShape.PlainTypeGroupRef("executionModeInputType"),
        CatalogFieldMetadataSupport.fieldShape(ExecutionModeInput.class));
    assertEquals(
        new FieldShape.NestedTypeGroupRef("textSourceTypes"),
        CatalogFieldMetadataSupport.fieldShape(TextSourceInput.class));
    assertEquals(
        new FieldShape.NestedTypeGroupRef("binarySourceTypes"),
        CatalogFieldMetadataSupport.fieldShape(BinarySourceInput.class));
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
        "Unsupported parameterized catalog field type: java.util.Map<java.lang.String, dev.erst.gridgrind.contract.dto.CellInput>",
        unsupportedParameterized.getMessage());

    IllegalStateException nestedMismatch =
        assertThrows(
            IllegalStateException.class,
            () ->
                CatalogFieldMetadataSupport.validateNestedTypeGroupMapping(
                    CellInput.class, "wrongGroup"));
    assertEquals(
        "Field-shape nested group mapping mismatch for dev.erst.gridgrind.contract.dto.CellInput: expected=wrongGroup, mapped=cellInputTypes",
        nestedMismatch.getMessage());

    IllegalStateException plainMismatch =
        assertThrows(
            IllegalStateException.class,
            () ->
                CatalogFieldMetadataSupport.validatePlainTypeGroupMapping(
                    ExecutionModeInput.class, "wrongGroup"));
    assertEquals(
        "Field-shape plain group mapping mismatch for dev.erst.gridgrind.contract.dto.ExecutionModeInput: expected=wrongGroup, mapped=executionModeInputType",
        plainMismatch.getMessage());

    IllegalStateException ambiguous =
        assertThrows(
            IllegalStateException.class,
            () ->
                CatalogFieldMetadataSupport.lookupAssignableGroup(
                    Map.of(Alpha.class, "alpha", Beta.class, "beta"), Gamma.class));
    assertTrue(
        ambiguous
            .getMessage()
            .startsWith(
                "Ambiguous catalog field group mapping for " + Gamma.class.getName() + ": "));
    assertTrue(ambiguous.getMessage().contains(Alpha.class.getName()));
    assertTrue(ambiguous.getMessage().contains(Beta.class.getName()));
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
    assertEquals(
        "alpha",
        CatalogFieldMetadataSupport.lookupAssignableGroup(
            Map.of(Alpha.class, "alpha"), Gamma.class));
  }

  @Test
  @SuppressWarnings("PMD.UseConcurrentHashMap")
  void lookupAssignableGroupListPrefersMostSpecificAssignableGroup()
      throws ReflectiveOperationException {
    Map<Class<?>, List<String>> groups = new LinkedHashMap<>();
    groups.put(dev.erst.gridgrind.contract.selector.Selector.class, List.of("selector"));
    groups.put(dev.erst.gridgrind.contract.selector.SheetSelector.class, List.of("nested"));

    assertEquals(
        List.of("nested"),
        CatalogFieldMetadataSupport.lookupAssignableGroupList(
            groups, dev.erst.gridgrind.contract.selector.SheetSelector.ByName.class));
    assertEquals(
        null,
        CatalogFieldMetadataSupport.lookupAssignableGroupList(
            groups, dev.erst.gridgrind.contract.selector.Selector.class));
    Map<Class<?>, List<String>> reversedGroups = new LinkedHashMap<>();
    reversedGroups.put(dev.erst.gridgrind.contract.selector.SheetSelector.class, List.of("nested"));
    reversedGroups.put(dev.erst.gridgrind.contract.selector.Selector.class, List.of("selector"));
    assertEquals(
        List.of("nested"),
        CatalogFieldMetadataSupport.lookupAssignableGroupList(
            reversedGroups, dev.erst.gridgrind.contract.selector.SheetSelector.ByName.class));
    assertEquals(
        "nested",
        CatalogFieldMetadataSupport.lookupAssignableGroup(
            orderedGroupMap(
                dev.erst.gridgrind.contract.selector.Selector.class,
                "selector",
                dev.erst.gridgrind.contract.selector.SheetSelector.class,
                "nested"),
            dev.erst.gridgrind.contract.selector.SheetSelector.ByName.class));
    assertEquals(
        "nested",
        CatalogFieldMetadataSupport.lookupAssignableGroup(
            orderedGroupMap(
                dev.erst.gridgrind.contract.selector.SheetSelector.class,
                "nested",
                dev.erst.gridgrind.contract.selector.Selector.class,
                "selector"),
            dev.erst.gridgrind.contract.selector.SheetSelector.ByName.class));
  }

  private record MetadataFixture(
      String name, List<String> notes, Map<String, CellInput> cells, WorkbookStep step) {}

  /** Broader fixture interface used for ambiguous assignable-group testing. */
  private interface Alpha {}

  /** Fixture type that matches both Alpha and Beta. */
  private static final class Gamma implements Alpha, Beta {}

  /** Second broader fixture interface used for ambiguous assignable-group testing. */
  private interface Beta {}

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

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static Map<Class<?>, String> orderedGroupMap(
      Class<?> firstType, String firstGroup, Class<?> secondType, String secondGroup) {
    Map<Class<?>, String> groups = new LinkedHashMap<>();
    groups.put(firstType, firstGroup);
    groups.put(secondType, secondGroup);
    return groups;
  }
}
