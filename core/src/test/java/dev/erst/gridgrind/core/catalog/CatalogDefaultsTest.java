package dev.erst.gridgrind.core.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.pivot.PivotTableInput;
import dev.erst.gridgrind.core.model.security.OoxmlEncryptionInput;
import dev.erst.gridgrind.core.model.table.TableInput;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Checks published defaults against explicit authoring contracts and immutable observations. */
class CatalogDefaultsTest {
  @Test
  void publishesConstantDefaultsForEnumsSourcesAndCollections() {
    assertEquals(
        Optional.of("AES_256"), field(OoxmlEncryptionInput.class, "cipher").defaultValue());
    assertEquals(Optional.of("SHA_512"), field(OoxmlEncryptionInput.class, "hash").defaultValue());
    assertEquals(
        Optional.of(Map.of("type", "INLINE", "text", "")),
        field(TableInput.class, "comment").defaultValue());
    assertEquals(Optional.of(List.of()), field(TableInput.class, "columns").defaultValue());
    assertEquals(Optional.of(true), field(TableInput.class, "hasAutofilter").defaultValue());
    assertEquals(Optional.of(""), field(TableInput.class, "headerRowCellStyle").defaultValue());
  }

  @Test
  void derivedDefaultsNameTheSourceWithoutInventingAConstant() {
    FieldEntry display = field(PivotTableInput.DataField.class, "displayName");
    assertEquals(FieldRequirement.OPTIONAL, display.requirement());
    assertEquals(Optional.of("sourceColumnName"), display.defaultFrom());
    assertEquals(Optional.empty(), display.defaultValue());
  }

  @Test
  @SuppressWarnings("PMD.UseConcurrentHashMap") // Mutable fixture verifies defensive copying.
  void publishedObjectDefaultsDoNotRetainMutableInputs() {
    List<String> labels = new ArrayList<>(List.of("original"));
    Map<String, Object> input = new LinkedHashMap<>();
    input.put("labels", labels);
    FieldEntry entry =
        new FieldEntry(
            "settings",
            FieldRequirement.OPTIONAL,
            new FieldShape.PlainTypeGroupRef("settings"),
            List.of(),
            Optional.of(input),
            Optional.empty(),
            List.of(),
            List.of(),
            false);
    labels.clear();
    input.clear();
    assertEquals(Optional.of(Map.of("labels", List.of("original"))), entry.defaultValue());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new FieldEntry(
                "settings",
                FieldRequirement.OPTIONAL,
                new FieldShape.PlainTypeGroupRef("settings"),
                List.of(),
                Optional.of(Map.of()),
                Optional.of("source"),
                List.of(),
                List.of(),
                false));
  }

  private static FieldEntry field(Class<? extends Record> type, String name) {
    var component =
        Arrays.stream(type.getRecordComponents())
            .filter(candidate -> candidate.getName().equals(name))
            .findFirst()
            .orElseThrow();
    return CatalogFieldMetadataSupport.fieldEntry(
        component, java.util.Set.copyOf(GridGrindProtocolContractSupport.optionalFieldNames(type)));
  }
}
