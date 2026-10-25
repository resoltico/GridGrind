package dev.erst.gridgrind.core.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentInput;
import dev.erst.gridgrind.core.plan.ExecutionPolicyInput;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.json.codec.GridGrindJson;
import dev.erst.gridgrind.json.codec.GridGrindJsonOutput;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.node.ArrayNode;
import tools.jackson.databind.node.ObjectNode;

/** Contract test for published executable step templates in the protocol catalog. */
class GridGrindProtocolCatalogStepTemplateContractTest {
  @Test
  void everyLiveCatalogStepTypeHasAnExecutableTemplate() {
    Catalog catalog = GridGrindProtocolCatalog.catalog();
    List<String> missing =
        Stream.of(
                catalog.mutationActionTypes().stream(),
                catalog.assertionTypes().stream(),
                catalog.inspectionQueryTypes().stream())
            .flatMap(stream -> stream)
            .filter(entry -> entry.stepTemplate().isEmpty())
            .map(TypeEntry::id)
            .toList();

    assertTrue(
        missing.isEmpty(),
        () ->
            "Every step type in the live catalog must carry an executable template; missing: "
                + missing);
  }

  @Test
  void publishedStepTemplatesRoundTripAsRequests() {
    Catalog catalog = GridGrindProtocolCatalog.catalog();
    List<String> failures =
        Stream.of(
                catalog.mutationActionTypes().stream(),
                catalog.assertionTypes().stream(),
                catalog.inspectionQueryTypes().stream())
            .flatMap(stream -> stream)
            .map(GridGrindProtocolCatalogStepTemplateContractTest::roundTripFailure)
            .flatMap(Optional::stream)
            .toList();

    assertTrue(
        failures.isEmpty(),
        () ->
            "Published step templates must decode as request steps:"
                + System.lineSeparator()
                + String.join(System.lineSeparator(), failures));
  }

  @Test
  void setRangeAndAppendRowTemplatesPreferTypedCellWrappers() {
    Catalog catalog = GridGrindProtocolCatalog.catalog();
    TypeEntry setRange = typeEntry(catalog.mutationActionTypes(), "SET_RANGE");
    TypeEntry appendRow = typeEntry(catalog.mutationActionTypes(), "APPEND_ROW");

    assertEquals(
        "TYPED",
        tree(setRange.stepTemplate().orElseThrow().template())
            .path("action")
            .path("rows")
            .path("type")
            .asText());
    assertTrue(
        tree(setRange.stepTemplate().orElseThrow().template())
            .path("action")
            .path("rows")
            .has("cells"));
    assertEquals(
        "TYPED",
        tree(appendRow.stepTemplate().orElseThrow().template())
            .path("action")
            .path("values")
            .path("type")
            .asText());
    assertTrue(
        tree(appendRow.stepTemplate().orElseThrow().template())
            .path("action")
            .path("values")
            .has("cells"));
  }

  private static Optional<String> roundTripFailure(TypeEntry entry) {
    try {
      GridGrindJson.readRequest(encodeRequest(entry.stepTemplate().orElseThrow()));
      return Optional.empty();
    } catch (RuntimeException | IOException exception) {
      return Optional.of(entry.id() + " -> " + exception.getMessage());
    }
  }

  private static byte[] encodeRequest(ProtocolStepTemplate template) throws IOException {
    ObjectNode request =
        GridGrindJsonOutput.requestTree(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                ExecutionPolicyInput.defaults(),
                FormulaEnvironmentInput.empty(),
                List.of()));
    ArrayNode steps = request.putArray("steps");
    steps.add(tree(template.template()));
    ByteArrayOutputStream buffer = new ByteArrayOutputStream();
    GridGrindJsonOutput.writeCatalogLookupValue(buffer, request);
    return buffer.toByteArray();
  }

  private static TypeEntry typeEntry(List<TypeEntry> entries, String id) {
    return entries.stream()
        .filter(entry -> entry.id().equals(id))
        .findFirst()
        .orElseThrow(() -> new IllegalStateException("Catalog is missing type " + id));
  }

  private static ObjectNode tree(Object value) {
    return tools.jackson.databind.json.JsonMapper.builder().build().valueToTree(value);
  }
}
