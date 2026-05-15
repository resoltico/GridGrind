package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
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
        GridGrindJson.requestTree(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                ExecutionPolicyInput.defaults(),
                FormulaEnvironmentInput.empty(),
                List.of()));
    ArrayNode steps = request.putArray("steps");
    steps.add(template.template());
    ByteArrayOutputStream buffer = new ByteArrayOutputStream();
    GridGrindJson.writeCatalogLookupValue(buffer, request);
    return buffer.toByteArray();
  }
}
