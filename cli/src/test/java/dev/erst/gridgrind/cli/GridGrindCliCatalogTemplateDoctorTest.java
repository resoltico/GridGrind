package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.catalog.Catalog;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.catalog.TypeEntry;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import tools.jackson.databind.node.ArrayNode;
import tools.jackson.databind.node.JsonNodeFactory;
import tools.jackson.databind.node.ObjectNode;

/**
 * Ensures every published step template is accepted by the same doctor surface that callers use.
 */
class GridGrindCliCatalogTemplateDoctorTest extends GridGrindCliTestSupport {
  private static final JsonNodeFactory JSON = JsonNodeFactory.instance;

  @TempDir Path workspace;

  @Test
  void everyPublishedStepTemplateDoctorsCleanly() throws Exception {
    Catalog catalog = GridGrindProtocolCatalog.catalog();
    List<TypeEntry> entries =
        Stream.of(
                catalog.mutationActionTypes(),
                catalog.assertionTypes(),
                catalog.inspectionQueryTypes())
            .flatMap(List::stream)
            .toList();

    GridGrindCli cli = new GridGrindCli();
    entries.forEach(entry -> doctorTemplate(cli, entry));
  }

  private void doctorTemplate(GridGrindCli cli, TypeEntry entry) {
    try {
      Path requestPath = workspace.resolve(entry.id().toLowerCase(java.util.Locale.ROOT) + ".json");
      Files.write(requestPath, GridGrindCliJson.writeBytes(requestFor(entry)));
      ByteArrayOutputStream stdout = new ByteArrayOutputStream();
      ByteArrayOutputStream stderr = new ByteArrayOutputStream();
      int exitCode =
          cli.run(
              new String[] {"--doctor-request", "--request", requestPath.toString()},
              InputStream.nullInputStream(),
              stdout,
              stderr);
      RequestDoctorReport report = doctorReport(stdout, stderr);

      assertEquals(0, exitCode, entry.id());
      assertTrue(report.valid(), () -> entry.id() + ": " + report.problems());
    } catch (java.io.IOException exception) {
      throw new AssertionError(entry.id(), exception);
    }
  }

  private static ObjectNode requestFor(TypeEntry entry) {
    ObjectNode request = JSON.objectNode();
    request.put("protocolVersion", "V2");
    request.set("source", type("NEW"));
    request.set("persistence", type("NONE"));
    ArrayNode steps = JSON.arrayNode();
    steps.add(ensureSheet());
    steps.add(entry.stepTemplate().orElseThrow().template());
    request.set("steps", steps);
    return request;
  }

  private static ObjectNode ensureSheet() {
    ObjectNode step = JSON.objectNode();
    step.put("stepId", "template-prerequisite-sheet");
    step.set("target", type("SHEET_BY_NAME").put("name", "Sheet1"));
    step.set("action", type("ENSURE_SHEET"));
    return step;
  }

  private static ObjectNode type(String type) {
    return JSON.objectNode().put("type", type);
  }
}
