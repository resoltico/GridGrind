package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.TreeSet;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.node.ArrayNode;
import tools.jackson.databind.node.ObjectNode;

/** Creates explicit fixture authority and records persisted facts using product inspections. */
final class GoldenWorkbook {
  private GoldenWorkbook() {}

  static ObjectNode execute(WorkbookLauncher launcher, Path workspace, Path request)
      throws IOException, InterruptedException {
    JsonNode plan = GoldenSnapshots.JSON.readTree(Files.readAllBytes(request));
    Path grant = grant(plan, workspace, request.getFileName() + "-grant.json");
    Path temporary = Files.createDirectories(workspace.resolve("temporary"));
    List<String> arguments =
        new ArrayList<>(
            List.of(
                "run",
                "--request",
                request.toString(),
                "--grant",
                grant.toString(),
                "--temp-root",
                temporary.toString()));
    Path provider = workspace.resolve("fixture-secrets");
    if (Files.isDirectory(provider)) {
      arguments.addAll(List.of("--secrets-provider", provider.toString()));
    }
    GoldenLauncher.Result result = launcher.run(workspace, arguments.toArray(String[]::new));
    assertEquals(0, result.exit(), () -> result.stdout() + result.stderr());
    ObjectNode observation = GoldenSnapshots.observation(result);
    JsonNode persistence = plan.path("persistence");
    String type = persistence.path("type").asString();
    if (!"NONE".equals(type)) {
      String path =
          "SAVE_AS".equals(type)
              ? persistence.path("path").asString()
              : plan.path("source").path("path").asString();
      Path output = workspace.resolve(path).normalize();
      assertTrue(Files.isRegularFile(output), () -> "Missing output workbook: " + output);
      observation.set("readback", readback(launcher, workspace, output));
    }
    return observation;
  }

  static Path grant(JsonNode plan, Path workspace, String fileName) throws IOException {
    Set<String> resources = new TreeSet<>();
    Set<String> operations = new TreeSet<>();
    Set<String> references = new TreeSet<>();
    collect(plan, "", workspace, resources, operations, references);
    ObjectNode grant = GoldenSnapshots.JSON.createObjectNode();
    ArrayNode readable = grant.putArray("readableResources");
    resources.forEach(path -> readable.addObject().put("type", "FILE").put("path", path));
    ArrayNode ids = grant.putArray("operationIds");
    operations.forEach(ids::add);
    grant.putObject("targetAuthority").put("type", "WORKBOOK_WIDE");
    ObjectNode publication = grant.putObject("publicationAuthority");
    JsonNode persistence = plan.path("persistence");
    switch (persistence.path("type").asString()) {
      case "NONE" -> publication.put("type", "NONE");
      case "OVERWRITE" -> publication.put("type", "OVERWRITE_SOURCE");
      case "SAVE_AS" -> {
        Path output = workspace.resolve(persistence.path("path").asString()).normalize();
        Files.createDirectories(output.getParent());
        publication.put("type", "SAVE_AS").put("path", output.toString());
        publication.set("ifExists", persistence.path("ifExists"));
      }
      default -> throw new AssertionError("Unknown fixture persistence: " + persistence);
    }
    ArrayNode secrets = grant.putArray("allowedSecretReferences");
    references.forEach(id -> secrets.addObject().put("id", id));
    grant.putObject("acceptancePolicy").put("type", "MINIMUM_ONLY");
    Path path = workspace.resolve(fileName);
    Files.writeString(path, GoldenSnapshots.JSON.writeValueAsString(grant));
    return path;
  }

  private static void collect(
      JsonNode node,
      String parent,
      Path workspace,
      Set<String> resources,
      Set<String> operations,
      Set<String> references) {
    if (node.isObject()) {
      String type = node.path("type").asString("");
      if (Set.of("action", "assertion", "query").contains(parent) && !type.isEmpty()) {
        operations.add(type);
      }
      if (Set.of("EXISTING", "UTF8_FILE").contains(type)
          || ("FILE".equals(type) && Set.of("source", "payload").contains(parent))) {
        resources.add(workspace.resolve(node.path("path").asString()).normalize().toString());
      }
      if (parent.endsWith("Ref") && node.has("id")) {
        references.add(node.path("id").asString());
      }
      node.properties()
          .forEach(
              entry ->
                  collect(
                      entry.getValue(),
                      entry.getKey(),
                      workspace,
                      resources,
                      operations,
                      references));
    } else if (node.isArray()) {
      node.forEach(value -> collect(value, parent, workspace, resources, operations, references));
    } else if (parent.endsWith("Ref") && node.isString()) {
      references.add(node.asString());
    }
  }

  private static ObjectNode readback(WorkbookLauncher launcher, Path workspace, Path output)
      throws IOException, InterruptedException {
    ObjectNode plan = readPlan(output);
    ArrayNode steps = (ArrayNode) plan.path("steps");
    inspect(steps, "summary", object("type", "WORKBOOK_CURRENT"), "GET_WORKBOOK_SUMMARY");
    JsonNode summary = runRead(launcher, workspace, plan, "summary");
    JsonNode workbook = summary.path("inspections").path(0).path("workbook");
    steps.removeAll();
    int sheetIndex = 0;
    for (JsonNode sheet : workbook.path("sheetNames")) {
      String id = "sheet-" + sheetIndex;
      sheetIndex++;
      String name = sheet.asString();
      ObjectNode target = object("type", "SHEET_BY_NAME").put("name", name);
      inspect(steps, id, target, "GET_SHEET_SUMMARY");
    }
    JsonNode bounds = runRead(launcher, workspace, plan, "bounds");
    steps.removeAll();
    for (JsonNode inspection : bounds.path("inspections")) {
      addSheetFacts(steps, inspection.path("sheet"));
    }
    inspect(steps, "tables", object("type", "TABLE_ALL"), "GET_TABLES");
    inspect(steps, "pivots", object("type", "PIVOT_TABLE_ALL"), "GET_PIVOT_TABLES");
    inspect(steps, "findings", object("type", "WORKBOOK_CURRENT"), "ANALYZE_WORKBOOK_FINDINGS");
    ObjectNode facts = GoldenSnapshots.JSON.createObjectNode();
    facts.set("summary", summary.path("inspections"));
    facts.set("bounds", bounds.path("inspections"));
    facts.set("facts", runRead(launcher, workspace, plan, "facts").path("inspections"));
    return facts;
  }

  private static void addSheetFacts(ArrayNode steps, JsonNode sheet) {
    String name = sheet.path("sheetName").asString();
    String prefix = "sheet-" + steps.size();
    ObjectNode target = object("type", "SHEET_BY_NAME").put("name", name);
    inspect(steps, prefix + "-layout", target, "GET_SHEET_LAYOUT");
    inspect(steps, prefix + "-print", target, "GET_PRINT_LAYOUT");
    inspect(
        steps,
        prefix + "-charts",
        object("type", "CHART_ALL_ON_SHEET").put("sheetName", name),
        "GET_CHARTS");
    inspect(
        steps,
        prefix + "-drawing",
        object("type", "DRAWING_OBJECT_ALL_ON_SHEET").put("sheetName", name),
        "GET_DRAWING_OBJECTS");
    int rows = sheet.path("lastRowIndex").asInt() + 1;
    int columns = sheet.path("lastColumnIndex").asInt() + 1;
    assertTrue((long) rows * columns <= 250_000, "Fixture readback exceeds the product cell limit");
    if (rows <= 0 || columns <= 0) {
      return;
    }
    ObjectNode cells = object("type", "CELL_BY_ADDRESSES").put("sheetName", name);
    ArrayNode addresses = cells.putArray("addresses");
    for (int row = 1; row <= rows; row++) {
      for (int column = 1; column <= columns; column++) {
        addresses.add(columnName(column) + row);
      }
    }
    ObjectNode query = inspect(steps, prefix + "-cells", cells, "GET_CELLS");
    ArrayNode facets = query.putObject("projection").putArray("facets");
    List.of(
            "VALUE",
            "STYLE",
            "FORMAT",
            "HYPERLINK",
            "COMMENT",
            "FORMULA",
            "RICH_TEXT_RUNS",
            "TEMPORAL")
        .forEach(facets::add);
    for (String operation : List.of("GET_DATA_VALIDATIONS", "GET_CONDITIONAL_FORMATTING")) {
      inspect(
          steps,
          prefix + "-" + operation,
          object("type", "RANGE_ALL_ON_SHEET").put("sheetName", name),
          operation);
    }
  }

  private static String columnName(int index) {
    StringBuilder name = new StringBuilder();
    for (int value = index; value > 0; value = (value - 1) / 26) {
      name.append((char) ('A' + (value - 1) % 26));
    }
    return name.reverse().toString();
  }

  private static ObjectNode readPlan(Path output) {
    ObjectNode plan = GoldenSnapshots.JSON.createObjectNode();
    plan.put("protocolVersion", 1);
    plan.putObject("source").put("type", "EXISTING").put("path", output.toString());
    plan.putObject("persistence").put("type", "NONE");
    plan.putArray("steps");
    return plan;
  }

  private static JsonNode runRead(
      WorkbookLauncher launcher, Path workspace, ObjectNode plan, String name)
      throws IOException, InterruptedException {
    Path request = workspace.resolve("readback-" + name + ".json");
    Files.writeString(request, GoldenSnapshots.JSON.writeValueAsString(plan));
    Path grant = grant(plan, workspace, "readback-grant.json");
    GoldenLauncher.Result result =
        launcher.run(
            workspace, "run", "--request", request.toString(), "--grant", grant.toString());
    assertEquals(
        0, result.exit(), () -> "Workbook readback failed: " + result.stdout() + result.stderr());
    return GoldenSnapshots.JSON.readTree(result.stdout());
  }

  private static ObjectNode inspect(
      ArrayNode steps, String id, ObjectNode target, String operation) {
    ObjectNode step = steps.addObject().put("stepId", id);
    step.set("target", target);
    return step.putObject("query").put("type", operation);
  }

  private static ObjectNode object(String field, String value) {
    return GoldenSnapshots.JSON.createObjectNode().put(field, value);
  }
}
