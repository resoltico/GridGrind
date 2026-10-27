package dev.erst.gridgrind.cli.command;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.render.GridGrindCliJson;
import dev.erst.gridgrind.core.discovery.RecipeKeywordMatchReport;
import dev.erst.gridgrind.core.discovery.RecipeRegistry;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.json.JsonMapper;

/** Focused contract tests for the CLI recipe-discovery surfaces. */
class GridGrindCliRecipeDiscoveryContractTest extends GridGrindCliTestSupport {
  @Test
  void printRecipeCatalogListUsesOneCompactRowShapeForBothTypes() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"recipe", "list"}, InputStream.nullInputStream(), stdout);

    JsonNode catalog = JsonMapper.builder().build().readTree(stdout.toByteArray());
    JsonNode example = recipeWithView(catalog.path("recipes"), "EXAMPLE");
    JsonNode task = recipeWithView(catalog.path("recipes"), "TASK_STARTER");

    assertEquals(0, exitCode);
    assertEquals(
        List.of("advisory", "id", "requestFileName", "requiredWorkspacePaths", "summary", "type"),
        fieldNames(example));
    assertEquals(fieldNames(example), fieldNames(task));
  }

  @Test
  void printRecipeCatalogLookupReturnsTaskStarterDetailSurface() throws IOException {
    JsonNode output = runRecipeCatalogLookup("DASHBOARD");

    assertEquals("DASHBOARD", output.path("id").asText());
    assertEquals("TASK_STARTER", output.path("type").asText());
    assertTrue(output.has("intentTags"), "task lookup must expose top-level intent tags");
    assertTrue(
        output.has("discoveryProfile"), "task lookup must expose discovery-profile metadata");
    assertTrue(output.has("workflow"), "task lookup must expose staged workflow detail");
    assertTrue(
        output.path("workflow").toString().contains("\"capabilityRefs\""),
        "task workflow detail must expose exact capability refs");
    assertTrue(
        output.has("requestProfile"), "task lookup must expose the exact runnable request profile");
    assertTrue(
        output.path("requestProfile").has("mutationActionTypes"),
        "request profile must expose exact mutation-action ids");
    assertTrue(
        output.path("discoveryProfile").path("discoveryTerms").isArray(),
        "task lookup must expose discovery terms without forcing nested intentTags");
  }

  @Test
  void printRecipeCatalogLookupReturnsExampleSpecificDetailSurface() throws IOException {
    JsonNode output = runRecipeCatalogLookup("BUDGET");

    assertEquals("BUDGET", output.path("id").asText());
    assertEquals("EXAMPLE", output.path("type").asText());
    assertTrue(output.has("intentTags"), "example lookup must expose intent tags");
    assertTrue(
        output.has("requestProfile"),
        "example lookup must expose the exact runnable request profile");
    assertEquals("NEW", output.path("requestProfile").path("sourceType").asText());
    assertTrue(output.path("requestProfile").path("stepCount").asInt() > 0);
    assertTrue(output.path("requestProfile").path("mutationActionTypes").isArray());
  }

  @Test
  void printRecipeKeywordMatchReturnsPublishedIntentTagsWhenNothingMatches() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"recipe", "match", "zzzz no such workflow"},
                InputStream.nullInputStream(),
                stdout);

    RecipeKeywordMatchReport report =
        GridGrindCliJson.readBytes(stdout.toByteArray(), RecipeKeywordMatchReport.class);

    assertEquals(0, exitCode);
    assertEquals(List.of("zzzz"), report.normalizedTerms());
    assertEquals(List.of("zzzz"), report.unmatchedTerms());
    assertEquals(publishedIntentTags(), report.suggestedIntentTags());
    assertEquals(List.of(), report.candidates());
  }

  private static JsonNode runRecipeCatalogLookup(String id) throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"recipe", "show", id, "--details"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    return JsonMapper.builder().build().readTree(stdout.toByteArray());
  }

  private static List<String> publishedIntentTags() {
    return RecipeRegistry.recipes().stream()
        .flatMap(recipe -> recipe.intentTags().stream())
        .distinct()
        .sorted()
        .toList();
  }

  private static JsonNode recipeWithView(JsonNode recipes, String view) {
    for (JsonNode recipe : recipes) {
      if (view.equals(recipe.path("type").asText())) {
        return recipe;
      }
    }
    throw new IllegalArgumentException("missing recipe view " + view);
  }

  private static List<String> fieldNames(JsonNode node) {
    java.util.List<String> names = new java.util.ArrayList<>();
    for (var property : node.properties()) {
      names.add(property.getKey());
    }
    return names.stream().sorted().toList();
  }
}
