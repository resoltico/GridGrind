package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.cli.discovery.RecipeKeywordMatchReport;
import dev.erst.gridgrind.cli.examples.GridGrindCliRecipeRegistry;
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
  void printRecipeCatalogListUsesOneCompactRowShapeForBothViews() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"--print-recipe-catalog"}, InputStream.nullInputStream(), stdout);

    JsonNode catalog = JsonMapper.builder().build().readTree(stdout.toByteArray());
    JsonNode example = recipeWithView(catalog.path("recipes"), "EXAMPLE");
    JsonNode task = recipeWithView(catalog.path("recipes"), "TASK_STARTER");

    assertEquals(0, exitCode);
    assertEquals(
        List.of(
            "id", "requestFileName", "requiredWorkspacePaths", "summary", "view", "workspaceMode"),
        fieldNames(example));
    assertEquals(fieldNames(example), fieldNames(task));
  }

  @Test
  void printRecipeCatalogLookupReturnsTaskStarterDetailSurface() throws IOException {
    JsonNode output = runRecipeCatalogLookup("DASHBOARD");

    assertEquals("DASHBOARD", output.path("id").asText());
    assertEquals("TASK_STARTER", output.path("view").asText());
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
    assertEquals("EXAMPLE", output.path("view").asText());
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
                new String[] {"--print-recipe-keyword-match", "--query", "zzzz no such workflow"},
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
                new String[] {"--print-recipe-catalog", "--lookup", id},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    return JsonMapper.builder().build().readTree(stdout.toByteArray());
  }

  private static List<String> publishedIntentTags() {
    return GridGrindCliRecipeRegistry.recipes().stream()
        .flatMap(recipe -> recipe.intentTags().stream())
        .distinct()
        .sorted()
        .toList();
  }

  private static JsonNode recipeWithView(JsonNode recipes, String view) {
    for (JsonNode recipe : recipes) {
      if (view.equals(recipe.path("view").asText())) {
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
