package dev.erst.gridgrind.cli.examples;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.RecipeAdvisory;
import dev.erst.gridgrind.cli.discovery.TaskEntry;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import org.junit.jupiter.api.Test;

/** Coverage for the helper guards that feed the canonical CLI recipe registry. */
class GridGrindRecipeSupportCoverageTest {
  @Test
  void cliRecipeRejectsInvalidIdsAndIntentTags() {
    WorkbookPlan plan = templatePlan();

    assertEquals(
        "id must use upper-case discovery tokens",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindCliRecipe(
                        "workbook_health",
                        dev.erst.gridgrind.cli.discovery.RecipeView.EXAMPLE,
                        "recipe.json",
                        "summary",
                        RecipeAdvisory.SELF_CONTAINED,
                        List.of(),
                        List.of("tag"),
                        plan))
            .getMessage());
    assertEquals(
        "requestFileName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindCliRecipe(
                        "WORKBOOK_HEALTH",
                        dev.erst.gridgrind.cli.discovery.RecipeView.EXAMPLE,
                        " ",
                        "summary",
                        RecipeAdvisory.SELF_CONTAINED,
                        List.of(),
                        List.of("tag"),
                        plan))
            .getMessage());
    assertEquals(
        "intentTags must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindCliRecipe(
                        "WORKBOOK_HEALTH",
                        dev.erst.gridgrind.cli.discovery.RecipeView.EXAMPLE,
                        "recipe.json",
                        "summary",
                        RecipeAdvisory.SELF_CONTAINED,
                        List.of(),
                        List.of(),
                        plan))
            .getMessage());
  }

  @Test
  void recipeDefinitionsRejectInvalidMetadataAndExposeCanonicalViews() {
    WorkbookPlan plan = templatePlan();

    assertEquals(
        "id must use upper-case discovery tokens",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindExampleRecipeDefinition(
                        "budget",
                        "budget.json",
                        "summary",
                        RecipeAdvisory.SELF_CONTAINED,
                        List.of(),
                        List.of("tag"),
                        plan,
                        plan))
            .getMessage());
    assertEquals(
        "requestFileName must be one portable file name, not a repository path",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindExampleRecipeDefinition(
                        "BUDGET",
                        "examples/budget.json",
                        "summary",
                        RecipeAdvisory.SELF_CONTAINED,
                        List.of(),
                        List.of("tag"),
                        plan,
                        plan))
            .getMessage());
    assertEquals(
        "requestFileName must be one portable file name, not a repository path",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindExampleRecipeDefinition(
                        "BUDGET",
                        "examples\\budget.json",
                        "summary",
                        RecipeAdvisory.SELF_CONTAINED,
                        List.of(),
                        List.of("tag"),
                        plan,
                        plan))
            .getMessage());
    assertEquals(
        "requestFileName must end with .json",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindExampleRecipeDefinition(
                        "BUDGET",
                        "budget.txt",
                        "summary",
                        RecipeAdvisory.SELF_CONTAINED,
                        List.of(),
                        List.of("tag"),
                        plan,
                        plan))
            .getMessage());
    assertEquals(
        "SELF_CONTAINED example recipes must not publish requiredWorkspacePaths",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindExampleRecipeDefinition(
                        "BUDGET",
                        "budget.json",
                        "summary",
                        RecipeAdvisory.SELF_CONTAINED,
                        List.of("asset.xlsx"),
                        List.of("tag"),
                        plan,
                        plan))
            .getMessage());
    assertEquals(
        "REQUIRES_EXAMPLE_ASSETS example recipes must publish requiredWorkspacePaths",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindExampleRecipeDefinition(
                        "CUSTOM_XML",
                        "custom-xml.json",
                        "summary",
                        RecipeAdvisory.REQUIRES_EXAMPLE_ASSETS,
                        List.of(),
                        List.of("tag"),
                        plan,
                        plan))
            .getMessage());
    assertEquals(
        "repositoryPlan must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new GridGrindExampleRecipeDefinition(
                        "BUDGET",
                        "budget.json",
                        "summary",
                        RecipeAdvisory.SELF_CONTAINED,
                        List.of(),
                        List.of("tag"),
                        plan,
                        null))
            .getMessage());
    assertEquals(
        "intentTags must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindExampleRecipeDefinition(
                        "BUDGET",
                        "budget.json",
                        "summary",
                        RecipeAdvisory.SELF_CONTAINED,
                        List.of(),
                        List.of(" "),
                        plan,
                        plan))
            .getMessage());
    assertEquals(
        "intentTags must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindExampleRecipeDefinition(
                        "BUDGET",
                        "budget.json",
                        "summary",
                        RecipeAdvisory.SELF_CONTAINED,
                        List.of(),
                        List.of(),
                        plan,
                        plan))
            .getMessage());
    List<GridGrindRecipeDefinition> definitions = GridGrindRecipeDefinitions.definitions();
    assertEquals(
        definitions.stream()
            .map(GridGrindRecipeSupportCoverageTest::definitionId)
            .distinct()
            .count(),
        definitions.size());

    GridGrindExampleRecipeDefinition definition = exampleDefinition(definitions, "BUDGET");
    GridGrindExampleRecipeDefinition customXmlDefinition =
        exampleDefinition(definitions, "CUSTOM_XML");
    GridGrindExampleRecipeDefinition packageSecurityDefinition =
        exampleDefinition(definitions, "PACKAGE_SECURITY_INSPECTION");
    GridGrindTaskRecipeDefinition dashboardDefinition = taskDefinition(definitions, "DASHBOARD");
    GridGrindTaskRecipeDefinition maintenanceDefinition =
        taskDefinition(definitions, "WORKBOOK_MAINTENANCE");
    assertEquals("BUDGET", definition.id());
    assertEquals(
        List.of("budget", "authoring", "table", "formula", "inspection"), definition.intentTags());
    assertEquals(RecipeAdvisory.SELF_CONTAINED, definition.advisory());
    assertEquals(List.of(), definition.requiredWorkspacePaths());
    assertEquals(RecipeAdvisory.REQUIRES_EXAMPLE_ASSETS, customXmlDefinition.advisory());
    assertEquals(
        List.of(
            "custom-xml-assets/custom-xml-mapping.xlsx", "custom-xml-assets/custom-xml-update.xml"),
        customXmlDefinition.requiredWorkspacePaths());
    assertEquals(
        List.of("package-security-assets/gridgrind-package-security.xlsx"),
        packageSecurityDefinition.requiredWorkspacePaths());
    assertEquals(dev.erst.gridgrind.cli.discovery.RecipeView.EXAMPLE, definition.recipe().view());
    assertEquals(
        dev.erst.gridgrind.cli.discovery.RecipeView.TASK_STARTER,
        dashboardDefinition.recipe().view());
    assertEquals("dashboard-request.json", dashboardDefinition.recipe().requestFileName());
    assertEquals("WORKBOOK_MAINTENANCE", maintenanceDefinition.task().id());
    assertEquals(
        definitions.stream()
            .filter(GridGrindExampleRecipeDefinition.class::isInstance)
            .map(GridGrindExampleRecipeDefinition.class::cast)
            .map(GridGrindExampleRecipeDefinition::id)
            .toList(),
        GridGrindShippedExamples.examples().stream()
            .map(GridGrindShippedExamples.ShippedExample::id)
            .toList());
    assertEquals(
        definitions.stream()
            .filter(GridGrindExampleRecipeDefinition.class::isInstance)
            .map(GridGrindExampleRecipeDefinition.class::cast)
            .map(GridGrindExampleRecipeDefinition::id)
            .toList(),
        GridGrindShippedExamples.repositoryExamples().stream()
            .map(GridGrindShippedExamples.ShippedExample::id)
            .toList());
    assertEquals(
        definitions.stream()
            .filter(GridGrindTaskRecipeDefinition.class::isInstance)
            .map(GridGrindTaskRecipeDefinition.class::cast)
            .map(definitionEntry -> definitionEntry.task().id())
            .toList(),
        GridGrindCliRecipeRegistry.taskRecipes().stream().map(GridGrindCliRecipe::id).toList());
  }

  private static GridGrindExampleRecipeDefinition exampleDefinition(
      List<GridGrindRecipeDefinition> definitions, String id) {
    return definitions.stream()
        .filter(GridGrindExampleRecipeDefinition.class::isInstance)
        .map(GridGrindExampleRecipeDefinition.class::cast)
        .filter(definition -> definition.id().equals(id))
        .findFirst()
        .orElseThrow(() -> new AssertionError("Missing example definition " + id));
  }

  private static GridGrindTaskRecipeDefinition taskDefinition(
      List<GridGrindRecipeDefinition> definitions, String id) {
    return definitions.stream()
        .filter(GridGrindTaskRecipeDefinition.class::isInstance)
        .map(GridGrindTaskRecipeDefinition.class::cast)
        .filter(definition -> definition.task().id().equals(id))
        .findFirst()
        .orElseThrow(() -> new AssertionError("Missing task definition " + id));
  }

  private static String definitionId(GridGrindRecipeDefinition definition) {
    if (definition instanceof GridGrindExampleRecipeDefinition example) {
      return example.id();
    }
    if (definition instanceof GridGrindTaskRecipeDefinition task) {
      return task.task().id();
    }
    throw new AssertionError("Unknown recipe definition type " + definition.getClass());
  }

  @Test
  void taskRecipeBuildersAndRegistryRejectDuplicateIdsAndExposeCanonicalEntries() {
    List<GridGrindRecipeDefinition> definitions = GridGrindRecipeDefinitions.definitions();
    GridGrindTaskRecipeDefinition dashboardDefinition = taskDefinition(definitions, "DASHBOARD");
    GridGrindTaskRecipeDefinition auditDefinition =
        taskDefinition(definitions, "AUDIT_EXISTING_WORKBOOK");
    GridGrindTaskRecipeDefinition maintenanceDefinition =
        taskDefinition(definitions, "WORKBOOK_MAINTENANCE");

    assertEquals("DASHBOARD", dashboardDefinition.task().id());
    assertEquals("dashboard-request.json", dashboardDefinition.recipe().requestFileName());
    assertEquals("DASHBOARD", dashboardDefinition.task().id());
    assertEquals("AUDIT_EXISTING_WORKBOOK", auditDefinition.task().id());
    assertEquals("WORKBOOK_MAINTENANCE", maintenanceDefinition.task().id());
    assertTrue(GridGrindCliRecipeRegistry.taskEntryFor("DASHBOARD").isPresent());
    assertTrue(
        GridGrindCliRecipeRegistry.taskRecipes().stream()
            .anyMatch(recipe -> "DASHBOARD".equals(recipe.id())));
    assertTrue(
        GridGrindCliRecipeRegistry.recipes().stream()
            .anyMatch(recipe -> "BUDGET".equals(recipe.id())));
    assertEquals(
        GridGrindCliRecipeRegistry.recipeFor("DASHBOARD").orElseThrow().intentTags(),
        dashboardDefinition.task().intentTags());
    assertEquals(
        GridGrindShippedExamples.repositoryExamples().stream()
            .filter(example -> "BUDGET".equals(example.id()))
            .findFirst()
            .orElseThrow()
            .plan(),
        GridGrindCliRecipeRegistry.repositoryExamplePlanFor("BUDGET"));
    assertEquals("DASHBOARD", GridGrindCliRecipeRegistry.requireTaskEntry("DASHBOARD").id());

    assertEquals(
        "id must not be null",
        assertThrows(NullPointerException.class, () -> GridGrindCliRecipeRegistry.recipeFor(null))
            .getMessage());
    assertEquals(
        "id must not be null",
        assertThrows(
                NullPointerException.class, () -> GridGrindCliRecipeRegistry.taskEntryFor(null))
            .getMessage());
    assertEquals(
        "Missing repository example plan for NO_SUCH_EXAMPLE",
        assertThrows(
                IllegalStateException.class,
                () -> GridGrindCliRecipeRegistry.repositoryExamplePlanFor("NO_SUCH_EXAMPLE"))
            .getMessage());
    assertEquals(
        "Missing task entry for recipe NO_SUCH_TASK",
        assertThrows(
                IllegalStateException.class,
                () -> GridGrindCliRecipeRegistry.requireTaskEntry("NO_SUCH_TASK"))
            .getMessage());
    assertEquals(
        Map.of("ONE", recipe("ONE")),
        GridGrindCliRecipeRegistry.recipesById(List.of(recipe("ONE"))));
    assertEquals(
        "Duplicate CLI recipe id DUPLICATE",
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindCliRecipeRegistry.recipesById(
                        List.of(recipe("DUPLICATE"), recipe("DUPLICATE"))))
            .getMessage());

    Map<String, TaskEntry> taskEntries = new ConcurrentHashMap<>();
    GridGrindCliRecipeRegistry.putUniqueAssociated(
        taskEntries, "DASHBOARD", dashboardDefinition.task(), "task entry");
    assertEquals("DASHBOARD", taskEntries.get("DASHBOARD").id());
    assertEquals(
        "Duplicate task entry for DASHBOARD",
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindCliRecipeRegistry.putUniqueAssociated(
                        taskEntries, "DASHBOARD", dashboardDefinition.task(), "task entry"))
            .getMessage());
  }

  @Test
  void taskStarterRecipeSupportRejectsBlankFileNamesAndTaskIds() {
    assertEquals(
        "fileName must not be null",
        assertThrows(
                NullPointerException.class, () -> TaskStarterRecipeSupport.taskStarterAsset(null))
            .getMessage());
    assertEquals(
        "fileName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> TaskStarterRecipeSupport.taskStarterAsset(" "))
            .getMessage());
    assertEquals(
        "taskId must not be null",
        assertThrows(
                NullPointerException.class,
                () -> TaskStarterRecipeSupport.taskRequestFileName(null))
            .getMessage());
    assertEquals(
        "taskId must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> TaskStarterRecipeSupport.taskRequestFileName(" "))
            .getMessage());
  }

  private static WorkbookPlan templatePlan() {
    return GridGrindProtocolCatalog.requestTemplate();
  }

  private static GridGrindCliRecipe recipe(String id) {
    return new GridGrindCliRecipe(
        id,
        dev.erst.gridgrind.cli.discovery.RecipeView.EXAMPLE,
        id.toLowerCase(Locale.ROOT) + ".json",
        "summary",
        RecipeAdvisory.SELF_CONTAINED,
        List.of(),
        List.of("tag"),
        templatePlan());
  }
}
