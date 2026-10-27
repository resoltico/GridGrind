package dev.erst.gridgrind.core.discovery;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import java.io.ByteArrayInputStream;
import java.util.Base64;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import javax.imageio.ImageIO;
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
                    new Recipe(
                        "workbook_health",
                        dev.erst.gridgrind.core.discovery.RecipeType.EXAMPLE,
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
                    new Recipe(
                        "WORKBOOK_HEALTH",
                        dev.erst.gridgrind.core.discovery.RecipeType.EXAMPLE,
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
                    new Recipe(
                        "WORKBOOK_HEALTH",
                        dev.erst.gridgrind.core.discovery.RecipeType.EXAMPLE,
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
                    new ExampleRecipeTemplate(
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
                    new ExampleRecipeTemplate(
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
                    new ExampleRecipeTemplate(
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
                    new ExampleRecipeTemplate(
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
                    new ExampleRecipeTemplate(
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
                    new ExampleRecipeTemplate(
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
                    new ExampleRecipeTemplate(
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
                    new ExampleRecipeTemplate(
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
                    new ExampleRecipeTemplate(
                        "BUDGET",
                        "budget.json",
                        "summary",
                        RecipeAdvisory.SELF_CONTAINED,
                        List.of(),
                        List.of(),
                        plan,
                        plan))
            .getMessage());
    List<RecipeTemplate> definitions = RecipeTemplates.definitions();
    assertEquals(
        definitions.stream()
            .map(GridGrindRecipeSupportCoverageTest::definitionId)
            .distinct()
            .count(),
        definitions.size());

    ExampleRecipeTemplate definition = exampleDefinition(definitions, "BUDGET");
    ExampleRecipeTemplate customXmlDefinition = exampleDefinition(definitions, "CUSTOM_XML");
    ExampleRecipeTemplate packageSecurityDefinition =
        exampleDefinition(definitions, "PACKAGE_SECURITY_INSPECTION");
    TaskRecipeTemplate dashboardDefinition = taskDefinition(definitions, "DASHBOARD");
    TaskRecipeTemplate maintenanceDefinition = taskDefinition(definitions, "WORKBOOK_MAINTENANCE");
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
    assertEquals(dev.erst.gridgrind.core.discovery.RecipeType.EXAMPLE, definition.recipe().type());
    assertEquals(
        dev.erst.gridgrind.core.discovery.RecipeType.TASK_STARTER,
        dashboardDefinition.recipe().type());
    assertEquals("dashboard-request.json", dashboardDefinition.recipe().requestFileName());
    assertEquals("WORKBOOK_MAINTENANCE", maintenanceDefinition.task().id());
    assertEquals(
        definitions.stream()
            .filter(ExampleRecipeTemplate.class::isInstance)
            .map(ExampleRecipeTemplate.class::cast)
            .map(ExampleRecipeTemplate::id)
            .toList(),
        ShippedExamples.examples().stream().map(ShippedExamples.ShippedExample::id).toList());
    assertEquals(
        definitions.stream()
            .filter(ExampleRecipeTemplate.class::isInstance)
            .map(ExampleRecipeTemplate.class::cast)
            .map(ExampleRecipeTemplate::id)
            .toList(),
        ShippedExamples.repositoryExamples().stream()
            .map(ShippedExamples.ShippedExample::id)
            .toList());
    assertEquals(
        definitions.stream()
            .filter(TaskRecipeTemplate.class::isInstance)
            .map(TaskRecipeTemplate.class::cast)
            .map(definitionEntry -> definitionEntry.task().id())
            .toList(),
        RecipeRegistry.taskRecipes().stream().map(Recipe::id).toList());
  }

  private static ExampleRecipeTemplate exampleDefinition(
      List<RecipeTemplate> definitions, String id) {
    return definitions.stream()
        .filter(ExampleRecipeTemplate.class::isInstance)
        .map(ExampleRecipeTemplate.class::cast)
        .filter(definition -> definition.id().equals(id))
        .findFirst()
        .orElseThrow(() -> new AssertionError("Missing example definition " + id));
  }

  private static TaskRecipeTemplate taskDefinition(List<RecipeTemplate> definitions, String id) {
    return definitions.stream()
        .filter(TaskRecipeTemplate.class::isInstance)
        .map(TaskRecipeTemplate.class::cast)
        .filter(definition -> definition.task().id().equals(id))
        .findFirst()
        .orElseThrow(() -> new AssertionError("Missing task definition " + id));
  }

  private static String definitionId(RecipeTemplate definition) {
    if (definition instanceof ExampleRecipeTemplate example) {
      return example.id();
    }
    if (definition instanceof TaskRecipeTemplate task) {
      return task.task().id();
    }
    throw new AssertionError("Unknown recipe definition type " + definition.getClass());
  }

  @Test
  void taskRecipeBuildersAndRegistryRejectDuplicateIdsAndExposeCanonicalEntries() {
    List<RecipeTemplate> definitions = RecipeTemplates.definitions();
    TaskRecipeTemplate dashboardDefinition = taskDefinition(definitions, "DASHBOARD");
    TaskRecipeTemplate auditDefinition = taskDefinition(definitions, "AUDIT_EXISTING_WORKBOOK");
    TaskRecipeTemplate maintenanceDefinition = taskDefinition(definitions, "WORKBOOK_MAINTENANCE");

    assertEquals("DASHBOARD", dashboardDefinition.task().id());
    assertEquals("dashboard-request.json", dashboardDefinition.recipe().requestFileName());
    assertEquals("DASHBOARD", dashboardDefinition.task().id());
    assertEquals("AUDIT_EXISTING_WORKBOOK", auditDefinition.task().id());
    assertEquals("WORKBOOK_MAINTENANCE", maintenanceDefinition.task().id());
    assertTrue(RecipeRegistry.taskEntryFor("DASHBOARD").isPresent());
    assertTrue(
        RecipeRegistry.taskRecipes().stream().anyMatch(recipe -> "DASHBOARD".equals(recipe.id())));
    assertTrue(RecipeRegistry.recipes().stream().anyMatch(recipe -> "BUDGET".equals(recipe.id())));
    assertEquals(
        RecipeRegistry.recipeFor("DASHBOARD").orElseThrow().intentTags(),
        dashboardDefinition.task().intentTags());
    assertEquals(
        ShippedExamples.repositoryExamples().stream()
            .filter(example -> "BUDGET".equals(example.id()))
            .findFirst()
            .orElseThrow()
            .plan(),
        RecipeRegistry.repositoryExamplePlanFor("BUDGET"));
    assertEquals("DASHBOARD", RecipeRegistry.requireTaskEntry("DASHBOARD").id());

    assertEquals(
        "id must not be null",
        assertThrows(NullPointerException.class, () -> RecipeRegistry.recipeFor(null))
            .getMessage());
    assertEquals(
        "id must not be null",
        assertThrows(NullPointerException.class, () -> RecipeRegistry.taskEntryFor(null))
            .getMessage());
    assertEquals(
        "Missing repository example plan for NO_SUCH_EXAMPLE",
        assertThrows(
                IllegalStateException.class,
                () -> RecipeRegistry.repositoryExamplePlanFor("NO_SUCH_EXAMPLE"))
            .getMessage());
    assertEquals(
        "Missing task entry for recipe NO_SUCH_TASK",
        assertThrows(
                IllegalStateException.class, () -> RecipeRegistry.requireTaskEntry("NO_SUCH_TASK"))
            .getMessage());
    assertEquals(Map.of("ONE", recipe("ONE")), RecipeRegistry.recipesById(List.of(recipe("ONE"))));
    assertEquals(
        "Duplicate recipe id DUPLICATE",
        assertThrows(
                IllegalStateException.class,
                () -> RecipeRegistry.recipesById(List.of(recipe("DUPLICATE"), recipe("DUPLICATE"))))
            .getMessage());

    Map<String, TaskEntry> taskEntries = new ConcurrentHashMap<>();
    RecipeRegistry.putUniqueAssociated(
        taskEntries, "DASHBOARD", dashboardDefinition.task(), "task entry");
    assertEquals("DASHBOARD", taskEntries.get("DASHBOARD").id());
    assertEquals(
        "Duplicate task entry for DASHBOARD",
        assertThrows(
                IllegalStateException.class,
                () ->
                    RecipeRegistry.putUniqueAssociated(
                        taskEntries, "DASHBOARD", dashboardDefinition.task(), "task entry"))
            .getMessage());
  }

  @Test
  void taskStarterRecipeSupportRejectsBlankFileNamesAndTaskIds() {
    assertEquals(
        "fileName must not be null",
        assertThrows(NullPointerException.class, () -> TaskStarters.taskStarterAsset(null))
            .getMessage());
    assertEquals(
        "fileName must not be blank",
        assertThrows(IllegalArgumentException.class, () -> TaskStarters.taskStarterAsset(" "))
            .getMessage());
    assertEquals(
        "taskId must not be null",
        assertThrows(NullPointerException.class, () -> TaskStarters.taskRequestFileName(null))
            .getMessage());
    assertEquals(
        "taskId must not be blank",
        assertThrows(IllegalArgumentException.class, () -> TaskStarters.taskRequestFileName(" "))
            .getMessage());
  }

  @Test
  void taskStarterPreviewIsAStandardsValidPng() throws java.io.IOException {
    byte[] png = Base64.getDecoder().decode(TaskStarters.onePixelPngBase64());

    assertNotNull(ImageIO.read(new ByteArrayInputStream(png)));
  }

  private static WorkbookPlan templatePlan() {
    return GridGrindProtocolCatalog.requestTemplate();
  }

  private static Recipe recipe(String id) {
    return new Recipe(
        id,
        dev.erst.gridgrind.core.discovery.RecipeType.EXAMPLE,
        id.toLowerCase(Locale.ROOT) + ".json",
        "summary",
        RecipeAdvisory.SELF_CONTAINED,
        List.of(),
        List.of("tag"),
        templatePlan());
  }
}
