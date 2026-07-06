package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

/** Coverage for direct catalog-command guidance helpers. */
class CliCatalogCommandSupportTest {
  @Test
  void unknownRecipeMessageCoversExampleTaskAndFallbackSuggestions() {
    assertTrue(
        CliCatalogCommandSupport.unknownRecipeMessage("workbook_health")
            .contains("did you mean WORKBOOK_HEALTH"));
    assertTrue(
        CliCatalogCommandSupport.unknownRecipeMessage("dashboard")
            .contains("did you mean DASHBOARD"));
    assertTrue(
        CliCatalogCommandSupport.unknownRecipeMessage("totally-unknown")
            .contains("discover valid recipe ids"));
  }
}
