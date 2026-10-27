package dev.erst.gridgrind.core.discovery;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

/** Coverage for task-recipe helper guards. */
class GridGrindTaskRecipeSupportTest {
  @Test
  void starterHelpersRejectNullAndBlankTaskIds() {
    NullPointerException nullTaskId =
        assertThrows(NullPointerException.class, () -> TaskRecipes.selfContainedStarter(null));
    assertEquals("taskId must not be null", nullTaskId.getMessage());

    IllegalArgumentException blankTaskId =
        assertThrows(
            IllegalArgumentException.class,
            () -> TaskRecipes.assetBackedStarter(" ", "asset.xlsx"));
    assertEquals("taskId must not be blank", blankTaskId.getMessage());
  }
}
