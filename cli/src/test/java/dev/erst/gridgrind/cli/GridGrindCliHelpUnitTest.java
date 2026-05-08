package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused unit coverage for CLI-owned help rendering helpers. */
class GridGrindCliHelpUnitTest {
  @Test
  void requestTemplateTextRendersUtf8Bytes() {
    assertEquals(
        "{\"protocolVersion\":\"V1\"}",
        GridGrindCliHelp.requestTemplateText(
            () -> "{\"protocolVersion\":\"V1\"}".getBytes(StandardCharsets.UTF_8)));
  }

  @Test
  void requestTemplateTextWrapsSerializationFailures() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindCliHelp.requestTemplateText(
                    () -> {
                      throw new IOException("synthetic failure");
                    }));

    assertEquals("Failed to render the built-in request template", failure.getMessage());
    assertEquals("synthetic failure", failure.getCause().getMessage());
  }

  @Test
  void renderHelpersCoverEmptyDefinitionAndReferenceSections() {
    assertEquals(
        "Limits:",
        GridGrindCliHelp.renderDefinitions(
            new CliSurface.CliDefinitionSection("Limits", List.of())));
    assertEquals(
        "Docs:",
        GridGrindCliHelp.renderReferences(
            new CliSurface.CliReferenceSection("Docs", List.of()), "https://example.invalid/root"));
  }

  @Test
  void cliSurfaceValueObjectsRejectBlankFields() {
    IllegalArgumentException blankSectionLabel =
        assertThrows(
            IllegalArgumentException.class, () -> new CliSurface.CliSection(" ", List.of("line")));
    assertEquals("label must not be blank", blankSectionLabel.getMessage());

    IllegalArgumentException blankCommandDescription =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new CliSurface.CliCommandExample("stdin", List.of("gridgrind"), Optional.of(" ")));
    assertEquals("description must not be blank", blankCommandDescription.getMessage());
  }
}
