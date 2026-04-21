package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for contract-owned CLI help rendering helpers. */
class GridGrindCliHelpTest {
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
}
