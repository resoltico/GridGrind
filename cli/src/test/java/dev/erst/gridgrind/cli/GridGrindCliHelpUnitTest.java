package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

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
        GridGrindCli.requestTemplateText(
            () -> "{\"protocolVersion\":\"V1\"}".getBytes(StandardCharsets.UTF_8)));
  }

  @Test
  void requestTemplateTextWrapsSerializationFailures() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindCli.requestTemplateText(
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
        GridGrindCliHelpRenderSupport.renderDefinitions(
            new CliSurface.CliDefinitionSection("Limits", List.of())));
    assertEquals(
        "Docs:",
        GridGrindCliHelpRenderSupport.renderReferences(
            new CliSurface.CliReferenceSection("Docs", List.of()), "https://example.invalid/root"));
  }

  @Test
  void renderHelpersWrapLongReferenceDescriptions() {
    String rendered =
        GridGrindCliHelpRenderSupport.renderReferences(
            new CliSurface.CliReferenceSection(
                "Docs",
                List.of(
                    new CliSurface.ReferenceEntry(
                        "Quick Start",
                        "QUICK_START.md",
                        "One deliberately long description that must wrap across multiple help"
                            + " lines for coverage because the renderer should emit at least one"
                            + " continuation line when public help prose exceeds the configured"
                            + " terminal width."))),
            "https://example.invalid/root");

    java.util.List<String> lines = rendered.lines().toList();
    assertTrue(lines.size() >= 4);
    assertTrue(lines.get(2).startsWith("    "));
    assertTrue(lines.get(3).startsWith("    "));
  }

  @Test
  void renderHelpersKeepOversizedSingleTokensOnTheirCurrentLine() {
    String longToken = "x".repeat(160);
    String rendered =
        GridGrindCliHelpRenderSupport.renderReferences(
            new CliSurface.CliReferenceSection(
                "Docs",
                List.of(new CliSurface.ReferenceEntry("Quick Start", "QUICK_START.md", longToken))),
            "https://example.invalid/root");

    assertTrue(rendered.contains("\n    " + longToken));
  }

  @Test
  void renderSectionWrapsBulletListsWithStableContinuationIndentation() {
    String rendered =
        GridGrindCliHelpRenderSupport.renderSection(
            new CliSurface.CliSection(
                "Notes",
                List.of(
                    "- One deliberately long bullet entry that needs continuation alignment for the"
                        + " public help surface so wrapped lines remain easy to scan.")));

    List<String> lines = rendered.lines().toList();
    assertEquals("Notes:", lines.getFirst());
    assertTrue(lines.get(1).startsWith("  - "));
    assertTrue(lines.stream().skip(2).allMatch(line -> line.startsWith("    ")));
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
