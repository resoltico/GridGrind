package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
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
  void renderCoordinateSystemsFallsBackToStackedLayoutWhenTerminalWidthIsTight() {
    String rendered =
        GridGrindCliHelpRenderSupport.renderCoordinateSystems(
            new CliSurface.CliTableSection(
                "Coordinate Systems",
                "Pattern",
                "Meaning",
                List.of(
                    new CliSurface.CoordinateSystemEntry(
                        "R1C1_REFERENCE_STYLE",
                        "One deliberately long convention description that cannot fit beside the"
                            + " pattern in a narrow terminal width."))),
            28);

    assertTrue(rendered.contains("Coordinate Systems:"));
    assertTrue(rendered.contains("  R1C1_REFERENCE_STYLE:"));
    assertTrue(rendered.contains("    deliberately long"));
    assertTrue(rendered.contains("    terminal width."));
    assertFalse(rendered.contains("Pattern Meaning"));
  }

  @Test
  void renderSectionPreservesCommandLabelsAndShellRedirections() {
    String rendered =
        GridGrindCliHelpRenderSupport.renderSection(
            new CliSurface.CliSection(
                "Examples",
                List.of(
                    "1. Print a minimal request: gridgrind --print-request-template < request.json"
                        + " > rendered.json",
                    "- Run via Docker: docker run --rm -i ghcr.io/resoltico/gridgrind:latest <"
                        + " request.json > response.json")),
            56);

    assertTrue(rendered.contains("1. Print a minimal request:"));
    assertTrue(rendered.contains("gridgrind --print-request-template"));
    assertTrue(rendered.contains("< request.json"));
    assertTrue(rendered.contains("> rendered.json"));
    assertTrue(rendered.contains("- Run via Docker:"));
    assertTrue(rendered.contains("docker run --rm -i"));
    assertTrue(rendered.contains("> response.json"));
  }

  @Test
  void wrappedIndentedLineSupportsOrderedListItemsWithoutCommandMarkers() {
    String rendered =
        GridGrindCliHelpRenderSupport.wrappedIndentedLine(
            "1. One deliberately long ordered help line that must wrap without becoming a command"
                + " block.",
            "  ",
            44);

    List<String> lines = rendered.lines().toList();
    assertTrue(lines.getFirst().startsWith("  1. "));
    assertTrue(lines.stream().skip(1).allMatch(line -> line.startsWith("     ")));
  }

  @Test
  void wrappingTokensHandlesBlankInputAndTerminalRedirectionTokens() {
    assertEquals(List.of(), GridGrindCliWrappingSupport.wrappingTokens(""));
    assertEquals(List.of(), GridGrindCliWrappingSupport.wrappingTokens("   "));
    assertEquals(
        List.of("gridgrind", "--print-request-template", ">"),
        GridGrindCliWrappingSupport.wrappingTokens("gridgrind --print-request-template >"));
    assertEquals(
        List.of("gridgrind", "< request.json", "> response.json"),
        GridGrindCliWrappingSupport.wrappingTokens("gridgrind < request.json > response.json"));
  }

  @Test
  void helpTextWidthParsesBlankInvalidAndClampedColumnHints() {
    assertEquals(88, GridGrindCliWrappingSupport.helpTextWidth(null));
    assertEquals(88, GridGrindCliWrappingSupport.helpTextWidth(" "));
    assertEquals(88, GridGrindCliWrappingSupport.helpTextWidth("wide enough?"));
    assertEquals(72, GridGrindCliWrappingSupport.helpTextWidth("40"));
    assertEquals(91, GridGrindCliWrappingSupport.helpTextWidth("91"));
    assertEquals(120, GridGrindCliWrappingSupport.helpTextWidth("240"));
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
