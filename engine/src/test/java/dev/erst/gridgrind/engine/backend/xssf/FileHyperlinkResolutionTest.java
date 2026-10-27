package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.engine.links.FileHyperlinkResolution;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Tests for local file hyperlink resolution result records. */
class FileHyperlinkResolutionTest {
  @Test
  void validatesResolutionRecordInputs() {
    assertThrows(
        IllegalArgumentException.class,
        () -> new FileHyperlinkResolution.UnresolvedRelativePath(" "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new FileHyperlinkResolution.ResolvedPath(" ", Path.of("/tmp/report.xlsx")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new FileHyperlinkResolution.MalformedPath(" ", "bad path"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new FileHyperlinkResolution.MalformedPath("report.xlsx", " "));
  }
}
