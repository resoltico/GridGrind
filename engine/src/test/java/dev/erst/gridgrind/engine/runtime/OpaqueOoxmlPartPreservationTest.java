package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.Preservation;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Proves byte preservation compares named OOXML package entries rather than semantic summaries. */
class OpaqueOoxmlPartPreservationTest {
  @TempDir Path root;

  @Test
  void establishesByteIdentityForEveryExplicitPart() throws Exception {
    Path source = packageWithPart("source.xlsx", "customXml/item1.xml", "<report>1</report>");
    Path staged = packageWithPart("staged.xlsx", "customXml/item1.xml", "<report>1</report>");

    Preservation.ByteEstablished established =
        OpaqueOoxmlPartPreservation.establish(source, staged, List.of("/customXml/item1.xml"));

    assertEquals(List.of("/customXml/item1.xml"), established.partNames());
  }

  @Test
  void rejectsChangedOrMissingBytePreservation() throws Exception {
    Path source = packageWithPart("source.xlsx", "customXml/item1.xml", "<report>1</report>");
    Path changed = packageWithPart("changed.xlsx", "customXml/item1.xml", "<report>2</report>");

    PreservationFailedException changedFailure =
        assertThrows(
            PreservationFailedException.class,
            () ->
                OpaqueOoxmlPartPreservation.establish(
                    source, changed, List.of("/customXml/item1.xml")));
    assertEquals(Preservation.Kind.BYTE, changedFailure.kind());
    assertEquals(List.of("/customXml/item1.xml"), changedFailure.identifiers());

    Path missing = packageWithPart("missing.xlsx", "customXml/item2.xml", "<report>1</report>");
    PreservationFailedException missingFailure =
        assertThrows(
            PreservationFailedException.class,
            () ->
                OpaqueOoxmlPartPreservation.establish(
                    source, missing, List.of("/customXml/item1.xml")));
    assertEquals(Preservation.Kind.BYTE, missingFailure.kind());

    Path differentSize =
        packageWithPart("different-size.xlsx", "customXml/item1.xml", "<report>longer</report>");
    PreservationFailedException sizeFailure =
        assertThrows(
            PreservationFailedException.class,
            () ->
                OpaqueOoxmlPartPreservation.establish(
                    source, differentSize, List.of("/customXml/item1.xml")));
    assertEquals(List.of("/customXml/item1.xml"), sizeFailure.identifiers());

    Path sourceMissing =
        packageWithPart("source-missing.xlsx", "customXml/item2.xml", "<report>1</report>");
    PreservationFailedException sourceMissingFailure =
        assertThrows(
            PreservationFailedException.class,
            () ->
                OpaqueOoxmlPartPreservation.establish(
                    sourceMissing, changed, List.of("/customXml/item1.xml")));
    assertEquals(List.of("/customXml/item1.xml"), sourceMissingFailure.identifiers());
  }

  @Test
  void treatsDifferentReadProgressAsDifferentByteStreams() throws Exception {
    try (var source = new ByteArrayInputStream(new byte[] {1, 2});
        var staged = new SingleByteReadInputStream(new byte[] {1, 2})) {
      assertFalse(OpaqueOoxmlPartPreservation.identicalBytes(source, staged));
    }
  }

  @Test
  void rejectsHostPartNamesThatCannotNameOneOoxmlPackagePart() {
    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.Requirement
                    .PreserveOpaqueOoxmlParts(List.of("../customXml/item1.xml")));

    assertEquals("partNames must contain normalized OOXML part names", exception.getMessage());
  }

  private Path packageWithPart(String fileName, String partName, String value) throws IOException {
    Path archive = root.resolve(fileName);
    try (ZipOutputStream output = new ZipOutputStream(Files.newOutputStream(archive))) {
      output.putNextEntry(new ZipEntry(partName));
      output.write(value.getBytes(StandardCharsets.UTF_8));
      output.closeEntry();
    }
    return archive;
  }
}
