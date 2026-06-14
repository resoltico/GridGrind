package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.zip.CRC32;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;
import org.junit.jupiter.api.Test;

/** Focused coverage for deterministic OOXML artifact ZIP normalization. */
class ExcelDeterministicWorkbookArtifactSupportTest {
  @Test
  void normalizeWorkbookPackagePreservesEntryKindsAndAppliesDeterministicMetadata()
      throws IOException {
    byte[] storedBytes = "<workbook/>".getBytes(java.nio.charset.StandardCharsets.UTF_8);
    byte[] deflatedBytes = "<coreProperties/>".getBytes(java.nio.charset.StandardCharsets.UTF_8);
    Path workbookPath =
        ExcelTempFiles.createManagedTempFile("gridgrind-deterministic-package-", ".xlsx");
    writePackageFixture(workbookPath, storedBytes, deflatedBytes);

    long expectedTime = ExcelDeterministicWorkbookArtifactSupport.deterministicZipTimeMillis();
    long storedCrc = crc32(storedBytes);

    ExcelDeterministicWorkbookArtifactSupport.normalizeWorkbookPackage(workbookPath);

    try (ZipFile zipFile = new ZipFile(workbookPath.toFile())) {
      ZipEntry directoryEntry = zipFile.getEntry("xl/");
      assertTrue(directoryEntry.isDirectory());
      assertEquals(expectedTime, directoryEntry.getTime());
      assertEquals(ZipEntry.STORED, directoryEntry.getMethod());
      assertEquals(0, directoryEntry.getSize());
      assertEquals(0, directoryEntry.getCompressedSize());
      assertEquals(0, directoryEntry.getCrc());
      assertEquals("dir", directoryEntry.getComment());

      ZipEntry storedEntry = zipFile.getEntry("xl/workbook.xml");
      assertEquals(expectedTime, storedEntry.getTime());
      assertEquals(ZipEntry.STORED, storedEntry.getMethod());
      assertEquals(storedBytes.length, storedEntry.getSize());
      assertEquals(storedBytes.length, storedEntry.getCompressedSize());
      assertEquals(storedCrc, storedEntry.getCrc());
      assertEquals("stored", storedEntry.getComment());
      assertArrayEquals(storedBytes, zipFile.getInputStream(storedEntry).readAllBytes());

      ZipEntry deflatedEntry = zipFile.getEntry("docProps/core.xml");
      assertEquals(expectedTime, deflatedEntry.getTime());
      assertEquals(ZipEntry.DEFLATED, deflatedEntry.getMethod());
      assertEquals("deflated", deflatedEntry.getComment());
      assertArrayEquals(deflatedBytes, zipFile.getInputStream(deflatedEntry).readAllBytes());
    }
  }

  @Test
  void normalizeWorkbookPackageDeletesItsTempOutputWhenRewriteFails() throws IOException {
    Path workspace =
        ExcelTempFiles.createManagedTempDirectory("gridgrind-deterministic-failure-workspace-");
    Path invalidWorkbookPath = workspace.resolve("invalid.xlsx");
    Files.writeString(invalidWorkbookPath, "not a zip workbook");

    IOException failure =
        assertThrows(
            IOException.class,
            () ->
                ExcelDeterministicWorkbookArtifactSupport.normalizeWorkbookPackage(
                    invalidWorkbookPath));
    assertFalse(failure.getMessage().isBlank());
    try (var entries = Files.list(workspace)) {
      assertEquals(
          0L,
          entries
              .filter(path -> path.getFileName().toString().startsWith("gridgrind-deterministic-"))
              .count());
    }
  }

  private static void writePackageFixture(
      Path workbookPath, byte[] storedBytes, byte[] deflatedBytes) throws IOException {
    try (ZipOutputStream outputStream = new ZipOutputStream(Files.newOutputStream(workbookPath))) {
      ZipEntry directoryEntry = new ZipEntry("xl/");
      directoryEntry.setComment("dir");
      directoryEntry.setMethod(ZipEntry.STORED);
      directoryEntry.setSize(0);
      directoryEntry.setCompressedSize(0);
      directoryEntry.setCrc(0);
      outputStream.putNextEntry(directoryEntry);
      outputStream.closeEntry();

      ZipEntry storedEntry = new ZipEntry("xl/workbook.xml");
      storedEntry.setComment("stored");
      storedEntry.setMethod(ZipEntry.STORED);
      storedEntry.setSize(storedBytes.length);
      storedEntry.setCompressedSize(storedBytes.length);
      storedEntry.setCrc(crc32(storedBytes));
      outputStream.putNextEntry(storedEntry);
      outputStream.write(storedBytes);
      outputStream.closeEntry();

      ZipEntry deflatedEntry = new ZipEntry("docProps/core.xml");
      deflatedEntry.setComment("deflated");
      deflatedEntry.setMethod(ZipEntry.DEFLATED);
      outputStream.putNextEntry(deflatedEntry);
      outputStream.write(deflatedBytes);
      outputStream.closeEntry();
    }
  }

  private static long crc32(byte[] bytes) {
    CRC32 crc32 = new CRC32();
    crc32.update(bytes);
    return crc32.getValue();
  }
}
