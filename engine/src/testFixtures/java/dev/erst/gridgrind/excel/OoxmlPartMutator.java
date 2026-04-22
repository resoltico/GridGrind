package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.Map;
import java.util.function.UnaryOperator;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

/** Rewrites one or more OOXML zip entries in a temporary workbook copy for tests. */
public final class OoxmlPartMutator {
  private OoxmlPartMutator() {}

  /** Returns one temporary workbook with the selected text entry rewritten. */
  public static Path rewriteEntry(
      Path sourceWorkbook, String entryName, UnaryOperator<String> transformer) throws IOException {
    return rewriteEntries(sourceWorkbook, Map.of(entryName, transformer));
  }

  /** Returns one temporary workbook with the selected text entries rewritten. */
  public static Path rewriteEntries(
      Path sourceWorkbook, Map<String, UnaryOperator<String>> transformers) throws IOException {
    List<ZipEntryBytes> entries = new ArrayList<>();
    try (ZipFile zipFile = new ZipFile(sourceWorkbook.toFile())) {
      Enumeration<? extends ZipEntry> enumeration = zipFile.entries();
      while (enumeration.hasMoreElements()) {
        ZipEntry entry = enumeration.nextElement();
        try (InputStream inputStream = zipFile.getInputStream(entry)) {
          byte[] bytes = inputStream.readAllBytes();
          UnaryOperator<String> transformer = transformers.get(entry.getName());
          entries.add(
              new ZipEntryBytes(entry.getName(), transformedEntryBytes(bytes, transformer)));
        }
      }
    }

    Path mutatedWorkbook = ExcelTempFiles.createManagedTempFile("gridgrind-mutated-", ".xlsx");
    try (ZipOutputStream outputStream =
        new ZipOutputStream(Files.newOutputStream(mutatedWorkbook))) {
      for (ZipEntryBytes entry : entries) {
        writeZipEntry(outputStream, entry.name(), entry.bytes());
      }
    }
    return mutatedWorkbook;
  }

  private static byte[] transformedEntryBytes(byte[] bytes, UnaryOperator<String> transformer) {
    if (transformer == null) {
      return bytes;
    }
    return transformer
        .apply(new String(bytes, StandardCharsets.UTF_8))
        .getBytes(StandardCharsets.UTF_8);
  }

  private static void writeZipEntry(ZipOutputStream outputStream, String name, byte[] bytes)
      throws IOException {
    outputStream.putNextEntry(new ZipEntry(name));
    outputStream.write(bytes);
    outputStream.closeEntry();
  }

  /** Immutable zip-entry payload used to preserve original OOXML entry order. */
  private static final class ZipEntryBytes {
    private final String name;
    private final byte[] bytes;

    private ZipEntryBytes(String name, byte[] bytes) {
      this.name = name;
      this.bytes = bytes.clone();
    }

    private String name() {
      return name;
    }

    private byte[] bytes() {
      return bytes.clone();
    }
  }
}
