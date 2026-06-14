package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.nio.file.StandardOpenOption;
import java.time.Instant;
import java.util.Date;
import java.util.Enumeration;
import java.util.Objects;
import java.util.Optional;
import java.util.zip.CRC32;
import java.util.zip.Deflater;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

/** Normalizes workbook artifact metadata and ZIP container timestamps for deterministic saves. */
public final class ExcelDeterministicWorkbookArtifactSupport {
  private static final Instant DEFAULT_ARTIFACT_INSTANT = Instant.parse("2000-01-01T00:00:00Z");

  private ExcelDeterministicWorkbookArtifactSupport() {}

  /** Returns the deterministic workbook timestamp derived from the ambient build environment. */
  public static Optional<Date> deterministicTimestamp() {
    return deterministicTimestamp(System.getenv("SOURCE_DATE_EPOCH"));
  }

  /** Returns the deterministic workbook timestamp derived from one explicit epoch-seconds value. */
  public static Optional<Date> deterministicTimestamp(String sourceDateEpoch) {
    return Optional.of(Date.from(deterministicInstant(sourceDateEpoch)));
  }

  /** Returns the deterministic ZIP-entry timestamp derived from the ambient build environment. */
  public static long deterministicZipTimeMillis() {
    return deterministicZipTimeMillis(System.getenv("SOURCE_DATE_EPOCH"));
  }

  /**
   * Returns the deterministic ZIP-entry timestamp derived from one explicit epoch-seconds value.
   */
  public static long deterministicZipTimeMillis(String sourceDateEpoch) {
    return deterministicInstant(sourceDateEpoch).toEpochMilli();
  }

  /**
   * Rewrites one saved workbook package so ZIP metadata is deterministic across equivalent saves.
   */
  public static void normalizeWorkbookPackage(Path workbookPath) throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    Path absolutePath = workbookPath.toAbsolutePath().normalize();
    Path parent =
        Objects.requireNonNull(
            absolutePath.getParent(), "workbookPath must not be a filesystem root");
    Path normalizedPath = Files.createTempFile(parent, "gridgrind-deterministic-", ".xlsx");
    try {
      rewriteWorkbookPackage(absolutePath, normalizedPath, deterministicZipTimeMillis());
      Files.move(
          normalizedPath,
          absolutePath,
          StandardCopyOption.REPLACE_EXISTING,
          StandardCopyOption.ATOMIC_MOVE);
    } catch (IOException exception) {
      Files.deleteIfExists(normalizedPath);
      throw exception;
    }
  }

  private static void rewriteWorkbookPackage(
      Path sourcePath, Path normalizedPath, long deterministicZipTimeMillis) throws IOException {
    try (ZipFile zipFile = new ZipFile(sourcePath.toFile());
        ZipOutputStream outputStream =
            new ZipOutputStream(
                Files.newOutputStream(
                    normalizedPath,
                    StandardOpenOption.TRUNCATE_EXISTING,
                    StandardOpenOption.WRITE))) {
      outputStream.setLevel(Deflater.DEFAULT_COMPRESSION);
      Enumeration<? extends ZipEntry> entries = zipFile.entries();
      while (entries.hasMoreElements()) {
        ZipEntry sourceEntry = entries.nextElement();
        byte[] content = entryBytes(zipFile, sourceEntry);
        ZipEntry normalizedEntry =
            normalizedEntry(sourceEntry, content, deterministicZipTimeMillis);
        outputStream.putNextEntry(normalizedEntry);
        if (!sourceEntry.isDirectory()) {
          outputStream.write(content);
        }
        outputStream.closeEntry();
      }
    }
  }

  private static byte[] entryBytes(ZipFile zipFile, ZipEntry entry) throws IOException {
    if (entry.isDirectory()) {
      return new byte[0];
    }
    try (InputStream inputStream = zipFile.getInputStream(entry)) {
      return inputStream.readAllBytes();
    }
  }

  private static ZipEntry normalizedEntry(
      ZipEntry sourceEntry, byte[] content, long deterministicZipTimeMillis) {
    ZipEntry normalizedEntry = new ZipEntry(sourceEntry.getName());
    normalizedEntry.setComment(sourceEntry.getComment());
    normalizedEntry.setTime(deterministicZipTimeMillis);
    if (sourceEntry.isDirectory()) {
      normalizedEntry.setMethod(ZipEntry.STORED);
      normalizedEntry.setSize(0);
      normalizedEntry.setCompressedSize(0);
      normalizedEntry.setCrc(0);
      return normalizedEntry;
    }
    if (sourceEntry.getMethod() == ZipEntry.STORED) {
      CRC32 crc32 = new CRC32();
      crc32.update(content);
      normalizedEntry.setMethod(ZipEntry.STORED);
      normalizedEntry.setSize(content.length);
      normalizedEntry.setCompressedSize(content.length);
      normalizedEntry.setCrc(crc32.getValue());
      return normalizedEntry;
    }
    normalizedEntry.setMethod(ZipEntry.DEFLATED);
    return normalizedEntry;
  }

  private static Instant deterministicInstant(String sourceDateEpoch) {
    if (sourceDateEpoch == null || sourceDateEpoch.isBlank()) {
      return DEFAULT_ARTIFACT_INSTANT;
    }
    try {
      long epochSeconds = Long.parseLong(sourceDateEpoch.trim());
      if (epochSeconds < 0) {
        throw new IllegalArgumentException("SOURCE_DATE_EPOCH must be >= 0");
      }
      return Instant.ofEpochSecond(epochSeconds);
    } catch (NumberFormatException exception) {
      throw new IllegalArgumentException(
          "SOURCE_DATE_EPOCH must be one integer Unix timestamp in whole seconds", exception);
    }
  }
}
