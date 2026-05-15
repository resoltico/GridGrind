package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayOutputStream;
import java.io.FilterInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Objects;
import tools.jackson.databind.DeserializationFeature;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.json.JsonMapper;

/** Shared JSON codec for CLI-owned discovery surfaces. */
public final class GridGrindCliJson {
  private static final JsonMapper JSON_MAPPER =
      JsonMapper.builder().enable(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES).build();

  private GridGrindCliJson() {}

  /** Reads one task catalog from UTF-8 JSON bytes. */
  public static TaskCatalog readTaskCatalog(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return JSON_MAPPER.readValue(bytes, TaskCatalog.class);
  }

  /** Reads one task catalog from a caller-owned input stream without closing it. */
  public static TaskCatalog readTaskCatalog(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return JSON_MAPPER.readValue(nonClosing(inputStream), TaskCatalog.class);
  }

  /** Reads one keyword-match report from UTF-8 JSON bytes. */
  public static TaskKeywordMatchReport readTaskKeywordMatchReport(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return JSON_MAPPER.readValue(bytes, TaskKeywordMatchReport.class);
  }

  /** Reads one keyword-match report from a caller-owned input stream without closing it. */
  public static TaskKeywordMatchReport readTaskKeywordMatchReport(InputStream inputStream)
      throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return JSON_MAPPER.readValue(nonClosing(inputStream), TaskKeywordMatchReport.class);
  }

  /** Reads one built-in example catalog from UTF-8 JSON bytes. */
  public static ShippedExampleCatalog readShippedExampleCatalog(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return JSON_MAPPER.readValue(bytes, ShippedExampleCatalog.class);
  }

  /** Reads one CLI failure report from UTF-8 JSON bytes. */
  public static CliFailureReport readCliFailureReport(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return JSON_MAPPER.readValue(bytes, CliFailureReport.class);
  }

  /** Reads one built-in example catalog from a caller-owned input stream without closing it. */
  public static ShippedExampleCatalog readShippedExampleCatalog(InputStream inputStream)
      throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return JSON_MAPPER.readValue(nonClosing(inputStream), ShippedExampleCatalog.class);
  }

  /** Reads one CLI failure report from a caller-owned input stream without closing it. */
  public static CliFailureReport readCliFailureReport(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return JSON_MAPPER.readValue(nonClosing(inputStream), CliFailureReport.class);
  }

  /** Reads one generic JSON tree from UTF-8 bytes. */
  public static JsonNode readTree(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return JSON_MAPPER.readTree(bytes);
  }

  /** Writes one task catalog to UTF-8 JSON bytes. */
  public static byte[] writeTaskCatalogBytes(TaskCatalog catalog) throws IOException {
    return writeBytes(catalog);
  }

  /** Writes one keyword-match report to UTF-8 JSON bytes. */
  public static byte[] writeTaskKeywordMatchReportBytes(TaskKeywordMatchReport report)
      throws IOException {
    return writeBytes(report);
  }

  /** Writes one built-in example catalog to UTF-8 JSON bytes. */
  public static byte[] writeShippedExampleCatalogBytes(ShippedExampleCatalog catalog)
      throws IOException {
    return writeBytes(catalog);
  }

  /** Writes one CLI failure report to UTF-8 JSON bytes. */
  public static byte[] writeCliFailureReportBytes(CliFailureReport report) throws IOException {
    return writeBytes(report);
  }

  /** Writes one task entry as JSON without closing the caller-owned output stream. */
  public static void writeTaskEntry(OutputStream outputStream, TaskEntry entry) throws IOException {
    writeValue(outputStream, entry);
  }

  /** Writes one task catalog as JSON without closing the caller-owned output stream. */
  public static void writeTaskCatalog(OutputStream outputStream, TaskCatalog catalog)
      throws IOException {
    writeValue(outputStream, catalog);
  }

  /** Writes one keyword-match report as JSON without closing the caller-owned output stream. */
  public static void writeTaskKeywordMatchReport(
      OutputStream outputStream, TaskKeywordMatchReport report) throws IOException {
    writeValue(outputStream, report);
  }

  /** Writes one built-in example catalog as JSON without closing the caller-owned output stream. */
  public static void writeShippedExampleCatalog(
      OutputStream outputStream, ShippedExampleCatalog catalog) throws IOException {
    writeValue(outputStream, catalog);
  }

  /** Writes one CLI failure report as JSON without closing the caller-owned output stream. */
  public static void writeCliFailureReport(OutputStream outputStream, CliFailureReport report)
      throws IOException {
    writeValue(outputStream, report);
  }

  private static byte[] writeBytes(Object value) throws IOException {
    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    writeValue(outputStream, value);
    return outputStream.toByteArray();
  }

  private static void writeValue(OutputStream outputStream, Object value) throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(value, "value must not be null");
    GridGrindJson.writeCatalogLookupValue(outputStream, value);
  }

  private static InputStream nonClosing(InputStream inputStream) {
    return new FilterInputStream(inputStream) {
      @Override
      public void close() {
        // Caller owns the lifecycle of streams passed into the CLI discovery codec.
      }
    };
  }
}
