package dev.erst.gridgrind.cli.discovery;

import java.io.IOException;
import java.io.OutputStream;

/** Shared JSON codec for CLI-owned discovery surfaces. */
public final class GridGrindCliJson {
  private GridGrindCliJson() {}

  /** Reads one task catalog from UTF-8 JSON bytes. */
  public static TaskCatalog readTaskCatalog(byte[] bytes) throws IOException {
    return GridGrindCliJsonCodecSupport.readBytes(bytes, TaskCatalog.class);
  }

  /** Reads one keyword-match report from UTF-8 JSON bytes. */
  public static TaskKeywordMatchReport readTaskKeywordMatchReport(byte[] bytes) throws IOException {
    return GridGrindCliJsonCodecSupport.readBytes(bytes, TaskKeywordMatchReport.class);
  }

  /** Reads one built-in example catalog from UTF-8 JSON bytes. */
  public static ShippedExampleCatalog readShippedExampleCatalog(byte[] bytes) throws IOException {
    return GridGrindCliJsonCodecSupport.readBytes(bytes, ShippedExampleCatalog.class);
  }

  /** Reads one protocol-catalog search report from UTF-8 JSON bytes. */
  public static ProtocolCatalogSearchReport readProtocolCatalogSearchReport(byte[] bytes)
      throws IOException {
    return GridGrindCliJsonCodecSupport.readBytes(bytes, ProtocolCatalogSearchReport.class);
  }

  /** Reads one CLI failure report from UTF-8 JSON bytes. */
  public static CliFailureReport readCliFailureReport(byte[] bytes) throws IOException {
    return GridGrindCliJsonCodecSupport.readBytes(bytes, CliFailureReport.class);
  }

  /** Writes one task catalog to UTF-8 JSON bytes. */
  public static byte[] writeTaskCatalogBytes(TaskCatalog catalog) throws IOException {
    return GridGrindCliJsonCodecSupport.writeBytes(catalog);
  }

  /** Writes one keyword-match report to UTF-8 JSON bytes. */
  public static byte[] writeTaskKeywordMatchReportBytes(TaskKeywordMatchReport report)
      throws IOException {
    return GridGrindCliJsonCodecSupport.writeBytes(report);
  }

  /** Writes one built-in example catalog to UTF-8 JSON bytes. */
  public static byte[] writeShippedExampleCatalogBytes(ShippedExampleCatalog catalog)
      throws IOException {
    return GridGrindCliJsonCodecSupport.writeBytes(catalog);
  }

  /** Writes one protocol-catalog search report to UTF-8 JSON bytes. */
  public static byte[] writeProtocolCatalogSearchReportBytes(ProtocolCatalogSearchReport report)
      throws IOException {
    return GridGrindCliJsonCodecSupport.writeBytes(report);
  }

  /** Writes one CLI failure report to UTF-8 JSON bytes. */
  public static byte[] writeCliFailureReportBytes(CliFailureReport report) throws IOException {
    return GridGrindCliJsonCodecSupport.writeBytes(report);
  }

  /** Writes one task entry as JSON without closing the caller-owned output stream. */
  public static void writeTaskEntry(OutputStream outputStream, TaskEntry entry) throws IOException {
    GridGrindCliJsonCodecSupport.writeValue(outputStream, entry);
  }

  /** Writes one task catalog as JSON without closing the caller-owned output stream. */
  public static void writeTaskCatalog(OutputStream outputStream, TaskCatalog catalog)
      throws IOException {
    GridGrindCliJsonCodecSupport.writeValue(outputStream, catalog);
  }

  /** Writes one keyword-match report as JSON without closing the caller-owned output stream. */
  public static void writeTaskKeywordMatchReport(
      OutputStream outputStream, TaskKeywordMatchReport report) throws IOException {
    GridGrindCliJsonCodecSupport.writeValue(outputStream, report);
  }

  /** Writes one built-in example catalog as JSON without closing the caller-owned output stream. */
  public static void writeShippedExampleCatalog(
      OutputStream outputStream, ShippedExampleCatalog catalog) throws IOException {
    GridGrindCliJsonCodecSupport.writeValue(outputStream, catalog);
  }

  /** Writes one protocol-catalog search report as JSON without closing the caller stream. */
  public static void writeProtocolCatalogSearchReport(
      OutputStream outputStream, ProtocolCatalogSearchReport report) throws IOException {
    GridGrindCliJsonCodecSupport.writeValue(outputStream, report);
  }

  /** Writes one CLI failure report as JSON without closing the caller-owned output stream. */
  public static void writeCliFailureReport(OutputStream outputStream, CliFailureReport report)
      throws IOException {
    GridGrindCliJsonCodecSupport.writeValue(outputStream, report);
  }
}
