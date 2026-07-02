package dev.erst.gridgrind.cli.discovery;

import java.io.IOException;
import java.io.OutputStream;

/** JSON codec for protocol-catalog discovery payloads. */
public final class ProtocolCatalogCliJson {
  private ProtocolCatalogCliJson() {}

  /** Reads one compact protocol-catalog index report from UTF-8 JSON bytes. */
  public static ProtocolCatalogIndexReport readProtocolCatalogIndexReport(byte[] bytes)
      throws IOException {
    return GridGrindCliJsonCodecSupport.readBytes(bytes, ProtocolCatalogIndexReport.class);
  }

  /** Reads one protocol-catalog search report from UTF-8 JSON bytes. */
  public static ProtocolCatalogSearchReport readProtocolCatalogSearchReport(byte[] bytes)
      throws IOException {
    return GridGrindCliJsonCodecSupport.readBytes(bytes, ProtocolCatalogSearchReport.class);
  }

  /** Writes one compact protocol-catalog index report to UTF-8 JSON bytes. */
  static byte[] writeProtocolCatalogIndexReportBytes(ProtocolCatalogIndexReport report)
      throws IOException {
    return writeProtocolCatalogIndexReportBytes(report, false);
  }

  /** Writes one compact protocol-catalog index report to UTF-8 JSON bytes. */
  static byte[] writeProtocolCatalogIndexReportBytes(
      ProtocolCatalogIndexReport report, boolean pretty) throws IOException {
    return GridGrindCliJsonCodecSupport.writeBytes(report, pretty);
  }

  /** Writes one protocol-catalog search report to UTF-8 JSON bytes. */
  static byte[] writeProtocolCatalogSearchReportBytes(ProtocolCatalogSearchReport report)
      throws IOException {
    return writeProtocolCatalogSearchReportBytes(report, false);
  }

  /** Writes one protocol-catalog search report to UTF-8 JSON bytes. */
  static byte[] writeProtocolCatalogSearchReportBytes(
      ProtocolCatalogSearchReport report, boolean pretty) throws IOException {
    return GridGrindCliJsonCodecSupport.writeBytes(report, pretty);
  }

  /** Writes one compact protocol-catalog index report as JSON without closing the caller stream. */
  public static void writeProtocolCatalogIndexReport(
      OutputStream outputStream, ProtocolCatalogIndexReport report, boolean pretty)
      throws IOException {
    GridGrindCliJsonCodecSupport.writeValue(outputStream, report, pretty);
  }

  /** Writes one protocol-catalog search report as JSON without closing the caller stream. */
  public static void writeProtocolCatalogSearchReport(
      OutputStream outputStream, ProtocolCatalogSearchReport report, boolean pretty)
      throws IOException {
    GridGrindCliJsonCodecSupport.writeValue(outputStream, report, pretty);
  }
}
