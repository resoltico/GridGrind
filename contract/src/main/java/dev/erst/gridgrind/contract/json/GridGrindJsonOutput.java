package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.catalog.Catalog;
import dev.erst.gridgrind.contract.catalog.TypeEntry;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Objects;
import tools.jackson.databind.node.ObjectNode;

/** Shared write-side JSON codec for the GridGrind protocol. */
public final class GridGrindJsonOutput {
  private GridGrindJsonOutput() {}

  /** Renders one request into its machine-readable object tree form without I/O. */
  public static ObjectNode requestTree(WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    return GridGrindJsonMapperSupport.WIRE_JSON_MAPPER.valueToTree(request);
  }

  /** Serializes a request to bytes. */
  public static byte[] writeRequestBytes(WorkbookPlan request) throws IOException {
    return writeRequestBytes(request, false);
  }

  /** Serializes a request to bytes. */
  public static byte[] writeRequestBytes(WorkbookPlan request, boolean pretty) throws IOException {
    Objects.requireNonNull(request, "request must not be null");
    return GridGrindJsonCodecSupport.writeBytes(wireWriteMapper(pretty), request);
  }

  /** Serializes a response to bytes. */
  public static byte[] writeResponseBytes(GridGrindResponse response) throws IOException {
    return writeResponseBytes(response, false);
  }

  /** Serializes a response to bytes. */
  public static byte[] writeResponseBytes(GridGrindResponse response, boolean pretty)
      throws IOException {
    Objects.requireNonNull(response, "response must not be null");
    return GridGrindJsonCodecSupport.writeBytes(wireWriteMapper(pretty), response);
  }

  /** Serializes a protocol catalog to bytes. */
  public static byte[] writeProtocolCatalogBytes(Catalog catalog) throws IOException {
    return writeProtocolCatalogBytes(catalog, false);
  }

  /** Serializes a protocol catalog to bytes. */
  public static byte[] writeProtocolCatalogBytes(Catalog catalog, boolean pretty)
      throws IOException {
    Objects.requireNonNull(catalog, "catalog must not be null");
    return GridGrindJsonCodecSupport.writeBytes(wireWriteMapper(pretty), catalog);
  }

  /** Serializes a request doctor report to bytes. */
  public static byte[] writeRequestDoctorReportBytes(RequestDoctorReport report)
      throws IOException {
    return writeRequestDoctorReportBytes(report, false);
  }

  /** Serializes a request doctor report to bytes. */
  public static byte[] writeRequestDoctorReportBytes(RequestDoctorReport report, boolean pretty)
      throws IOException {
    Objects.requireNonNull(report, "report must not be null");
    return GridGrindJsonCodecSupport.writeBytes(wireWriteMapper(pretty), report);
  }

  /** Writes a request to an output stream without closing the caller-owned stream. */
  public static void writeRequest(OutputStream outputStream, WorkbookPlan request, boolean pretty)
      throws IOException {
    writeValue(outputStream, request, pretty);
  }

  /** Writes a response to an output stream without closing the caller-owned stream. */
  public static void writeResponse(
      OutputStream outputStream, GridGrindResponse response, boolean pretty) throws IOException {
    writeValue(outputStream, response, pretty);
  }

  /** Writes a protocol catalog to an output stream without closing the caller-owned stream. */
  public static void writeProtocolCatalog(
      OutputStream outputStream, Catalog catalog, boolean pretty) throws IOException {
    writeValue(outputStream, catalog, pretty);
  }

  /** Writes a request doctor report to an output stream without closing the caller-owned stream. */
  public static void writeRequestDoctorReport(
      OutputStream outputStream, RequestDoctorReport report, boolean pretty) throws IOException {
    writeValue(outputStream, report, pretty);
  }

  /**
   * Writes a single catalog type entry to an output stream without closing the caller-owned stream.
   */
  public static void writeTypeEntry(OutputStream outputStream, TypeEntry entry) throws IOException {
    writeTypeEntry(outputStream, entry, false);
  }

  /**
   * Writes a single catalog type entry to an output stream without closing the caller-owned stream.
   */
  public static void writeTypeEntry(OutputStream outputStream, TypeEntry entry, boolean pretty)
      throws IOException {
    writeValue(outputStream, entry, pretty);
  }

  /** Writes one protocol-catalog lookup value to an output stream without closing it. */
  public static void writeCatalogLookupValue(OutputStream outputStream, Object value)
      throws IOException {
    writeCatalogLookupValue(outputStream, value, false);
  }

  /** Writes one protocol-catalog lookup value to an output stream without closing it. */
  public static void writeCatalogLookupValue(
      OutputStream outputStream, Object value, boolean pretty) throws IOException {
    writeValue(outputStream, value, pretty);
  }

  /**
   * Writes one catalog lookup result as a JSON object with protocolVersion prepended to the root.
   */
  public static void writeCatalogLookupResult(
      OutputStream outputStream, GridGrindProtocolVersion protocolVersion, Object value)
      throws IOException {
    writeCatalogLookupResult(outputStream, protocolVersion, value, false);
  }

  /**
   * Writes one protocol-catalog lookup result as a JSON object with protocolVersion prepended to
   * the root.
   */
  public static void writeCatalogLookupResult(
      OutputStream outputStream,
      GridGrindProtocolVersion protocolVersion,
      Object value,
      boolean pretty)
      throws IOException {
    Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    Objects.requireNonNull(value, "value must not be null");
    var mapper = wireWriteMapper(pretty);
    ObjectNode valueNode = mapper.valueToTree(value);
    ObjectNode envelope = mapper.createObjectNode();
    envelope.put("protocolVersion", protocolVersion.name());
    for (var field : valueNode.properties()) {
      envelope.set(field.getKey(), field.getValue());
    }
    writeValue(outputStream, envelope, pretty);
  }

  private static void writeValue(OutputStream outputStream, Object value, boolean pretty)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(value, "value must not be null");
    wireWriteMapper(pretty).writeValue(outputStream, value);
  }

  private static tools.jackson.databind.json.JsonMapper wireWriteMapper(boolean pretty) {
    return pretty
        ? GridGrindJsonMapperSupport.PRETTY_WIRE_JSON_MAPPER
        : GridGrindJsonMapperSupport.WIRE_JSON_MAPPER;
  }
}
