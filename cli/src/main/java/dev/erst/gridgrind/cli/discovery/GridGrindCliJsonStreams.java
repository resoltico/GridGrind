package dev.erst.gridgrind.cli.discovery;

import java.io.IOException;
import java.io.InputStream;
import tools.jackson.databind.JsonNode;

/** Internal stream-oriented helpers for tests that exercise caller-owned stream lifecycles. */
final class GridGrindCliJsonStreams {
  private GridGrindCliJsonStreams() {}

  static TaskCatalog readTaskCatalog(InputStream inputStream) throws IOException {
    return GridGrindCliJsonCodecSupport.readStream(inputStream, TaskCatalog.class);
  }

  static TaskKeywordMatchReport readTaskKeywordMatchReport(InputStream inputStream)
      throws IOException {
    return GridGrindCliJsonCodecSupport.readStream(inputStream, TaskKeywordMatchReport.class);
  }

  static ShippedExampleCatalog readShippedExampleCatalog(InputStream inputStream)
      throws IOException {
    return GridGrindCliJsonCodecSupport.readStream(inputStream, ShippedExampleCatalog.class);
  }

  static ProtocolCatalogSearchReport readProtocolCatalogSearchReport(InputStream inputStream)
      throws IOException {
    return GridGrindCliJsonCodecSupport.readStream(inputStream, ProtocolCatalogSearchReport.class);
  }

  static CliFailureReport readCliFailureReport(InputStream inputStream) throws IOException {
    return GridGrindCliJsonCodecSupport.readStream(inputStream, CliFailureReport.class);
  }

  static JsonNode readTree(byte[] bytes) throws IOException {
    return GridGrindCliJsonCodecSupport.readTree(bytes);
  }
}
