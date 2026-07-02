package dev.erst.gridgrind.cli.discovery;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Residual coverage for discovery helper branches and protocol-catalog index codecs. */
class CliDiscoveryResidualCoverageTest {
  @Test
  void unavailableLocationsCollapseToEmptyOptionalFacts() {
    CliFailureLocation unavailable = CliFailureLocation.unavailable();
    CliFailureLocation cursorOnly =
        CliFailureLocation.from(Optional.empty(), Optional.of(3), Optional.of(7)).orElseThrow();

    assertFalse(unavailable.isAvailable());
    assertEquals(Optional.empty(), cursorOnly.jsonPath());
    assertEquals(Optional.of(3), cursorOnly.jsonLine());
    assertEquals(Optional.empty(), CliFailureLocation.from(new RuntimeException("boom")));
    assertEquals(
        Optional.empty(),
        CliDiscoveryValidation.copyOptionalLocation(Optional.empty(), "location"));
    assertEquals(
        Optional.empty(),
        CliDiscoveryValidation.copyOptionalLocation(Optional.of(unavailable), "location"));
  }

  @Test
  void protocolCatalogIndexCodecRoundTripsAndGroupIndicesExposeEntryCounts() throws IOException {
    ProtocolCatalogIndexReport report =
        new ProtocolCatalogIndexReport(
            GridGrindProtocolVersion.current(),
            "type",
            "SET_CELL",
            List.of(new ProtocolCatalogGroupIndex("topLevel", List.of("A", "B"))),
            List.of(new ProtocolCatalogGroupIndex("nested", List.of("C"))),
            List.of(new ProtocolCatalogGroupIndex("plain", List.of("NUMBER"))),
            List.of(
                new ProtocolCatalogFieldMetadataKey(
                    "projectedByFacets",
                    "Field is present only when one listed facet is requested.")),
            List.of(new ProtocolCatalogLookupNamespace("<group>:<id>", "Use stable catalog ids.")));

    assertEquals(2, report.topLevelGroups().getFirst().entryCount());

    byte[] bytes = ProtocolCatalogCliJson.writeProtocolCatalogIndexReportBytes(report);
    assertEquals(report, ProtocolCatalogCliJson.readProtocolCatalogIndexReport(bytes));

    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    ProtocolCatalogCliJson.writeProtocolCatalogIndexReport(outputStream, report, false);
    assertEquals(
        report, ProtocolCatalogCliJson.readProtocolCatalogIndexReport(outputStream.toByteArray()));
  }

  @Test
  void genericDiscoveryCodecConvenienceWritersDefaultToCompactJson() throws IOException {
    ProtocolCatalogSearchReport report =
        new ProtocolCatalogSearchReport(
            GridGrindProtocolVersion.current(),
            "chart title",
            List.of(
                new ProtocolCatalogSearchHit(
                    "mutationActionTypes",
                    "SET_CHART",
                    "mutationActionTypes:SET_CHART",
                    "ENTRY",
                    "Create or mutate one supported simple chart on one sheet.",
                    List.of("SET_CHART"),
                    List.of("chartInputType:ChartInput"))));

    byte[] bytes = GridGrindCliJsonCodecSupport.writeBytes(report);
    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    GridGrindCliJsonCodecSupport.writeValue(outputStream, report);

    assertEquals(report, ProtocolCatalogCliJson.readProtocolCatalogSearchReport(bytes));
    assertEquals(
        report, ProtocolCatalogCliJson.readProtocolCatalogSearchReport(outputStream.toByteArray()));
  }
}
