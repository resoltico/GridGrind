package dev.erst.gridgrind.protocol.parity;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.List;
import org.junit.jupiter.api.Test;

/** End-to-end parity-ledger verification for the Apache POI XSSF Phase 1 harness. */
final class XlsxParityTest {
  @Test
  void ledgerAndCorpusResourcesLoadAndCoverTheKnownScenarioRegistry() {
    XlsxParityLedger.Ledger ledger =
        XlsxParitySupport.call("load parity ledger", XlsxParityLedger::loadLedger);
    XlsxParityLedger.CorpusManifest corpus =
        XlsxParitySupport.call("load parity corpus manifest", XlsxParityLedger::loadCorpus);

    assertEquals("5.5.1", ledger.poiVersion());
    assertEquals("5.5.1", corpus.poiVersion());
    assertEquals(
        List.of(
            XlsxParityScenarios.CORE_WORKBOOK,
            XlsxParityScenarios.ADVANCED_NONDRAWING,
            XlsxParityScenarios.EXTERNAL_FORMULA,
            XlsxParityScenarios.UDF_FORMULA,
            XlsxParityScenarios.DRAWING_IMAGE,
            XlsxParityScenarios.CHART,
            XlsxParityScenarios.PIVOT,
            XlsxParityScenarios.EMBEDDED_OBJECT,
            XlsxParityScenarios.LARGE_SHEET,
            XlsxParityScenarios.SIGNED_WORKBOOK,
            XlsxParityScenarios.ENCRYPTED_WORKBOOK),
        corpus.scenarios().stream().map(XlsxParityLedger.Scenario::id).toList());
    assertTrue(
        ledger.rows().stream()
            .allMatch(
                row ->
                    row.scenarioIds().stream()
                        .allMatch(id -> corpus.scenariosById().containsKey(id))));
  }

  @Test
  void everyLedgerRowMatchesItsObservedProbeOutcome() {
    XlsxParityLedger.Ledger ledger =
        XlsxParitySupport.call("load parity ledger", XlsxParityLedger::loadLedger);
    Path temporaryRoot =
        XlsxParitySupport.call(
            "create parity temporary directory",
            () -> Files.createTempDirectory("gridgrind-parity-"));
    try {
      XlsxParityProbeRegistry.ProbeContext context =
          new XlsxParityProbeRegistry.ProbeContext(temporaryRoot);
      List<ProbeOutcomeEntry> outcomes = new java.util.ArrayList<>();

      for (XlsxParityLedger.Row row : ledger.rows()) {
        XlsxParityProbeRegistry.ProbeResult outcome = outcomeFor(row.probeId(), context, outcomes);
        assertEquals(
            row.expectedOutcome(),
            outcome.outcome(),
            "%s expected %s but observed %s via %s"
                .formatted(row.id(), row.expectedOutcome(), outcome.outcome(), outcome.detail()));
      }
    } finally {
      deleteRecursively(temporaryRoot);
    }
  }

  @Test
  void everyCorpusScenarioMaterializesToAnXlsxFile() {
    XlsxParityLedger.CorpusManifest corpus =
        XlsxParitySupport.call("load parity corpus manifest", XlsxParityLedger::loadCorpus);
    Path temporaryRoot =
        XlsxParitySupport.call(
            "create parity corpus temporary directory",
            () -> Files.createTempDirectory("gridgrind-parity-corpus-"));
    try {
      for (XlsxParityLedger.Scenario scenario : corpus.scenarios()) {
        XlsxParityScenarios.MaterializedScenario materialized =
            XlsxParityScenarios.materialize(scenario.id(), temporaryRoot);
        assertTrue(Files.exists(materialized.workbookPath()), scenario.id());
        assertTrue(
            materialized.workbookPath().getFileName().toString().endsWith(".xlsx"), scenario.id());
      }
    } finally {
      deleteRecursively(temporaryRoot);
    }
  }

  private static void deleteRecursively(Path root) {
    if (!Files.exists(root)) {
      return;
    }
    XlsxParitySupport.call(
        "delete parity temporary directory " + root,
        () -> {
          try (var paths = Files.walk(root)) {
            paths.sorted(Comparator.reverseOrder()).forEach(XlsxParityTest::deletePath);
          }
          return null;
        });
  }

  private static void deletePath(Path path) {
    try {
      Files.deleteIfExists(path);
    } catch (java.io.IOException exception) {
      throw new XlsxParityException("Failed to delete " + path, exception);
    }
  }

  private static XlsxParityProbeRegistry.ProbeResult outcomeFor(
      String probeId,
      XlsxParityProbeRegistry.ProbeContext context,
      List<ProbeOutcomeEntry> outcomes) {
    return outcomes.stream()
        .filter(entry -> entry.probeId().equals(probeId))
        .findFirst()
        .map(ProbeOutcomeEntry::result)
        .orElseGet(
            () -> {
              XlsxParityProbeRegistry.ProbeResult result =
                  XlsxParityProbeRegistry.runProbe(probeId, context);
              outcomes.add(new ProbeOutcomeEntry(probeId, result));
              return result;
            });
  }

  private record ProbeOutcomeEntry(String probeId, XlsxParityProbeRegistry.ProbeResult result) {}
}
