package dev.erst.gridgrind.engine.runtime.parity;

import static dev.erst.gridgrind.engine.runtime.parity.ParityPlanSupport.inspect;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.CalculationPolicySpec;
import dev.erst.gridgrind.core.model.formula.CalculationStrategySpec;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.formula.FormulaExternalWorkbookSpec;
import dev.erst.gridgrind.core.model.formula.FormulaMissingWorkbookPolicy;
import dev.erst.gridgrind.core.model.formula.ResolvedExternalWorkbook;
import dev.erst.gridgrind.core.model.formula.ResolvedFormulaEnvironment;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import dev.erst.gridgrind.engine.runtime.DefaultGridGrindRequestExecutor;
import dev.erst.gridgrind.engine.runtime.ExecutionProgressSink;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.List;
import org.apache.poi.poifs.crypt.HashAlgorithm;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** End-to-end parity-ledger verification for the Apache POI XSSF Phase 1 harness. */
final class XlsxParityTest {
  @Test
  void ledgerAndCorpusResourcesLoadAndCoverTheKnownScenarioRegistry() {
    XlsxParityLedger.Ledger ledger =
        XlsxParitySupport.call("load parity ledger", XlsxParityLedger::loadLedger);
    XlsxParityLedger.CorpusManifest corpus =
        XlsxParitySupport.call("load parity corpus manifest", XlsxParityLedger::loadCorpus);

    String runtimePoiVersion = org.apache.poi.Version.getVersion();
    assertEquals(runtimePoiVersion, ledger.poiVersion(), "Ledger must name the tested POI runtime");
    assertEquals(runtimePoiVersion, corpus.poiVersion(), "Corpus must name the tested POI runtime");
    assertEquals(
        List.of(
            XlsxParityScenarios.CORE_WORKBOOK,
            XlsxParityScenarios.ADVANCED_NONDRAWING,
            XlsxParityScenarios.EXTERNAL_FORMULA,
            XlsxParityScenarios.UDF_FORMULA,
            XlsxParityScenarios.FORMULA_LIFECYCLE,
            XlsxParityScenarios.DRAWING_AUTHORING,
            XlsxParityScenarios.DRAWING_IMAGE,
            XlsxParityScenarios.DRAWING_COMMENTS,
            XlsxParityScenarios.DRAWING_MERGED_IMAGE,
            XlsxParityScenarios.CHART,
            XlsxParityScenarios.CHART_AUTHORING,
            XlsxParityScenarios.CHART_UNSUPPORTED,
            XlsxParityScenarios.PIVOT,
            XlsxParityScenarios.PIVOT_AUTHORING,
            XlsxParityScenarios.EMBEDDED_OBJECT,
            XlsxParityScenarios.LARGE_SHEET,
            XlsxParityScenarios.SIGNED_WORKBOOK,
            XlsxParityScenarios.INVALID_SIGNATURE_WORKBOOK,
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

  @Test
  void workbookProtectionOracleRecognizesModernWorkbookAndRevisionsHashes() {
    Path temporaryRoot =
        XlsxParitySupport.call(
            "create workbook-protection oracle temporary directory",
            () -> Files.createTempDirectory("gridgrind-parity-protection-"));
    try {
      Path workbookPath = temporaryRoot.resolve("workbook-protection.xlsx");
      XlsxParitySupport.call(
          "write workbook-protection parity workbook",
          () -> {
            try (XSSFWorkbook workbook = new XSSFWorkbook();
                OutputStream outputStream = Files.newOutputStream(workbookPath)) {
              workbook.createSheet("Protection");
              workbook.lockStructure();
              workbook.lockWindows();
              workbook.lockRevision();
              workbook.setWorkbookPassword(
                  XlsxParityScenarios.WORKBOOK_PROTECTION_PASSWORD, HashAlgorithm.sha512);
              workbook.setRevisionsPassword("gridgrind-phase3-revisions", HashAlgorithm.sha512);
              workbook.write(outputStream);
            }
            return null;
          });

      XlsxParityOracle.WorkbookProtectionSnapshot protection =
          XlsxParityOracle.workbookProtection(workbookPath);

      assertTrue(protection.structureLocked());
      assertTrue(protection.windowsLocked());
      assertTrue(protection.revisionsLocked());
      assertTrue(protection.workbookPasswordHashPresent());
      assertTrue(protection.revisionsPasswordHashPresent());
      assertTrue(protection.passwordMatches());
      assertFalse(
          XlsxParitySupport.call(
              "validate mismatched workbook-protection password",
              () -> {
                try (XSSFWorkbook workbook = new XSSFWorkbook(workbookPath.toFile())) {
                  return workbook.validateWorkbookPassword("not-the-phase-password");
                }
              }));
    } finally {
      deleteRecursively(temporaryRoot);
    }
  }

  @Test
  void externalFormulaCorpusScenarioOpensThroughGridGrindFormulaEnvironment() {
    Path temporaryRoot =
        XlsxParitySupport.call(
            "create external-formula corpus temporary directory",
            () -> Files.createTempDirectory("gridgrind-parity-external-open-"));
    try {
      XlsxParityScenarios.MaterializedScenario scenario =
          XlsxParityScenarios.materialize(XlsxParityScenarios.EXTERNAL_FORMULA, temporaryRoot);
      assertDoesNotThrow(
          () -> {
            try (XssfWorkbook workbook =
                XssfWorkbooks.open(
                    scenario.workbookPath(),
                    new ResolvedFormulaEnvironment(
                        List.of(
                            new ResolvedExternalWorkbook(
                                "referenced.xlsx", scenario.attachment("referencedWorkbook"))),
                        FormulaMissingWorkbookPolicy.ERROR,
                        List.of()),
                    XlsxParitySupport.tempFileFactory(temporaryRoot))) {
              workbook.formulas().evaluateAll();
            }
          });
    } finally {
      deleteRecursively(temporaryRoot);
    }
  }

  @Test
  void externalFormulaCorpusScenarioExecutesThroughProtocolExecutor() {
    Path temporaryRoot =
        XlsxParitySupport.call(
            "create external-formula protocol temporary directory",
            () -> Files.createTempDirectory("gridgrind-parity-external-protocol-"));
    try {
      XlsxParityScenarios.MaterializedScenario scenario =
          XlsxParityScenarios.materialize(XlsxParityScenarios.EXTERNAL_FORMULA, temporaryRoot);
      Path outputPath = temporaryRoot.resolve("output.xlsx");
      WorkbookPlan request =
          ParityPlanSupport.request(
              new WorkbookPlan.WorkbookSource.ExistingFile(scenario.workbookPath().toString()),
              new WorkbookPlan.WorkbookPersistence.SaveAs(
                  outputPath.toString(),
                  WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
                  dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec.none()),
              ExecutionPolicySpec.calculation(
                  CalculationPolicySpec.strategy(new CalculationStrategySpec.EvaluateAll())),
              new FormulaEnvironmentSpec(
                  List.of(
                      new FormulaExternalWorkbookSpec(
                          "referenced.xlsx", scenario.attachment("referencedWorkbook").toString())),
                  dev.erst.gridgrind.core.model.formula.FormulaMissingWorkbookPolicy.ERROR,
                  List.of()),
              List.of(),
              List.of(
                  inspect(
                      "cells",
                      new CellSelector.ByAddresses(
                          new SheetName("Ops"), List.of(new CellAddress("B1"))),
                      XlsxParityProbeRegistry.allFacetCellsQuery())));
      WorkbookResult response =
          new DefaultGridGrindRequestExecutor()
              .execute(
                  request,
                  XlsxParitySupport.bindings(temporaryRoot, request),
                  ExecutionProgressSink.NOOP);
      assertInstanceOf(WorkbookResult.Success.class, response);
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
