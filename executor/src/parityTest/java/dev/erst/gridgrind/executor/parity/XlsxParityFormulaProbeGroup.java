package dev.erst.gridgrind.executor.parity;

import static dev.erst.gridgrind.executor.parity.ParityPlanSupport.calculateAll;
import static dev.erst.gridgrind.executor.parity.ParityPlanSupport.calculateTargets;
import static dev.erst.gridgrind.executor.parity.ParityPlanSupport.clearFormulaCaches;
import static dev.erst.gridgrind.executor.parity.ParityPlanSupport.executionPolicy;
import static dev.erst.gridgrind.executor.parity.ParityPlanSupport.inspect;
import static dev.erst.gridgrind.executor.parity.XlsxParityProbeRegistry.*;

import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.executor.parity.XlsxParityProbeRegistry.ProbeContext;
import dev.erst.gridgrind.executor.parity.XlsxParityProbeRegistry.ProbeResult;
import java.nio.file.Path;
import java.util.List;

/** Formula-focused parity probes. */
final class XlsxParityFormulaProbeGroup {
  private XlsxParityFormulaProbeGroup() {}

  static ProbeResult probeFormulaCore(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario core =
        context.copiedScenario(XlsxParityScenarios.CORE_WORKBOOK, "formula-core-source");
    Path outputPath = context.derivedWorkbook("formula-core");
    XlsxParityGridGrind.mutateWorkbook(
        core.workbookPath(), outputPath, executionPolicy(calculateAll()), null, List.of());
    double value = XlsxParityOracle.evaluateFormula(outputPath, "Ops", "D3");
    return value == 25.0d
        ? pass("GridGrind evaluates core in-workbook formulas correctly.")
        : fail("GridGrind formula evaluation diverged on the core workbook.");
  }

  static ProbeResult probeFormulaExternal(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario external =
        context.scenario(XlsxParityScenarios.EXTERNAL_FORMULA);
    double poiValue =
        XlsxParityOracle.evaluateExternalFormula(
            external.workbookPath(), external.attachment("referencedWorkbook"));
    Path outputPath = context.derivedWorkbook("formula-external");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.mutateWorkbook(
            external.workbookPath(),
            outputPath,
            executionPolicy(calculateAll()),
            externalFormulaEnvironment(external.attachment("referencedWorkbook")),
            List.of(),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Ops", List.of("B1")),
                new InspectionQuery.GetCells()));
    dev.erst.gridgrind.contract.dto.CellReport.FormulaReport formula =
        cast(
            dev.erst.gridgrind.contract.dto.CellReport.FormulaReport.class,
            XlsxParityGridGrind.read(success, "cells", InspectionResult.CellsResult.class)
                .cells()
                .getFirst());
    dev.erst.gridgrind.contract.dto.CellReport.NumberReport evaluation =
        cast(dev.erst.gridgrind.contract.dto.CellReport.NumberReport.class, formula.evaluation());
    String cachedValue = XlsxParityOracle.cachedFormulaRawValue(outputPath, "Ops", "B1");
    return poiValue == 7.5d && evaluation.numberValue() == poiValue && "7.5".equals(cachedValue)
        ? pass("External-workbook formula bindings evaluate and persist cached results.")
        : fail(
            "External-workbook formula parity mismatch."
                + " poiValue="
                + poiValue
                + " gridValue="
                + evaluation.numberValue()
                + " cachedValue="
                + cachedValue);
  }

  static ProbeResult probeFormulaMissingWorkbookPolicy(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario external =
        context.scenario(XlsxParityScenarios.EXTERNAL_FORMULA);
    boolean poiStrictFails =
        XlsxParityOracle.externalFormulaFailsWithoutBinding(external.workbookPath());
    double poiCachedValue =
        XlsxParityOracle.evaluateExternalFormulaUsingCachedValue(external.workbookPath());
    GridGrindResponse.Failure strictFailure =
        XlsxParityGridGrind.mutateWorkbookExpectingFailure(
            external.workbookPath(), null, executionPolicy(calculateAll()), null, List.of());
    GridGrindResponse.Success cachedSuccess =
        XlsxParityGridGrind.readWorkbook(
            external.workbookPath(),
            missingWorkbookCachedValueEnvironment(),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Ops", List.of("B1")),
                new InspectionQuery.GetCells()));
    dev.erst.gridgrind.contract.dto.CellReport.FormulaReport cachedFormula =
        cast(
            dev.erst.gridgrind.contract.dto.CellReport.FormulaReport.class,
            XlsxParityGridGrind.read(cachedSuccess, "cells", InspectionResult.CellsResult.class)
                .cells()
                .getFirst());
    dev.erst.gridgrind.contract.dto.CellReport.NumberReport cachedEvaluation =
        cast(
            dev.erst.gridgrind.contract.dto.CellReport.NumberReport.class,
            cachedFormula.evaluation());
    return poiStrictFails
            && poiCachedValue == 7.5d
            && strictFailure.problem().code() == GridGrindProblemCode.MISSING_EXTERNAL_WORKBOOK
            && cachedEvaluation.numberValue() == poiCachedValue
        ? pass("Missing-workbook policy control matches POI strict and cached-value behavior.")
        : fail(
            "Missing-workbook policy parity mismatch."
                + " poiStrictFails="
                + poiStrictFails
                + " poiCachedValue="
                + poiCachedValue
                + " gridStrictCode="
                + strictFailure.problem().code()
                + " gridCachedValue="
                + cachedEvaluation.numberValue());
  }

  static ProbeResult probeFormulaUdf(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario udf =
        context.scenario(XlsxParityScenarios.UDF_FORMULA);
    double poiValue = XlsxParityOracle.evaluateUdfFormula(udf.workbookPath());
    Path outputPath = context.derivedWorkbook("formula-udf");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.mutateWorkbook(
            udf.workbookPath(),
            outputPath,
            executionPolicy(calculateAll()),
            udfFormulaEnvironment(),
            List.of(),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Ops", List.of("B1")),
                new InspectionQuery.GetCells()));
    dev.erst.gridgrind.contract.dto.CellReport.FormulaReport formula =
        cast(
            dev.erst.gridgrind.contract.dto.CellReport.FormulaReport.class,
            XlsxParityGridGrind.read(success, "cells", InspectionResult.CellsResult.class)
                .cells()
                .getFirst());
    dev.erst.gridgrind.contract.dto.CellReport.NumberReport evaluation =
        cast(dev.erst.gridgrind.contract.dto.CellReport.NumberReport.class, formula.evaluation());
    String cachedValue = XlsxParityOracle.cachedFormulaRawValue(outputPath, "Ops", "B1");
    return poiValue == 42.0d && evaluation.numberValue() == poiValue && "42.0".equals(cachedValue)
        ? pass("Template-backed UDF toolpacks evaluate and persist cached results.")
        : fail(
            "UDF-backed formula parity mismatch."
                + " poiValue="
                + poiValue
                + " gridValue="
                + evaluation.numberValue()
                + " cachedValue="
                + cachedValue);
  }

  static ProbeResult probeFormulaLifecycle(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario lifecycle =
        context.scenario(XlsxParityScenarios.FORMULA_LIFECYCLE);
    String originalB1 =
        XlsxParityOracle.cachedFormulaRawValue(lifecycle.workbookPath(), "Budget", "B1");
    String originalC1 =
        XlsxParityOracle.cachedFormulaRawValue(lifecycle.workbookPath(), "Budget", "C1");

    Path targetedPath = context.derivedWorkbook("formula-lifecycle-targeted");
    XlsxParityGridGrind.mutateWorkbook(
        lifecycle.workbookPath(),
        targetedPath,
        executionPolicy(calculateTargets(new CellSelector.QualifiedAddress("Budget", "B1"))),
        null,
        List.of());
    String targetedB1 = XlsxParityOracle.cachedFormulaRawValue(targetedPath, "Budget", "B1");
    String targetedC1 = XlsxParityOracle.cachedFormulaRawValue(targetedPath, "Budget", "C1");

    Path clearedPath = context.derivedWorkbook("formula-lifecycle-cleared");
    XlsxParityGridGrind.mutateWorkbook(
        lifecycle.workbookPath(),
        clearedPath,
        executionPolicy(clearFormulaCaches()),
        null,
        List.of());
    String clearedB1 = XlsxParityOracle.cachedFormulaRawValue(clearedPath, "Budget", "B1");
    String clearedC1 = XlsxParityOracle.cachedFormulaRawValue(clearedPath, "Budget", "C1");

    boolean hasDedicatedPolicies =
        XlsxParityGridGrind.hasCalculationStrategyType("EVALUATE_TARGETS")
            && XlsxParityGridGrind.hasCalculationStrategyType("CLEAR_CACHES_ONLY");
    return hasDedicatedPolicies
            && "4.0".equals(originalB1)
            && "6.0".equals(originalC1)
            && "8.0".equals(targetedB1)
            && "6.0".equals(targetedC1)
            && clearedB1 == null
            && clearedC1 == null
        ? pass(
            "Formula lifecycle controls support targeted evaluation and explicit cache clearing.")
        : fail(
            "Formula lifecycle parity mismatch."
                + " hasDedicatedPolicies="
                + hasDedicatedPolicies
                + " originalB1="
                + originalB1
                + " originalC1="
                + originalC1
                + " targetedB1="
                + targetedB1
                + " targetedC1="
                + targetedC1
                + " clearedB1="
                + clearedB1
                + " clearedC1="
                + clearedC1);
  }
}
