package dev.erst.gridgrind.executor.parity;

import static dev.erst.gridgrind.executor.parity.ParityPlanSupport.calculateAll;
import static dev.erst.gridgrind.executor.parity.ParityPlanSupport.calculateTargets;
import static dev.erst.gridgrind.executor.parity.ParityPlanSupport.clearFormulaCaches;
import static dev.erst.gridgrind.executor.parity.ParityPlanSupport.executionPolicy;
import static dev.erst.gridgrind.executor.parity.ParityPlanSupport.inspect;
import static dev.erst.gridgrind.executor.parity.ParityPlanSupport.mutate;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.excel.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.ExcelFillPattern;
import dev.erst.gridgrind.excel.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.excel.ExcelPrintOrientation;
import dev.erst.gridgrind.executor.parity.ParityPlanSupport.PendingMutation;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import java.util.function.Function;
import java.util.stream.Collectors;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Registry of executable parity probes keyed by stable ledger probe identifiers. */
final class XlsxParityProbeRegistry {
  private XlsxParityProbeRegistry() {}

  static ProbeResult runProbe(String probeId, ProbeContext context) {
    return switch (probeId) {
      case "probe-core-read" -> probeCoreRead(context);
      case "probe-named-range-read" -> probeNamedRangeRead(context);
      case "probe-workbook-protection-read-gap" -> probeWorkbookProtectionReadGap(context);
      case "probe-advanced-print-read-gap" -> probeAdvancedPrintReadGap(context);
      case "probe-advanced-style-read-gap" -> probeAdvancedStyleReadGap(context);
      case "probe-rich-comment-read-gap" -> probeRichCommentReadGap(context);
      case "probe-data-validation-analysis-gap" -> probeDataValidationAnalysisGap(context);
      case "probe-autofilter-criteria-gap" -> probeAutofilterCriteriaReadGap(context);
      case "probe-table-advanced-gap" -> probeTableAdvancedReadGap(context);
      case "probe-conditional-formatting-modeled-read" ->
          probeConditionalFormattingModeledRead(context);
      case "probe-conditional-formatting-unmodeled-read-gap" ->
          probeConditionalFormattingUnmodeledReadGap(context);
      case "probe-core-mutation" -> probeCoreMutation(context);
      case "probe-workbook-protection-mutation-gap" -> probeWorkbookProtectionMutationGap(context);
      case "probe-sheet-password-mutation-gap" -> probeSheetPasswordMutationGap(context);
      case "probe-sheet-copy-gap" -> probeSheetCopyGap(context);
      case "probe-advanced-print-mutation-gap" -> probeAdvancedPrintMutationGap(context);
      case "probe-advanced-style-mutation-gap" -> probeAdvancedStyleMutationGap(context);
      case "probe-rich-comment-mutation-gap" -> probeRichCommentMutationGap(context);
      case "probe-named-range-formula-mutation-gap" -> probeNamedRangeFormulaMutationGap(context);
      case "probe-autofilter-criteria-mutation-gap" -> probeAutofilterCriteriaMutationGap(context);
      case "probe-table-advanced-mutation-gap" -> probeTableAdvancedMutationGap(context);
      case "probe-conditional-formatting-advanced-mutation-gap" ->
          probeConditionalFormattingAdvancedMutationGap(context);
      case "probe-formula-core" -> probeFormulaCore(context);
      case "probe-formula-external" -> probeFormulaExternal(context);
      case "probe-formula-missing-workbook-policy" -> probeFormulaMissingWorkbookPolicy(context);
      case "probe-formula-udf" -> probeFormulaUdf(context);
      case "probe-formula-lifecycle" -> probeFormulaLifecycle(context);
      case "probe-drawing-readback" -> probeDrawingReadback(context);
      case "probe-drawing-authoring" -> probeDrawingAuthoring(context);
      case "probe-drawing-comment-coexistence" -> probeDrawingCommentCoexistence(context);
      case "probe-drawing-merged-image-preservation" ->
          probeDrawingMergedImagePreservation(context);
      case "probe-embedded-object-platform" -> probeEmbeddedObjectPlatform(context);
      case "probe-chart-readback" -> probeChartReadback(context);
      case "probe-chart-authoring" -> probeChartAuthoring(context);
      case "probe-chart-mutation" -> probeChartMutation(context);
      case "probe-chart-preservation" -> probeChartPreservation(context);
      case "probe-chart-unsupported" -> probeChartUnsupported(context);
      case "probe-pivot-readback" -> probePivotReadback(context);
      case "probe-pivot-authoring" -> probePivotAuthoring(context);
      case "probe-pivot-mutation" -> probePivotMutation(context);
      case "probe-pivot-preservation" -> probePivotPreservation(context);
      case "probe-event-model-gap" -> probeEventModelGap(context);
      case "probe-sxssf-gap" -> probeSxssfGap(context);
      case "probe-encryption-gap" -> probeEncryptionGap(context);
      case "probe-encryption-password-errors" -> probeEncryptionPasswordErrors(context);
      case "probe-signing-gap" -> probeSigningGap(context);
      case "probe-signing-invalid-signature" -> probeSigningInvalidSignature(context);
      default -> throw new IllegalArgumentException("Unknown parity probe id: " + probeId);
    };
  }

  private static ProbeResult probeCoreRead(ProbeContext context) {
    CoreReadObservation observation = readCoreObservation(context);
    List<String> mismatches = new ArrayList<>();
    compareCoreWorkbookSummary(observation, mismatches);
    compareCoreWorkbookCells(observation, mismatches);
    compareCoreWorkbookLinksAndComments(observation, mismatches);
    compareCoreWorkbookLayoutAndStructures(observation, mismatches);
    return mismatches.isEmpty()
        ? pass("Core workbook reads match the direct POI oracle.")
        : fail("Core workbook factual reads diverged: " + String.join("; ", mismatches));
  }

  private static CoreReadObservation readCoreObservation(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario scenario =
        context.scenario(XlsxParityScenarios.CORE_WORKBOOK);
    XlsxParityOracle.CoreWorkbookSnapshot direct =
        XlsxParityOracle.coreWorkbook(scenario.workbookPath());
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            scenario.workbookPath(),
            inspect(
                "summary",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetWorkbookSummary()),
            inspect(
                "ops-summary",
                new SheetSelector.ByName("Ops"),
                new InspectionQuery.GetSheetSummary()),
            inspect(
                "queue-summary",
                new SheetSelector.ByName("Queue"),
                new InspectionQuery.GetSheetSummary()),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Ops", List.of("A3", "B3", "C3", "D3")),
                new InspectionQuery.GetCells()),
            inspect(
                "merged", new SheetSelector.ByName("Ops"), new InspectionQuery.GetMergedRegions()),
            inspect(
                "links",
                new CellSelector.ByAddresses("Ops", List.of("C4")),
                new InspectionQuery.GetHyperlinks()),
            inspect(
                "comments",
                new CellSelector.ByAddresses("Ops", List.of("C4")),
                new InspectionQuery.GetComments()),
            inspect(
                "layout", new SheetSelector.ByName("Ops"), new InspectionQuery.GetSheetLayout()),
            inspect("print", new SheetSelector.ByName("Ops"), new InspectionQuery.GetPrintLayout()),
            inspect(
                "validations",
                new RangeSelector.AllOnSheet("Ops"),
                new InspectionQuery.GetDataValidations()),
            inspect(
                "formatting",
                new RangeSelector.AllOnSheet("Ops"),
                new InspectionQuery.GetConditionalFormatting()),
            inspect(
                "filters-ops",
                new SheetSelector.ByName("Ops"),
                new InspectionQuery.GetAutofilters()),
            inspect(
                "filters-queue",
                new SheetSelector.ByName("Queue"),
                new InspectionQuery.GetAutofilters()),
            inspect("tables", new TableSelector.All(), new InspectionQuery.GetTables()));

    InspectionResult.WorkbookSummaryResult summaryResult =
        XlsxParityGridGrind.read(success, "summary", InspectionResult.WorkbookSummaryResult.class);
    if (!(summaryResult.workbook()
        instanceof GridGrindResponse.WorkbookSummary.WithSheets summary)) {
      throw new IllegalStateException(
          "Core workbook summary did not surface active and selected sheets.");
    }

    return new CoreReadObservation(
        direct,
        summary,
        XlsxParityGridGrind.read(success, "ops-summary", InspectionResult.SheetSummaryResult.class),
        XlsxParityGridGrind.read(
            success, "queue-summary", InspectionResult.SheetSummaryResult.class),
        XlsxParityGridGrind.read(success, "cells", InspectionResult.CellsResult.class),
        XlsxParityGridGrind.read(success, "merged", InspectionResult.MergedRegionsResult.class),
        XlsxParityGridGrind.read(success, "links", InspectionResult.HyperlinksResult.class),
        XlsxParityGridGrind.read(success, "comments", InspectionResult.CommentsResult.class),
        XlsxParityGridGrind.read(success, "layout", InspectionResult.SheetLayoutResult.class),
        XlsxParityGridGrind.read(success, "print", InspectionResult.PrintLayoutResult.class),
        XlsxParityGridGrind.read(
            success, "validations", InspectionResult.DataValidationsResult.class),
        XlsxParityGridGrind.read(
            success, "formatting", InspectionResult.ConditionalFormattingResult.class),
        XlsxParityGridGrind.read(success, "filters-ops", InspectionResult.AutofiltersResult.class),
        XlsxParityGridGrind.read(
            success, "filters-queue", InspectionResult.AutofiltersResult.class),
        XlsxParityGridGrind.read(success, "tables", InspectionResult.TablesResult.class));
  }

  private static void compareCoreWorkbookSummary(
      CoreReadObservation observation, List<String> mismatches) {
    GridGrindResponse.WorkbookSummary.WithSheets summary = observation.summary();
    XlsxParityOracle.CoreWorkbookSnapshot direct = observation.direct();
    if (!summary.sheetNames().equals(direct.sheetNames())) {
      mismatches.add(
          "sheetNames=%s direct=%s".formatted(summary.sheetNames(), direct.sheetNames()));
    }
    if (!summary.activeSheetName().equals(direct.activeSheetName())) {
      mismatches.add(
          "activeSheet=%s direct=%s"
              .formatted(summary.activeSheetName(), direct.activeSheetName()));
    }
    if (!summary.selectedSheetNames().equals(direct.selectedSheetNames())) {
      mismatches.add(
          "selectedSheets=%s direct=%s"
              .formatted(summary.selectedSheetNames(), direct.selectedSheetNames()));
    }
    if (summary.forceFormulaRecalculationOnOpen() != direct.forceFormulaRecalculation()) {
      mismatches.add(
          "forceFormulaRecalculation=%s direct=%s"
              .formatted(
                  summary.forceFormulaRecalculationOnOpen(), direct.forceFormulaRecalculation()));
    }
    if (!(observation.opsSummary().sheet().protection()
        instanceof GridGrindResponse.SheetProtectionReport.Protected)) {
      mismatches.add("ops sheet protection did not report Protected");
    }
    if (observation.queueSummary().sheet().visibility()
        != dev.erst.gridgrind.excel.ExcelSheetVisibility.HIDDEN) {
      mismatches.add(
          "queue visibility=%s".formatted(observation.queueSummary().sheet().visibility()));
    }
  }

  private static void compareCoreWorkbookCells(
      CoreReadObservation observation, List<String> mismatches) {
    Map<String, GridGrindResponse.CellReport> cellsByAddress =
        byAddress(observation.cells().cells());
    compareCoreMergedRegion(observation, mismatches);
    compareCoreRichTextCell(cellsByAddress, mismatches);
    compareCorePrimitiveCells(cellsByAddress, mismatches);
    compareCoreFormulaCell(observation, cellsByAddress, mismatches);
  }

  private static void compareCoreMergedRegion(
      CoreReadObservation observation, List<String> mismatches) {
    if (observation.merged().mergedRegions().size() != 1) {
      mismatches.add("mergedRegionCount=%d".formatted(observation.merged().mergedRegions().size()));
    } else if (!observation
        .merged()
        .mergedRegions()
        .getFirst()
        .range()
        .equals(observation.direct().mergedRegion())) {
      mismatches.add(
          "mergedRegion=%s direct=%s"
              .formatted(
                  observation.merged().mergedRegions().getFirst().range(),
                  observation.direct().mergedRegion()));
    }
  }

  private static void compareCoreRichTextCell(
      Map<String, GridGrindResponse.CellReport> cellsByAddress, List<String> mismatches) {
    GridGrindResponse.CellReport.TextReport richTextCell =
        cast(GridGrindResponse.CellReport.TextReport.class, cellsByAddress.get("A3"));
    if (!"Ada Lovelace".equals(richTextCell.stringValue())) {
      mismatches.add("richTextCell.stringValue=%s".formatted(richTextCell.stringValue()));
    }
    if (richTextCell.richText() == null || richTextCell.richText().size() != 2) {
      mismatches.add("richText runs=%s".formatted(richTextCell.richText()));
    }
  }

  private static void compareCorePrimitiveCells(
      Map<String, GridGrindResponse.CellReport> cellsByAddress, List<String> mismatches) {
    GridGrindResponse.CellReport.NumberReport numericCell =
        cast(GridGrindResponse.CellReport.NumberReport.class, cellsByAddress.get("B3"));
    GridGrindResponse.CellReport.BooleanReport booleanCell =
        cast(GridGrindResponse.CellReport.BooleanReport.class, cellsByAddress.get("C3"));
    if (numericCell.numberValue() != 12.5d) {
      mismatches.add("numericCell=%s".formatted(numericCell.numberValue()));
    }
    if (!Boolean.TRUE.equals(booleanCell.booleanValue())) {
      mismatches.add("booleanCell=%s".formatted(booleanCell.booleanValue()));
    }
  }

  private static void compareCoreFormulaCell(
      CoreReadObservation observation,
      Map<String, GridGrindResponse.CellReport> cellsByAddress,
      List<String> mismatches) {
    GridGrindResponse.CellReport.FormulaReport formulaCell =
        cast(GridGrindResponse.CellReport.FormulaReport.class, cellsByAddress.get("D3"));
    GridGrindResponse.CellReport.NumberReport formulaEvaluation =
        cast(GridGrindResponse.CellReport.NumberReport.class, formulaCell.evaluation());
    if (!formulaCell.formula().equals(observation.direct().formulaText())) {
      mismatches.add(
          "formulaText=%s direct=%s"
              .formatted(formulaCell.formula(), observation.direct().formulaText()));
    }
    if (formulaEvaluation.numberValue() != observation.direct().formulaValue()) {
      mismatches.add(
          "formulaValue=%s direct=%s"
              .formatted(formulaEvaluation.numberValue(), observation.direct().formulaValue()));
    }
  }

  private static void compareCoreWorkbookLinksAndComments(
      CoreReadObservation observation, List<String> mismatches) {
    if (observation.links().hyperlinks().size() != 1) {
      mismatches.add("hyperlinkCount=%d".formatted(observation.links().hyperlinks().size()));
    } else if (!observation
        .links()
        .hyperlinks()
        .getFirst()
        .hyperlink()
        .equals(new HyperlinkTarget.Url(observation.direct().hyperlinkTarget()))) {
      mismatches.add(
          "hyperlink=%s direct=%s"
              .formatted(
                  observation.links().hyperlinks().getFirst().hyperlink(),
                  observation.direct().hyperlinkTarget()));
    }
    if (observation.comments().comments().size() != 1) {
      mismatches.add("commentCount=%d".formatted(observation.comments().comments().size()));
      return;
    }
    GridGrindResponse.CommentReport comment =
        observation.comments().comments().getFirst().comment();
    if (!comment.text().equals(observation.direct().commentText())) {
      mismatches.add(
          "commentText=%s direct=%s".formatted(comment.text(), observation.direct().commentText()));
    }
    if (!comment.author().equals(observation.direct().commentAuthor())) {
      mismatches.add(
          "commentAuthor=%s direct=%s"
              .formatted(comment.author(), observation.direct().commentAuthor()));
    }
    if (comment.visible() != observation.direct().commentVisible()) {
      mismatches.add(
          "commentVisible=%s direct=%s"
              .formatted(comment.visible(), observation.direct().commentVisible()));
    }
  }

  private static void compareCoreWorkbookLayoutAndStructures(
      CoreReadObservation observation, List<String> mismatches) {
    compareCoreLayout(observation, mismatches);
    compareCoreValidationsAndFormatting(observation, mismatches);
    compareCoreAutofilters(observation, mismatches);
    compareCoreTables(observation, mismatches);
  }

  private static void compareCoreLayout(CoreReadObservation observation, List<String> mismatches) {
    if (observation.layout().layout().zoomPercent() != observation.direct().zoomPercent()) {
      mismatches.add(
          "zoomPercent=%s direct=%s"
              .formatted(
                  observation.layout().layout().zoomPercent(), observation.direct().zoomPercent()));
    }
    if (!(observation.layout().layout().pane()
        instanceof dev.erst.gridgrind.contract.dto.PaneReport.Frozen)) {
      mismatches.add("pane=%s".formatted(observation.layout().layout().pane()));
    }
    if (observation.print().layout().orientation()
        != dev.erst.gridgrind.excel.ExcelPrintOrientation.LANDSCAPE) {
      mismatches.add("printOrientation=%s".formatted(observation.print().layout().orientation()));
    }
  }

  private static void compareCoreValidationsAndFormatting(
      CoreReadObservation observation, List<String> mismatches) {
    if (observation.validations().validations().size() != 1) {
      mismatches.add(
          "validationCount=%d".formatted(observation.validations().validations().size()));
    } else if (!(observation.validations().validations().getFirst()
        instanceof DataValidationEntryReport.Supported)) {
      mismatches.add(
          "validationEntry=%s".formatted(observation.validations().validations().getFirst()));
    }
    if (observation.formatting().conditionalFormattingBlocks().size()
        != observation.direct().conditionalFormattingBlockCount()) {
      mismatches.add(
          "conditionalFormattingBlocks=%s direct=%s"
              .formatted(
                  observation.formatting().conditionalFormattingBlocks().size(),
                  observation.direct().conditionalFormattingBlockCount()));
    }
  }

  private static void compareCoreAutofilters(
      CoreReadObservation observation, List<String> mismatches) {
    if (observation.filtersOps().autofilters().size() != 1) {
      mismatches.add(
          "opsAutofilterCount=%d".formatted(observation.filtersOps().autofilters().size()));
    } else if (!observation
        .filtersOps()
        .autofilters()
        .getFirst()
        .range()
        .equals(observation.direct().sheetAutofilterRange())) {
      mismatches.add(
          "opsAutofilterRange=%s direct=%s"
              .formatted(
                  observation.filtersOps().autofilters().getFirst().range(),
                  observation.direct().sheetAutofilterRange()));
    }
    if (observation.filtersQueue().autofilters().size() != 1) {
      mismatches.add(
          "queueAutofilterCount=%d".formatted(observation.filtersQueue().autofilters().size()));
    } else if (!(observation.filtersQueue().autofilters().getFirst()
        instanceof AutofilterEntryReport.TableOwned)) {
      mismatches.add(
          "queueAutofilter=%s".formatted(observation.filtersQueue().autofilters().getFirst()));
    }
  }

  private static void compareCoreTables(CoreReadObservation observation, List<String> mismatches) {
    if (observation.tables().tables().size() != 1) {
      mismatches.add("tableCount=%d".formatted(observation.tables().tables().size()));
    } else {
      if (!observation
          .tables()
          .tables()
          .getFirst()
          .name()
          .equals(observation.direct().tableName())) {
        mismatches.add(
            "tableName=%s direct=%s"
                .formatted(
                    observation.tables().tables().getFirst().name(),
                    observation.direct().tableName()));
      }
      if (!observation
          .tables()
          .tables()
          .getFirst()
          .columnNames()
          .equals(observation.direct().tableColumns())) {
        mismatches.add(
            "tableColumns=%s direct=%s"
                .formatted(
                    observation.tables().tables().getFirst().columnNames(),
                    observation.direct().tableColumns()));
      }
    }
    if (!observation.direct().dataValidationKinds().equals(List.of("WHOLE"))) {
      mismatches.add(
          "directValidationKinds=%s".formatted(observation.direct().dataValidationKinds()));
    }
  }

  private static ProbeResult probeNamedRangeRead(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario core =
        context.scenario(XlsxParityScenarios.CORE_WORKBOOK);
    XlsxParityScenarios.MaterializedScenario advanced =
        context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING);
    List<XlsxParityOracle.NamedRangeSnapshot> directCore =
        XlsxParityOracle.namedRanges(core.workbookPath());
    List<XlsxParityOracle.NamedRangeSnapshot> directAdvanced =
        XlsxParityOracle.namedRanges(advanced.workbookPath());

    InspectionResult.NamedRangesResult coreRanges =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.readWorkbook(
                core.workbookPath(),
                inspect(
                    "names",
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                    new InspectionQuery.GetNamedRanges())),
            "names",
            InspectionResult.NamedRangesResult.class);
    InspectionResult.NamedRangesResult advancedRanges =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.readWorkbook(
                advanced.workbookPath(),
                inspect(
                    "names",
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                    new InspectionQuery.GetNamedRanges())),
            "names",
            InspectionResult.NamedRangesResult.class);

    boolean coreOk =
        coreRanges.namedRanges().size() == 2
            && coreRanges.namedRanges().stream()
                .allMatch(GridGrindResponse.NamedRangeReport.RangeReport.class::isInstance);
    boolean advancedOk =
        advancedRanges.namedRanges().size() == 2
            && advancedRanges.namedRanges().stream()
                .allMatch(GridGrindResponse.NamedRangeReport.FormulaReport.class::isInstance);
    boolean directOk =
        directCore.stream()
                .filter(snapshot -> !snapshot.name().startsWith("_xlnm."))
                .noneMatch(XlsxParityOracle.NamedRangeSnapshot::formulaDefined)
            && directAdvanced.stream()
                .anyMatch(XlsxParityOracle.NamedRangeSnapshot::formulaDefined);
    return coreOk && advancedOk && directOk
        ? pass("Explicit and formula-defined named-range reads match POI-backed scenarios.")
        : fail(
            "Named-range reads did not cover both explicit and formula-defined workbook state."
                + " coreReports="
                + coreRanges.namedRanges().stream()
                    .map(report -> report.getClass().getSimpleName() + ":" + report.name())
                    .toList()
                + " advancedReports="
                + advancedRanges.namedRanges().stream()
                    .map(report -> report.getClass().getSimpleName() + ":" + report.name())
                    .toList()
                + " directCore="
                + directCore.stream()
                    .map(snapshot -> snapshot.name() + ":" + snapshot.formulaDefined())
                    .toList()
                + " directAdvanced="
                + directAdvanced.stream()
                    .map(snapshot -> snapshot.name() + ":" + snapshot.formulaDefined())
                    .toList());
  }

  private static ProbeResult probeWorkbookProtectionReadGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario advanced =
        context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING);
    XlsxParityOracle.WorkbookProtectionSnapshot direct =
        XlsxParityOracle.workbookProtection(advanced.workbookPath());
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            advanced.workbookPath(),
            inspect(
                "protection",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetWorkbookProtection()));
    InspectionResult.WorkbookProtectionResult protection =
        XlsxParityGridGrind.read(
            success, "protection", InspectionResult.WorkbookProtectionResult.class);
    WorkbookProtectionReport observed = protection.protection();
    boolean parityAchieved =
        XlsxParityGridGrind.hasInspectionQueryType("GET_WORKBOOK_PROTECTION")
            && direct.structureLocked() == observed.structureLocked()
            && direct.windowsLocked() == observed.windowsLocked()
            && direct.revisionsLocked() == observed.revisionsLocked()
            && direct.workbookPasswordHashPresent() == observed.workbookPasswordHashPresent()
            && direct.revisionsPasswordHashPresent() == observed.revisionsPasswordHashPresent();
    return parityAchieved
        ? pass("Workbook-protection read parity is present.")
        : fail(
            "Workbook-protection read mismatch." + " direct=" + direct + " observed=" + observed);
  }

  private static ProbeResult probeAdvancedPrintReadGap(ProbeContext context) {
    XlsxParityOracle.AdvancedPrintSnapshot direct =
        XlsxParityOracle.advancedPrint(
            context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING).workbookPath());
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING).workbookPath(),
            inspect(
                "print",
                new SheetSelector.ByName("Advanced"),
                new InspectionQuery.GetPrintLayout()));
    InspectionResult.PrintLayoutResult print =
        XlsxParityGridGrind.read(success, "print", InspectionResult.PrintLayoutResult.class);
    boolean parityAchieved =
        approximatelyEquals(direct.leftMargin(), print.layout().setup().margins().left())
            && approximatelyEquals(direct.rightMargin(), print.layout().setup().margins().right())
            && approximatelyEquals(direct.topMargin(), print.layout().setup().margins().top())
            && approximatelyEquals(direct.bottomMargin(), print.layout().setup().margins().bottom())
            && direct.printGridlines() == print.layout().setup().printGridlines()
            && direct.horizontallyCentered() == print.layout().setup().horizontallyCentered()
            && direct.verticallyCentered() == print.layout().setup().verticallyCentered()
            && direct.paperSize() == print.layout().setup().paperSize()
            && direct.draft() == print.layout().setup().draft()
            && direct.noColor() == print.layout().setup().blackAndWhite()
            && direct.copies() == print.layout().setup().copies()
            && direct.usePage() == print.layout().setup().useFirstPageNumber()
            && direct.pageStart() == print.layout().setup().firstPageNumber()
            && direct.rowBreaks().equals(print.layout().setup().rowBreaks())
            && direct.columnBreaks().equals(print.layout().setup().columnBreaks());
    return parityAchieved
        ? pass("Advanced print-layout read parity is present.")
        : fail(
            "Advanced print-layout read mismatch."
                + " direct="
                + direct
                + " observed="
                + print.layout().setup());
  }

  private static ProbeResult probeAdvancedStyleReadGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario advanced =
        context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING);
    XlsxParityOracle.StyleSnapshot themed =
        XlsxParityOracle.style(advanced.workbookPath(), "Advanced", "A3");
    XlsxParityOracle.StyleSnapshot gradient =
        XlsxParityOracle.style(advanced.workbookPath(), "Advanced", "A4");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            advanced.workbookPath(),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Advanced", List.of("A3", "A4")),
                new InspectionQuery.GetCells()));
    InspectionResult.CellsResult cells =
        XlsxParityGridGrind.read(success, "cells", InspectionResult.CellsResult.class);
    Map<String, GridGrindResponse.CellReport> byAddress = byAddress(cells.cells());
    GridGrindResponse.CellReport themedCell = byAddress.get("A3");
    GridGrindResponse.CellReport gradientCell = byAddress.get("A4");
    boolean themedOk = matchesThemedStyle(themedCell, themed);
    boolean gradientOk = matchesGradientStyle(gradientCell, gradient);
    return themedOk && gradientOk
        ? pass("Advanced style read parity is present.")
        : fail(
            "Advanced style read mismatch."
                + " themedDirect="
                + themed
                + " themedObserved="
                + (themedCell == null ? null : themedCell.style())
                + " gradientDirect="
                + gradient
                + " gradientObserved="
                + (gradientCell == null ? null : gradientCell.style().fill()));
  }

  private static boolean matchesThemedStyle(
      GridGrindResponse.CellReport themedCell, XlsxParityOracle.StyleSnapshot themed) {
    return themedCell != null
        && matchesColorDescriptor(
            themedCell.style().font().fontColor(), themed.fontColorDescriptor())
        && matchesColorDescriptor(
            themedCell.style().fill().foregroundColor(), themed.fillColorDescriptor())
        && matchesColorDescriptor(
            themedCell.style().border().bottom().color(), themed.borderColorDescriptor());
  }

  private static boolean matchesGradientStyle(
      GridGrindResponse.CellReport gradientCell, XlsxParityOracle.StyleSnapshot gradient) {
    return gradient.gradientFill()
        && gradientCell != null
        && gradientCell.style().fill().gradient() != null
        && approximatelyEquals(45.0d, gradientCell.style().fill().gradient().degree())
        && gradientCell.style().fill().gradient().stops().size() == 2
        && matchesColorDescriptor(
            gradientCell.style().fill().gradient().stops().get(1).color(), "theme=4|tint=0.45");
  }

  private static ProbeResult probeRichCommentReadGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario advanced =
        context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING);
    XlsxParityOracle.CommentSnapshot direct =
        XlsxParityOracle.comment(advanced.workbookPath(), "Advanced", "E2");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            advanced.workbookPath(),
            inspect(
                "comments",
                new CellSelector.ByAddresses("Advanced", List.of("E2")),
                new InspectionQuery.GetComments()));
    InspectionResult.CommentsResult comments =
        XlsxParityGridGrind.read(success, "comments", InspectionResult.CommentsResult.class);
    boolean parityAchieved =
        comments.comments().size() == 1
            && comments.comments().getFirst().comment().author().equals(direct.author())
            && comments.comments().getFirst().comment().visible() == direct.visible()
            && comments.comments().getFirst().comment().runs() != null
            && comments.comments().getFirst().comment().runs().size() == direct.runCount()
            && comments.comments().getFirst().comment().anchor() != null
            && comments.comments().getFirst().comment().anchor().firstColumn() == direct.col1()
            && comments.comments().getFirst().comment().anchor().firstRow() == direct.row1()
            && comments.comments().getFirst().comment().anchor().lastColumn() == direct.col2()
            && comments.comments().getFirst().comment().anchor().lastRow() == direct.row2();
    return parityAchieved
        ? pass("Rich comment read parity is present.")
        : fail(
            "Rich comment read mismatch."
                + " direct="
                + direct
                + " observed="
                + comments.comments());
  }

  private static ProbeResult probeDataValidationAnalysisGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario advanced =
        context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING);
    List<String> directKinds =
        XlsxParityOracle.dataValidationKinds(advanced.workbookPath(), "Advanced");
    GridGrindResponse response =
        XlsxParityGridGrind.executeReadWorkbook(
            advanced.workbookPath(),
            inspect(
                "health",
                new SheetSelector.ByNames(List.of("Advanced")),
                new InspectionQuery.AnalyzeDataValidationHealth()));
    if (response instanceof GridGrindResponse.Failure failure) {
      return fail(
          "GridGrind data-validation analysis fails on malformed POI-authored rules: "
              + failure.problem().message());
    }
    GridGrindResponse.Success success = cast(GridGrindResponse.Success.class, response);
    InspectionResult.DataValidationHealthResult health =
        XlsxParityGridGrind.read(
            success, "health", InspectionResult.DataValidationHealthResult.class);
    List<String> findingCodes =
        health.analysis().findings().stream()
            .map(finding -> finding.code().name())
            .distinct()
            .toList();
    boolean parityAchieved =
        directKinds.containsAll(List.of("EMPTY_EXPLICIT_LIST", "MISSING_FORMULA"))
            && findingCodes.contains("DATA_VALIDATION_EMPTY_EXPLICIT_LIST")
            && findingCodes.contains("DATA_VALIDATION_MALFORMED_RULE");
    return parityAchieved
        ? pass("Fine-grained data-validation diagnostics are present.")
        : fail(
            "GridGrind analysis collapses distinct malformed validation states into generic findings.");
  }

  private static ProbeResult probeAutofilterCriteriaReadGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario advanced =
        context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING);
    XlsxParityOracle.AutofilterSnapshot direct =
        XlsxParityOracle.autofilter(advanced.workbookPath(), "Advanced");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            advanced.workbookPath(),
            inspect(
                "filters",
                new SheetSelector.ByName("Advanced"),
                new InspectionQuery.GetAutofilters()));
    InspectionResult.AutofiltersResult filters =
        XlsxParityGridGrind.read(success, "filters", InspectionResult.AutofiltersResult.class);
    if (direct.filterColumnCount() == 0 || !direct.hasSortState()) {
      return fail(
          "Advanced autofilter scenario did not retain criteria/sort-state metadata: " + direct);
    }
    boolean parityAchieved =
        filters.autofilters().size() == 1
            && filters.autofilters().getFirst() instanceof AutofilterEntryReport.SheetOwned observed
            && observed.range().equals(direct.range())
            && observed.filterColumns().size() == direct.filterColumns().size()
            && observed.filterColumns().getFirst().columnId()
                == direct.filterColumns().getFirst().columnId()
            && observed.filterColumns().getFirst().criterion()
                instanceof AutofilterFilterCriterionReport.Values values
            && values.values().equals(direct.filterColumns().getFirst().values())
            && values.includeBlank() == direct.filterColumns().getFirst().includeBlank()
            && observed.sortState() != null
            && observed.sortState().range().equals(direct.sortState().range())
            && observed.sortState().conditions().size() == direct.sortState().conditions().size()
            && sortConditionsMatch(
                observed.sortState().conditions().getFirst(),
                direct.sortState().conditions().getFirst());
    return parityAchieved
        ? pass("Autofilter criteria and sort-state read parity is present.")
        : fail(
            "Autofilter criteria and sort-state read mismatch."
                + " direct="
                + direct
                + " observed="
                + filters.autofilters());
  }

  private static ProbeResult probeTableAdvancedReadGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario advanced =
        context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING);
    XlsxParityOracle.TableSnapshot direct =
        XlsxParityOracle.table(advanced.workbookPath(), "AdvancedTable");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            advanced.workbookPath(),
            inspect("tables", new TableSelector.All(), new InspectionQuery.GetTables()));
    InspectionResult.TablesResult tables =
        XlsxParityGridGrind.read(success, "tables", InspectionResult.TablesResult.class);
    TableEntryReport observed = findTable(tables.tables(), "AdvancedTable");
    boolean parityAchieved =
        observed != null
            && direct.comment().equals(observed.comment())
            && direct.published() == observed.published()
            && direct.insertRow() == observed.insertRow()
            && direct.insertRowShift() == observed.insertRowShift()
            && direct.headerRowCellStyle().equals(observed.headerRowCellStyle())
            && direct.dataCellStyle().equals(observed.dataCellStyle())
            && direct.totalsRowCellStyle().equals(observed.totalsRowCellStyle())
            && observed.columns().size() >= 3
            && direct.totalsRowLabel().equals(observed.columns().get(0).totalsRowLabel())
            && direct.totalsRowFunction().equals(observed.columns().get(1).totalsRowFunction())
            && direct
                .calculatedColumnFormula()
                .equals(observed.columns().get(1).calculatedColumnFormula())
            && direct.uniqueName().equals(observed.columns().get(2).uniqueName());
    return parityAchieved
        ? pass("Advanced table read parity is present.")
        : fail("Advanced table read mismatch." + " direct=" + direct + " observed=" + observed);
  }

  private static ProbeResult probeConditionalFormattingModeledRead(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario advanced =
        context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING);
    List<String> directKinds =
        XlsxParityOracle.conditionalFormattingKinds(advanced.workbookPath(), "Advanced");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            advanced.workbookPath(),
            inspect(
                "formatting",
                new RangeSelector.AllOnSheet("Advanced"),
                new InspectionQuery.GetConditionalFormatting()));
    InspectionResult.ConditionalFormattingResult formatting =
        XlsxParityGridGrind.read(
            success, "formatting", InspectionResult.ConditionalFormattingResult.class);
    Set<String> actualKinds =
        formatting.conditionalFormattingBlocks().stream()
            .flatMap(block -> block.rules().stream())
            .map(rule -> rule.getClass().getSimpleName())
            .collect(Collectors.toSet());
    boolean ok =
        directKinds.stream().anyMatch(kind -> "colorScale".equalsIgnoreCase(kind))
            && directKinds.stream().anyMatch(kind -> "dataBar".equalsIgnoreCase(kind))
            && directKinds.stream().anyMatch(kind -> "iconSet".equalsIgnoreCase(kind))
            && actualKinds.containsAll(Set.of("ColorScaleRule", "DataBarRule", "IconSetRule"));
    return ok
        ? pass("GridGrind reads the modeled advanced conditional-format rule families.")
        : fail("GridGrind did not read the modeled advanced conditional-format rule families.");
  }

  private static ProbeResult probeConditionalFormattingUnmodeledReadGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario advanced =
        context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING);
    List<String> directKinds =
        XlsxParityOracle.conditionalFormattingKinds(advanced.workbookPath(), "Advanced");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            advanced.workbookPath(),
            inspect(
                "formatting",
                new RangeSelector.AllOnSheet("Advanced"),
                new InspectionQuery.GetConditionalFormatting()));
    InspectionResult.ConditionalFormattingResult formatting =
        XlsxParityGridGrind.read(
            success, "formatting", InspectionResult.ConditionalFormattingResult.class);
    boolean hasUnsupportedTop10 =
        formatting.conditionalFormattingBlocks().stream()
            .flatMap(block -> block.rules().stream())
            .filter(ConditionalFormattingRuleReport.UnsupportedRule.class::isInstance)
            .map(ConditionalFormattingRuleReport.UnsupportedRule.class::cast)
            .anyMatch(rule -> "TOP_10".equals(rule.kind()));
    boolean hasTop10Rule =
        formatting.conditionalFormattingBlocks().stream()
            .flatMap(block -> block.rules().stream())
            .anyMatch(ConditionalFormattingRuleReport.Top10Rule.class::isInstance);
    return directKinds.stream().anyMatch(kind -> "top10".equalsIgnoreCase(kind))
            && hasTop10Rule
            && !hasUnsupportedTop10
        ? pass("Full conditional-format read parity is present.")
        : fail(
            "Conditional-format read mismatch."
                + " directKinds="
                + directKinds
                + " hasTop10Rule="
                + hasTop10Rule
                + " hasUnsupportedTop10="
                + hasUnsupportedTop10);
  }

  private static ProbeResult probeCoreMutation(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario scenario =
        context.scenario(XlsxParityScenarios.CORE_WORKBOOK);
    XlsxParityOracle.CoreWorkbookSnapshot direct =
        XlsxParityOracle.coreWorkbook(scenario.workbookPath());
    List<XlsxParityOracle.NamedRangeSnapshot> names =
        XlsxParityOracle.namedRanges(scenario.workbookPath()).stream()
            .filter(snapshot -> !snapshot.name().startsWith("_xlnm."))
            .toList();
    boolean ok =
        direct.sheetNames().equals(List.of("Ops", "Queue"))
            && direct.forceFormulaRecalculation()
            && direct.formulaValue() == 25.0d
            && "QueueTable".equals(direct.tableName())
            && names.stream().allMatch(snapshot -> !snapshot.formulaDefined())
            && names.size() == 2;
    return ok
        ? pass("Core GridGrind mutations materialize the expected workbook semantics.")
        : fail("Core GridGrind mutations did not produce the expected workbook semantics.");
  }

  private static ProbeResult probeWorkbookProtectionMutationGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario source =
        context.copiedScenario(
            XlsxParityScenarios.ADVANCED_NONDRAWING, "workbook-protection-mutation-source");
    XlsxParityOracle.WorkbookProtectionSnapshot expected =
        XlsxParityOracle.workbookProtection(source.workbookPath());

    Path clearedPath = context.derivedWorkbook("workbook-protection-cleared");
    GridGrindResponse.Success clearedSuccess =
        XlsxParityGridGrind.mutateWorkbook(
            source.workbookPath(),
            clearedPath,
            List.of(
                mutate(
                    new WorkbookSelector.Current(), new MutationAction.ClearWorkbookProtection())),
            inspect(
                "protection",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetWorkbookProtection()));
    XlsxParityOracle.WorkbookProtectionSnapshot clearedDirect =
        XlsxParityOracle.workbookProtection(clearedPath);
    WorkbookProtectionReport clearedObserved =
        XlsxParityGridGrind.read(
                clearedSuccess, "protection", InspectionResult.WorkbookProtectionResult.class)
            .protection();
    boolean clearedOk =
        clearedDirect.equals(
                new XlsxParityOracle.WorkbookProtectionSnapshot(
                    false, false, false, false, false, false))
            && matchesWorkbookProtectionReport(clearedObserved, clearedDirect);

    Path restoredPath = context.derivedWorkbook("workbook-protection-restored");
    GridGrindResponse.Success restoredSuccess =
        XlsxParityGridGrind.mutateWorkbook(
            clearedPath,
            restoredPath,
            List.of(
                mutate(
                    new WorkbookSelector.Current(),
                    new MutationAction.SetWorkbookProtection(
                        new WorkbookProtectionInput(
                            true,
                            true,
                            true,
                            XlsxParityScenarios.WORKBOOK_PROTECTION_PASSWORD,
                            null)))),
            inspect(
                "protection",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetWorkbookProtection()));
    XlsxParityOracle.WorkbookProtectionSnapshot restoredDirect =
        XlsxParityOracle.workbookProtection(restoredPath);
    WorkbookProtectionReport restoredObserved =
        XlsxParityGridGrind.read(
                restoredSuccess, "protection", InspectionResult.WorkbookProtectionResult.class)
            .protection();
    boolean restoredOk =
        restoredDirect.equals(expected)
            && matchesWorkbookProtectionReport(restoredObserved, expected);

    return clearedOk && restoredOk
        ? pass("Workbook-protection mutation parity is present.")
        : fail(
            "Workbook-protection mutation mismatch."
                + " expected="
                + expected
                + " cleared="
                + clearedDirect
                + " restored="
                + restoredDirect
                + " observed="
                + restoredObserved);
  }

  private static ProbeResult probeSheetPasswordMutationGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario advanced =
        context.copiedScenario(
            XlsxParityScenarios.ADVANCED_NONDRAWING, "sheet-password-mutation-source");
    InspectionResult.SheetSummaryResult expectedSummary =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.readWorkbook(
                advanced.workbookPath(),
                inspect(
                    "summary",
                    new SheetSelector.ByName("Advanced"),
                    new InspectionQuery.GetSheetSummary())),
            "summary",
            InspectionResult.SheetSummaryResult.class);
    SheetProtectionSettings expectedSettings = sheetProtectionSettings(expectedSummary);
    if (expectedSettings == null
        || !sheetPasswordMatches(
            advanced.workbookPath(), "Advanced", XlsxParityScenarios.SHEET_PROTECTION_PASSWORD)) {
      return fail("Advanced parity workbook did not retain its expected protected-sheet password.");
    }

    Path clearedPath = context.derivedWorkbook("sheet-protection-cleared");
    InspectionResult.SheetSummaryResult clearedSummary =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.mutateWorkbook(
                advanced.workbookPath(),
                clearedPath,
                List.of(
                    mutate(
                        new SheetSelector.ByName("Advanced"),
                        new MutationAction.ClearSheetProtection())),
                inspect(
                    "summary",
                    new SheetSelector.ByName("Advanced"),
                    new InspectionQuery.GetSheetSummary())),
            "summary",
            InspectionResult.SheetSummaryResult.class);
    boolean clearedOk =
        clearedSummary.sheet().protection()
                instanceof GridGrindResponse.SheetProtectionReport.Unprotected
            && !sheetPasswordMatches(
                clearedPath, "Advanced", XlsxParityScenarios.SHEET_PROTECTION_PASSWORD);

    Path restoredPath = context.derivedWorkbook("sheet-protection-restored");
    InspectionResult.SheetSummaryResult restoredSummary =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.mutateWorkbook(
                clearedPath,
                restoredPath,
                List.of(
                    mutate(
                        new SheetSelector.ByName("Advanced"),
                        new MutationAction.SetSheetProtection(
                            expectedSettings, XlsxParityScenarios.SHEET_PROTECTION_PASSWORD))),
                inspect(
                    "summary",
                    new SheetSelector.ByName("Advanced"),
                    new InspectionQuery.GetSheetSummary())),
            "summary",
            InspectionResult.SheetSummaryResult.class);
    boolean restoredOk =
        restoredSummary.sheet().protection()
                instanceof GridGrindResponse.SheetProtectionReport.Protected protectedReport
            && protectedReport.settings().equals(expectedSettings)
            && sheetPasswordMatches(
                restoredPath, "Advanced", XlsxParityScenarios.SHEET_PROTECTION_PASSWORD);

    return clearedOk && restoredOk
        ? pass("Password-bearing sheet-protection mutation parity is present.")
        : fail(
            "Sheet-protection password mutation mismatch."
                + " expectedSettings="
                + expectedSettings
                + " cleared="
                + clearedSummary.sheet().protection()
                + " restored="
                + restoredSummary.sheet().protection());
  }

  private static ProbeResult probeSheetCopyGap(ProbeContext context) {
    SheetCopySourceObservation source = readSheetCopySource(context);
    if (source.sourceTable() == null) {
      return fail("Advanced parity workbook did not retain its expected source table.");
    }
    SheetCopyCopiedObservation copied = copySheetObservation(context, source);
    List<String> mismatches = new ArrayList<>();
    compareSheetCopyWorkbook(copied, mismatches);
    compareSheetCopyComments(source, copied, mismatches);
    compareSheetCopyPrint(source, copied, mismatches);
    compareSheetCopyAutofilter(source, copied, mismatches);
    compareSheetCopyFormatting(source, copied, mismatches);
    compareSheetCopyProtection(source, copied, mismatches);
    compareSheetCopyLocalNames(copied, mismatches);
    compareSheetCopyTable(source, copied, mismatches);
    return mismatches.isEmpty()
        ? pass("Sheet-copy parity is present.")
        : fail("Sheet-copy mutation mismatch: " + String.join("; ", mismatches));
  }

  private static SheetCopySourceObservation readSheetCopySource(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario source =
        context.copiedScenario(XlsxParityScenarios.ADVANCED_NONDRAWING, "sheet-copy-source");
    InspectionResult.CommentsResult sourceComments =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.readWorkbook(
                source.workbookPath(),
                inspect(
                    "comments",
                    new CellSelector.ByAddresses("Advanced", List.of("E2")),
                    new InspectionQuery.GetComments())),
            "comments",
            InspectionResult.CommentsResult.class);
    InspectionResult.AutofiltersResult sourceAutofilters =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.readWorkbook(
                source.workbookPath(),
                inspect(
                    "filters",
                    new SheetSelector.ByName("Advanced"),
                    new InspectionQuery.GetAutofilters())),
            "filters",
            InspectionResult.AutofiltersResult.class);
    InspectionResult.ConditionalFormattingResult sourceFormatting =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.readWorkbook(
                source.workbookPath(),
                inspect(
                    "formatting",
                    new RangeSelector.AllOnSheet("Advanced"),
                    new InspectionQuery.GetConditionalFormatting())),
            "formatting",
            InspectionResult.ConditionalFormattingResult.class);
    InspectionResult.TablesResult sourceTables =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.readWorkbook(
                source.workbookPath(),
                inspect("tables", new TableSelector.All(), new InspectionQuery.GetTables())),
            "tables",
            InspectionResult.TablesResult.class);
    InspectionResult.SheetSummaryResult sourceSummary =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.readWorkbook(
                source.workbookPath(),
                inspect(
                    "summary",
                    new SheetSelector.ByName("Advanced"),
                    new InspectionQuery.GetSheetSummary())),
            "summary",
            InspectionResult.SheetSummaryResult.class);
    return new SheetCopySourceObservation(
        source,
        sourceComments,
        sourceAutofilters,
        sourceFormatting,
        sourceSummary,
        findTable(sourceTables.tables(), "AdvancedTable"),
        XlsxParityOracle.comment(source.workbookPath(), "Advanced", "E2"),
        XlsxParityOracle.advancedPrint(source.workbookPath(), "Advanced"),
        XlsxParityOracle.autofilter(source.workbookPath(), "Advanced"),
        XlsxParityOracle.table(source.workbookPath(), "AdvancedTable"));
  }

  private static SheetCopyCopiedObservation copySheetObservation(
      ProbeContext context, SheetCopySourceObservation source) {
    Path copiedPath = context.derivedWorkbook("sheet-copy-advanced");
    GridGrindResponse.Success copiedSuccess =
        XlsxParityGridGrind.mutateWorkbook(
            source.source().workbookPath(),
            copiedPath,
            List.of(
                mutate(
                    new SheetSelector.ByName("Advanced"),
                    new MutationAction.CopySheet(
                        "Advanced Replica", new SheetCopyPosition.AppendAtEnd()))),
            inspect(
                "workbook",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetWorkbookSummary()),
            inspect(
                "summary",
                new SheetSelector.ByName("Advanced Replica"),
                new InspectionQuery.GetSheetSummary()),
            inspect(
                "comments",
                new CellSelector.ByAddresses("Advanced Replica", List.of("E2")),
                new InspectionQuery.GetComments()),
            inspect(
                "filters",
                new SheetSelector.ByName("Advanced Replica"),
                new InspectionQuery.GetAutofilters()),
            inspect(
                "formatting",
                new RangeSelector.AllOnSheet("Advanced Replica"),
                new InspectionQuery.GetConditionalFormatting()),
            inspect("tables", new TableSelector.All(), new InspectionQuery.GetTables()));
    InspectionResult.WorkbookSummaryResult copiedWorkbook =
        XlsxParityGridGrind.read(
            copiedSuccess, "workbook", InspectionResult.WorkbookSummaryResult.class);
    InspectionResult.SheetSummaryResult copiedSummary =
        XlsxParityGridGrind.read(
            copiedSuccess, "summary", InspectionResult.SheetSummaryResult.class);
    InspectionResult.CommentsResult copiedComments =
        XlsxParityGridGrind.read(copiedSuccess, "comments", InspectionResult.CommentsResult.class);
    InspectionResult.AutofiltersResult copiedAutofilters =
        XlsxParityGridGrind.read(
            copiedSuccess, "filters", InspectionResult.AutofiltersResult.class);
    InspectionResult.ConditionalFormattingResult copiedFormatting =
        XlsxParityGridGrind.read(
            copiedSuccess, "formatting", InspectionResult.ConditionalFormattingResult.class);
    InspectionResult.TablesResult copiedTables =
        XlsxParityGridGrind.read(copiedSuccess, "tables", InspectionResult.TablesResult.class);

    TableEntryReport copiedTable = findTableOnSheet(copiedTables.tables(), "Advanced Replica");
    return new SheetCopyCopiedObservation(
        copiedPath,
        copiedWorkbook,
        copiedSummary,
        copiedComments,
        copiedAutofilters,
        copiedFormatting,
        copiedTable,
        XlsxParityOracle.comment(copiedPath, "Advanced Replica", "E2"),
        XlsxParityOracle.advancedPrint(copiedPath, "Advanced Replica"),
        XlsxParityOracle.autofilter(copiedPath, "Advanced Replica"),
        copiedTable == null ? null : XlsxParityOracle.table(copiedPath, copiedTable.name()));
  }

  private static void compareSheetCopyWorkbook(
      SheetCopyCopiedObservation copied, List<String> mismatches) {
    if (!(copied.copiedWorkbook().workbook()
            instanceof GridGrindResponse.WorkbookSummary.WithSheets workbook)
        || !workbook.sheetNames().contains("Advanced Replica")) {
      mismatches.add("workbookSummary=" + copied.copiedWorkbook().workbook());
    }
  }

  private static void compareSheetCopyComments(
      SheetCopySourceObservation source,
      SheetCopyCopiedObservation copied,
      List<String> mismatches) {
    if (!source.sourceComments().comments().equals(copied.copiedComments().comments())
        || !source.sourceComment().equals(copied.copiedComment())) {
      mismatches.add(
          "comments=%s copied=%s oracle=%s copiedOracle=%s"
              .formatted(
                  source.sourceComments().comments(),
                  copied.copiedComments().comments(),
                  source.sourceComment(),
                  copied.copiedComment()));
    }
  }

  private static void compareSheetCopyPrint(
      SheetCopySourceObservation source,
      SheetCopyCopiedObservation copied,
      List<String> mismatches) {
    if (!source.sourcePrint().equals(copied.copiedPrint())) {
      mismatches.add("print=%s copied=%s".formatted(source.sourcePrint(), copied.copiedPrint()));
    }
  }

  private static void compareSheetCopyAutofilter(
      SheetCopySourceObservation source,
      SheetCopyCopiedObservation copied,
      List<String> mismatches) {
    if (!source.sourceAutofilters().autofilters().equals(copied.copiedAutofilters().autofilters())
        || !source.sourceAutofilter().equals(copied.copiedAutofilter())) {
      mismatches.add(
          "autofilter=%s copied=%s oracle=%s copiedOracle=%s"
              .formatted(
                  source.sourceAutofilters().autofilters(),
                  copied.copiedAutofilters().autofilters(),
                  source.sourceAutofilter(),
                  copied.copiedAutofilter()));
    }
  }

  private static void compareSheetCopyFormatting(
      SheetCopySourceObservation source,
      SheetCopyCopiedObservation copied,
      List<String> mismatches) {
    List<String> sourceKinds =
        XlsxParityOracle.conditionalFormattingKinds(source.source().workbookPath(), "Advanced");
    List<String> copiedKinds =
        XlsxParityOracle.conditionalFormattingKinds(copied.copiedPath(), "Advanced Replica");
    if (!source
            .sourceFormatting()
            .conditionalFormattingBlocks()
            .equals(copied.copiedFormatting().conditionalFormattingBlocks())
        || !sourceKinds.equals(copiedKinds)) {
      mismatches.add(
          "formatting=%s copied=%s kinds=%s copiedKinds=%s"
              .formatted(
                  source.sourceFormatting().conditionalFormattingBlocks(),
                  copied.copiedFormatting().conditionalFormattingBlocks(),
                  sourceKinds,
                  copiedKinds));
    }
  }

  private static void compareSheetCopyProtection(
      SheetCopySourceObservation source,
      SheetCopyCopiedObservation copied,
      List<String> mismatches) {
    if (!(copied.copiedSummary().sheet().protection()
            instanceof GridGrindResponse.SheetProtectionReport.Protected protectedReport)
        || !(source.sourceSummary().sheet().protection()
            instanceof GridGrindResponse.SheetProtectionReport.Protected sourceProtected)
        || !protectedReport.settings().equals(sourceProtected.settings())
        || !sheetPasswordMatches(
            copied.copiedPath(),
            "Advanced Replica",
            XlsxParityScenarios.SHEET_PROTECTION_PASSWORD)) {
      mismatches.add(
          "protection=%s copied=%s"
              .formatted(
                  source.sourceSummary().sheet().protection(),
                  copied.copiedSummary().sheet().protection()));
    }
  }

  private static void compareSheetCopyLocalNames(
      SheetCopyCopiedObservation copied, List<String> mismatches) {
    if (!hasCopiedSheetScopedFormulaName(copied.copiedPath())) {
      mismatches.add("localNamedRanges=" + XlsxParityOracle.namedRanges(copied.copiedPath()));
    }
  }

  private static boolean hasCopiedSheetScopedFormulaName(Path copiedPath) {
    return XlsxParityOracle.namedRanges(copiedPath).stream()
        .anyMatch(
            snapshot ->
                "SheetScopedFormulaBudget".equals(snapshot.name())
                    && "Advanced Replica".equals(snapshot.scope())
                    && snapshot.formulaDefined()
                    && snapshot.refersToFormula().contains("Advanced Replica")
                    && snapshot.refersToFormula().contains("$C$2:$C$5"));
  }

  private static void compareSheetCopyTable(
      SheetCopySourceObservation source,
      SheetCopyCopiedObservation copied,
      List<String> mismatches) {
    TableEntryReport copiedTable = copied.copiedTable();
    TableEntryReport sourceTable = source.sourceTable();
    if (copiedTable == null
        || !tableMatchesReadReport(copiedTable, sourceTable, null, "Advanced Replica")
        || copiedTable.name().equals(sourceTable.name())
        || !source.sourceDirectTable().equals(copied.copiedDirectTable())) {
      mismatches.add(
          "table=%s copied=%s direct=%s copiedDirect=%s"
              .formatted(
                  sourceTable,
                  copiedTable,
                  source.sourceDirectTable(),
                  copied.copiedDirectTable()));
    }
  }

  private static ProbeResult probeAdvancedPrintMutationGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario source =
        context.copiedScenario(
            XlsxParityScenarios.ADVANCED_NONDRAWING, "advanced-print-mutation-source");
    XlsxParityOracle.AdvancedPrintSnapshot expectedDirect =
        XlsxParityOracle.advancedPrint(source.workbookPath(), "Advanced");
    InspectionResult.PrintLayoutResult expectedRead =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.readWorkbook(
                source.workbookPath(),
                inspect(
                    "print",
                    new SheetSelector.ByName("Advanced"),
                    new InspectionQuery.GetPrintLayout())),
            "print",
            InspectionResult.PrintLayoutResult.class);

    Path clearedPath = context.derivedWorkbook("advanced-print-cleared");
    InspectionResult.PrintLayoutResult clearedRead =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.mutateWorkbook(
                source.workbookPath(),
                clearedPath,
                List.of(
                    mutate(
                        new SheetSelector.ByName("Advanced"),
                        new MutationAction.ClearPrintLayout())),
                inspect(
                    "print",
                    new SheetSelector.ByName("Advanced"),
                    new InspectionQuery.GetPrintLayout())),
            "print",
            InspectionResult.PrintLayoutResult.class);
    boolean clearedOk = clearedRead.layout().equals(defaultPrintLayoutReport("Advanced"));

    Path restoredPath = context.derivedWorkbook("advanced-print-restored");
    InspectionResult.PrintLayoutResult restoredRead =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.mutateWorkbook(
                clearedPath,
                restoredPath,
                List.of(
                    mutate(
                        new SheetSelector.ByName("Advanced"),
                        new MutationAction.SetPrintLayout(advancedPrintLayoutInput()))),
                inspect(
                    "print",
                    new SheetSelector.ByName("Advanced"),
                    new InspectionQuery.GetPrintLayout())),
            "print",
            InspectionResult.PrintLayoutResult.class);
    boolean restoredOk =
        restoredRead.layout().equals(expectedRead.layout())
            && XlsxParityOracle.advancedPrint(restoredPath, "Advanced").equals(expectedDirect);

    return clearedOk && restoredOk
        ? pass("Advanced print-layout mutation parity is present.")
        : fail(
            "Advanced print-layout mutation mismatch."
                + " cleared="
                + clearedRead.layout()
                + " restored="
                + restoredRead.layout()
                + " expected="
                + expectedRead.layout());
  }

  private static ProbeResult probeAdvancedStyleMutationGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario advanced =
        context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING);
    XlsxParityScenarios.MaterializedScenario source =
        context.copiedScenario(XlsxParityScenarios.CORE_WORKBOOK, "advanced-style-mutation-source");
    XlsxParityOracle.StyleSnapshot expectedThemed =
        XlsxParityOracle.style(advanced.workbookPath(), "Advanced", "A3");
    XlsxParityOracle.StyleSnapshot expectedGradient =
        XlsxParityOracle.style(advanced.workbookPath(), "Advanced", "A4");

    Path styledPath = context.derivedWorkbook("advanced-style-authored");
    InspectionResult.CellsResult styledCells =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.mutateWorkbook(
                source.workbookPath(),
                styledPath,
                List.of(
                    mutate(
                        new RangeSelector.ByRange("Ops", "A5:A6"),
                        new MutationAction.SetRange(
                            List.of(
                                List.of(text("ThemeTintStyle")),
                                List.of(text("GradientFillStyle"))))),
                    mutate(
                        new RangeSelector.ByRange("Ops", "A5"),
                        new MutationAction.ApplyStyle(advancedThemedStyleInput())),
                    mutate(
                        new RangeSelector.ByRange("Ops", "A6"),
                        new MutationAction.ApplyStyle(advancedGradientStyleInput()))),
                inspect(
                    "cells",
                    new CellSelector.ByAddresses("Ops", List.of("A5", "A6")),
                    new InspectionQuery.GetCells())),
            "cells",
            InspectionResult.CellsResult.class);
    Map<String, GridGrindResponse.CellReport> styledByAddress = byAddress(styledCells.cells());
    boolean authoredOk =
        XlsxParityOracle.style(styledPath, "Ops", "A5").equals(expectedThemed)
            && XlsxParityOracle.style(styledPath, "Ops", "A6").equals(expectedGradient)
            && matchesThemedStyle(styledByAddress.get("A5"), expectedThemed)
            && matchesGradientStyle(styledByAddress.get("A6"), expectedGradient);

    Path clearedPath = context.derivedWorkbook("advanced-style-cleared");
    InspectionResult.CellsResult clearedCells =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.mutateWorkbook(
                styledPath,
                clearedPath,
                List.of(
                    mutate(
                        new RangeSelector.ByRange("Ops", "A5:A6"),
                        new MutationAction.ClearRange())),
                inspect(
                    "cells",
                    new CellSelector.ByAddresses("Ops", List.of("A5", "A6")),
                    new InspectionQuery.GetCells())),
            "cells",
            InspectionResult.CellsResult.class);
    boolean clearedOk =
        clearedCells.cells().stream()
            .allMatch(GridGrindResponse.CellReport.BlankReport.class::isInstance);

    return authoredOk && clearedOk
        ? pass("Advanced style mutation parity is present.")
        : fail(
            "Advanced style mutation mismatch."
                + " authoredOk="
                + authoredOk
                + " clearedOk="
                + clearedOk
                + " styled="
                + styledCells.cells()
                + " cleared="
                + clearedCells.cells());
  }

  private static ProbeResult probeRichCommentMutationGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario source =
        context.copiedScenario(
            XlsxParityScenarios.ADVANCED_NONDRAWING, "rich-comment-mutation-source");
    XlsxParityOracle.CommentSnapshot expectedDirect =
        XlsxParityOracle.comment(source.workbookPath(), "Advanced", "E2");
    InspectionResult.CommentsResult expectedRead =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.readWorkbook(
                source.workbookPath(),
                inspect(
                    "comments",
                    new CellSelector.ByAddresses("Advanced", List.of("E2")),
                    new InspectionQuery.GetComments())),
            "comments",
            InspectionResult.CommentsResult.class);

    Path clearedPath = context.derivedWorkbook("rich-comment-cleared");
    InspectionResult.CommentsResult clearedRead =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.mutateWorkbook(
                source.workbookPath(),
                clearedPath,
                List.of(
                    mutate(
                        new CellSelector.ByAddress("Advanced", "E2"),
                        new MutationAction.ClearComment())),
                inspect(
                    "comments",
                    new CellSelector.ByAddresses("Advanced", List.of("E2")),
                    new InspectionQuery.GetComments())),
            "comments",
            InspectionResult.CommentsResult.class);
    boolean clearedOk = clearedRead.comments().isEmpty();

    Path restoredPath = context.derivedWorkbook("rich-comment-restored");
    InspectionResult.CommentsResult restoredRead =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.mutateWorkbook(
                clearedPath,
                restoredPath,
                List.of(
                    mutate(
                        new CellSelector.ByAddress("Advanced", "E2"),
                        new MutationAction.SetComment(advancedCommentInput()))),
                inspect(
                    "comments",
                    new CellSelector.ByAddresses("Advanced", List.of("E2")),
                    new InspectionQuery.GetComments())),
            "comments",
            InspectionResult.CommentsResult.class);
    boolean restoredOk =
        restoredRead.comments().equals(expectedRead.comments())
            && XlsxParityOracle.comment(restoredPath, "Advanced", "E2").equals(expectedDirect);

    return clearedOk && restoredOk
        ? pass("Rich comment mutation parity is present.")
        : fail(
            "Rich comment mutation mismatch."
                + " cleared="
                + clearedRead.comments()
                + " restored="
                + restoredRead.comments()
                + " expected="
                + expectedRead.comments());
  }

  private static ProbeResult probeNamedRangeFormulaMutationGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario source =
        context.copiedScenario(
            XlsxParityScenarios.ADVANCED_NONDRAWING, "named-range-formula-mutation-source");
    List<XlsxParityOracle.NamedRangeSnapshot> expectedDirect =
        XlsxParityOracle.namedRanges(source.workbookPath()).stream()
            .filter(XlsxParityOracle.NamedRangeSnapshot::formulaDefined)
            .toList();
    InspectionResult.NamedRangesResult expectedRead =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.readWorkbook(
                source.workbookPath(),
                inspect(
                    "names",
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                    new InspectionQuery.GetNamedRanges())),
            "names",
            InspectionResult.NamedRangesResult.class);

    Path clearedPath = context.derivedWorkbook("formula-named-ranges-cleared");
    InspectionResult.NamedRangesResult clearedRead =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.mutateWorkbook(
                source.workbookPath(),
                clearedPath,
                List.of(
                    mutate(
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope(
                            "WorkbookFormulaBudget"),
                        new MutationAction.DeleteNamedRange()),
                    mutate(
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope(
                            "SheetScopedFormulaBudget", "Advanced"),
                        new MutationAction.DeleteNamedRange())),
                inspect(
                    "names",
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                    new InspectionQuery.GetNamedRanges())),
            "names",
            InspectionResult.NamedRangesResult.class);
    boolean clearedOk =
        clearedRead.namedRanges().stream()
            .noneMatch(
                namedRange ->
                    "WorkbookFormulaBudget".equals(namedRange.name())
                        || "SheetScopedFormulaBudget".equals(namedRange.name()));

    Path restoredPath = context.derivedWorkbook("formula-named-ranges-restored");
    InspectionResult.NamedRangesResult restoredRead =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.mutateWorkbook(
                clearedPath,
                restoredPath,
                List.of(
                    mutate(
                        new MutationAction.SetNamedRange(
                            "WorkbookFormulaBudget",
                            new NamedRangeScope.Workbook(),
                            new NamedRangeTarget("SUM(Advanced!$B$2:$B$5)"))),
                    mutate(
                        new MutationAction.SetNamedRange(
                            "SheetScopedFormulaBudget",
                            new NamedRangeScope.Sheet("Advanced"),
                            new NamedRangeTarget("SUM(Advanced!$C$2:$C$5)")))),
                inspect(
                    "names",
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                    new InspectionQuery.GetNamedRanges())),
            "names",
            InspectionResult.NamedRangesResult.class);
    boolean restoredOk =
        restoredRead.namedRanges().equals(expectedRead.namedRanges())
            && XlsxParityOracle.namedRanges(restoredPath).stream()
                .filter(XlsxParityOracle.NamedRangeSnapshot::formulaDefined)
                .toList()
                .equals(expectedDirect);

    return clearedOk && restoredOk
        ? pass("Formula-defined named-range mutation parity is present.")
        : fail(
            "Formula-defined named-range mutation mismatch."
                + " cleared="
                + clearedRead.namedRanges()
                + " restored="
                + restoredRead.namedRanges());
  }

  private static ProbeResult probeAutofilterCriteriaMutationGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario source =
        context.copiedScenario(
            XlsxParityScenarios.ADVANCED_NONDRAWING, "autofilter-criteria-mutation-source");
    XlsxParityOracle.AutofilterSnapshot expectedDirect =
        XlsxParityOracle.autofilter(source.workbookPath(), "Advanced");
    InspectionResult.AutofiltersResult expectedRead =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.readWorkbook(
                source.workbookPath(),
                inspect(
                    "filters",
                    new SheetSelector.ByName("Advanced"),
                    new InspectionQuery.GetAutofilters())),
            "filters",
            InspectionResult.AutofiltersResult.class);

    Path clearedPath = context.derivedWorkbook("autofilter-cleared");
    InspectionResult.AutofiltersResult clearedRead =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.mutateWorkbook(
                source.workbookPath(),
                clearedPath,
                List.of(
                    mutate(
                        new SheetSelector.ByName("Advanced"),
                        new MutationAction.ClearAutofilter())),
                inspect(
                    "filters",
                    new SheetSelector.ByName("Advanced"),
                    new InspectionQuery.GetAutofilters())),
            "filters",
            InspectionResult.AutofiltersResult.class);
    boolean clearedOk =
        clearedRead.autofilters().isEmpty()
            && XlsxParityOracle.autofilter(clearedPath, "Advanced")
                .equals(new XlsxParityOracle.AutofilterSnapshot("", List.of(), null));

    Path restoredPath = context.derivedWorkbook("autofilter-restored");
    InspectionResult.AutofiltersResult restoredRead =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.mutateWorkbook(
                clearedPath,
                restoredPath,
                List.of(
                    mutate(
                        new RangeSelector.ByRange("Advanced", "A1:C5"),
                        new MutationAction.SetAutofilter(
                            List.of(
                                new AutofilterFilterColumnInput(
                                    0L,
                                    true,
                                    new AutofilterFilterCriterionInput.Values(
                                        List.of("R1C0"), false))),
                            new AutofilterSortStateInput(
                                "A2:C5",
                                false,
                                false,
                                "",
                                List.of(
                                    new AutofilterSortConditionInput(
                                        "B2:B5", true, "", null, null)))))),
                inspect(
                    "filters",
                    new SheetSelector.ByName("Advanced"),
                    new InspectionQuery.GetAutofilters())),
            "filters",
            InspectionResult.AutofiltersResult.class);
    boolean restoredOk =
        restoredRead.autofilters().equals(expectedRead.autofilters())
            && XlsxParityOracle.autofilter(restoredPath, "Advanced").equals(expectedDirect);

    return clearedOk && restoredOk
        ? pass("Autofilter criteria mutation parity is present.")
        : fail(
            "Autofilter criteria mutation mismatch."
                + " cleared="
                + clearedRead.autofilters()
                + " restored="
                + restoredRead.autofilters()
                + " expected="
                + expectedRead.autofilters());
  }

  private static ProbeResult probeTableAdvancedMutationGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario source =
        context.copiedScenario(
            XlsxParityScenarios.ADVANCED_NONDRAWING, "advanced-table-mutation-source");
    XlsxParityOracle.TableSnapshot expectedDirect =
        XlsxParityOracle.table(source.workbookPath(), "AdvancedTable");
    InspectionResult.TablesResult expectedRead =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.readWorkbook(
                source.workbookPath(),
                inspect("tables", new TableSelector.All(), new InspectionQuery.GetTables())),
            "tables",
            InspectionResult.TablesResult.class);
    TableEntryReport expectedTable = findTable(expectedRead.tables(), "AdvancedTable");
    if (expectedTable == null) {
      return fail("Advanced parity workbook did not retain its expected advanced table.");
    }

    Path clearedPath = context.derivedWorkbook("advanced-table-cleared");
    InspectionResult.TablesResult clearedRead =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.mutateWorkbook(
                source.workbookPath(),
                clearedPath,
                List.of(
                    mutate(
                        new TableSelector.ByNameOnSheet("AdvancedTable", "Advanced"),
                        new MutationAction.DeleteTable())),
                inspect("tables", new TableSelector.All(), new InspectionQuery.GetTables())),
            "tables",
            InspectionResult.TablesResult.class);
    boolean clearedOk = findTable(clearedRead.tables(), "AdvancedTable") == null;

    Path restoredPath = context.derivedWorkbook("advanced-table-restored");
    InspectionResult.TablesResult restoredRead =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.mutateWorkbook(
                clearedPath,
                restoredPath,
                List.of(mutate(new MutationAction.SetTable(advancedTableInput()))),
                inspect("tables", new TableSelector.All(), new InspectionQuery.GetTables())),
            "tables",
            InspectionResult.TablesResult.class);
    TableEntryReport restoredTable = findTable(restoredRead.tables(), "AdvancedTable");
    XlsxParityOracle.TableSnapshot restoredDirect =
        XlsxParityOracle.table(restoredPath, "AdvancedTable");
    boolean restoredOk =
        restoredTable != null
            && tableMatchesReadReport(restoredTable, expectedTable, "AdvancedTable", "Advanced")
            && restoredDirect.equals(expectedDirect);

    return clearedOk && restoredOk
        ? pass("Advanced table mutation parity is present.")
        : fail(
            "Advanced table mutation mismatch."
                + " cleared="
                + clearedRead.tables()
                + " restored="
                + restoredRead.tables()
                + " expected="
                + expectedRead.tables()
                + " expectedDirect="
                + expectedDirect
                + " restoredDirect="
                + restoredDirect);
  }

  private static ProbeResult probeConditionalFormattingAdvancedMutationGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario source =
        context.copiedScenario(
            XlsxParityScenarios.ADVANCED_NONDRAWING,
            "advanced-conditional-formatting-mutation-source");
    InspectionResult.ConditionalFormattingResult expectedRead =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.readWorkbook(
                source.workbookPath(),
                inspect(
                    "formatting",
                    new RangeSelector.AllOnSheet("Advanced"),
                    new InspectionQuery.GetConditionalFormatting())),
            "formatting",
            InspectionResult.ConditionalFormattingResult.class);
    List<String> expectedKinds =
        XlsxParityOracle.conditionalFormattingKinds(source.workbookPath(), "Advanced");

    Path clearedPath = context.derivedWorkbook("advanced-conditional-formatting-cleared");
    InspectionResult.ConditionalFormattingResult clearedRead =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.mutateWorkbook(
                source.workbookPath(),
                clearedPath,
                List.of(
                    mutate(
                        new RangeSelector.AllOnSheet("Advanced"),
                        new MutationAction.ClearConditionalFormatting())),
                inspect(
                    "formatting",
                    new RangeSelector.AllOnSheet("Advanced"),
                    new InspectionQuery.GetConditionalFormatting())),
            "formatting",
            InspectionResult.ConditionalFormattingResult.class);
    boolean clearedOk = clearedRead.conditionalFormattingBlocks().isEmpty();

    Path restoredPath = context.derivedWorkbook("advanced-conditional-formatting-restored");
    InspectionResult.ConditionalFormattingResult restoredRead =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.mutateWorkbook(
                clearedPath,
                restoredPath,
                advancedConditionalFormattingMutations(),
                inspect(
                    "formatting",
                    new RangeSelector.AllOnSheet("Advanced"),
                    new InspectionQuery.GetConditionalFormatting())),
            "formatting",
            InspectionResult.ConditionalFormattingResult.class);
    boolean restoredOk =
        restoredRead
                .conditionalFormattingBlocks()
                .equals(expectedRead.conditionalFormattingBlocks())
            && XlsxParityOracle.conditionalFormattingKinds(restoredPath, "Advanced")
                .equals(expectedKinds);

    return clearedOk && restoredOk
        ? pass("Advanced conditional-formatting mutation parity is present.")
        : fail(
            "Advanced conditional-formatting mutation mismatch."
                + " cleared="
                + clearedRead.conditionalFormattingBlocks()
                + " restored="
                + restoredRead.conditionalFormattingBlocks()
                + " expected="
                + expectedRead.conditionalFormattingBlocks());
  }

  private static ProbeResult probeFormulaCore(ProbeContext context) {
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

  private static ProbeResult probeFormulaExternal(ProbeContext context) {
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
    GridGrindResponse.CellReport.FormulaReport formula =
        cast(
            GridGrindResponse.CellReport.FormulaReport.class,
            XlsxParityGridGrind.read(success, "cells", InspectionResult.CellsResult.class)
                .cells()
                .getFirst());
    GridGrindResponse.CellReport.NumberReport evaluation =
        cast(GridGrindResponse.CellReport.NumberReport.class, formula.evaluation());
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

  private static ProbeResult probeFormulaMissingWorkbookPolicy(ProbeContext context) {
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
    GridGrindResponse.CellReport.FormulaReport cachedFormula =
        cast(
            GridGrindResponse.CellReport.FormulaReport.class,
            XlsxParityGridGrind.read(cachedSuccess, "cells", InspectionResult.CellsResult.class)
                .cells()
                .getFirst());
    GridGrindResponse.CellReport.NumberReport cachedEvaluation =
        cast(GridGrindResponse.CellReport.NumberReport.class, cachedFormula.evaluation());
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

  private static ProbeResult probeFormulaUdf(ProbeContext context) {
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
    GridGrindResponse.CellReport.FormulaReport formula =
        cast(
            GridGrindResponse.CellReport.FormulaReport.class,
            XlsxParityGridGrind.read(success, "cells", InspectionResult.CellsResult.class)
                .cells()
                .getFirst());
    GridGrindResponse.CellReport.NumberReport evaluation =
        cast(GridGrindResponse.CellReport.NumberReport.class, formula.evaluation());
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

  private static ProbeResult probeFormulaLifecycle(ProbeContext context) {
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

  private static ProbeResult probeDrawingReadback(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario drawing =
        context.scenario(XlsxParityScenarios.DRAWING_IMAGE);
    XlsxParityOracle.DrawingSheetSnapshot direct =
        XlsxParityOracle.drawingSheet(drawing.workbookPath(), "Ops");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            drawing.workbookPath(),
            inspect(
                "drawing",
                new DrawingObjectSelector.AllOnSheet("Ops"),
                new InspectionQuery.GetDrawingObjects()),
            inspect(
                "payload",
                new DrawingObjectSelector.ByName("Ops", "OpsPicture"),
                new InspectionQuery.GetDrawingObjectPayload()));
    InspectionResult.DrawingObjectsResult drawingObjects =
        XlsxParityGridGrind.read(success, "drawing", InspectionResult.DrawingObjectsResult.class);
    InspectionResult.DrawingObjectPayloadResult payload =
        XlsxParityGridGrind.read(
            success, "payload", InspectionResult.DrawingObjectPayloadResult.class);
    if (!XlsxParityGridGrind.hasInspectionQueryType("GET_DRAWING_OBJECTS")
        || !XlsxParityGridGrind.hasInspectionQueryType("GET_DRAWING_OBJECT_PAYLOAD")) {
      return fail("GridGrind is missing the Phase 5 drawing read types.");
    }
    if (direct.objects().size() != 1
        || !(direct.objects().getFirst()
            instanceof XlsxParityOracle.PictureDrawingObjectSnapshot picture)
        || !direct.mergedRegions().isEmpty()
        || !direct.comments().isEmpty()) {
      return fail("Phase 5 drawing corpus workbook is not in the expected picture-only shape.");
    }
    if (drawingObjects.drawingObjects().size() != 1
        || !(drawingObjects.drawingObjects().getFirst()
            instanceof DrawingObjectReport.Picture reportPicture)
        || !(payload.payload() instanceof DrawingObjectPayloadReport.Picture picturePayload)) {
      return fail("GridGrind did not return the expected picture drawing reports.");
    }
    boolean matches =
        "OpsPicture".equals(reportPicture.name())
            && matchesAnchor(reportPicture.anchor(), picture.anchor())
            && picture.pictureDigest().equals(reportPicture.sha256())
            && picture.pictureDigest().equals(picturePayload.sha256());
    return matches
        ? pass("GridGrind reads existing picture-backed drawing objects and payloads losslessly.")
        : fail("GridGrind drawing readback diverged from the direct POI picture oracle.");
  }

  private static ProbeResult probeDrawingAuthoring(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario drawing =
        context.scenario(XlsxParityScenarios.DRAWING_AUTHORING);
    XlsxParityOracle.DrawingSheetSnapshot direct =
        XlsxParityOracle.drawingSheet(drawing.workbookPath(), "Ops");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            drawing.workbookPath(),
            inspect(
                "drawing",
                new DrawingObjectSelector.AllOnSheet("Ops"),
                new InspectionQuery.GetDrawingObjects()),
            inspect(
                "picture-payload",
                new DrawingObjectSelector.ByName("Ops", "OpsPicture"),
                new InspectionQuery.GetDrawingObjectPayload()),
            inspect(
                "embed-payload",
                new DrawingObjectSelector.ByName("Ops", "OpsEmbed"),
                new InspectionQuery.GetDrawingObjectPayload()));
    InspectionResult.DrawingObjectsResult drawingObjects =
        XlsxParityGridGrind.read(success, "drawing", InspectionResult.DrawingObjectsResult.class);
    InspectionResult.DrawingObjectPayloadResult picturePayload =
        XlsxParityGridGrind.read(
            success, "picture-payload", InspectionResult.DrawingObjectPayloadResult.class);
    InspectionResult.DrawingObjectPayloadResult embedPayload =
        XlsxParityGridGrind.read(
            success, "embed-payload", InspectionResult.DrawingObjectPayloadResult.class);
    if (!XlsxParityGridGrind.hasMutationActionType("SET_PICTURE")
        || !XlsxParityGridGrind.hasMutationActionType("SET_SHAPE")
        || !XlsxParityGridGrind.hasMutationActionType("SET_EMBEDDED_OBJECT")
        || !XlsxParityGridGrind.hasMutationActionType("SET_DRAWING_OBJECT_ANCHOR")
        || !XlsxParityGridGrind.hasMutationActionType("DELETE_DRAWING_OBJECT")
        || !XlsxParityGridGrind.hasInspectionQueryType("GET_DRAWING_OBJECTS")
        || !XlsxParityGridGrind.hasInspectionQueryType("GET_DRAWING_OBJECT_PAYLOAD")) {
      return fail("GridGrind is missing the Phase 5 drawing mutation or read contract.");
    }
    if (!direct.mergedRegions().isEmpty() || !direct.comments().isEmpty()) {
      return fail("GridGrind-authored Phase 5 drawing workbook contains unexpected extra state.");
    }
    List<String> directNames =
        direct.objects().stream().map(XlsxParityOracle.DirectDrawingObjectSnapshot::name).toList();
    List<String> reportedNames =
        drawingObjects.drawingObjects().stream().map(DrawingObjectReport::name).toList();
    if (!directNames.equals(List.of("OpsPicture", "OpsShape", "OpsEmbed"))
        || !directNames.equals(reportedNames)) {
      return fail(
          "GridGrind-authored drawing workbook did not persist the expected final drawing order."
              + " direct="
              + directNames
              + " reported="
              + reportedNames);
    }

    XlsxParityOracle.PictureDrawingObjectSnapshot directPicture =
        directObject(direct, "OpsPicture", XlsxParityOracle.PictureDrawingObjectSnapshot.class);
    XlsxParityOracle.ShapeDrawingObjectSnapshot directShape =
        directObject(direct, "OpsShape", XlsxParityOracle.ShapeDrawingObjectSnapshot.class);
    XlsxParityOracle.EmbeddedObjectDrawingObjectSnapshot directEmbedded =
        directObject(
            direct, "OpsEmbed", XlsxParityOracle.EmbeddedObjectDrawingObjectSnapshot.class);
    DrawingObjectReport.Picture reportPicture =
        drawingObjectReport(drawingObjects, "OpsPicture", DrawingObjectReport.Picture.class);
    DrawingObjectReport.Shape reportShape =
        drawingObjectReport(drawingObjects, "OpsShape", DrawingObjectReport.Shape.class);
    DrawingObjectReport.EmbeddedObject reportEmbedded =
        drawingObjectReport(drawingObjects, "OpsEmbed", DrawingObjectReport.EmbeddedObject.class);
    if (!(picturePayload.payload()
            instanceof DrawingObjectPayloadReport.Picture picturePayloadReport)
        || !(embedPayload.payload()
            instanceof DrawingObjectPayloadReport.EmbeddedObject embedPayloadReport)) {
      return fail("GridGrind-authored drawing payload reads returned unexpected report types.");
    }

    boolean matches =
        matchesAnchor(reportPicture.anchor(), directPicture.anchor())
            && directPicture.pictureDigest().equals(reportPicture.sha256())
            && directPicture.pictureDigest().equals(picturePayloadReport.sha256())
            && matchesAnchor(reportShape.anchor(), directShape.anchor())
            && "SIMPLE_SHAPE".equals(directShape.kind())
            && reportShape.kind() == dev.erst.gridgrind.excel.ExcelDrawingShapeKind.SIMPLE_SHAPE
            && Objects.equals(directShape.presetGeometryToken(), reportShape.presetGeometryToken())
            && Objects.equals(directShape.text(), reportShape.text())
            && matchesAnchor(reportEmbedded.anchor(), directEmbedded.anchor())
            && directEmbedded.objectDigest().equals(reportEmbedded.sha256())
            && directEmbedded.objectDigest().equals(embedPayloadReport.sha256())
            && Objects.equals(directEmbedded.fileName(), reportEmbedded.fileName());
    return matches
        ? pass(
            "GridGrind authors, mutates, persists, and rereads Phase 5 drawing objects with direct POI agreement.")
        : fail("GridGrind-authored drawing workbook diverged from the direct POI oracle.");
  }

  private static ProbeResult probeDrawingCommentCoexistence(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario drawing =
        context.copiedScenario(XlsxParityScenarios.DRAWING_COMMENTS, "drawing-comments-source");
    XlsxParityOracle.DrawingSheetSnapshot before =
        XlsxParityOracle.drawingSheet(drawing.workbookPath(), "Ops");
    Path outputPath = context.derivedWorkbook("drawing-comments-preserved");
    XlsxParityGridGrind.mutateWorkbook(
        drawing.workbookPath(),
        outputPath,
        List.of(
            mutate(
                new CellSelector.ByAddress("Ops", "B2"),
                new MutationAction.SetComment(comment("Transient", "GridGrind", false))),
            mutate(new CellSelector.ByAddress("Ops", "B2"), new MutationAction.ClearComment())));
    XlsxParityOracle.DrawingSheetSnapshot after = XlsxParityOracle.drawingSheet(outputPath, "Ops");
    return before.equals(after)
        ? pass(
            "Comment operations coexist with drawing parts without corrupting pictures or comments.")
        : fail("Comment operations changed drawing-backed or comment-backed sheet state.");
  }

  private static ProbeResult probeDrawingMergedImagePreservation(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario drawing =
        context.copiedScenario(
            XlsxParityScenarios.DRAWING_MERGED_IMAGE, "drawing-merged-image-source");
    XlsxParityOracle.DrawingSheetSnapshot before =
        XlsxParityOracle.drawingSheet(drawing.workbookPath(), "Ops");
    Path outputPath = context.derivedWorkbook("drawing-merged-image-preserved");
    XlsxParityGridGrind.mutateWorkbook(
        drawing.workbookPath(),
        outputPath,
        List.of(
            mutate(
                new CellSelector.ByAddress("Ops", "F8"),
                new MutationAction.SetCell(text("Touch")))));
    XlsxParityOracle.DrawingSheetSnapshot after = XlsxParityOracle.drawingSheet(outputPath, "Ops");
    return !before.mergedRegions().isEmpty() && before.equals(after)
        ? pass("Unrelated edits preserve pictures that coexist with merged regions.")
        : fail("GridGrind changed merged-region or picture-backed drawing state.");
  }

  private static ProbeResult probeEmbeddedObjectPlatform(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario embedded =
        context.copiedScenario(XlsxParityScenarios.EMBEDDED_OBJECT, "embedded-object-source");
    XlsxParityOracle.DrawingSheetSnapshot before =
        XlsxParityOracle.drawingSheet(embedded.workbookPath(), "Objects");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            embedded.workbookPath(),
            inspect(
                "drawing",
                new DrawingObjectSelector.AllOnSheet("Objects"),
                new InspectionQuery.GetDrawingObjects()),
            inspect(
                "payload",
                new DrawingObjectSelector.ByName("Objects", "OpsEmbed"),
                new InspectionQuery.GetDrawingObjectPayload()));
    InspectionResult.DrawingObjectsResult drawingObjects =
        XlsxParityGridGrind.read(success, "drawing", InspectionResult.DrawingObjectsResult.class);
    InspectionResult.DrawingObjectPayloadResult payload =
        XlsxParityGridGrind.read(
            success, "payload", InspectionResult.DrawingObjectPayloadResult.class);
    Path outputPath = context.derivedWorkbook("embedded-object-preserved");
    XlsxParityGridGrind.mutateWorkbook(
        embedded.workbookPath(),
        outputPath,
        List.of(
            mutate(
                new CellSelector.ByAddress("Objects", "G1"),
                new MutationAction.SetCell(text("Touch")))));
    XlsxParityOracle.DrawingSheetSnapshot after =
        XlsxParityOracle.drawingSheet(outputPath, "Objects");
    if (!XlsxParityGridGrind.hasInspectionQueryType("GET_DRAWING_OBJECTS")
        || !XlsxParityGridGrind.hasInspectionQueryType("GET_DRAWING_OBJECT_PAYLOAD")) {
      return fail("GridGrind is missing the Phase 5 embedded-object read types.");
    }
    if (before.objects().size() != 1
        || !(before.objects().getFirst()
            instanceof XlsxParityOracle.EmbeddedObjectDrawingObjectSnapshot directEmbedded)
        || !(drawingObjects.drawingObjects().getFirst()
            instanceof DrawingObjectReport.EmbeddedObject reportEmbedded)
        || !(payload.payload()
            instanceof DrawingObjectPayloadReport.EmbeddedObject payloadEmbedded)) {
      return fail("Embedded-object parity corpus did not surface the expected single OLE object.");
    }
    boolean readMatches =
        "OpsEmbed".equals(reportEmbedded.name())
            && matchesAnchor(reportEmbedded.anchor(), directEmbedded.anchor())
            && directEmbedded.objectDigest().equals(reportEmbedded.sha256())
            && directEmbedded.objectDigest().equals(payloadEmbedded.sha256())
            && Objects.equals(directEmbedded.fileName(), reportEmbedded.fileName());
    return readMatches && before.equals(after)
        ? pass("GridGrind reads and preserves embedded OLE payloads without package drift.")
        : fail("GridGrind embedded-object parity diverged from the direct POI oracle.");
  }

  private static ProbeResult probeChartReadback(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario chart = context.scenario(XlsxParityScenarios.CHART);
    List<ChartReport> directCharts = XlsxParityOracle.charts(chart.workbookPath(), "Chart");
    XlsxParityOracle.DrawingSheetSnapshot directDrawing =
        XlsxParityOracle.drawingSheet(chart.workbookPath(), "Chart");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            chart.workbookPath(),
            inspect(
                "charts", new ChartSelector.AllOnSheet("Chart"), new InspectionQuery.GetCharts()),
            inspect(
                "drawing",
                new DrawingObjectSelector.AllOnSheet("Chart"),
                new InspectionQuery.GetDrawingObjects()));
    InspectionResult.ChartsResult charts =
        XlsxParityGridGrind.read(success, "charts", InspectionResult.ChartsResult.class);
    InspectionResult.DrawingObjectsResult drawing =
        XlsxParityGridGrind.read(success, "drawing", InspectionResult.DrawingObjectsResult.class);
    if (!XlsxParityGridGrind.hasInspectionQueryType("GET_CHARTS")
        || !XlsxParityGridGrind.hasInspectionQueryType("GET_DRAWING_OBJECTS")) {
      return fail("GridGrind is missing the Phase 6 chart read types.");
    }
    if (directCharts.size() != 1
        || !hasSinglePlot(directCharts.getFirst(), ChartReport.Bar.class)
        || directDrawing.objects().size() != 1
        || !(directDrawing.objects().getFirst()
            instanceof XlsxParityOracle.ChartDrawingObjectSnapshot directDrawingChart)) {
      return fail(
          "Phase 6 simple-chart corpus workbook is not in the expected single-chart shape.");
    }
    if (charts.charts().size() != 1
        || drawing.drawingObjects().size() != 1
        || !(drawing.drawingObjects().getFirst()
            instanceof DrawingObjectReport.Chart reportChart)) {
      return fail("GridGrind did not return the expected chart readback shape.");
    }
    boolean matches =
        charts.charts().equals(directCharts)
            && chartDrawingMatches(reportChart, directDrawingChart)
            && drawing.drawingObjects().size() == 1;
    return matches
        ? pass(
            "GridGrind reads supported POI-authored simple charts and chart drawing inventory losslessly.")
        : fail("GridGrind chart readback diverged from the direct POI oracle.");
  }

  private static ProbeResult probeChartAuthoring(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario chart =
        context.scenario(XlsxParityScenarios.CHART_AUTHORING);
    List<ChartReport> directCharts = XlsxParityOracle.charts(chart.workbookPath(), "Chart");
    XlsxParityOracle.DrawingSheetSnapshot directDrawing =
        XlsxParityOracle.drawingSheet(chart.workbookPath(), "Chart");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            chart.workbookPath(),
            inspect(
                "charts", new ChartSelector.AllOnSheet("Chart"), new InspectionQuery.GetCharts()),
            inspect(
                "drawing",
                new DrawingObjectSelector.AllOnSheet("Chart"),
                new InspectionQuery.GetDrawingObjects()));
    InspectionResult.ChartsResult charts =
        XlsxParityGridGrind.read(success, "charts", InspectionResult.ChartsResult.class);
    InspectionResult.DrawingObjectsResult drawing =
        XlsxParityGridGrind.read(success, "drawing", InspectionResult.DrawingObjectsResult.class);
    if (!XlsxParityGridGrind.hasMutationActionType("SET_CHART")
        || !XlsxParityGridGrind.hasInspectionQueryType("GET_CHARTS")
        || !XlsxParityGridGrind.hasInspectionQueryType("GET_DRAWING_OBJECTS")) {
      return fail("GridGrind is missing the Phase 6 chart authoring contract.");
    }
    if (directCharts.size() != 1
        || !hasSinglePlot(directCharts.getFirst(), ChartReport.Bar.class)
        || directDrawing.objects().size() != 1
        || !(directDrawing.objects().getFirst()
            instanceof XlsxParityOracle.ChartDrawingObjectSnapshot directDrawingChart)) {
      return fail(
          "GridGrind-authored chart corpus workbook is not in the expected single-bar-chart shape.");
    }
    ChartReport directChart = directCharts.getFirst();
    ChartReport.Bar directChartPlot = onlyPlot(directChart, ChartReport.Bar.class);
    if (directChartPlot.series().size() != 2
        || !(directChartPlot.series().get(1).categories()
            instanceof ChartReport.DataSource.StringReference categories)
        || !(directChartPlot.series().get(1).values()
            instanceof ChartReport.DataSource.NumericReference values)
        || !"ChartCategories".equals(categories.formula())
        || !"ChartActual".equals(values.formula())) {
      return fail(
          "GridGrind-authored chart corpus workbook did not retain named-range-backed series binding.");
    }
    if (charts.charts().size() != 1
        || drawing.drawingObjects().size() != 1
        || !(drawing.drawingObjects().getFirst()
            instanceof DrawingObjectReport.Chart reportChart)) {
      return fail("GridGrind-authored chart workbook did not surface the expected chart reports.");
    }
    boolean matches =
        charts.charts().equals(directCharts)
            && chartDrawingMatches(reportChart, directDrawingChart)
            && drawing.drawingObjects().size() == 1;
    return matches
        ? pass(
            "GridGrind authors named-range-backed simple charts and rereads them with direct POI agreement.")
        : fail("GridGrind-authored chart workbook diverged from the direct POI oracle.");
  }

  private static ProbeResult probeChartMutation(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario chart =
        context.copiedScenario(XlsxParityScenarios.CHART, "chart-mutation-source");
    Path outputPath = context.derivedWorkbook("chart-mutated");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.mutateWorkbook(
            chart.workbookPath(),
            outputPath,
            List.of(
                mutate(
                    new CellSelector.ByAddress("Chart", "F1"),
                    new MutationAction.SetCell(text("Touch"))),
                mutate(
                    new SheetSelector.ByName("Chart"),
                    new MutationAction.SetChart(
                        new ChartInput(
                            "OpsChart",
                            twoCellAnchorInput(
                                4, 1, 11, 17, ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                            chartTitle("Actual focus"),
                            new ChartInput.Legend.Hidden(),
                            ExcelChartDisplayBlanksAs.ZERO,
                            true,
                            List.of(
                                new ChartInput.Bar(
                                    false,
                                    ExcelChartBarDirection.BAR,
                                    null,
                                    null,
                                    null,
                                    null,
                                    List.of(
                                        new ChartInput.Series(
                                            new ChartInput.Title.Formula("C1"),
                                            new ChartInput.DataSource.Reference("A2:A4"),
                                            new ChartInput.DataSource.Reference("C2:C4"),
                                            null,
                                            null,
                                            null,
                                            null)))))))),
            inspect(
                "charts", new ChartSelector.AllOnSheet("Chart"), new InspectionQuery.GetCharts()),
            inspect(
                "drawing",
                new DrawingObjectSelector.AllOnSheet("Chart"),
                new InspectionQuery.GetDrawingObjects()));
    List<ChartReport> directCharts = XlsxParityOracle.charts(outputPath, "Chart");
    XlsxParityOracle.DrawingSheetSnapshot directDrawing =
        XlsxParityOracle.drawingSheet(outputPath, "Chart");
    InspectionResult.ChartsResult charts =
        XlsxParityGridGrind.read(success, "charts", InspectionResult.ChartsResult.class);
    InspectionResult.DrawingObjectsResult drawing =
        XlsxParityGridGrind.read(success, "drawing", InspectionResult.DrawingObjectsResult.class);
    if (directCharts.size() != 1
        || !hasSinglePlot(directCharts.getFirst(), ChartReport.Bar.class)
        || directDrawing.objects().size() != 1
        || !(directDrawing.objects().getFirst()
            instanceof XlsxParityOracle.ChartDrawingObjectSnapshot directDrawingChart)
        || charts.charts().size() != 1
        || drawing.drawingObjects().size() != 1
        || !(drawing.drawingObjects().getFirst()
            instanceof DrawingObjectReport.Chart reportChart)) {
      return fail(
          "Chart mutation parity did not produce the expected single updated supported chart.");
    }
    ChartReport directChart = directCharts.getFirst();
    ChartReport.Bar directChartPlot = onlyPlot(directChart, ChartReport.Bar.class);
    boolean matches =
        charts.charts().equals(directCharts)
            && chartDrawingMatches(reportChart, directDrawingChart)
            && directChartPlot.barDirection() == ExcelChartBarDirection.BAR
            && directChart.plotOnlyVisibleCells()
            && !directChartPlot.varyColors()
            && directChart.legend() instanceof ChartReport.Legend.Hidden
            && directChart.title().equals(new ChartReport.Title.Text("Actual focus"));
    return matches
        ? pass(
            "GridGrind mutates existing simple charts after workbook-core edits with direct POI agreement.")
        : fail("GridGrind simple-chart mutation diverged from the direct POI oracle.");
  }

  private static ProbeResult probeChartPreservation(ProbeContext context) {
    List<String> mismatches = new ArrayList<>();
    for (String scenarioId :
        List.of(
            XlsxParityScenarios.CHART,
            XlsxParityScenarios.CHART_AUTHORING,
            XlsxParityScenarios.CHART_UNSUPPORTED)) {
      XlsxParityScenarios.MaterializedScenario chart =
          context.copiedScenario(scenarioId, "chart-preservation-" + scenarioId);
      List<ChartReport> beforeCharts = XlsxParityOracle.charts(chart.workbookPath(), "Chart");
      XlsxParityOracle.DrawingSheetSnapshot beforeDrawing =
          XlsxParityOracle.drawingSheet(chart.workbookPath(), "Chart");
      Path outputPath = context.derivedWorkbook("chart-preserved-" + scenarioId.replace('-', '_'));
      XlsxParityGridGrind.mutateWorkbook(
          chart.workbookPath(),
          outputPath,
          List.of(
              mutate(
                  new CellSelector.ByAddress("Chart", "F8"),
                  new MutationAction.SetCell(text("Touch")))));
      List<ChartReport> afterCharts = XlsxParityOracle.charts(outputPath, "Chart");
      XlsxParityOracle.DrawingSheetSnapshot afterDrawing =
          XlsxParityOracle.drawingSheet(outputPath, "Chart");
      if (!beforeCharts.equals(afterCharts) || !beforeDrawing.equals(afterDrawing)) {
        mismatches.add(scenarioId);
      }
    }
    return mismatches.isEmpty()
        ? pass(
            "All Phase 6 chart corpus workbooks preserve chart relations and anchors across unrelated edits.")
        : fail("Unrelated GridGrind edits changed chart state for scenarios: " + mismatches);
  }

  private static ProbeResult probeChartUnsupported(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario chart =
        context.copiedScenario(XlsxParityScenarios.CHART_UNSUPPORTED, "chart-unsupported-source");
    List<ChartReport> directCharts = XlsxParityOracle.charts(chart.workbookPath(), "Chart");
    XlsxParityOracle.DrawingSheetSnapshot directDrawing =
        XlsxParityOracle.drawingSheet(chart.workbookPath(), "Chart");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            chart.workbookPath(),
            inspect(
                "charts", new ChartSelector.AllOnSheet("Chart"), new InspectionQuery.GetCharts()),
            inspect(
                "drawing",
                new DrawingObjectSelector.AllOnSheet("Chart"),
                new InspectionQuery.GetDrawingObjects()));
    InspectionResult.ChartsResult charts =
        XlsxParityGridGrind.read(success, "charts", InspectionResult.ChartsResult.class);
    InspectionResult.DrawingObjectsResult drawing =
        XlsxParityGridGrind.read(success, "drawing", InspectionResult.DrawingObjectsResult.class);
    Path replacementPath = context.derivedWorkbook("chart-combo-replaced");
    XlsxParityGridGrind.mutateWorkbook(
        chart.workbookPath(),
        replacementPath,
        List.of(
            mutate(
                new SheetSelector.ByName("Chart"),
                new MutationAction.SetChart(
                    new ChartInput(
                        "ComboChart",
                        twoCellAnchorInput(
                            4, 1, 11, 16, ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                        chartTitle("Roadmap"),
                        new ChartInput.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
                        ExcelChartDisplayBlanksAs.SPAN,
                        false,
                        List.of(
                            new ChartInput.Bar(
                                true,
                                ExcelChartBarDirection.COLUMN,
                                null,
                                null,
                                null,
                                null,
                                List.of(
                                    new ChartInput.Series(
                                        new ChartInput.Title.Formula("B1"),
                                        new ChartInput.DataSource.Reference("A2:A4"),
                                        new ChartInput.DataSource.Reference("B2:B4"),
                                        null,
                                        null,
                                        null,
                                        null)))))))));
    if (directCharts.size() != 1
        || directCharts.getFirst().plots().size() != 2
        || directDrawing.objects().size() != 1
        || !(directDrawing.objects().getFirst()
            instanceof XlsxParityOracle.ChartDrawingObjectSnapshot directDrawingChart)
        || charts.charts().size() != 1
        || drawing.drawingObjects().size() != 1
        || !(drawing.drawingObjects().getFirst()
            instanceof DrawingObjectReport.Chart reportChart)) {
      return fail(
          "Combo chart corpus workbook did not surface the expected multi-plot chart shape.");
    }
    InspectionResult.ChartsResult replacedCharts =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.readWorkbook(
                replacementPath,
                inspect(
                    "charts",
                    new ChartSelector.AllOnSheet("Chart"),
                    new InspectionQuery.GetCharts())),
            "charts",
            InspectionResult.ChartsResult.class);
    boolean matches =
        charts.charts().equals(directCharts)
            && chartDrawingMatches(reportChart, directDrawingChart)
            && plotTypeTokens(directCharts.getFirst()).equals(List.of("BAR", "LINE"))
            && replacedCharts.charts().size() == 1
            && hasSinglePlot(replacedCharts.charts().getFirst(), ChartReport.Bar.class);
    return matches
        ? pass(
            "Combo charts are surfaced truthfully, preserved losslessly, and can be authoritatively replaced.")
        : fail("Combo-chart handling diverged from the direct POI oracle or replacement contract.");
  }

  private static ProbeResult probePivotPreservation(ProbeContext context) {
    List<String> mismatches = new ArrayList<>();
    for (String scenarioId :
        List.of(XlsxParityScenarios.PIVOT, XlsxParityScenarios.PIVOT_AUTHORING)) {
      XlsxParityScenarios.MaterializedScenario pivot =
          context.copiedScenario(scenarioId, "pivot-preservation-" + scenarioId);
      List<PivotTableReport> before = XlsxParityOracle.pivotTables(pivot.workbookPath());
      Path outputPath = context.derivedWorkbook("pivot-preserved-" + scenarioId.replace('-', '_'));
      XlsxParityGridGrind.mutateWorkbook(
          pivot.workbookPath(),
          outputPath,
          List.of(
              mutate(new SheetSelector.ByName("Scratch"), new MutationAction.EnsureSheet()),
              mutate(
                  new CellSelector.ByAddress("Scratch", "A1"),
                  new MutationAction.SetCell(text("Touch")))));
      List<PivotTableReport> after = XlsxParityOracle.pivotTables(outputPath);
      if (!before.equals(after)) {
        mismatches.add(scenarioId);
      }
    }
    return mismatches.isEmpty()
        ? pass(
            "Phase 7 pivot corpus workbooks preserve pivot parts and sources across unrelated edits.")
        : fail("Unrelated GridGrind edits changed pivot state for scenarios: " + mismatches);
  }

  private static ProbeResult probePivotReadback(ProbeContext context) {
    if (!XlsxParityGridGrind.hasInspectionQueryType("GET_PIVOT_TABLES")
        || !XlsxParityGridGrind.hasInspectionQueryType("ANALYZE_PIVOT_TABLE_HEALTH")) {
      return fail("GridGrind is missing the Phase 7 pivot read contract.");
    }
    XlsxParityScenarios.MaterializedScenario pivot = context.scenario(XlsxParityScenarios.PIVOT);
    List<PivotTableReport> directPivots = XlsxParityOracle.pivotTables(pivot.workbookPath());
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            pivot.workbookPath(),
            inspect("pivots", new PivotTableSelector.All(), new InspectionQuery.GetPivotTables()),
            inspect(
                "health",
                new PivotTableSelector.All(),
                new InspectionQuery.AnalyzePivotTableHealth()));
    InspectionResult.PivotTablesResult pivots =
        XlsxParityGridGrind.read(success, "pivots", InspectionResult.PivotTablesResult.class);
    InspectionResult.PivotTableHealthResult health =
        XlsxParityGridGrind.read(success, "health", InspectionResult.PivotTableHealthResult.class);
    if (directPivots.size() != 1
        || !(directPivots.getFirst() instanceof PivotTableReport.Supported supported)
        || !(supported.source() instanceof PivotTableReport.Source.Range)
        || !supported.rowLabels().equals(List.of(new PivotTableReport.Field(0, "Region")))) {
      return fail(
          "Direct POI pivot corpus workbook is not in the expected supported single-pivot shape.");
    }
    boolean matches = pivots.pivotTables().equals(directPivots) && pivotHealthClean(health, 1);
    return matches
        ? pass("GridGrind reads supported POI-authored pivot tables with direct POI agreement.")
        : fail("GridGrind pivot readback diverged from the direct POI oracle.");
  }

  private static ProbeResult probePivotAuthoring(ProbeContext context) {
    if (!XlsxParityGridGrind.hasMutationActionType("SET_PIVOT_TABLE")
        || !XlsxParityGridGrind.hasInspectionQueryType("GET_PIVOT_TABLES")
        || !XlsxParityGridGrind.hasInspectionQueryType("ANALYZE_PIVOT_TABLE_HEALTH")) {
      return fail("GridGrind is missing the Phase 7 pivot authoring contract.");
    }
    XlsxParityScenarios.MaterializedScenario pivot =
        context.scenario(XlsxParityScenarios.PIVOT_AUTHORING);
    List<PivotTableReport> directPivots = XlsxParityOracle.pivotTables(pivot.workbookPath());
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            pivot.workbookPath(),
            inspect("pivots", new PivotTableSelector.All(), new InspectionQuery.GetPivotTables()),
            inspect(
                "health",
                new PivotTableSelector.All(),
                new InspectionQuery.AnalyzePivotTableHealth()));
    InspectionResult.PivotTablesResult pivots =
        XlsxParityGridGrind.read(success, "pivots", InspectionResult.PivotTablesResult.class);
    InspectionResult.PivotTableHealthResult health =
        XlsxParityGridGrind.read(success, "health", InspectionResult.PivotTableHealthResult.class);
    if (directPivots.size() != 3
        || supportedPivotByName(directPivots, "Sales Pivot") == null
        || !(supportedPivotByName(directPivots, "Named Pivot").source()
            instanceof PivotTableReport.Source.NamedRange)
        || !(supportedPivotByName(directPivots, "Table Pivot").source()
            instanceof PivotTableReport.Source.Table)) {
      return fail(
          "GridGrind-authored pivot corpus workbook did not retain the expected range, named-range, and table-backed pivot shapes.");
    }
    boolean matches =
        pivots.pivotTables().equals(directPivots) && pivotHealthClean(health, directPivots.size());
    return matches
        ? pass(
            "GridGrind authors range-, named-range-, and table-backed pivot tables and rereads them with direct POI agreement.")
        : fail("GridGrind-authored pivot workbook diverged from the direct POI oracle.");
  }

  private static ProbeResult probePivotMutation(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario pivot =
        context.copiedScenario(XlsxParityScenarios.PIVOT_AUTHORING, "pivot-mutation-source");
    Path outputPath = context.derivedWorkbook("pivot-mutated");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.mutateWorkbook(
            pivot.workbookPath(),
            outputPath,
            List.of(
                mutate(
                    new MutationAction.SetPivotTable(
                        new PivotTableInput(
                            "Sales Pivot",
                            "RangeReport",
                            new PivotTableInput.Source.Range("Data", "A1:D5"),
                            new PivotTableInput.Anchor("D6"),
                            List.of("Stage"),
                            List.of("Region"),
                            List.of(),
                            List.of(
                                new PivotTableInput.DataField(
                                    "Amount",
                                    dev.erst.gridgrind.excel.ExcelPivotDataConsolidateFunction.SUM,
                                    "Total Amount",
                                    "#,##0.00"))))),
                mutate(
                    new PivotTableSelector.ByNameOnSheet("Table Pivot", "TableReport"),
                    new MutationAction.DeletePivotTable())),
            inspect("pivots", new PivotTableSelector.All(), new InspectionQuery.GetPivotTables()),
            inspect(
                "health",
                new PivotTableSelector.All(),
                new InspectionQuery.AnalyzePivotTableHealth()));
    List<PivotTableReport> directPivots = XlsxParityOracle.pivotTables(outputPath);
    InspectionResult.PivotTablesResult pivots =
        XlsxParityGridGrind.read(success, "pivots", InspectionResult.PivotTablesResult.class);
    InspectionResult.PivotTableHealthResult health =
        XlsxParityGridGrind.read(success, "health", InspectionResult.PivotTableHealthResult.class);
    PivotTableReport.Supported salesPivot = supportedPivotByName(directPivots, "Sales Pivot");
    if (directPivots.size() != 2
        || salesPivot == null
        || !"D6".equals(salesPivot.anchor().topLeftAddress())
        || !salesPivot.rowLabels().equals(List.of(new PivotTableReport.Field(1, "Stage")))
        || !salesPivot.columnLabels().equals(List.of(new PivotTableReport.Field(0, "Region")))
        || supportedPivotByName(directPivots, "Table Pivot") != null) {
      return fail(
          "Pivot mutation parity did not produce the expected updated supported pivot set.");
    }
    boolean matches =
        pivots.pivotTables().equals(directPivots) && pivotHealthClean(health, directPivots.size());
    return matches
        ? pass("GridGrind mutates and deletes supported pivot tables with direct POI agreement.")
        : fail("GridGrind pivot mutation diverged from the direct POI oracle.");
  }

  private static boolean pivotHealthClean(
      InspectionResult.PivotTableHealthResult health, int expectedPivotCount) {
    return health.analysis().checkedPivotTableCount() == expectedPivotCount
        && health.analysis().summary().totalCount() == 0
        && health.analysis().findings().isEmpty();
  }

  private static PivotTableReport.Supported supportedPivotByName(
      List<PivotTableReport> pivots, String name) {
    return pivots.stream()
        .filter(PivotTableReport.Supported.class::isInstance)
        .map(PivotTableReport.Supported.class::cast)
        .filter(pivot -> pivot.name().equals(name))
        .findFirst()
        .orElse(null);
  }

  private static ProbeResult probeEventModelGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario largeSheet =
        context.scenario(XlsxParityScenarios.LARGE_SHEET);
    int poiRows = XlsxParityOracle.eventModelRowCount(largeSheet.workbookPath());
    GridGrindResponse.Success full =
        XlsxParityGridGrind.readWorkbook(
            largeSheet.workbookPath(),
            inspect(
                "workbook",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetWorkbookSummary()),
            inspect(
                "sheet", new SheetSelector.ByName("Large"), new InspectionQuery.GetSheetSummary()));
    GridGrindResponse response =
        XlsxParityGridGrind.executeReadWorkbook(
            largeSheet.workbookPath(),
            new ExecutionModeInput(
                ExecutionModeInput.ReadMode.EVENT_READ, ExecutionModeInput.WriteMode.FULL_XSSF),
            inspect(
                "workbook",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetWorkbookSummary()),
            inspect(
                "sheet", new SheetSelector.ByName("Large"), new InspectionQuery.GetSheetSummary()));
    if (response instanceof GridGrindResponse.Failure failure) {
      return fail("GridGrind event-read request failed: " + failure.problem().message());
    }
    GridGrindResponse.Success event = (GridGrindResponse.Success) response;
    InspectionResult.SheetSummaryResult sheetSummary =
        XlsxParityGridGrind.read(event, "sheet", InspectionResult.SheetSummaryResult.class);
    if (sheetSummary.sheet().physicalRowCount() != poiRows
        || sheetSummary.sheet().lastColumnIndex() != 19) {
      return fail(
          "GridGrind event-read sheet summary diverged from the large-sheet oracle: rows="
              + sheetSummary.sheet().physicalRowCount()
              + ", lastColumnIndex="
              + sheetSummary.sheet().lastColumnIndex());
    }
    return full.inspections().equals(event.inspections())
        ? pass("Event-model parity is present and matches full-XSSF summary reads.")
        : fail("GridGrind event-read summaries diverged from the full-XSSF summaries.");
  }

  private static ProbeResult probeSxssfGap(ProbeContext context) {
    Path streamedWorkbook = XlsxParityOracle.sxssfWriteWorkbook(context.derivedDirectory("sxssf"));
    boolean poiSucceeded =
        XlsxParitySupport.call(
            "inspect streamed parity workbook size",
            () -> Files.exists(streamedWorkbook) && Files.size(streamedWorkbook) > 0);
    Path gridGrindWorkbook = context.derivedWorkbook("streaming-gridgrind");
    List<PendingMutation> operations = new ArrayList<>();
    operations.add(mutate(new SheetSelector.ByName("Streamed"), new MutationAction.EnsureSheet()));
    for (int rowIndex = 0; rowIndex < 1_500; rowIndex++) {
      operations.add(
          mutate(
              new SheetSelector.ByName("Streamed"),
              new MutationAction.AppendRow(
                  List.of(text("R" + rowIndex), new CellInput.Numeric((double) rowIndex)))));
    }
    GridGrindResponse.Success streamed =
        XlsxParityGridGrind.writeNewWorkbook(
            gridGrindWorkbook,
            new ExecutionModeInput(
                ExecutionModeInput.ReadMode.FULL_XSSF,
                ExecutionModeInput.WriteMode.STREAMING_WRITE),
            operations,
            inspect(
                "workbook",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetWorkbookSummary()),
            inspect(
                "sheet",
                new SheetSelector.ByName("Streamed"),
                new InspectionQuery.GetSheetSummary()));
    GridGrindResponse.Success reopened =
        XlsxParityGridGrind.readWorkbook(
            gridGrindWorkbook,
            inspect(
                "workbook",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetWorkbookSummary()),
            inspect(
                "sheet",
                new SheetSelector.ByName("Streamed"),
                new InspectionQuery.GetSheetSummary()));
    InspectionResult.SheetSummaryResult sheetSummary =
        XlsxParityGridGrind.read(streamed, "sheet", InspectionResult.SheetSummaryResult.class);
    return poiSucceeded
            && Files.exists(gridGrindWorkbook)
            && sheetSummary.sheet().physicalRowCount() == 1_500
            && sheetSummary.sheet().lastColumnIndex() == 1
            && streamed.inspections().equals(reopened.inspections())
        ? pass("SXSSF streaming-write parity is present.")
        : fail("GridGrind streaming-write parity did not match the expected low-memory workbook.");
  }

  private static ProbeResult probeEncryptionGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario encrypted =
        context.scenario(XlsxParityScenarios.ENCRYPTED_WORKBOOK);
    OoxmlOpenSecurityInput sourceSecurity = encryptedOpenSecurity();
    boolean poiSourceOpens = XlsxParityOracle.encryptedWorkbookOpens(encrypted.workbookPath());

    GridGrindResponse.Success sourceRead =
        XlsxParityGridGrind.readWorkbook(
            encrypted.workbookPath(),
            sourceSecurity,
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Encrypted", List.of("A1")),
                new InspectionQuery.GetCells()));
    InspectionResult.PackageSecurityResult sourceSecurityResult =
        XlsxParityGridGrind.read(
            sourceRead, "security", InspectionResult.PackageSecurityResult.class);
    InspectionResult.CellsResult sourceCells =
        XlsxParityGridGrind.read(sourceRead, "cells", InspectionResult.CellsResult.class);

    Path preservedOutput = context.derivedWorkbook("encrypted-preserved");
    GridGrindResponse.Success preservedSave =
        XlsxParityGridGrind.mutateWorkbook(
            encrypted.workbookPath(),
            sourceSecurity,
            preservedOutput,
            null,
            null,
            List.of(
                mutate(
                    new CellSelector.ByAddress("Encrypted", "A2"),
                    new MutationAction.SetCell(text("Preserved encryption")))));
    boolean poiPreservedOpens = XlsxParityOracle.encryptedWorkbookOpens(preservedOutput);
    String poiPreservedText =
        XlsxParityOracle.encryptedStringCell(preservedOutput, "Encrypted", "A2");
    GridGrindResponse.Success preservedRead =
        XlsxParityGridGrind.readWorkbook(
            preservedOutput,
            sourceSecurity,
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Encrypted", List.of("A1", "A2")),
                new InspectionQuery.GetCells()));
    InspectionResult.PackageSecurityResult preservedSecurityResult =
        XlsxParityGridGrind.read(
            preservedRead, "security", InspectionResult.PackageSecurityResult.class);
    InspectionResult.CellsResult preservedCells =
        XlsxParityGridGrind.read(preservedRead, "cells", InspectionResult.CellsResult.class);

    Path authoredOutput = context.derivedWorkbook("encrypted-authored");
    GridGrindResponse.Success authoredSave =
        XlsxParityGridGrind.writeNewWorkbook(
            authoredOutput,
            new OoxmlPersistenceSecurityInput(
                new OoxmlEncryptionInput(XlsxParityScenarios.ENCRYPTION_PASSWORD, null), null),
            (ExecutionModeInput) null,
            List.of(
                mutate(new SheetSelector.ByName("Secure"), new MutationAction.EnsureSheet()),
                mutate(
                    new CellSelector.ByAddress("Secure", "A1"),
                    new MutationAction.SetCell(text("Authored encrypted")))));
    String poiAuthoredText = XlsxParityOracle.encryptedStringCell(authoredOutput, "Secure", "A1");
    GridGrindResponse.Success authoredRead =
        XlsxParityGridGrind.readWorkbook(
            authoredOutput,
            sourceSecurity,
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Secure", List.of("A1")),
                new InspectionQuery.GetCells()));
    InspectionResult.PackageSecurityResult authoredSecurityResult =
        XlsxParityGridGrind.read(
            authoredRead, "security", InspectionResult.PackageSecurityResult.class);
    InspectionResult.CellsResult authoredCells =
        XlsxParityGridGrind.read(authoredRead, "cells", InspectionResult.CellsResult.class);

    boolean preservedSavedToExpectedPath =
        XlsxParityGridGrind.savedPath(preservedSave)
            .equals(preservedOutput.toAbsolutePath().toString());
    boolean authoredSavedToExpectedPath =
        XlsxParityGridGrind.savedPath(authoredSave)
            .equals(authoredOutput.toAbsolutePath().toString());

    return poiSourceOpens
            && hasEncryptedAgilePackage(sourceSecurityResult.security())
            && "Encrypted workbook".equals(textCell(sourceCells, "A1"))
            && preservedSavedToExpectedPath
            && poiPreservedOpens
            && "Preserved encryption".equals(poiPreservedText)
            && hasEncryptedAgilePackage(preservedSecurityResult.security())
            && "Encrypted workbook".equals(textCell(preservedCells, "A1"))
            && "Preserved encryption".equals(textCell(preservedCells, "A2"))
            && authoredSavedToExpectedPath
            && "Authored encrypted".equals(poiAuthoredText)
            && hasEncryptedAgilePackage(authoredSecurityResult.security())
            && "Authored encrypted".equals(textCell(authoredCells, "A1"))
        ? pass(
            "OOXML encryption parity is present for encrypted open, preserved encrypted save-as, and authored encrypted save-as.")
        : fail(
            "OOXML encryption parity mismatch."
                + " poiSourceOpens="
                + poiSourceOpens
                + " sourceEncrypted="
                + hasEncryptedAgilePackage(sourceSecurityResult.security())
                + " preservedSaved="
                + preservedSavedToExpectedPath
                + " poiPreservedOpens="
                + poiPreservedOpens
                + " poiPreservedText="
                + poiPreservedText
                + " preservedEncrypted="
                + hasEncryptedAgilePackage(preservedSecurityResult.security())
                + " authoredSaved="
                + authoredSavedToExpectedPath
                + " poiAuthoredText="
                + poiAuthoredText
                + " authoredEncrypted="
                + hasEncryptedAgilePackage(authoredSecurityResult.security()));
  }

  private static ProbeResult probeEncryptionPasswordErrors(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario encrypted =
        context.scenario(XlsxParityScenarios.ENCRYPTED_WORKBOOK);

    GridGrindResponse missingPasswordResponse =
        XlsxParityGridGrind.executeReadWorkbook(
            encrypted.workbookPath(),
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()));
    GridGrindResponse wrongPasswordResponse =
        XlsxParityGridGrind.executeReadWorkbook(
            encrypted.workbookPath(),
            new OoxmlOpenSecurityInput("gridgrind-phase9-wrong-password"),
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()));

    if (!(missingPasswordResponse instanceof GridGrindResponse.Failure missingPassword)) {
      return fail("Encrypted workbook open without a password unexpectedly succeeded.");
    }
    if (!(wrongPasswordResponse instanceof GridGrindResponse.Failure wrongPassword)) {
      return fail("Encrypted workbook open with the wrong password unexpectedly succeeded.");
    }

    return missingPassword.problem().code() == GridGrindProblemCode.WORKBOOK_PASSWORD_REQUIRED
            && missingPassword.problem().category() == GridGrindProblemCategory.SECURITY
            && wrongPassword.problem().code() == GridGrindProblemCode.INVALID_WORKBOOK_PASSWORD
            && wrongPassword.problem().category() == GridGrindProblemCategory.SECURITY
        ? pass(
            "Encrypted workbook password failures map to stable security problem codes for missing and invalid passwords.")
        : fail(
            "Encrypted workbook password failure parity mismatch."
                + " missingCode="
                + missingPassword.problem().code()
                + " missingCategory="
                + missingPassword.problem().category()
                + " wrongCode="
                + wrongPassword.problem().code()
                + " wrongCategory="
                + wrongPassword.problem().category());
  }

  private static ProbeResult probeSigningGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario signed =
        context.scenario(XlsxParityScenarios.SIGNED_WORKBOOK);
    boolean poiSourceValid = XlsxParityOracle.signatureValid(signed.workbookPath());
    GridGrindResponse.Success sourceRead =
        XlsxParityGridGrind.readWorkbook(
            signed.workbookPath(),
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Signed", List.of("A1")),
                new InspectionQuery.GetCells()));
    InspectionResult.PackageSecurityResult sourceSecurityResult =
        XlsxParityGridGrind.read(
            sourceRead, "security", InspectionResult.PackageSecurityResult.class);
    InspectionResult.CellsResult sourceCells =
        XlsxParityGridGrind.read(sourceRead, "cells", InspectionResult.CellsResult.class);

    Path preservedOutput = context.derivedWorkbook("signed-preserved");
    GridGrindResponse.Success preservedSave =
        XlsxParityGridGrind.mutateWorkbook(signed.workbookPath(), preservedOutput, List.of());
    boolean poiPreservedValid = XlsxParityOracle.signatureValid(preservedOutput);
    GridGrindResponse.Success preservedRead =
        XlsxParityGridGrind.readWorkbook(
            preservedOutput,
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()));
    InspectionResult.PackageSecurityResult preservedSecurityResult =
        XlsxParityGridGrind.read(
            preservedRead, "security", InspectionResult.PackageSecurityResult.class);

    XlsxParityScenarios.MaterializedScenario unsignedMutationSource =
        context.copiedScenario(XlsxParityScenarios.SIGNED_WORKBOOK, "signed-unsigned-mutation");
    Path unsignedOutput = context.derivedWorkbook("signed-unsigned-mutation");
    GridGrindResponse unsignedMutationResponse =
        XlsxParityGridGrind.executeMutateWorkbook(
            unsignedMutationSource.workbookPath(),
            unsignedOutput,
            List.of(
                mutate(
                    new CellSelector.ByAddress("Signed", "C1"),
                    new MutationAction.SetCell(text("Touch")))));
    if (!(unsignedMutationResponse instanceof GridGrindResponse.Failure unsignedMutationFailure)) {
      return fail("Mutating a signed workbook without explicit re-sign unexpectedly succeeded.");
    }

    XlsxParityScenarios.MaterializedScenario resignSource =
        context.copiedScenario(XlsxParityScenarios.SIGNED_WORKBOOK, "signed-resigned-source");
    Path resignedOutput = context.derivedWorkbook("signed-resigned");
    GridGrindResponse.Success resignedSave =
        XlsxParityGridGrind.mutateWorkbook(
            resignSource.workbookPath(),
            null,
            resignedOutput,
            new OoxmlPersistenceSecurityInput(
                null,
                signingInput(
                    resignSource.attachment(XlsxParityScenarios.SIGNING_PKCS12_ATTACHMENT),
                    "GridGrind parity re-sign")),
            null,
            List.of(
                mutate(
                    new CellSelector.ByAddress("Signed", "C1"),
                    new MutationAction.SetCell(text("Re-signed")))));
    boolean poiResignedValid = XlsxParityOracle.signatureValid(resignedOutput);
    GridGrindResponse.Success resignedRead =
        XlsxParityGridGrind.readWorkbook(
            resignedOutput,
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Signed", List.of("C1")),
                new InspectionQuery.GetCells()));
    InspectionResult.PackageSecurityResult resignedSecurityResult =
        XlsxParityGridGrind.read(
            resignedRead, "security", InspectionResult.PackageSecurityResult.class);
    InspectionResult.CellsResult resignedCells =
        XlsxParityGridGrind.read(resignedRead, "cells", InspectionResult.CellsResult.class);

    Path authoredOutput = context.derivedWorkbook("signed-authored");
    GridGrindResponse.Success authoredSave =
        XlsxParityGridGrind.writeNewWorkbook(
            authoredOutput,
            new OoxmlPersistenceSecurityInput(
                null,
                signingInput(
                    signed.attachment(XlsxParityScenarios.SIGNING_PKCS12_ATTACHMENT),
                    "GridGrind parity authored signature")),
            (ExecutionModeInput) null,
            List.of(
                mutate(new SheetSelector.ByName("Signed"), new MutationAction.EnsureSheet()),
                mutate(
                    new CellSelector.ByAddress("Signed", "A1"),
                    new MutationAction.SetCell(text("Authored signed")))));
    boolean poiAuthoredValid = XlsxParityOracle.signatureValid(authoredOutput);
    GridGrindResponse.Success authoredRead =
        XlsxParityGridGrind.readWorkbook(
            authoredOutput,
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Signed", List.of("A1")),
                new InspectionQuery.GetCells()));
    InspectionResult.PackageSecurityResult authoredSecurityResult =
        XlsxParityGridGrind.read(
            authoredRead, "security", InspectionResult.PackageSecurityResult.class);
    InspectionResult.CellsResult authoredCells =
        XlsxParityGridGrind.read(authoredRead, "cells", InspectionResult.CellsResult.class);

    return poiSourceValid
            && hasSingleSignatureState(
                sourceSecurityResult.security(), ExcelOoxmlSignatureState.VALID)
            && "Signed workbook".equals(textCell(sourceCells, "A1"))
            && XlsxParityGridGrind.savedPath(preservedSave)
                .equals(preservedOutput.toAbsolutePath().toString())
            && poiPreservedValid
            && hasSingleSignatureState(
                preservedSecurityResult.security(), ExcelOoxmlSignatureState.VALID)
            && unsignedMutationFailure.problem().code() == GridGrindProblemCode.INVALID_REQUEST
            && unsignedMutationFailure
                .problem()
                .message()
                .contains("persistence.security.signature")
            && !Files.exists(unsignedOutput)
            && XlsxParityGridGrind.savedPath(resignedSave)
                .equals(resignedOutput.toAbsolutePath().toString())
            && poiResignedValid
            && hasSingleSignatureState(
                resignedSecurityResult.security(), ExcelOoxmlSignatureState.VALID)
            && "Re-signed".equals(textCell(resignedCells, "C1"))
            && XlsxParityGridGrind.savedPath(authoredSave)
                .equals(authoredOutput.toAbsolutePath().toString())
            && poiAuthoredValid
            && hasSingleSignatureState(
                authoredSecurityResult.security(), ExcelOoxmlSignatureState.VALID)
            && "Authored signed".equals(textCell(authoredCells, "A1"))
        ? pass(
            "OOXML signing parity is present for signature validation, unchanged signature preservation, explicit re-signing, and authored signed save-as.")
        : fail(
            "OOXML signing parity mismatch."
                + " poiSourceValid="
                + poiSourceValid
                + " sourceValid="
                + hasSingleSignatureState(
                    sourceSecurityResult.security(), ExcelOoxmlSignatureState.VALID)
                + " poiPreservedValid="
                + poiPreservedValid
                + " preservedValid="
                + hasSingleSignatureState(
                    preservedSecurityResult.security(), ExcelOoxmlSignatureState.VALID)
                + " unsignedFailureCode="
                + unsignedMutationFailure.problem().code()
                + " unsignedOutputExists="
                + Files.exists(unsignedOutput)
                + " poiResignedValid="
                + poiResignedValid
                + " resignedValid="
                + hasSingleSignatureState(
                    resignedSecurityResult.security(), ExcelOoxmlSignatureState.VALID)
                + " poiAuthoredValid="
                + poiAuthoredValid
                + " authoredValid="
                + hasSingleSignatureState(
                    authoredSecurityResult.security(), ExcelOoxmlSignatureState.VALID));
  }

  private static ProbeResult probeSigningInvalidSignature(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario invalid =
        context.scenario(XlsxParityScenarios.INVALID_SIGNATURE_WORKBOOK);
    boolean poiInvalid = !XlsxParityOracle.signatureValid(invalid.workbookPath());
    GridGrindResponse.Success read =
        XlsxParityGridGrind.readWorkbook(
            invalid.workbookPath(),
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Signed", List.of("A1")),
                new InspectionQuery.GetCells()));
    InspectionResult.PackageSecurityResult security =
        XlsxParityGridGrind.read(read, "security", InspectionResult.PackageSecurityResult.class);
    InspectionResult.CellsResult cells =
        XlsxParityGridGrind.read(read, "cells", InspectionResult.CellsResult.class);

    return poiInvalid
            && hasSingleSignatureState(security.security(), ExcelOoxmlSignatureState.INVALID)
            && "Signed workbook".equals(textCell(cells, "A1"))
        ? pass(
            "Invalid OOXML package signatures degrade into truthful INVALID signature facts instead of aborting the read.")
        : fail(
            "Invalid OOXML signature parity mismatch."
                + " poiInvalid="
                + poiInvalid
                + " gridState="
                + security.security().signatures());
  }

  private static FormulaEnvironmentInput externalFormulaEnvironment(Path referencedWorkbookPath) {
    return new FormulaEnvironmentInput(
        List.of(
            new FormulaExternalWorkbookInput(
                "referenced.xlsx", referencedWorkbookPath.toAbsolutePath().toString())),
        FormulaMissingWorkbookPolicy.ERROR,
        List.of());
  }

  private static FormulaEnvironmentInput missingWorkbookCachedValueEnvironment() {
    return new FormulaEnvironmentInput(
        List.of(), FormulaMissingWorkbookPolicy.USE_CACHED_VALUE, List.of());
  }

  private static FormulaEnvironmentInput udfFormulaEnvironment() {
    return new FormulaEnvironmentInput(
        List.of(),
        FormulaMissingWorkbookPolicy.ERROR,
        List.of(
            new FormulaUdfToolpackInput(
                "math", List.of(new FormulaUdfFunctionInput("DOUBLE", 1, null, "ARG1*2")))));
  }

  private static OoxmlOpenSecurityInput encryptedOpenSecurity() {
    return new OoxmlOpenSecurityInput(XlsxParityScenarios.ENCRYPTION_PASSWORD);
  }

  private static OoxmlSignatureInput signingInput(Path pkcs12Path, String description) {
    return new OoxmlSignatureInput(
        pkcs12Path.toAbsolutePath().toString(),
        XlsxParityScenarios.SIGNING_KEYSTORE_PASSWORD,
        XlsxParityScenarios.SIGNING_KEY_PASSWORD,
        XlsxParityScenarios.SIGNING_KEY_ALIAS,
        null,
        description);
  }

  private static <T> T cast(Class<T> type, Object value) {
    return type.cast(value);
  }

  private static boolean hasEncryptedAgilePackage(OoxmlPackageSecurityReport security) {
    return security.encryption().encrypted()
        && security.encryption().mode() == ExcelOoxmlEncryptionMode.AGILE
        && security.signatures().isEmpty();
  }

  private static boolean hasSingleSignatureState(
      OoxmlPackageSecurityReport security, ExcelOoxmlSignatureState state) {
    return !security.encryption().encrypted()
        && security.signatures().size() == 1
        && security.signatures().getFirst().state() == state;
  }

  private static String textCell(InspectionResult.CellsResult cells, String address) {
    return cast(
            GridGrindResponse.CellReport.TextReport.class, byAddress(cells.cells()).get(address))
        .stringValue();
  }

  private static Map<String, GridGrindResponse.CellReport> byAddress(
      List<GridGrindResponse.CellReport> cells) {
    return cells.stream()
        .collect(
            Collectors.toUnmodifiableMap(
                GridGrindResponse.CellReport::address, Function.identity()));
  }

  private static boolean matchesWorkbookProtectionReport(
      WorkbookProtectionReport observed, XlsxParityOracle.WorkbookProtectionSnapshot direct) {
    return observed.structureLocked() == direct.structureLocked()
        && observed.windowsLocked() == direct.windowsLocked()
        && observed.revisionsLocked() == direct.revisionsLocked()
        && observed.workbookPasswordHashPresent() == direct.workbookPasswordHashPresent()
        && observed.revisionsPasswordHashPresent() == direct.revisionsPasswordHashPresent();
  }

  private static SheetProtectionSettings sheetProtectionSettings(
      InspectionResult.SheetSummaryResult summary) {
    return switch (summary.sheet().protection()) {
      case GridGrindResponse.SheetProtectionReport.Unprotected _ -> null;
      case GridGrindResponse.SheetProtectionReport.Protected protectedReport ->
          protectedReport.settings();
    };
  }

  private static boolean sheetPasswordMatches(
      Path workbookPath, String sheetName, String password) {
    return XlsxParitySupport.call(
        "validate protected-sheet password " + sheetName + "@" + workbookPath.getFileName(),
        () -> {
          try (XSSFWorkbook workbook =
              (XSSFWorkbook) WorkbookFactory.create(workbookPath.toFile())) {
            return workbook.getSheet(sheetName).validateSheetPassword(password);
          }
        });
  }

  private static PrintLayoutReport defaultPrintLayoutReport(String sheetName) {
    return new PrintLayoutReport(
        sheetName,
        new PrintAreaReport.None(),
        ExcelPrintOrientation.PORTRAIT,
        new PrintScalingReport.Automatic(),
        new PrintTitleRowsReport.None(),
        new PrintTitleColumnsReport.None(),
        new HeaderFooterTextReport("", "", ""),
        new HeaderFooterTextReport("", "", ""),
        PrintSetupReport.defaults());
  }

  private static PrintLayoutInput advancedPrintLayoutInput() {
    return new PrintLayoutInput(
        new PrintAreaInput.None(),
        ExcelPrintOrientation.PORTRAIT,
        new PrintScalingInput.Automatic(),
        new PrintTitleRowsInput.None(),
        new PrintTitleColumnsInput.None(),
        HeaderFooterTextInput.blank(),
        HeaderFooterTextInput.blank(),
        new PrintSetupInput(
            new PrintMarginsInput(0.35d, 0.55d, 0.60d, 0.45d, 0.3d, 0.3d),
            true,
            true,
            true,
            8,
            true,
            true,
            2,
            true,
            4,
            List.of(6),
            List.of(3)));
  }

  private static CellStyleInput advancedThemedStyleInput() {
    return new CellStyleInput(
        null,
        null,
        new CellFontInput(null, true, null, null, null, 6, null, -0.35d, null, null),
        new CellFillInput(
            ExcelFillPattern.SOLID, null, 3, null, 0.30d, null, null, null, null, null),
        new CellBorderInput(
            null,
            null,
            null,
            new CellBorderSideInput(
                ExcelBorderStyle.THIN,
                null,
                null,
                Short.toUnsignedInt(IndexedColors.DARK_RED.getIndex()),
                null),
            null),
        null);
  }

  private static CellStyleInput advancedGradientStyleInput() {
    return new CellStyleInput(
        null,
        null,
        null,
        new CellFillInput(
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            new CellGradientFillInput(
                "LINEAR",
                45.0d,
                null,
                null,
                null,
                null,
                List.of(
                    new CellGradientStopInput(0.0d, new ColorInput("#1F497D")),
                    new CellGradientStopInput(1.0d, new ColorInput(null, 4, null, 0.45d))))),
        null,
        null);
  }

  private static CommentInput advancedCommentInput() {
    return comment(
        "Lead review scheduled",
        "GridGrind",
        false,
        List.of(
            richTextRun(
                "Lead",
                new CellFontInput(true, null, null, null, null, 4, null, -0.20d, null, null)),
            richTextRun(" ", null),
            richTextRun(
                "review scheduled",
                new CellFontInput(
                    null,
                    true,
                    null,
                    null,
                    null,
                    null,
                    Short.toUnsignedInt(IndexedColors.DARK_GREEN.getIndex()),
                    null,
                    null,
                    null))),
        new CommentAnchorInput(4, 1, 8, 7));
  }

  private static DrawingAnchorInput.TwoCell twoCellAnchorInput(
      int fromColumn, int fromRow, int toColumn, int toRow, ExcelDrawingAnchorBehavior behavior) {
    return new DrawingAnchorInput.TwoCell(
        new DrawingMarkerInput(fromColumn, fromRow, 0, 0),
        new DrawingMarkerInput(toColumn, toRow, 0, 0),
        behavior);
  }

  private static TableInput advancedTableInput() {
    return new TableInput(
        "AdvancedTable",
        "Advanced",
        "H1:J5",
        true,
        false,
        new TableStyleInput.None(),
        inlineText("Parity advanced table"),
        true,
        true,
        true,
        "ParityHeader",
        "ParityData",
        "ParityTotals",
        List.of(
            new TableColumnInput(0, "", "Total", "", ""),
            new TableColumnInput(1, "", "", "SUM", "[@Sales]*2"),
            new TableColumnInput(2, "owner-unique", "", "", "")));
  }

  private static List<PendingMutation> advancedConditionalFormattingMutations() {
    PendingMutation colorScale =
        mutate(
            new SheetSelector.ByName("Advanced"),
            new MutationAction.SetConditionalFormatting(
                new ConditionalFormattingBlockInput(
                    List.of("L2:L5"),
                    List.of(
                        new ConditionalFormattingRuleInput.ColorScaleRule(
                            false,
                            List.of(
                                new ConditionalFormattingThresholdInput(
                                    ExcelConditionalFormattingThresholdType.MIN, null, null),
                                new ConditionalFormattingThresholdInput(
                                    ExcelConditionalFormattingThresholdType.PERCENTILE,
                                    null,
                                    50.0d),
                                new ConditionalFormattingThresholdInput(
                                    ExcelConditionalFormattingThresholdType.MAX, null, null)),
                            List.of(
                                new ColorInput("#AA2211"),
                                new ColorInput("#FFDD55"),
                                new ColorInput("#11CC66")))))));
    PendingMutation dataBar =
        mutate(
            new SheetSelector.ByName("Advanced"),
            new MutationAction.SetConditionalFormatting(
                new ConditionalFormattingBlockInput(
                    List.of("M2:M5"),
                    List.of(
                        new ConditionalFormattingRuleInput.DataBarRule(
                            false,
                            new ColorInput("#123456"),
                            true,
                            10,
                            90,
                            new ConditionalFormattingThresholdInput(
                                ExcelConditionalFormattingThresholdType.MIN, null, null),
                            new ConditionalFormattingThresholdInput(
                                ExcelConditionalFormattingThresholdType.MAX, null, null))))));
    PendingMutation iconSet =
        mutate(
            new SheetSelector.ByName("Advanced"),
            new MutationAction.SetConditionalFormatting(
                new ConditionalFormattingBlockInput(
                    List.of("N2:N5"),
                    List.of(
                        new ConditionalFormattingRuleInput.IconSetRule(
                            false,
                            ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
                            true,
                            true,
                            List.of(
                                new ConditionalFormattingThresholdInput(
                                    ExcelConditionalFormattingThresholdType.PERCENT, null, 0.0d),
                                new ConditionalFormattingThresholdInput(
                                    ExcelConditionalFormattingThresholdType.PERCENT, null, 33.0d),
                                new ConditionalFormattingThresholdInput(
                                    ExcelConditionalFormattingThresholdType.PERCENT,
                                    null,
                                    67.0d)))))));
    PendingMutation top10 =
        mutate(
            new SheetSelector.ByName("Advanced"),
            new MutationAction.SetConditionalFormatting(
                new ConditionalFormattingBlockInput(
                    List.of("K2:K5"),
                    List.of(
                        new ConditionalFormattingRuleInput.Top10Rule(
                            false, 10, false, false, null)))));
    return List.of(colorScale, dataBar, iconSet, top10);
  }

  private static ProbeResult pass(String detail) {
    return new ProbeResult(XlsxParityLedger.ExpectedOutcome.PASS, detail);
  }

  private static ProbeResult fail(String detail) {
    return new ProbeResult(XlsxParityLedger.ExpectedOutcome.FAIL, detail);
  }

  private static boolean sortConditionsMatch(
      AutofilterSortConditionReport observed, XlsxParityOracle.SortConditionSnapshot direct) {
    return observed.range().equals(direct.range())
        && observed.descending() == direct.descending()
        && observed.sortBy().equals(direct.sortBy());
  }

  private static TableEntryReport findTable(List<TableEntryReport> tables, String name) {
    return tables.stream().filter(table -> name.equals(table.name())).findFirst().orElse(null);
  }

  private static TableEntryReport findTableOnSheet(
      List<TableEntryReport> tables, String sheetName) {
    return tables.stream()
        .filter(table -> sheetName.equals(table.sheetName()))
        .findFirst()
        .orElse(null);
  }

  private static boolean tableMatchesReadReport(
      TableEntryReport observed,
      TableEntryReport expected,
      String expectedName,
      String expectedSheetName) {
    return observed != null
        && (expectedName == null || observed.name().equals(expectedName))
        && (expectedSheetName == null || observed.sheetName().equals(expectedSheetName))
        && observed.range().equals(expected.range())
        && observed.headerRowCount() == expected.headerRowCount()
        && observed.totalsRowCount() == expected.totalsRowCount()
        && observed.columnNames().equals(expected.columnNames())
        && normalizedTableColumns(observed.columns())
            .equals(normalizedTableColumns(expected.columns()))
        && observed.style().equals(expected.style())
        && observed.hasAutofilter() == expected.hasAutofilter()
        && observed.comment().equals(expected.comment())
        && observed.published() == expected.published()
        && observed.insertRow() == expected.insertRow()
        && observed.insertRowShift() == expected.insertRowShift()
        && observed.headerRowCellStyle().equals(expected.headerRowCellStyle())
        && observed.dataCellStyle().equals(expected.dataCellStyle())
        && observed.totalsRowCellStyle().equals(expected.totalsRowCellStyle());
  }

  private static List<String> normalizedTableColumns(List<TableColumnReport> columns) {
    return columns.stream()
        .map(
            column ->
                "%s|%s|%s|%s|%s"
                    .formatted(
                        column.name(),
                        column.uniqueName(),
                        column.totalsRowLabel(),
                        column.totalsRowFunction(),
                        column.calculatedColumnFormula()))
        .toList();
  }

  private static boolean matchesAnchor(
      DrawingAnchorReport anchor, XlsxParityOracle.DrawingAnchorSnapshot direct) {
    if (!(anchor instanceof DrawingAnchorReport.TwoCell twoCell)) {
      return false;
    }
    return twoCell.from().columnIndex() == direct.col1()
        && twoCell.from().rowIndex() == direct.row1()
        && twoCell.to().columnIndex() == direct.col2()
        && twoCell.to().rowIndex() == direct.row2();
  }

  private static boolean chartDrawingMatches(
      DrawingObjectReport.Chart observed, XlsxParityOracle.ChartDrawingObjectSnapshot direct) {
    return observed.name().equals(direct.name())
        && matchesAnchor(observed.anchor(), direct.anchor())
        && observed.supported() == direct.supported()
        && observed.plotTypeTokens().equals(direct.plotTypeTokens())
        && observed.title().equals(direct.title());
  }

  private static boolean hasSinglePlot(
      ChartReport chart, Class<? extends ChartReport.Plot> plotType) {
    return chart.plots().size() == 1 && plotType.isInstance(chart.plots().getFirst());
  }

  private static <T extends ChartReport.Plot> T onlyPlot(ChartReport chart, Class<T> plotType) {
    if (!hasSinglePlot(chart, plotType)) {
      throw new IllegalStateException(
          "Expected one " + plotType.getSimpleName() + " plot on chart " + chart.name());
    }
    return plotType.cast(chart.plots().getFirst());
  }

  private static List<String> plotTypeTokens(ChartReport chart) {
    return chart.plots().stream()
        .map(
            plot ->
                switch (plot) {
                  case ChartReport.Area _ -> "AREA";
                  case ChartReport.Area3D _ -> "AREA_3D";
                  case ChartReport.Bar _ -> "BAR";
                  case ChartReport.Bar3D _ -> "BAR_3D";
                  case ChartReport.Doughnut _ -> "DOUGHNUT";
                  case ChartReport.Line _ -> "LINE";
                  case ChartReport.Line3D _ -> "LINE_3D";
                  case ChartReport.Pie _ -> "PIE";
                  case ChartReport.Pie3D _ -> "PIE_3D";
                  case ChartReport.Radar _ -> "RADAR";
                  case ChartReport.Scatter _ -> "SCATTER";
                  case ChartReport.Surface _ -> "SURFACE";
                  case ChartReport.Surface3D _ -> "SURFACE_3D";
                  case ChartReport.Unsupported unsupported -> unsupported.plotTypeToken();
                })
        .toList();
  }

  private static <T extends XlsxParityOracle.DirectDrawingObjectSnapshot> T directObject(
      XlsxParityOracle.DrawingSheetSnapshot snapshot, String name, Class<T> type) {
    return snapshot.objects().stream()
        .filter(candidate -> candidate.name().equals(name))
        .findFirst()
        .filter(type::isInstance)
        .map(type::cast)
        .orElseThrow(
            () ->
                new IllegalStateException(
                    "Missing direct oracle drawing object "
                        + name
                        + " as "
                        + type.getSimpleName()));
  }

  private static <T extends DrawingObjectReport> T drawingObjectReport(
      InspectionResult.DrawingObjectsResult result, String name, Class<T> type) {
    return result.drawingObjects().stream()
        .filter(candidate -> candidate.name().equals(name))
        .findFirst()
        .filter(type::isInstance)
        .map(type::cast)
        .orElseThrow(
            () ->
                new IllegalStateException(
                    "Missing GridGrind drawing report " + name + " as " + type.getSimpleName()));
  }

  private static boolean matchesColorDescriptor(
      dev.erst.gridgrind.contract.dto.CellColorReport color, String descriptor) {
    if (descriptor == null || descriptor.isBlank()) {
      return color == null;
    }
    if (color == null) {
      return false;
    }
    List<String> parts = Arrays.asList(descriptor.split("\\|"));
    for (String part : parts) {
      if (part.startsWith("rgb=")) {
        String expected = "#" + part.substring("rgb=".length());
        if (!expected.equals(color.rgb())) {
          return false;
        }
      } else if (part.startsWith("theme=")) {
        if (!Integer.valueOf(part.substring("theme=".length())).equals(color.theme())) {
          return false;
        }
      } else if (part.startsWith("indexed=")) {
        if (!Integer.valueOf(part.substring("indexed=".length())).equals(color.indexed())) {
          return false;
        }
      } else if (part.startsWith("tint=")
          && (color.tint() == null
              || !approximatelyEquals(
                  Double.parseDouble(part.substring("tint=".length())), color.tint()))) {
        return false;
      }
    }
    return true;
  }

  private static boolean approximatelyEquals(double left, double right) {
    return Math.abs(left - right) <= 0.000001d;
  }

  private static CellInput.Text text(String value) {
    return new CellInput.Text(TextSourceInput.inline(value));
  }

  private static ChartInput.Title.Text chartTitle(String value) {
    return new ChartInput.Title.Text(TextSourceInput.inline(value));
  }

  private static RichTextRunInput richTextRun(String value, CellFontInput font) {
    return new RichTextRunInput(TextSourceInput.inline(value), font);
  }

  private static CommentInput comment(String text, String author, boolean visible) {
    return new CommentInput(TextSourceInput.inline(text), author, visible);
  }

  private static CommentInput comment(
      String text,
      String author,
      boolean visible,
      List<RichTextRunInput> runs,
      CommentAnchorInput anchor) {
    return new CommentInput(TextSourceInput.inline(text), author, visible, runs, anchor);
  }

  private static TextSourceInput inlineText(String value) {
    return TextSourceInput.inline(value);
  }

  record ProbeResult(XlsxParityLedger.ExpectedOutcome outcome, String detail) {}

  private record CoreReadObservation(
      XlsxParityOracle.CoreWorkbookSnapshot direct,
      GridGrindResponse.WorkbookSummary.WithSheets summary,
      InspectionResult.SheetSummaryResult opsSummary,
      InspectionResult.SheetSummaryResult queueSummary,
      InspectionResult.CellsResult cells,
      InspectionResult.MergedRegionsResult merged,
      InspectionResult.HyperlinksResult links,
      InspectionResult.CommentsResult comments,
      InspectionResult.SheetLayoutResult layout,
      InspectionResult.PrintLayoutResult print,
      InspectionResult.DataValidationsResult validations,
      InspectionResult.ConditionalFormattingResult formatting,
      InspectionResult.AutofiltersResult filtersOps,
      InspectionResult.AutofiltersResult filtersQueue,
      InspectionResult.TablesResult tables) {}

  private record SheetCopySourceObservation(
      XlsxParityScenarios.MaterializedScenario source,
      InspectionResult.CommentsResult sourceComments,
      InspectionResult.AutofiltersResult sourceAutofilters,
      InspectionResult.ConditionalFormattingResult sourceFormatting,
      InspectionResult.SheetSummaryResult sourceSummary,
      TableEntryReport sourceTable,
      XlsxParityOracle.CommentSnapshot sourceComment,
      XlsxParityOracle.AdvancedPrintSnapshot sourcePrint,
      XlsxParityOracle.AutofilterSnapshot sourceAutofilter,
      XlsxParityOracle.TableSnapshot sourceDirectTable) {}

  private record SheetCopyCopiedObservation(
      Path copiedPath,
      InspectionResult.WorkbookSummaryResult copiedWorkbook,
      InspectionResult.SheetSummaryResult copiedSummary,
      InspectionResult.CommentsResult copiedComments,
      InspectionResult.AutofiltersResult copiedAutofilters,
      InspectionResult.ConditionalFormattingResult copiedFormatting,
      TableEntryReport copiedTable,
      XlsxParityOracle.CommentSnapshot copiedComment,
      XlsxParityOracle.AdvancedPrintSnapshot copiedPrint,
      XlsxParityOracle.AutofilterSnapshot copiedAutofilter,
      XlsxParityOracle.TableSnapshot copiedDirectTable) {}

  /** Temporary-workspace cache and copy helper for materialized parity scenarios. */
  static final class ProbeContext {
    private final Path temporaryRoot;
    private final Map<String, XlsxParityScenarios.MaterializedScenario> scenarioCache =
        new ConcurrentHashMap<>();

    ProbeContext(Path temporaryRoot) {
      this.temporaryRoot = temporaryRoot;
    }

    XlsxParityScenarios.MaterializedScenario scenario(String scenarioId) {
      XlsxParityScenarios.MaterializedScenario existing = scenarioCache.get(scenarioId);
      if (existing != null) {
        return existing;
      }
      XlsxParityScenarios.MaterializedScenario materialized =
          XlsxParityScenarios.materialize(scenarioId, temporaryRoot.resolve("corpus"));
      scenarioCache.put(scenarioId, materialized);
      return materialized;
    }

    XlsxParityScenarios.MaterializedScenario copiedScenario(String scenarioId, String copyName) {
      return XlsxParitySupport.call(
          "copy materialized parity scenario " + scenarioId,
          () -> {
            XlsxParityScenarios.MaterializedScenario source = scenario(scenarioId);
            Path directory =
                Files.createDirectories(temporaryRoot.resolve("scenario-copies").resolve(copyName));
            Path workbookCopy = directory.resolve(source.workbookPath().getFileName().toString());
            Files.copy(source.workbookPath(), workbookCopy, StandardCopyOption.REPLACE_EXISTING);
            Map<String, Path> attachmentCopies = new ConcurrentHashMap<>();
            for (Map.Entry<String, Path> entry : source.attachments().entrySet()) {
              Path attachmentCopy = directory.resolve(entry.getValue().getFileName().toString());
              Files.copy(entry.getValue(), attachmentCopy, StandardCopyOption.REPLACE_EXISTING);
              attachmentCopies.put(entry.getKey(), attachmentCopy);
            }
            return new XlsxParityScenarios.MaterializedScenario(workbookCopy, attachmentCopies);
          });
    }

    Path derivedWorkbook(String name) {
      Path directory = derivedDirectory("workbooks");
      return directory.resolve(name + ".xlsx");
    }

    Path derivedDirectory(String name) {
      return XlsxParitySupport.call(
          "create derived parity directory " + name,
          () -> Files.createDirectories(temporaryRoot.resolve(name)));
    }
  }
}
