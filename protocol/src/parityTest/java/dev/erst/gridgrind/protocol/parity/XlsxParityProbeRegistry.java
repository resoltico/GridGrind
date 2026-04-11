package dev.erst.gridgrind.protocol.parity;

import dev.erst.gridgrind.protocol.catalog.TypeEntry;
import dev.erst.gridgrind.protocol.dto.AutofilterEntryReport;
import dev.erst.gridgrind.protocol.dto.CellFillReport;
import dev.erst.gridgrind.protocol.dto.CellFontReport;
import dev.erst.gridgrind.protocol.dto.CellInput;
import dev.erst.gridgrind.protocol.dto.CellSelection;
import dev.erst.gridgrind.protocol.dto.CommentInput;
import dev.erst.gridgrind.protocol.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.protocol.dto.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.protocol.dto.DataValidationEntryReport;
import dev.erst.gridgrind.protocol.dto.GridGrindResponse;
import dev.erst.gridgrind.protocol.dto.HyperlinkTarget;
import dev.erst.gridgrind.protocol.dto.NamedRangeSelection;
import dev.erst.gridgrind.protocol.dto.RangeSelection;
import dev.erst.gridgrind.protocol.dto.SheetCopyPosition;
import dev.erst.gridgrind.protocol.dto.SheetSelection;
import dev.erst.gridgrind.protocol.dto.TableEntryReport;
import dev.erst.gridgrind.protocol.dto.TableSelection;
import dev.erst.gridgrind.protocol.exec.DefaultGridGrindRequestExecutor;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadResult;
import java.lang.reflect.RecordComponent;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import java.util.function.Function;
import java.util.stream.Collectors;
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
      case "probe-formula-external-gap" -> probeFormulaExternalGap(context);
      case "probe-formula-udf-gap" -> probeFormulaUdfGap(context);
      case "probe-formula-lifecycle-gap" -> probeFormulaLifecycleGap();
      case "probe-drawing-preservation" -> probeDrawingPreservation(context);
      case "probe-embedded-object-preservation" -> probeEmbeddedObjectPreservation(context);
      case "probe-chart-preservation" -> probeChartPreservation(context);
      case "probe-pivot-preservation" -> probePivotPreservation(context);
      case "probe-event-model-gap" -> probeEventModelGap(context);
      case "probe-sxssf-gap" -> probeSxssfGap(context);
      case "probe-encryption-gap" -> probeEncryptionGap(context);
      case "probe-signing-gap" -> probeSigningGap(context);
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
            new WorkbookReadOperation.GetWorkbookSummary("summary"),
            new WorkbookReadOperation.GetSheetSummary("ops-summary", "Ops"),
            new WorkbookReadOperation.GetSheetSummary("queue-summary", "Queue"),
            new WorkbookReadOperation.GetCells("cells", "Ops", List.of("A3", "B3", "C3", "D3")),
            new WorkbookReadOperation.GetMergedRegions("merged", "Ops"),
            new WorkbookReadOperation.GetHyperlinks(
                "links", "Ops", new CellSelection.Selected(List.of("C4"))),
            new WorkbookReadOperation.GetComments(
                "comments", "Ops", new CellSelection.Selected(List.of("C4"))),
            new WorkbookReadOperation.GetSheetLayout("layout", "Ops"),
            new WorkbookReadOperation.GetPrintLayout("print", "Ops"),
            new WorkbookReadOperation.GetDataValidations(
                "validations", "Ops", new RangeSelection.All()),
            new WorkbookReadOperation.GetConditionalFormatting(
                "formatting", "Ops", new RangeSelection.All()),
            new WorkbookReadOperation.GetAutofilters("filters-ops", "Ops"),
            new WorkbookReadOperation.GetAutofilters("filters-queue", "Queue"),
            new WorkbookReadOperation.GetTables("tables", new TableSelection.All()));

    WorkbookReadResult.WorkbookSummaryResult summaryResult =
        XlsxParityGridGrind.read(
            success, "summary", WorkbookReadResult.WorkbookSummaryResult.class);
    if (!(summaryResult.workbook()
        instanceof GridGrindResponse.WorkbookSummary.WithSheets summary)) {
      throw new IllegalStateException(
          "Core workbook summary did not surface active and selected sheets.");
    }

    return new CoreReadObservation(
        direct,
        summary,
        XlsxParityGridGrind.read(
            success, "ops-summary", WorkbookReadResult.SheetSummaryResult.class),
        XlsxParityGridGrind.read(
            success, "queue-summary", WorkbookReadResult.SheetSummaryResult.class),
        XlsxParityGridGrind.read(success, "cells", WorkbookReadResult.CellsResult.class),
        XlsxParityGridGrind.read(success, "merged", WorkbookReadResult.MergedRegionsResult.class),
        XlsxParityGridGrind.read(success, "links", WorkbookReadResult.HyperlinksResult.class),
        XlsxParityGridGrind.read(success, "comments", WorkbookReadResult.CommentsResult.class),
        XlsxParityGridGrind.read(success, "layout", WorkbookReadResult.SheetLayoutResult.class),
        XlsxParityGridGrind.read(success, "print", WorkbookReadResult.PrintLayoutResult.class),
        XlsxParityGridGrind.read(
            success, "validations", WorkbookReadResult.DataValidationsResult.class),
        XlsxParityGridGrind.read(
            success, "formatting", WorkbookReadResult.ConditionalFormattingResult.class),
        XlsxParityGridGrind.read(
            success, "filters-ops", WorkbookReadResult.AutofiltersResult.class),
        XlsxParityGridGrind.read(
            success, "filters-queue", WorkbookReadResult.AutofiltersResult.class),
        XlsxParityGridGrind.read(success, "tables", WorkbookReadResult.TablesResult.class));
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
        instanceof dev.erst.gridgrind.protocol.dto.PaneReport.Frozen)) {
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

    WorkbookReadResult.NamedRangesResult coreRanges =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.readWorkbook(
                core.workbookPath(),
                new WorkbookReadOperation.GetNamedRanges("names", new NamedRangeSelection.All())),
            "names",
            WorkbookReadResult.NamedRangesResult.class);
    WorkbookReadResult.NamedRangesResult advancedRanges =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.readWorkbook(
                advanced.workbookPath(),
                new WorkbookReadOperation.GetNamedRanges("names", new NamedRangeSelection.All())),
            "names",
            WorkbookReadResult.NamedRangesResult.class);

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
    boolean parityAchieved =
        direct.structureLocked()
            && direct.windowsLocked()
            && direct.revisionLocked()
            && direct.passwordMatches()
            && XlsxParityGridGrind.hasReadType("GET_WORKBOOK_PROTECTION");
    return parityAchieved
        ? pass("Workbook-protection read parity is present.")
        : fail(
            "Workbook protection exists in POI, but GridGrind exposes no dedicated read surface.");
  }

  private static ProbeResult probeAdvancedPrintReadGap(ProbeContext context) {
    XlsxParityOracle.AdvancedPrintSnapshot direct =
        XlsxParityOracle.advancedPrint(
            context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING).workbookPath());
    Set<String> fields = componentNames(dev.erst.gridgrind.protocol.dto.PrintLayoutReport.class);
    boolean surfaceCapable =
        direct.leftMargin() > 0.0d
            && fields.contains("leftMargin")
            && fields.contains("rightMargin")
            && fields.contains("paperSize")
            && fields.contains("copies");
    return surfaceCapable
        ? pass("Advanced print-layout read parity is present.")
        : fail("POI exposes advanced print setup, but GridGrind's read report is core-only.");
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
            new WorkbookReadOperation.GetCells("cells", "Advanced", List.of("A3", "A4")));
    WorkbookReadResult.CellsResult cells =
        XlsxParityGridGrind.read(success, "cells", WorkbookReadResult.CellsResult.class);
    Map<String, GridGrindResponse.CellReport> byAddress = byAddress(cells.cells());
    Set<String> fontFields = componentNames(CellFontReport.class);
    Set<String> fillFields = componentNames(CellFillReport.class);
    boolean directAdvanced =
        themed.fontColorDescriptor().contains("theme=")
            || themed.fillColorDescriptor().contains("theme=")
            || themed.borderColorDescriptor().contains("indexed=")
            || gradient.gradientFill();
    boolean surfaceCapable =
        fontFields.contains("theme")
            && fontFields.contains("tint")
            && fillFields.contains("gradient")
            && fillFields.contains("theme");
    boolean observedCells = byAddress.size() == 2;
    return directAdvanced && observedCells && !surfaceCapable
        ? fail("GridGrind style reads normalize away theme/indexed/tint/gradient semantics.")
        : pass("Advanced style read parity is present.");
  }

  private static ProbeResult probeRichCommentReadGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario advanced =
        context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING);
    XlsxParityOracle.CommentSnapshot direct =
        XlsxParityOracle.comment(advanced.workbookPath(), "Advanced", "E2");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            advanced.workbookPath(),
            new WorkbookReadOperation.GetComments(
                "comments", "Advanced", new CellSelection.Selected(List.of("E2"))));
    WorkbookReadResult.CommentsResult comments =
        XlsxParityGridGrind.read(success, "comments", WorkbookReadResult.CommentsResult.class);
    boolean commentVisible =
        comments.comments().size() == 1
            && comments.comments().getFirst().comment().author().equals(direct.author());
    boolean surfaceCapable =
        componentNames(GridGrindResponse.CommentReport.class).contains("anchor")
            && componentNames(GridGrindResponse.CommentReport.class).contains("runs");
    return direct.runCount() > 1 && commentVisible && !surfaceCapable
        ? fail("GridGrind comment reads are plain-text only and omit anchor metadata.")
        : pass("Rich comment read parity is present.");
  }

  private static ProbeResult probeDataValidationAnalysisGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario advanced =
        context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING);
    List<String> directKinds =
        XlsxParityOracle.dataValidationKinds(advanced.workbookPath(), "Advanced");
    GridGrindResponse response =
        XlsxParityGridGrind.executeReadWorkbook(
            advanced.workbookPath(),
            new WorkbookReadOperation.AnalyzeDataValidationHealth(
                "health", new SheetSelection.Selected(List.of("Advanced"))));
    if (response instanceof GridGrindResponse.Failure failure) {
      return fail(
          "GridGrind data-validation analysis fails on malformed POI-authored rules: "
              + failure.problem().message());
    }
    GridGrindResponse.Success success = cast(GridGrindResponse.Success.class, response);
    WorkbookReadResult.DataValidationHealthResult health =
        XlsxParityGridGrind.read(
            success, "health", WorkbookReadResult.DataValidationHealthResult.class);
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
            new WorkbookReadOperation.GetAutofilters("filters", "Advanced"));
    WorkbookReadResult.AutofiltersResult filters =
        XlsxParityGridGrind.read(success, "filters", WorkbookReadResult.AutofiltersResult.class);
    boolean observed =
        filters.autofilters().size() == 1
            && filters.autofilters().getFirst().range().equals(direct.range());
    boolean surfaceCapable =
        componentNames(AutofilterEntryReport.SheetOwned.class).contains("filterColumns")
            && componentNames(AutofilterEntryReport.SheetOwned.class).contains("sortState");
    boolean directHasAdvancedState = direct.filterColumnCount() > 0 && direct.hasSortState();
    if (!directHasAdvancedState) {
      return fail(
          "Advanced autofilter scenario did not retain criteria/sort-state metadata: " + direct);
    }
    return observed && surfaceCapable
        ? pass("Autofilter criteria and sort-state read parity is present.")
        : fail(
            "GridGrind reads autofilter ownership but not criteria or sort state."
                + " direct="
                + direct
                + " observed="
                + observed
                + " surfaceCapable="
                + surfaceCapable);
  }

  private static ProbeResult probeTableAdvancedReadGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario advanced =
        context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING);
    XlsxParityOracle.TableSnapshot direct =
        XlsxParityOracle.table(advanced.workbookPath(), "AdvancedTable");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            advanced.workbookPath(),
            new WorkbookReadOperation.GetTables("tables", new TableSelection.All()));
    WorkbookReadResult.TablesResult tables =
        XlsxParityGridGrind.read(success, "tables", WorkbookReadResult.TablesResult.class);
    boolean observed =
        tables.tables().stream().anyMatch(table -> "AdvancedTable".equals(table.name()));
    boolean surfaceCapable =
        componentNames(TableEntryReport.class).contains("comment")
            && componentNames(TableEntryReport.class).contains("published")
            && componentNames(TableEntryReport.class).contains("totalsRowFunction");
    boolean directAdvanced =
        direct.comment() != null
            && direct.published()
            && direct.insertRow()
            && !direct.uniqueName().isBlank();
    return directAdvanced && observed && !surfaceCapable
        ? fail("GridGrind reads core table facts only and omits advanced table metadata.")
        : pass("Advanced table read parity is present.");
  }

  private static ProbeResult probeConditionalFormattingModeledRead(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario advanced =
        context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING);
    List<String> directKinds =
        XlsxParityOracle.conditionalFormattingKinds(advanced.workbookPath(), "Advanced");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            advanced.workbookPath(),
            new WorkbookReadOperation.GetConditionalFormatting(
                "formatting", "Advanced", new RangeSelection.All()));
    WorkbookReadResult.ConditionalFormattingResult formatting =
        XlsxParityGridGrind.read(
            success, "formatting", WorkbookReadResult.ConditionalFormattingResult.class);
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
            new WorkbookReadOperation.GetConditionalFormatting(
                "formatting", "Advanced", new RangeSelection.All()));
    WorkbookReadResult.ConditionalFormattingResult formatting =
        XlsxParityGridGrind.read(
            success, "formatting", WorkbookReadResult.ConditionalFormattingResult.class);
    boolean hasUnsupportedTop10 =
        formatting.conditionalFormattingBlocks().stream()
            .flatMap(block -> block.rules().stream())
            .filter(ConditionalFormattingRuleReport.UnsupportedRule.class::isInstance)
            .map(ConditionalFormattingRuleReport.UnsupportedRule.class::cast)
            .anyMatch(rule -> "TOP_10".equals(rule.kind()));
    return directKinds.stream().anyMatch(kind -> "top10".equalsIgnoreCase(kind))
            && hasUnsupportedTop10
        ? fail(
            "GridGrind still degrades unmodeled XSSF conditional-format families to UnsupportedRule.")
        : pass("Full conditional-format read parity is present.");
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
    XlsxParityOracle.WorkbookProtectionSnapshot direct =
        XlsxParityOracle.workbookProtection(
            context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING).workbookPath());
    boolean parityAchieved =
        direct.passwordMatches() && XlsxParityGridGrind.hasOperationType("SET_WORKBOOK_PROTECTION");
    return parityAchieved
        ? pass("Workbook-protection mutation parity is present.")
        : fail(
            "POI can author workbook protection, but GridGrind exposes no write operation for it.");
  }

  private static ProbeResult probeSheetPasswordMutationGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario advanced =
        context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING);
    boolean poiHasPassword =
        XlsxParitySupport.call(
            "inspect sheet-protection password in advanced parity workbook",
            () -> {
              try (XSSFWorkbook workbook =
                  (XSSFWorkbook) WorkbookFactory.create(advanced.workbookPath().toFile())) {
                return workbook
                    .getSheet("Advanced")
                    .validateSheetPassword(XlsxParityScenarios.SHEET_PROTECTION_PASSWORD);
              }
            });
    TypeEntry setSheetProtection = operationType("SET_SHEET_PROTECTION");
    boolean hasPasswordField = setSheetProtection.field("password") != null;
    return poiHasPassword && !hasPasswordField
        ? fail(
            "GridGrind can toggle sheet-protection flags but cannot set or preserve sheet passwords.")
        : pass("Password-bearing sheet-protection mutation parity is present.");
  }

  private static ProbeResult probeSheetCopyGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario core =
        context.copiedScenario(XlsxParityScenarios.CORE_WORKBOOK, "sheet-copy-core");
    boolean poiCopySucceeded =
        XlsxParitySupport.call(
            "clone table-bearing parity sheet with POI",
            () -> {
              try (XSSFWorkbook workbook =
                  (XSSFWorkbook) WorkbookFactory.create(core.workbookPath().toFile())) {
                workbook.cloneSheet(workbook.getSheetIndex("Queue"));
                return workbook.getNumberOfSheets() == 3;
              }
            });
    GridGrindResponse.Failure failure =
        XlsxParityGridGrind.mutateWorkbookExpectingFailure(
            core.workbookPath(),
            List.of(
                new WorkbookOperation.CopySheet(
                    "Queue", "Queue Replica", new SheetCopyPosition.AppendAtEnd())));
    return poiCopySucceeded
        ? fail(
            "POI can clone the table-bearing sheet, but GridGrind still rejects it: "
                + failure.problem().message())
        : pass("Sheet-copy parity is present.");
  }

  private static ProbeResult probeAdvancedPrintMutationGap(ProbeContext context) {
    XlsxParityOracle.AdvancedPrintSnapshot direct =
        XlsxParityOracle.advancedPrint(
            context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING).workbookPath());
    Set<String> fields = componentNames(dev.erst.gridgrind.protocol.dto.PrintLayoutInput.class);
    boolean surfaceCapable =
        fields.contains("leftMargin")
            && fields.contains("paperSize")
            && fields.contains("copies")
            && fields.contains("rowBreaks");
    return direct.copies() > 0 && !surfaceCapable
        ? fail("GridGrind has no authoring surface for advanced print setup.")
        : pass("Advanced print-layout mutation parity is present.");
  }

  private static ProbeResult probeAdvancedStyleMutationGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario advanced =
        context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING);
    XlsxParityOracle.StyleSnapshot themed =
        XlsxParityOracle.style(advanced.workbookPath(), "Advanced", "A3");
    XlsxParityOracle.StyleSnapshot gradient =
        XlsxParityOracle.style(advanced.workbookPath(), "Advanced", "A4");
    boolean directAdvanced =
        themed.fontColorDescriptor().contains("theme=")
            || themed.fillColorDescriptor().contains("theme=")
            || themed.borderColorDescriptor().contains("indexed=")
            || gradient.gradientFill();
    boolean surfaceCapable =
        componentNames(dev.erst.gridgrind.protocol.dto.CellFontInput.class).contains("theme")
            && componentNames(dev.erst.gridgrind.protocol.dto.CellFillInput.class)
                .contains("gradient");
    return directAdvanced && !surfaceCapable
        ? fail("GridGrind style mutation input only supports RGB-oriented style patches.")
        : pass("Advanced style mutation parity is present.");
  }

  private static ProbeResult probeRichCommentMutationGap(ProbeContext context) {
    XlsxParityOracle.CommentSnapshot direct =
        XlsxParityOracle.comment(
            context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING).workbookPath(),
            "Advanced",
            "E2");
    boolean surfaceCapable =
        componentNames(CommentInput.class).contains("anchor")
            && componentNames(CommentInput.class).contains("runs");
    return direct.runCount() > 1 && !surfaceCapable
        ? fail("GridGrind comment mutation input is plain-text only.")
        : pass("Rich comment mutation parity is present.");
  }

  private static ProbeResult probeNamedRangeFormulaMutationGap(ProbeContext context) {
    boolean directHasFormulaName =
        XlsxParityOracle.namedRanges(
                context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING).workbookPath())
            .stream()
            .anyMatch(XlsxParityOracle.NamedRangeSnapshot::formulaDefined);
    boolean hasFormulaField =
        componentNames(dev.erst.gridgrind.protocol.dto.NamedRangeTarget.class).contains("formula");
    return directHasFormulaName && !hasFormulaField
        ? fail("GridGrind can read formula-defined names but cannot author them.")
        : pass("Formula-defined named-range mutation parity is present.");
  }

  private static ProbeResult probeAutofilterCriteriaMutationGap(ProbeContext context) {
    XlsxParityOracle.AutofilterSnapshot direct =
        XlsxParityOracle.autofilter(
            context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING).workbookPath(), "Advanced");
    boolean surfaceCapable =
        componentNames(WorkbookOperation.SetAutofilter.class).contains("criteria")
            && componentNames(WorkbookOperation.SetAutofilter.class).contains("sortState");
    boolean directHasAdvancedState = direct.filterColumnCount() > 0 && direct.hasSortState();
    if (!directHasAdvancedState) {
      return fail(
          "Advanced autofilter scenario did not retain criteria/sort-state metadata: " + direct);
    }
    return surfaceCapable
        ? pass("Autofilter criteria mutation parity is present.")
        : fail(
            "GridGrind can author only bare autofilter ranges, not filter criteria or sort state.");
  }

  private static ProbeResult probeTableAdvancedMutationGap(ProbeContext context) {
    XlsxParityOracle.TableSnapshot direct =
        XlsxParityOracle.table(
            context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING).workbookPath(),
            "AdvancedTable");
    boolean surfaceCapable =
        componentNames(dev.erst.gridgrind.protocol.dto.TableInput.class).contains("comment")
            && componentNames(dev.erst.gridgrind.protocol.dto.TableInput.class)
                .contains("published")
            && componentNames(dev.erst.gridgrind.protocol.dto.TableInput.class)
                .contains("calculatedColumns");
    boolean directAdvanced =
        direct.comment() != null
            && direct.published()
            && !direct.calculatedColumnFormula().isBlank();
    return directAdvanced && !surfaceCapable
        ? fail("GridGrind table mutation input covers only the core table contract.")
        : pass("Advanced table mutation parity is present.");
  }

  private static ProbeResult probeConditionalFormattingAdvancedMutationGap(ProbeContext context) {
    boolean directHasUnmodeledFamily =
        XlsxParityOracle.conditionalFormattingKinds(
                context.scenario(XlsxParityScenarios.ADVANCED_NONDRAWING).workbookPath(),
                "Advanced")
            .stream()
            .anyMatch(kind -> "top10".equalsIgnoreCase(kind));
    Set<String> authoredKinds =
        Set.of(
            ConditionalFormattingRuleInput.FormulaRule.class.getSimpleName(),
            ConditionalFormattingRuleInput.CellValueRule.class.getSimpleName());
    return directHasUnmodeledFamily && authoredKinds.size() == 2
        ? fail("GridGrind authors only formula and cell-value conditional-format rules.")
        : pass("Advanced conditional-format mutation parity is present.");
  }

  private static ProbeResult probeFormulaCore(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario core =
        context.copiedScenario(XlsxParityScenarios.CORE_WORKBOOK, "formula-core-source");
    Path outputPath = context.derivedWorkbook("formula-core");
    XlsxParityGridGrind.mutateWorkbook(
        core.workbookPath(), outputPath, List.of(new WorkbookOperation.EvaluateFormulas()));
    double value = XlsxParityOracle.evaluateFormula(outputPath, "Ops", "D3");
    return value == 25.0d
        ? pass("GridGrind evaluates core in-workbook formulas correctly.")
        : fail("GridGrind formula evaluation diverged on the core workbook.");
  }

  private static ProbeResult probeFormulaExternalGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario external =
        context.copiedScenario(XlsxParityScenarios.EXTERNAL_FORMULA, "formula-external-source");
    double poiValue =
        XlsxParityOracle.evaluateExternalFormula(
            external.workbookPath(), external.attachment("referencedWorkbook"));
    GridGrindResponse.Failure failure =
        XlsxParityGridGrind.mutateWorkbookExpectingFailure(
            external.workbookPath(), List.of(new WorkbookOperation.EvaluateFormulas()));
    return poiValue == 7.5d
        ? fail(
            "POI can evaluate external workbook references, but GridGrind cannot configure them: "
                + failure.problem().message())
        : pass("External-workbook formula parity is present.");
  }

  private static ProbeResult probeFormulaUdfGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario udf =
        context.copiedScenario(XlsxParityScenarios.UDF_FORMULA, "formula-udf-source");
    double poiValue = XlsxParityOracle.evaluateUdfFormula(udf.workbookPath());
    GridGrindResponse.Failure failure =
        XlsxParityGridGrind.mutateWorkbookExpectingFailure(
            udf.workbookPath(), List.of(new WorkbookOperation.EvaluateFormulas()));
    return poiValue == 42.0d
        ? fail(
            "POI can evaluate registered UDFs, but GridGrind has no UDF registration seam: "
                + failure.problem().message())
        : pass("UDF-backed formula parity is present.");
  }

  private static ProbeResult probeFormulaLifecycleGap() {
    boolean poiHasLifecycleControls =
        XlsxParitySupport.call(
            "inspect POI formula-evaluator lifecycle methods",
            () ->
                org.apache.poi.ss.usermodel.FormulaEvaluator.class
                        .getMethod("clearAllCachedResultValues")
                        .getReturnType()
                        .equals(void.class)
                    && org.apache.poi.ss.usermodel.FormulaEvaluator.class
                        .getMethod("evaluateFormulaCell", org.apache.poi.ss.usermodel.Cell.class)
                        .getReturnType()
                        .getSimpleName()
                        .contains("CellType"));
    boolean gridGrindHasDedicatedOps =
        XlsxParityGridGrind.hasOperationType("EVALUATE_FORMULA_CELL")
            || XlsxParityGridGrind.hasOperationType("CLEAR_FORMULA_CACHES");
    return poiHasLifecycleControls && !gridGrindHasDedicatedOps
        ? fail("POI exposes targeted evaluator lifecycle controls that GridGrind does not surface.")
        : pass("Formula evaluator lifecycle parity is present.");
  }

  private static ProbeResult probeDrawingPreservation(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario drawing =
        context.copiedScenario(XlsxParityScenarios.DRAWING_IMAGE, "drawing-source");
    XlsxParityOracle.DrawingSnapshot before = XlsxParityOracle.drawing(drawing.workbookPath());
    Path outputPath = context.derivedWorkbook("drawing-preserved");
    XlsxParityGridGrind.mutateWorkbook(
        drawing.workbookPath(),
        outputPath,
        List.of(
            new WorkbookOperation.EnsureSheet("Scratch"),
            new WorkbookOperation.SetCell("Scratch", "A1", new CellInput.Text("Touch"))));
    XlsxParityOracle.DrawingSnapshot after = XlsxParityOracle.drawing(outputPath);
    return before.equals(after)
        ? pass("Unrelated GridGrind edits preserve existing pictures and anchors.")
        : fail("Unrelated GridGrind edits changed picture-backed drawing state.");
  }

  private static ProbeResult probeEmbeddedObjectPreservation(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario embedded =
        context.copiedScenario(XlsxParityScenarios.EMBEDDED_OBJECT, "embedded-object-source");
    XlsxParityOracle.EmbeddedObjectSnapshot before =
        XlsxParityOracle.embeddedObject(embedded.workbookPath());
    Path outputPath = context.derivedWorkbook("embedded-preserved");
    XlsxParityGridGrind.mutateWorkbook(
        embedded.workbookPath(),
        outputPath,
        List.of(
            new WorkbookOperation.EnsureSheet("Scratch"),
            new WorkbookOperation.SetCell("Scratch", "A1", new CellInput.Text("Touch"))));
    XlsxParityOracle.EmbeddedObjectSnapshot after = XlsxParityOracle.embeddedObject(outputPath);
    return before.equals(after)
        ? pass("Unrelated GridGrind edits preserve embedded OLE objects.")
        : fail("Unrelated GridGrind edits changed embedded-object package state.");
  }

  private static ProbeResult probeChartPreservation(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario chart =
        context.copiedScenario(XlsxParityScenarios.CHART, "chart-source");
    XlsxParityOracle.ChartSnapshot before = XlsxParityOracle.chart(chart.workbookPath());
    Path outputPath = context.derivedWorkbook("chart-preserved");
    XlsxParityGridGrind.mutateWorkbook(
        chart.workbookPath(),
        outputPath,
        List.of(new WorkbookOperation.SetCell("Chart", "A6", new CellInput.Text("Touch"))));
    XlsxParityOracle.ChartSnapshot after = XlsxParityOracle.chart(outputPath);
    return before.equals(after)
        ? pass("Unrelated GridGrind edits preserve chart parts.")
        : fail("Unrelated GridGrind edits changed chart package state.");
  }

  private static ProbeResult probePivotPreservation(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario pivot =
        context.copiedScenario(XlsxParityScenarios.PIVOT, "pivot-source");
    XlsxParityOracle.PivotSnapshot before = XlsxParityOracle.pivot(pivot.workbookPath());
    Path outputPath = context.derivedWorkbook("pivot-preserved");
    XlsxParityGridGrind.mutateWorkbook(
        pivot.workbookPath(),
        outputPath,
        List.of(new WorkbookOperation.SetCell("Data", "C1", new CellInput.Text("Touch"))));
    XlsxParityOracle.PivotSnapshot after = XlsxParityOracle.pivot(outputPath);
    return before.equals(after)
        ? pass("Unrelated GridGrind edits preserve pivot-table parts.")
        : fail("Unrelated GridGrind edits changed pivot-table package state.");
  }

  private static ProbeResult probeEventModelGap(ProbeContext context) {
    int poiRows =
        XlsxParityOracle.eventModelRowCount(
            context.scenario(XlsxParityScenarios.LARGE_SHEET).workbookPath());
    boolean gridGrindHasEventSource = XlsxParityGridGrind.hasSourceType("EVENT_READ");
    return poiRows >= 2000 && !gridGrindHasEventSource
        ? fail(
            "POI can read the workbook through the XSSF event model, but GridGrind has no such mode.")
        : pass("Event-model parity is present.");
  }

  private static ProbeResult probeSxssfGap(ProbeContext context) {
    Path streamedWorkbook = XlsxParityOracle.sxssfWriteWorkbook(context.derivedDirectory("sxssf"));
    boolean poiSucceeded =
        XlsxParitySupport.call(
            "inspect streamed parity workbook size",
            () -> Files.exists(streamedWorkbook) && Files.size(streamedWorkbook) > 0);
    boolean gridGrindHasStreamingWrite =
        XlsxParityGridGrind.hasPersistenceType("STREAMING_SAVE_AS")
            || XlsxParityGridGrind.hasOperationType("WRITE_STREAMING");
    return poiSucceeded && !gridGrindHasStreamingWrite
        ? fail("POI exposes SXSSF streaming write, but GridGrind does not.")
        : pass("SXSSF streaming-write parity is present.");
  }

  private static ProbeResult probeEncryptionGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario encrypted =
        context.scenario(XlsxParityScenarios.ENCRYPTED_WORKBOOK);
    boolean poiSucceeded = XlsxParityOracle.encryptedWorkbookOpens(encrypted.workbookPath());
    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                new dev.erst.gridgrind.protocol.dto.GridGrindRequest(
                    new dev.erst.gridgrind.protocol.dto.GridGrindRequest.WorkbookSource
                        .ExistingFile(encrypted.workbookPath().toString()),
                    new dev.erst.gridgrind.protocol.dto.GridGrindRequest.WorkbookPersistence.None(),
                    List.of(),
                    List.of()));
    boolean gridGrindFailed = response instanceof GridGrindResponse.Failure;
    return poiSucceeded && gridGrindFailed
        ? fail("POI can open the encrypted workbook with a password, but GridGrind cannot.")
        : pass("OOXML encryption parity is present.");
  }

  private static ProbeResult probeSigningGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario signed =
        context.copiedScenario(XlsxParityScenarios.SIGNED_WORKBOOK, "signed-source");
    boolean beforeValid = XlsxParityOracle.signatureValid(signed.workbookPath());
    Path outputPath = context.derivedWorkbook("signed-mutated");
    XlsxParityGridGrind.mutateWorkbook(
        signed.workbookPath(),
        outputPath,
        List.of(new WorkbookOperation.SetCell("Signed", "C1", new CellInput.Text("Touch"))));
    boolean afterValid = XlsxParityOracle.signatureValid(outputPath);
    boolean gridGrindHasSigningSurface =
        XlsxParityGridGrind.hasOperationType("SIGN_WORKBOOK")
            || XlsxParityGridGrind.hasReadType("GET_SIGNATURES");
    return beforeValid && !afterValid && !gridGrindHasSigningSurface
        ? fail(
            "GridGrind preserves the file format but cannot maintain or author OOXML signatures.")
        : pass("OOXML signing parity is present.");
  }

  private static <T> T cast(Class<T> type, Object value) {
    return type.cast(value);
  }

  private static Map<String, GridGrindResponse.CellReport> byAddress(
      List<GridGrindResponse.CellReport> cells) {
    return cells.stream()
        .collect(
            Collectors.toUnmodifiableMap(
                GridGrindResponse.CellReport::address, Function.identity()));
  }

  private static TypeEntry operationType(String id) {
    return XlsxParityGridGrind.catalog().operationTypes().stream()
        .filter(entry -> entry.id().equals(id))
        .findFirst()
        .orElseThrow(() -> new IllegalArgumentException("Unknown operation type " + id));
  }

  private static Set<String> componentNames(Class<?> recordType) {
    return Arrays.stream(recordType.getRecordComponents())
        .map(RecordComponent::getName)
        .collect(Collectors.toUnmodifiableSet());
  }

  private static ProbeResult pass(String detail) {
    return new ProbeResult(XlsxParityLedger.ExpectedOutcome.PASS, detail);
  }

  private static ProbeResult fail(String detail) {
    return new ProbeResult(XlsxParityLedger.ExpectedOutcome.FAIL, detail);
  }

  record ProbeResult(XlsxParityLedger.ExpectedOutcome outcome, String detail) {}

  private record CoreReadObservation(
      XlsxParityOracle.CoreWorkbookSnapshot direct,
      GridGrindResponse.WorkbookSummary.WithSheets summary,
      WorkbookReadResult.SheetSummaryResult opsSummary,
      WorkbookReadResult.SheetSummaryResult queueSummary,
      WorkbookReadResult.CellsResult cells,
      WorkbookReadResult.MergedRegionsResult merged,
      WorkbookReadResult.HyperlinksResult links,
      WorkbookReadResult.CommentsResult comments,
      WorkbookReadResult.SheetLayoutResult layout,
      WorkbookReadResult.PrintLayoutResult print,
      WorkbookReadResult.DataValidationsResult validations,
      WorkbookReadResult.ConditionalFormattingResult formatting,
      WorkbookReadResult.AutofiltersResult filtersOps,
      WorkbookReadResult.AutofiltersResult filtersQueue,
      WorkbookReadResult.TablesResult tables) {}

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
