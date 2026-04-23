package dev.erst.gridgrind.executor.parity;

import static dev.erst.gridgrind.executor.parity.ParityPlanSupport.inspect;
import static dev.erst.gridgrind.executor.parity.ParityPlanSupport.mutate;
import static dev.erst.gridgrind.executor.parity.XlsxParityProbeRegistry.*;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.executor.parity.XlsxParityProbeRegistry.ProbeContext;
import dev.erst.gridgrind.executor.parity.XlsxParityProbeRegistry.ProbeResult;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

/** Core workbook and mutation-gap parity probes. */
final class XlsxParityCoreProbeGroup {
  private XlsxParityCoreProbeGroup() {}

  static ProbeResult probeCoreRead(ProbeContext context) {
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
        != dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility.HIDDEN) {
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
        != dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation.LANDSCAPE) {
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

  static ProbeResult probeNamedRangeRead(ProbeContext context) {
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

  static ProbeResult probeWorkbookProtectionReadGap(ProbeContext context) {
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

  static ProbeResult probeAdvancedPrintReadGap(ProbeContext context) {
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

  static ProbeResult probeAdvancedStyleReadGap(ProbeContext context) {
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

  static ProbeResult probeRichCommentReadGap(ProbeContext context) {
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

  static ProbeResult probeDataValidationAnalysisGap(ProbeContext context) {
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

  static ProbeResult probeAutofilterCriteriaReadGap(ProbeContext context) {
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

  static ProbeResult probeTableAdvancedReadGap(ProbeContext context) {
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

  static ProbeResult probeConditionalFormattingModeledRead(ProbeContext context) {
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

  static ProbeResult probeConditionalFormattingUnmodeledReadGap(ProbeContext context) {
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

  static ProbeResult probeCoreMutation(ProbeContext context) {
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

  static ProbeResult probeWorkbookProtectionMutationGap(ProbeContext context) {
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

  static ProbeResult probeSheetPasswordMutationGap(ProbeContext context) {
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

  static ProbeResult probeSheetCopyGap(ProbeContext context) {
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

  static ProbeResult probeAdvancedPrintMutationGap(ProbeContext context) {
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

  static ProbeResult probeAdvancedStyleMutationGap(ProbeContext context) {
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

  static ProbeResult probeRichCommentMutationGap(ProbeContext context) {
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

  static ProbeResult probeNamedRangeFormulaMutationGap(ProbeContext context) {
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

  static ProbeResult probeAutofilterCriteriaMutationGap(ProbeContext context) {
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

  static ProbeResult probeTableAdvancedMutationGap(ProbeContext context) {
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

  static ProbeResult probeConditionalFormattingAdvancedMutationGap(ProbeContext context) {
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

  record CoreReadObservation(
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

  record SheetCopySourceObservation(
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

  record SheetCopyCopiedObservation(
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
}
