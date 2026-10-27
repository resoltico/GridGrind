package dev.erst.gridgrind.fuzz.support;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookExecutionEngine;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/** Owns workbook-surface expectation capture for `.xlsx` round-trip verification. */
final class XlsxRoundTripExpectedWorkbookSurfaceSupport {
  private XlsxRoundTripExpectedWorkbookSurfaceSupport() {}

  static Map<String, Map<XlsxRoundTripExpectedStateSupport.CellCoordinate, CellStyleReport>>
      expectedStyles(
          Map<String, List<CellReport>> candidateSnapshots, CellStyleReport defaultStyle) {
    LinkedHashMap<String, Map<XlsxRoundTripExpectedStateSupport.CellCoordinate, CellStyleReport>>
        expectedStylesBySheet = new LinkedHashMap<>();
    for (Map.Entry<String, List<CellReport>> entry : candidateSnapshots.entrySet()) {
      LinkedHashMap<XlsxRoundTripExpectedStateSupport.CellCoordinate, CellStyleReport>
          expectedStyles = new LinkedHashMap<>();
      for (CellReport snapshot : entry.getValue()) {
        if (!snapshot.style().orElseThrow().equals(defaultStyle)) {
          expectedStyles.put(
              XlsxRoundTripExpectedStateSupport.CellCoordinate.fromAddress(
                  snapshot.address().value()),
              snapshot.style().orElseThrow());
        }
      }
      if (!expectedStyles.isEmpty()) {
        expectedStylesBySheet.put(entry.getKey(), Map.copyOf(expectedStyles));
      }
    }
    return Map.copyOf(expectedStylesBySheet);
  }

  static Map<
          String,
          Map<
              XlsxRoundTripExpectedStateSupport.CellCoordinate,
              XlsxRoundTripExpectedStateSupport.ExpectedCellMetadata>>
      expectedMetadata(Map<String, List<CellReport>> candidateSnapshots) {
    LinkedHashMap<
            String,
            Map<
                XlsxRoundTripExpectedStateSupport.CellCoordinate,
                XlsxRoundTripExpectedStateSupport.ExpectedCellMetadata>>
        expectedMetadataBySheet = new LinkedHashMap<>();
    for (Map.Entry<String, List<CellReport>> entry : candidateSnapshots.entrySet()) {
      LinkedHashMap<
              XlsxRoundTripExpectedStateSupport.CellCoordinate,
              XlsxRoundTripExpectedStateSupport.ExpectedCellMetadata>
          expectedMetadata = new LinkedHashMap<>();
      for (CellReport snapshot : entry.getValue()) {
        if (hasMetadata(snapshot)) {
          expectedMetadata.put(
              XlsxRoundTripExpectedStateSupport.CellCoordinate.fromAddress(
                  snapshot.address().value()),
              XlsxRoundTripExpectedStateSupport.ExpectedCellMetadata.from(snapshot));
        }
      }
      if (!expectedMetadata.isEmpty()) {
        expectedMetadataBySheet.put(entry.getKey(), Map.copyOf(expectedMetadata));
      }
    }
    return Map.copyOf(expectedMetadataBySheet);
  }

  static Map<String, Map<XlsxRoundTripExpectedStateSupport.CellCoordinate, List<RichTextRunReport>>>
      expectedRichText(Map<String, List<CellReport>> candidateSnapshots) {
    LinkedHashMap<
            String, Map<XlsxRoundTripExpectedStateSupport.CellCoordinate, List<RichTextRunReport>>>
        expectedRichTextBySheet = new LinkedHashMap<>();
    for (Map.Entry<String, List<CellReport>> entry : candidateSnapshots.entrySet()) {
      LinkedHashMap<XlsxRoundTripExpectedStateSupport.CellCoordinate, List<RichTextRunReport>>
          expectedRichText = new LinkedHashMap<>();
      for (CellReport snapshot : entry.getValue()) {
        if (snapshot instanceof CellReport.TextReport text && text.runs().isPresent()) {
          expectedRichText.put(
              XlsxRoundTripExpectedStateSupport.CellCoordinate.fromAddress(text.address().value()),
              text.runs().orElseThrow());
        }
      }
      if (!expectedRichText.isEmpty()) {
        expectedRichTextBySheet.put(entry.getKey(), Map.copyOf(expectedRichText));
      }
    }
    return Map.copyOf(expectedRichTextBySheet);
  }

  static Map<
          XlsxRoundTripExpectedStateSupport.NamedRangeKey,
          XlsxRoundTripExpectedStateSupport.ExpectedNamedRange>
      expectedNamedRanges(XssfWorkbook workbook) {
    LinkedHashMap<
            XlsxRoundTripExpectedStateSupport.NamedRangeKey,
            XlsxRoundTripExpectedStateSupport.ExpectedNamedRange>
        expectedNamedRanges = new LinkedHashMap<>();
    for (NamedRangeReport namedRange : workbook.names().namedRanges()) {
      NamedRangeTarget target =
          switch (namedRange) {
            case NamedRangeReport.RangeReport rangeSnapshot -> rangeSnapshot.target();
            case NamedRangeReport.FormulaReport _ -> null;
          };
      expectedNamedRanges.put(
          new XlsxRoundTripExpectedStateSupport.NamedRangeKey(
              namedRange.name().value(), namedRange.scope()),
          new XlsxRoundTripExpectedStateSupport.ExpectedNamedRange(
              namedRange.scope(), namedRange.refersToFormula(), target));
    }
    return Map.copyOf(expectedNamedRanges);
  }

  static dev.erst.gridgrind.core.result.WorkbookSummary expectedWorkbookSummary(
      XssfWorkbook workbook) {
    WorkbookExecutionEngine readExecutor = new WorkbookExecutionEngine();
    return ((dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookSummaryResult)
            readExecutor
                .read(
                    workbook,
                    new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                        new StepId("workbook-summary"),
                        new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                        new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                            .GetWorkbookSummary()))
                .getFirst())
        .workbook();
  }

  static Map<String, dev.erst.gridgrind.core.result.SheetSummaryReport> expectedSheetSummaries(
      XssfWorkbook workbook) {
    WorkbookExecutionEngine readExecutor = new WorkbookExecutionEngine();
    LinkedHashMap<String, dev.erst.gridgrind.core.result.SheetSummaryReport> expected =
        new LinkedHashMap<>();
    for (String sheetName : workbook.sheets().sheetNames()) {
      expected.put(
          sheetName,
          ((dev.erst.gridgrind.core.plan.SheetInspectionResult.SheetSummaryResult)
                  readExecutor
                      .read(
                          workbook,
                          new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                              new StepId("sheet-summary"),
                              new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                                  new dev.erst.gridgrind.core.vocabulary.SheetName(sheetName)),
                              new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery
                                  .GetSheetSummary()))
                      .getFirst())
              .sheet());
    }
    return Map.copyOf(expected);
  }

  static Map<String, List<DataValidationEntryReport>> expectedDataValidations(
      XssfWorkbook workbook) {
    LinkedHashMap<String, List<DataValidationEntryReport>> expected = new LinkedHashMap<>();
    for (String sheetName : workbook.sheets().sheetNames()) {
      List<DataValidationEntryReport> validations =
          workbook.sheet(sheetName).metadata().dataValidations();
      if (!validations.isEmpty()) {
        expected.put(sheetName, validations);
      }
    }
    return Map.copyOf(expected);
  }

  static Map<String, List<ConditionalFormattingEntryReport>> expectedConditionalFormatting(
      XssfWorkbook workbook) {
    WorkbookExecutionEngine readExecutor = new WorkbookExecutionEngine();
    LinkedHashMap<String, List<ConditionalFormattingEntryReport>> expected = new LinkedHashMap<>();
    for (String sheetName : workbook.sheets().sheetNames()) {
      var result =
          (dev.erst.gridgrind.core.plan.SheetInspectionResult.ConditionalFormattingResult)
              readExecutor
                  .read(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("conditionalFormatting"),
                          new dev.erst.gridgrind.core.plan.RangeSelector.AllOnSheet(
                              new dev.erst.gridgrind.core.vocabulary.SheetName(sheetName)),
                          new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery
                              .GetConditionalFormatting()))
                  .getFirst();
      if (!result.conditionalFormattingBlocks().isEmpty()) {
        expected.put(sheetName, result.conditionalFormattingBlocks());
      }
    }
    return Map.copyOf(expected);
  }

  static Map<String, List<AutofilterEntryReport>> expectedAutofilters(XssfWorkbook workbook) {
    WorkbookExecutionEngine readExecutor = new WorkbookExecutionEngine();
    LinkedHashMap<String, List<AutofilterEntryReport>> expected = new LinkedHashMap<>();
    for (String sheetName : workbook.sheets().sheetNames()) {
      var result =
          (dev.erst.gridgrind.core.plan.SheetInspectionResult.AutofiltersResult)
              readExecutor
                  .read(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("autofilters"),
                          new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                              new dev.erst.gridgrind.core.vocabulary.SheetName(sheetName)),
                          new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery
                              .GetAutofilters()))
                  .getFirst();
      if (!result.autofilters().isEmpty()) {
        expected.put(sheetName, result.autofilters());
      }
    }
    return Map.copyOf(expected);
  }

  static Map<String, List<DrawingObjectReport>> expectedDrawingObjects(XssfWorkbook workbook) {
    LinkedHashMap<String, List<DrawingObjectReport>> expected = new LinkedHashMap<>();
    for (String sheetName : workbook.sheets().sheetNames()) {
      expected.put(sheetName, List.copyOf(workbook.sheet(sheetName).drawings().drawingObjects()));
    }
    return Map.copyOf(expected);
  }

  static Map<String, List<ChartReport>> expectedCharts(XssfWorkbook workbook) {
    LinkedHashMap<String, List<ChartReport>> expected = new LinkedHashMap<>();
    for (String sheetName : workbook.sheets().sheetNames()) {
      expected.put(sheetName, List.copyOf(workbook.sheet(sheetName).drawings().charts()));
    }
    return Map.copyOf(expected);
  }

  static List<PivotTableReport> expectedPivots(XssfWorkbook workbook) {
    WorkbookExecutionEngine readExecutor = new WorkbookExecutionEngine();
    var result =
        (dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.PivotTablesResult)
            readExecutor
                .read(
                    workbook,
                    new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                        new StepId("pivots"),
                        new PivotTableSelector.All(),
                        new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery
                            .GetPivotTables()))
                .getFirst();
    return List.copyOf(result.pivotTables());
  }

  static List<TableEntryReport> expectedTables(XssfWorkbook workbook) {
    WorkbookExecutionEngine readExecutor = new WorkbookExecutionEngine();
    var result =
        (dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.TablesResult)
            readExecutor
                .read(
                    workbook,
                    new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                        new StepId("tables"),
                        new TableSelector.All(),
                        new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery
                            .GetTables()))
                .getFirst();
    return List.copyOf(result.tables());
  }

  static boolean hasMetadata(CellReport metadata) {
    return metadata.hyperlink().isPresent() || metadata.comment().isPresent();
  }
}
