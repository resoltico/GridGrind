package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.ExcelAutofilterSnapshot;
import dev.erst.gridgrind.excel.ExcelCellMetadataSnapshot;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelCellStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelChartSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockSnapshot;
import dev.erst.gridgrind.excel.ExcelDataValidationSnapshot;
import dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot;
import dev.erst.gridgrind.excel.ExcelNamedRangeSnapshot;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelPivotTableSelection;
import dev.erst.gridgrind.excel.ExcelPivotTableSnapshot;
import dev.erst.gridgrind.excel.ExcelRangeSelection;
import dev.erst.gridgrind.excel.ExcelRichTextSnapshot;
import dev.erst.gridgrind.excel.ExcelTableSelection;
import dev.erst.gridgrind.excel.ExcelTableSnapshot;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/** Owns workbook-surface expectation capture for `.xlsx` round-trip verification. */
final class XlsxRoundTripExpectedWorkbookSurfaceSupport {
  private XlsxRoundTripExpectedWorkbookSurfaceSupport() {}

  static Map<String, Map<XlsxRoundTripExpectedStateSupport.CellCoordinate, ExcelCellStyleSnapshot>>
      expectedStyles(
          Map<String, List<ExcelCellSnapshot>> candidateSnapshots,
          ExcelCellStyleSnapshot defaultStyle) {
    LinkedHashMap<
            String, Map<XlsxRoundTripExpectedStateSupport.CellCoordinate, ExcelCellStyleSnapshot>>
        expectedStylesBySheet = new LinkedHashMap<>();
    for (Map.Entry<String, List<ExcelCellSnapshot>> entry : candidateSnapshots.entrySet()) {
      LinkedHashMap<XlsxRoundTripExpectedStateSupport.CellCoordinate, ExcelCellStyleSnapshot>
          expectedStyles = new LinkedHashMap<>();
      for (ExcelCellSnapshot snapshot : entry.getValue()) {
        if (!snapshot.style().equals(defaultStyle)) {
          expectedStyles.put(
              XlsxRoundTripExpectedStateSupport.CellCoordinate.fromAddress(snapshot.address()),
              snapshot.style());
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
      expectedMetadata(Map<String, List<ExcelCellSnapshot>> candidateSnapshots) {
    LinkedHashMap<
            String,
            Map<
                XlsxRoundTripExpectedStateSupport.CellCoordinate,
                XlsxRoundTripExpectedStateSupport.ExpectedCellMetadata>>
        expectedMetadataBySheet = new LinkedHashMap<>();
    for (Map.Entry<String, List<ExcelCellSnapshot>> entry : candidateSnapshots.entrySet()) {
      LinkedHashMap<
              XlsxRoundTripExpectedStateSupport.CellCoordinate,
              XlsxRoundTripExpectedStateSupport.ExpectedCellMetadata>
          expectedMetadata = new LinkedHashMap<>();
      for (ExcelCellSnapshot snapshot : entry.getValue()) {
        if (hasMetadata(snapshot.metadata())) {
          expectedMetadata.put(
              XlsxRoundTripExpectedStateSupport.CellCoordinate.fromAddress(snapshot.address()),
              XlsxRoundTripExpectedStateSupport.ExpectedCellMetadata.from(snapshot.metadata()));
        }
      }
      if (!expectedMetadata.isEmpty()) {
        expectedMetadataBySheet.put(entry.getKey(), Map.copyOf(expectedMetadata));
      }
    }
    return Map.copyOf(expectedMetadataBySheet);
  }

  static Map<String, Map<XlsxRoundTripExpectedStateSupport.CellCoordinate, ExcelRichTextSnapshot>>
      expectedRichText(Map<String, List<ExcelCellSnapshot>> candidateSnapshots) {
    LinkedHashMap<
            String, Map<XlsxRoundTripExpectedStateSupport.CellCoordinate, ExcelRichTextSnapshot>>
        expectedRichTextBySheet = new LinkedHashMap<>();
    for (Map.Entry<String, List<ExcelCellSnapshot>> entry : candidateSnapshots.entrySet()) {
      LinkedHashMap<XlsxRoundTripExpectedStateSupport.CellCoordinate, ExcelRichTextSnapshot>
          expectedRichText = new LinkedHashMap<>();
      for (ExcelCellSnapshot snapshot : entry.getValue()) {
        if (snapshot instanceof ExcelCellSnapshot.TextSnapshot text && text.richText() != null) {
          expectedRichText.put(
              XlsxRoundTripExpectedStateSupport.CellCoordinate.fromAddress(text.address()),
              text.richText());
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
      expectedNamedRanges(ExcelWorkbook workbook) {
    LinkedHashMap<
            XlsxRoundTripExpectedStateSupport.NamedRangeKey,
            XlsxRoundTripExpectedStateSupport.ExpectedNamedRange>
        expectedNamedRanges = new LinkedHashMap<>();
    for (ExcelNamedRangeSnapshot namedRange : workbook.namedRanges()) {
      ExcelNamedRangeTarget target =
          switch (namedRange) {
            case ExcelNamedRangeSnapshot.RangeSnapshot rangeSnapshot -> rangeSnapshot.target();
            case ExcelNamedRangeSnapshot.FormulaSnapshot _ -> null;
          };
      expectedNamedRanges.put(
          new XlsxRoundTripExpectedStateSupport.NamedRangeKey(
              namedRange.name(), namedRange.scope()),
          new XlsxRoundTripExpectedStateSupport.ExpectedNamedRange(
              namedRange.scope(), namedRange.refersToFormula(), target));
    }
    return Map.copyOf(expectedNamedRanges);
  }

  static dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary expectedWorkbookSummary(
      ExcelWorkbook workbook) {
    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    return ((dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummaryResult)
            readExecutor
                .apply(workbook, new WorkbookReadCommand.GetWorkbookSummary("workbook-summary"))
                .getFirst())
        .workbook();
  }

  static Map<String, dev.erst.gridgrind.excel.WorkbookSheetResult.SheetSummary>
      expectedSheetSummaries(ExcelWorkbook workbook) {
    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    LinkedHashMap<String, dev.erst.gridgrind.excel.WorkbookSheetResult.SheetSummary> expected =
        new LinkedHashMap<>();
    for (String sheetName : workbook.sheetNames()) {
      expected.put(
          sheetName,
          ((dev.erst.gridgrind.excel.WorkbookSheetResult.SheetSummaryResult)
                  readExecutor
                      .apply(
                          workbook,
                          new WorkbookReadCommand.GetSheetSummary(
                              "sheet-summary-" + sheetName, sheetName))
                      .getFirst())
              .sheet());
    }
    return Map.copyOf(expected);
  }

  static Map<String, List<ExcelDataValidationSnapshot>> expectedDataValidations(
      ExcelWorkbook workbook) {
    LinkedHashMap<String, List<ExcelDataValidationSnapshot>> expected = new LinkedHashMap<>();
    for (String sheetName : workbook.sheetNames()) {
      List<ExcelDataValidationSnapshot> validations =
          workbook.sheet(sheetName).dataValidations(new ExcelRangeSelection.All());
      if (!validations.isEmpty()) {
        expected.put(sheetName, validations);
      }
    }
    return Map.copyOf(expected);
  }

  static Map<String, List<ExcelConditionalFormattingBlockSnapshot>> expectedConditionalFormatting(
      ExcelWorkbook workbook) {
    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    LinkedHashMap<String, List<ExcelConditionalFormattingBlockSnapshot>> expected =
        new LinkedHashMap<>();
    for (String sheetName : workbook.sheetNames()) {
      var result =
          (dev.erst.gridgrind.excel.WorkbookRuleResult.ConditionalFormattingResult)
              readExecutor
                  .apply(
                      workbook,
                      new WorkbookReadCommand.GetConditionalFormatting(
                          "conditionalFormatting", sheetName, new ExcelRangeSelection.All()))
                  .getFirst();
      if (!result.conditionalFormattingBlocks().isEmpty()) {
        expected.put(sheetName, result.conditionalFormattingBlocks());
      }
    }
    return Map.copyOf(expected);
  }

  static Map<String, List<ExcelAutofilterSnapshot>> expectedAutofilters(ExcelWorkbook workbook) {
    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    LinkedHashMap<String, List<ExcelAutofilterSnapshot>> expected = new LinkedHashMap<>();
    for (String sheetName : workbook.sheetNames()) {
      var result =
          (dev.erst.gridgrind.excel.WorkbookRuleResult.AutofiltersResult)
              readExecutor
                  .apply(workbook, new WorkbookReadCommand.GetAutofilters("autofilters", sheetName))
                  .getFirst();
      if (!result.autofilters().isEmpty()) {
        expected.put(sheetName, result.autofilters());
      }
    }
    return Map.copyOf(expected);
  }

  static Map<String, List<ExcelDrawingObjectSnapshot>> expectedDrawingObjects(
      ExcelWorkbook workbook) {
    LinkedHashMap<String, List<ExcelDrawingObjectSnapshot>> expected = new LinkedHashMap<>();
    for (String sheetName : workbook.sheetNames()) {
      expected.put(sheetName, List.copyOf(workbook.sheet(sheetName).drawingObjects()));
    }
    return Map.copyOf(expected);
  }

  static Map<String, List<ExcelChartSnapshot>> expectedCharts(ExcelWorkbook workbook) {
    LinkedHashMap<String, List<ExcelChartSnapshot>> expected = new LinkedHashMap<>();
    for (String sheetName : workbook.sheetNames()) {
      expected.put(sheetName, List.copyOf(workbook.sheet(sheetName).charts()));
    }
    return Map.copyOf(expected);
  }

  static List<ExcelPivotTableSnapshot> expectedPivots(ExcelWorkbook workbook) {
    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    var result =
        (dev.erst.gridgrind.excel.WorkbookDrawingResult.PivotTablesResult)
            readExecutor
                .apply(
                    workbook,
                    new WorkbookReadCommand.GetPivotTables(
                        "pivots", new ExcelPivotTableSelection.All()))
                .getFirst();
    return List.copyOf(result.pivotTables());
  }

  static List<ExcelTableSnapshot> expectedTables(ExcelWorkbook workbook) {
    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    var result =
        (dev.erst.gridgrind.excel.WorkbookRuleResult.TablesResult)
            readExecutor
                .apply(
                    workbook,
                    new WorkbookReadCommand.GetTables("tables", new ExcelTableSelection.All()))
                .getFirst();
    return List.copyOf(result.tables());
  }

  static boolean hasMetadata(ExcelCellMetadataSnapshot metadata) {
    return metadata.hyperlink().isPresent() || metadata.comment().isPresent();
  }
}
