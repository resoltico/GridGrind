package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator;
import dev.erst.gridgrind.contract.dto.DrawingAnchorReport;
import dev.erst.gridgrind.contract.dto.DrawingMarkerReport;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelBorderSideSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFontSnapshot;
import dev.erst.gridgrind.excel.ExcelColorSnapshot;
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Covers delegation-only converter wrappers introduced by executor seam refactors. */
class ExecutorConverterDelegationCoverageTest {
  @Test
  void workbookCommandConverterDelegatesRemainCovered() {
    RichTextRunInput run = new RichTextRunInput(TextSourceInput.inline("Bold"), null);

    assertEquals(
        7L,
        WorkbookCommandConverter.toExcelCustomXmlMappingLocator(
                new CustomXmlMappingLocator(7L, "CourseMap"))
            .mapId());
    assertEquals(
        "Bold",
        WorkbookCommandConverter.toExcelRichText(new CellInput.RichText(List.of(run)))
            .runs()
            .getFirst()
            .text());
    assertEquals("Bold", WorkbookCommandConverter.toExcelRichTextRun(run).text());
    assertTrue(WorkbookCommandConverter.toExcelCellAlignment(null).isEmpty());
    assertTrue(WorkbookCommandConverter.toExcelCellFont(null).isEmpty());
    assertTrue(WorkbookCommandConverter.toExcelCellFill(null).isEmpty());
    assertTrue(WorkbookCommandConverter.toExcelCellProtection(null).isEmpty());
    assertTrue(WorkbookCommandConverter.toExcelBorderSide(null).isEmpty());
    assertTrue(WorkbookCommandConverter.toExcelDifferentialStyle(null).isEmpty());
  }

  @Test
  void inspectionResultConversionSupportDelegatesRemainCovered() {
    ExcelDrawingMarker marker = new ExcelDrawingMarker(1, 2, 3, 4);
    DrawingMarkerReport markerReport =
        InspectionResultDrawingReportSupport.toDrawingMarkerReport(marker);
    DrawingAnchorReport.TwoCell anchorReport =
        assertInstanceOf(
            DrawingAnchorReport.TwoCell.class,
            InspectionResultDrawingReportSupport.toDrawingAnchorReport(
                new ExcelDrawingAnchor.TwoCell(
                    marker,
                    new ExcelDrawingMarker(5, 6, 7, 8),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE)));

    assertEquals(1, markerReport.columnIndex());
    assertEquals(5, anchorReport.to().columnIndex());
    assertEquals(
        "Calibri",
        InspectionResultCellReportSupport.toCellFontReport(
                new ExcelCellFontSnapshot(
                    true,
                    false,
                    "Calibri",
                    new ExcelFontHeight(220),
                    new ExcelColorSnapshot("#112233"),
                    false,
                    false))
            .fontName());
    assertEquals(
        ExcelBorderStyle.THIN,
        InspectionResultCellReportSupport.toCellBorderSideReport(
                new ExcelBorderSideSnapshot(
                    ExcelBorderStyle.THIN, new ExcelColorSnapshot("#112233")))
            .style());
  }

  @Test
  void familyMutationConvertersReturnNullForOutOfFamilyActions() {
    MutationAction.SetCell setCell = new MutationAction.SetCell(new CellInput.Blank());

    assertTrue(
        WorkbookCommandWorkbookMutationConverter.toCommand(
                new CellSelector.ByAddress("Budget", "A1"), setCell)
            .isEmpty());
    assertTrue(
        WorkbookCommandCellMutationConverter.toCommand(
                new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet())
            .isEmpty());
    assertTrue(
        WorkbookCommandDrawingMutationConverter.toCommand(
                new CellSelector.ByAddress("Budget", "A1"), setCell)
            .isEmpty());
    assertTrue(
        WorkbookCommandStructuredMutationConverter.toCommand(
                new CellSelector.ByAddress("Budget", "A1"), setCell)
            .isEmpty());
  }
}
