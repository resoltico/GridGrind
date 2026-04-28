package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.CellColorReport;
import dev.erst.gridgrind.contract.dto.CellFillInput;
import dev.erst.gridgrind.contract.dto.CellFillReport;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ColorInput;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator;
import dev.erst.gridgrind.contract.dto.DrawingAnchorReport;
import dev.erst.gridgrind.contract.dto.DrawingMarkerReport;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelBorderSideSnapshot;
import dev.erst.gridgrind.excel.ExcelBorderSnapshot;
import dev.erst.gridgrind.excel.ExcelCellAlignmentSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFill;
import dev.erst.gridgrind.excel.ExcelCellFillSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFontSnapshot;
import dev.erst.gridgrind.excel.ExcelCellProtectionSnapshot;
import dev.erst.gridgrind.excel.ExcelCellStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelColor;
import dev.erst.gridgrind.excel.ExcelColorSnapshot;
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
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
                    ExcelColorSnapshot.rgb("#112233"),
                    false,
                    false))
            .fontName());
    assertEquals(
        ExcelBorderStyle.THIN,
        InspectionResultCellReportSupport.toCellBorderSideReport(
                new ExcelBorderSideSnapshot(
                    ExcelBorderStyle.THIN, ExcelColorSnapshot.rgb("#112233")))
            .style());
  }

  @Test
  void styleConvertersCoverPatternOnlyAndBackgroundOnlyVariants() {
    ExcelCellFill.PatternOnly patternOnly =
        assertInstanceOf(
            ExcelCellFill.PatternOnly.class,
            WorkbookCommandConverter.toExcelCellFill(CellFillInput.pattern(ExcelFillPattern.NONE))
                .orElseThrow());
    ExcelCellFill.PatternBackground backgroundOnlyInput =
        assertInstanceOf(
            ExcelCellFill.PatternBackground.class,
            WorkbookCommandConverter.toExcelCellFill(
                    CellFillInput.patternBackground(
                        ExcelFillPattern.BRICKS, ColorInput.indexed(10)))
                .orElseThrow());
    GridGrindResponse.CellStyleReport backgroundOnlyReport =
        InspectionResultCellReportSupport.toCellStyleReport(
            new ExcelCellStyleSnapshot(
                "General",
                new ExcelCellAlignmentSnapshot(
                    false, ExcelHorizontalAlignment.LEFT, ExcelVerticalAlignment.BOTTOM, 0, 0),
                new ExcelCellFontSnapshot(
                    false, false, "Aptos", new ExcelFontHeight(220), null, false, false),
                ExcelCellFillSnapshot.patternBackground(
                    ExcelFillPattern.BRICKS, ExcelColorSnapshot.indexed(10)),
                new ExcelBorderSnapshot(
                    new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
                    new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
                    new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
                    new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null)),
                new ExcelCellProtectionSnapshot(true, false)));
    CellFillReport.PatternBackground backgroundOnlyFill =
        assertInstanceOf(CellFillReport.PatternBackground.class, backgroundOnlyReport.fill());

    assertEquals(ExcelFillPattern.NONE, patternOnly.pattern());
    assertEquals(ExcelFillPattern.BRICKS, backgroundOnlyInput.pattern());
    assertEquals(ExcelColor.indexed(10), backgroundOnlyInput.backgroundColor());
    assertEquals(ExcelFillPattern.BRICKS, backgroundOnlyFill.pattern());
    assertEquals(CellColorReport.indexed(10), backgroundOnlyFill.backgroundColor());
    assertNull(ProtocolStyleTestAccess.fillForegroundColor(backgroundOnlyReport.fill()));
  }

  @Test
  void familyMutationConvertersHandleRepresentativeInFamilyActions() {
    MutationAction.SetCell setCell = new MutationAction.SetCell(new CellInput.Blank());

    WorkbookCommand.CreateSheet createSheet =
        assertInstanceOf(
            WorkbookCommand.CreateSheet.class,
            WorkbookCommandWorkbookMutationConverter.toCommand(
                new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()));
    WorkbookCommand.SetCell convertedCell =
        assertInstanceOf(
            WorkbookCommand.SetCell.class,
            WorkbookCommandCellMutationConverter.toCommand(
                new CellSelector.ByAddress("Budget", "A1"), setCell));
    WorkbookCommand.DeleteDrawingObject deleteDrawingObject =
        assertInstanceOf(
            WorkbookCommand.DeleteDrawingObject.class,
            WorkbookCommandDrawingMutationConverter.toCommand(
                new DrawingObjectSelector.ByName("Budget", "Logo"),
                new MutationAction.DeleteDrawingObject()));
    WorkbookCommand.ClearAutofilter clearAutofilter =
        assertInstanceOf(
            WorkbookCommand.ClearAutofilter.class,
            WorkbookCommandStructuredMutationConverter.toCommand(
                new SheetSelector.ByName("Budget"), new MutationAction.ClearAutofilter()));

    assertEquals("Budget", createSheet.sheetName());
    assertEquals("Budget", convertedCell.sheetName());
    assertEquals("A1", convertedCell.address());
    assertEquals("Budget", deleteDrawingObject.sheetName());
    assertEquals("Logo", deleteDrawingObject.objectName());
    assertEquals("Budget", clearAutofilter.sheetName());
  }
}
