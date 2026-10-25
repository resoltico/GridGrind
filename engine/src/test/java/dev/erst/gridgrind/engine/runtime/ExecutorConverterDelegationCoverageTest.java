package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFillInput;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.ColorInput;
import dev.erst.gridgrind.core.model.style.RichTextRunInput;
import dev.erst.gridgrind.core.model.value.CellInput;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.source.TextSourceInput;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.excel.*;
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
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingMarker;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Covers delegation-only converter wrappers introduced by executor seam refactors. */
class ExecutorConverterDelegationCoverageTest {
  @Test
  void workbookCommandConverterDelegatesRemainCovered() {
    RichTextRunInput run =
        new RichTextRunInput(TextSourceInput.inline("Bold"), java.util.Optional.empty());

    assertEquals(
        7L,
        WorkbookCommandDrawingInputConverter.toExcelCustomXmlMappingLocator(
                new CustomXmlMappingLocator(7L, "CourseMap"))
            .mapId());
    assertEquals(
        "Bold",
        WorkbookCommandCellInputConverter.toExcelRichText(new CellInput.RichText(List.of(run)))
            .runs()
            .getFirst()
            .text());
    assertEquals("Bold", WorkbookCommandCellInputConverter.toExcelRichTextRun(run).text());
    assertTrue(WorkbookCommandCellInputConverter.toExcelCellAlignment(null).isEmpty());
    assertTrue(WorkbookCommandCellInputConverter.toExcelCellFont(null).isEmpty());
    assertTrue(WorkbookCommandCellInputConverter.toExcelCellFill(null).isEmpty());
    assertTrue(WorkbookCommandCellInputConverter.toExcelCellProtection(null).isEmpty());
    assertTrue(WorkbookCommandCellInputConverter.toExcelBorderSide(null).isEmpty());
    assertTrue(WorkbookCommandStructuredInputConverter.toExcelDifferentialStyle(null).isEmpty());
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
        InspectionResultCellStyleReportSupport.toCellFontReport(
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
        assertInstanceOf(
                CellBorderSideReport.Colored.class,
                InspectionResultCellStyleReportSupport.toCellBorderSideReport(
                    new ExcelBorderSideSnapshot(
                        ExcelBorderStyle.THIN, ExcelColorSnapshot.rgb("#112233"))))
            .style());
  }

  @Test
  void styleConvertersCoverPatternOnlyAndBackgroundOnlyVariants() {
    ExcelCellFill.PatternOnly patternOnly =
        assertInstanceOf(
            ExcelCellFill.PatternOnly.class,
            WorkbookCommandCellInputConverter.toExcelCellFill(
                    CellFillInput.pattern(ExcelFillPattern.NONE))
                .orElseThrow());
    ExcelCellFill.PatternBackground backgroundOnlyInput =
        assertInstanceOf(
            ExcelCellFill.PatternBackground.class,
            WorkbookCommandCellInputConverter.toExcelCellFill(
                    CellFillInput.patternBackground(
                        ExcelFillPattern.BRICKS, ColorInput.indexed(10)))
                .orElseThrow());
    CellStyleReport backgroundOnlyReport =
        InspectionResultCellStyleReportSupport.toCellStyleReport(
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
    CellMutationAction.SetCell setCell = new CellMutationAction.SetCell(new CellInput.Blank());

    WorkbookSheetCommand.CreateSheet createSheet =
        assertInstanceOf(
            WorkbookSheetCommand.CreateSheet.class,
            WorkbookCommandWorkbookMutationConverter.toCommand(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.EnsureSheet()));
    WorkbookCellCommand.SetCell convertedCell =
        assertInstanceOf(
            WorkbookCellCommand.SetCell.class,
            WorkbookCommandCellMutationConverter.toCommand(
                new CellSelector.ByAddress(new SheetName("Budget"), "A1"), setCell));
    WorkbookDrawingCommand.DeleteDrawingObject deleteDrawingObject =
        assertInstanceOf(
            WorkbookDrawingCommand.DeleteDrawingObject.class,
            WorkbookCommandDrawingMutationConverter.toCommand(
                new DrawingObjectSelector.ByName(new SheetName("Budget"), "Logo"),
                new DrawingMutationAction.DeleteDrawingObject()));
    WorkbookTabularCommand.ClearAutofilter clearAutofilter =
        assertInstanceOf(
            WorkbookTabularCommand.ClearAutofilter.class,
            WorkbookCommandStructuredMutationConverter.toCommand(
                new SheetSelector.ByName(new SheetName("Budget")),
                new StructuredMutationAction.ClearAutofilter()));

    assertEquals("Budget", createSheet.sheetName());
    assertEquals("Budget", convertedCell.sheetName());
    assertEquals("A1", convertedCell.address());
    assertEquals("Budget", deleteDrawingObject.sheetName());
    assertEquals("Logo", deleteDrawingObject.objectName());
    assertEquals("Budget", clearAutofilter.sheetName());
  }
}
