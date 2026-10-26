package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerReport;
import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingMarker;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Covers delegation-only converter wrappers introduced by executor seam refactors. */
class ExecutorConverterDelegationCoverageTest {
  @Test
  void workbookCommandConverterDelegatesRemainCovered() {
    RichTextRunSpec run =
        new RichTextRunSpec(TextSourceSpec.inline("Bold"), java.util.Optional.empty());

    assertEquals(
        "Bold",
        WorkbookCommandCellValueSpecConverter.toResolvedRichText(
                new CellValueSpec.RichText(List.of(run)))
            .runs()
            .getFirst()
            .text());
    assertEquals("Bold", WorkbookCommandCellValueSpecConverter.toResolvedRichTextRun(run).text());
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
  }

  @Test
  void familyMutationConvertersHandleRepresentativeInFamilyActions() {
    CellMutationAction.SetCell setCell = new CellMutationAction.SetCell(new CellValueSpec.Blank());

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
                new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                setCell));
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
