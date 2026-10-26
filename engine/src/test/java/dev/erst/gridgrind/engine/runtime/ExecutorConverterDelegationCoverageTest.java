package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

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
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.excel.*;
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
