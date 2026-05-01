package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookDrawingCommand;

/** Converts drawing-backed mutation families into engine commands. */
final class WorkbookCommandDrawingMutationConverter {
  private WorkbookCommandDrawingMutationConverter() {}

  static WorkbookCommand toCommand(Selector target, DrawingMutationAction action) {
    return switch (action) {
      case DrawingMutationAction.SetPicture setPicture ->
          new WorkbookDrawingCommand.SetPicture(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandConverter.toExcelPictureDefinition(setPicture.picture()));
      case DrawingMutationAction.SetSignatureLine setSignatureLine ->
          new WorkbookDrawingCommand.SetSignatureLine(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandConverter.toExcelSignatureLineDefinition(
                  setSignatureLine.signatureLine()));
      case DrawingMutationAction.SetChart setChart ->
          new WorkbookDrawingCommand.SetChart(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandConverter.toExcelChartDefinition(setChart.chart()));
      case DrawingMutationAction.SetShape setShape ->
          new WorkbookDrawingCommand.SetShape(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandConverter.toExcelShapeDefinition(setShape.shape()));
      case DrawingMutationAction.SetEmbeddedObject setEmbeddedObject ->
          new WorkbookDrawingCommand.SetEmbeddedObject(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandConverter.toExcelEmbeddedObjectDefinition(
                  setEmbeddedObject.embeddedObject()));
      case DrawingMutationAction.SetDrawingObjectAnchor setDrawingObjectAnchor -> {
        DrawingObjectSelector.ByName selector =
            WorkbookCommandSelectorSupport.drawingObjectByName(target, action);
        yield new WorkbookDrawingCommand.SetDrawingObjectAnchor(
            selector.sheetName(),
            selector.objectName(),
            WorkbookCommandConverter.toExcelDrawingAnchor(setDrawingObjectAnchor.anchor()));
      }
      case DrawingMutationAction.DeleteDrawingObject _ -> {
        DrawingObjectSelector.ByName selector =
            WorkbookCommandSelectorSupport.drawingObjectByName(target, action);
        yield new WorkbookDrawingCommand.DeleteDrawingObject(
            selector.sheetName(), selector.objectName());
      }
    };
  }
}
