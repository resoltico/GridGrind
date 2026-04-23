package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.excel.WorkbookCommand;

/** Converts drawing-backed mutation families into engine commands. */
final class WorkbookCommandDrawingMutationConverter {
  private WorkbookCommandDrawingMutationConverter() {}

  static WorkbookCommand toCommand(Selector target, MutationAction action) {
    return switch (action) {
      case MutationAction.SetPicture setPicture ->
          new WorkbookCommand.SetPicture(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandConverter.toExcelPictureDefinition(setPicture.picture()));
      case MutationAction.SetSignatureLine setSignatureLine ->
          new WorkbookCommand.SetSignatureLine(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandConverter.toExcelSignatureLineDefinition(
                  setSignatureLine.signatureLine()));
      case MutationAction.SetChart setChart ->
          new WorkbookCommand.SetChart(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandConverter.toExcelChartDefinition(setChart.chart()));
      case MutationAction.SetShape setShape ->
          new WorkbookCommand.SetShape(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandConverter.toExcelShapeDefinition(setShape.shape()));
      case MutationAction.SetEmbeddedObject setEmbeddedObject ->
          new WorkbookCommand.SetEmbeddedObject(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandConverter.toExcelEmbeddedObjectDefinition(
                  setEmbeddedObject.embeddedObject()));
      case MutationAction.SetDrawingObjectAnchor setDrawingObjectAnchor -> {
        DrawingObjectSelector.ByName selector =
            WorkbookCommandSelectorSupport.drawingObjectByName(target, action);
        yield new WorkbookCommand.SetDrawingObjectAnchor(
            selector.sheetName(),
            selector.objectName(),
            WorkbookCommandConverter.toExcelDrawingAnchor(setDrawingObjectAnchor.anchor()));
      }
      case MutationAction.DeleteDrawingObject _ -> {
        DrawingObjectSelector.ByName selector =
            WorkbookCommandSelectorSupport.drawingObjectByName(target, action);
        yield new WorkbookCommand.DeleteDrawingObject(selector.sheetName(), selector.objectName());
      }
      default -> null;
    };
  }
}
