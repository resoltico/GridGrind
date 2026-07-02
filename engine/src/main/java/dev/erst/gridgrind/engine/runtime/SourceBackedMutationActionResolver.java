package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.SourceBackedResolutionIdentitySupport.sameReference;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.CellGridInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellRowInput;
import java.io.IOException;
import java.util.List;

/** Resolves source-backed mutation payloads into inline canonical action values. */
final class SourceBackedMutationActionResolver {
  private SourceBackedMutationActionResolver() {}

  static MutationAction resolve(MutationAction action, ExecutionInputBindings bindings)
      throws IOException {
    return switch (action) {
      case CellMutationAction cellAction -> resolveCellAction(cellAction, bindings);
      case DrawingMutationAction drawingAction -> resolveDrawingAction(drawingAction, bindings);
      case StructuredMutationAction structuredAction ->
          resolveStructuredAction(structuredAction, bindings);
      case WorkbookMutationAction workbookAction -> resolveWorkbookAction(workbookAction, bindings);
    };
  }

  private static CellMutationAction resolveCellAction(
      CellMutationAction action, ExecutionInputBindings bindings) throws IOException {
    return switch (action) {
      case CellMutationAction.SetCell setCell -> resolveSetCell(setCell, bindings);
      case CellMutationAction.SetRange setRange -> resolveSetRange(setRange, bindings);
      case CellMutationAction.SetComment setComment -> resolveSetComment(setComment, bindings);
      case CellMutationAction.AppendRow appendRow -> resolveAppendRow(appendRow, bindings);
      default -> action;
    };
  }

  private static DrawingMutationAction resolveDrawingAction(
      DrawingMutationAction action, ExecutionInputBindings bindings) throws IOException {
    return switch (action) {
      case DrawingMutationAction.SetPicture setPicture -> resolveSetPicture(setPicture, bindings);
      case DrawingMutationAction.SetSignatureLine setSignatureLine ->
          resolveSetSignatureLine(setSignatureLine, bindings);
      case DrawingMutationAction.SetChart setChart -> resolveSetChart(setChart, bindings);
      case DrawingMutationAction.SetShape setShape -> resolveSetShape(setShape, bindings);
      case DrawingMutationAction.SetEmbeddedObject setEmbeddedObject ->
          resolveSetEmbeddedObject(setEmbeddedObject, bindings);
      default -> action;
    };
  }

  private static StructuredMutationAction resolveStructuredAction(
      StructuredMutationAction action, ExecutionInputBindings bindings) throws IOException {
    return switch (action) {
      case StructuredMutationAction.SetDataValidation setDataValidation ->
          resolveSetDataValidation(setDataValidation, bindings);
      case StructuredMutationAction.SetTable setTable -> resolveSetTable(setTable, bindings);
      case StructuredMutationAction.ImportCustomXmlMapping importCustomXmlMapping ->
          resolveImportCustomXmlMapping(importCustomXmlMapping, bindings);
      default -> action;
    };
  }

  private static WorkbookMutationAction resolveWorkbookAction(
      WorkbookMutationAction action, ExecutionInputBindings bindings) throws IOException {
    if (action instanceof WorkbookMutationAction.SetPrintLayout setPrintLayout) {
      return resolveSetPrintLayout(setPrintLayout, bindings);
    }
    return action;
  }

  private static CellMutationAction resolveSetCell(
      CellMutationAction.SetCell setCell, ExecutionInputBindings bindings) throws IOException {
    CellInput resolvedValue = SourceBackedPlanResolver.resolveCellInput(setCell.value(), bindings);
    return sameReference(resolvedValue, setCell.value())
        ? setCell
        : new CellMutationAction.SetCell(resolvedValue);
  }

  private static CellMutationAction resolveSetRange(
      CellMutationAction.SetRange setRange, ExecutionInputBindings bindings) throws IOException {
    if (!(setRange.rows() instanceof CellGridInput.Typed typedRows)) {
      return setRange;
    }
    List<List<CellInput>> resolvedRows =
        SourceBackedPlanResolver.resolveRows(typedRows.cells(), bindings);
    return sameReference(resolvedRows, typedRows.cells())
        ? setRange
        : new CellMutationAction.SetRange(new CellGridInput.Typed(resolvedRows));
  }

  private static CellMutationAction resolveSetComment(
      CellMutationAction.SetComment setComment, ExecutionInputBindings bindings)
      throws IOException {
    var resolvedComment =
        SourceBackedStructuredInputResolver.resolveComment(setComment.comment(), bindings);
    return sameReference(resolvedComment, setComment.comment())
        ? setComment
        : new CellMutationAction.SetComment(resolvedComment);
  }

  private static CellMutationAction resolveAppendRow(
      CellMutationAction.AppendRow appendRow, ExecutionInputBindings bindings) throws IOException {
    if (!(appendRow.values() instanceof CellRowInput.Typed typedValues)) {
      return appendRow;
    }
    List<CellInput> resolvedValues =
        SourceBackedPlanResolver.resolveCells(typedValues.cells(), bindings);
    return sameReference(resolvedValues, typedValues.cells())
        ? appendRow
        : new CellMutationAction.AppendRow(new CellRowInput.Typed(resolvedValues));
  }

  private static DrawingMutationAction resolveSetPicture(
      DrawingMutationAction.SetPicture setPicture, ExecutionInputBindings bindings)
      throws IOException {
    var resolvedPicture =
        SourceBackedStructuredInputResolver.resolvePicture(setPicture.picture(), bindings);
    return sameReference(resolvedPicture, setPicture.picture())
        ? setPicture
        : new DrawingMutationAction.SetPicture(resolvedPicture);
  }

  private static DrawingMutationAction resolveSetSignatureLine(
      DrawingMutationAction.SetSignatureLine setSignatureLine, ExecutionInputBindings bindings)
      throws IOException {
    var resolvedSignatureLine =
        SourceBackedStructuredInputResolver.resolveSignatureLine(
            setSignatureLine.signatureLine(), bindings);
    return sameReference(resolvedSignatureLine, setSignatureLine.signatureLine())
        ? setSignatureLine
        : new DrawingMutationAction.SetSignatureLine(resolvedSignatureLine);
  }

  private static DrawingMutationAction resolveSetChart(
      DrawingMutationAction.SetChart setChart, ExecutionInputBindings bindings) throws IOException {
    var resolvedChart =
        SourceBackedStructuredInputResolver.resolveChart(setChart.chart(), bindings);
    return sameReference(resolvedChart, setChart.chart())
        ? setChart
        : new DrawingMutationAction.SetChart(resolvedChart);
  }

  private static DrawingMutationAction resolveSetShape(
      DrawingMutationAction.SetShape setShape, ExecutionInputBindings bindings) throws IOException {
    var resolvedShape =
        SourceBackedStructuredInputResolver.resolveShape(setShape.shape(), bindings);
    return sameReference(resolvedShape, setShape.shape())
        ? setShape
        : new DrawingMutationAction.SetShape(resolvedShape);
  }

  private static DrawingMutationAction resolveSetEmbeddedObject(
      DrawingMutationAction.SetEmbeddedObject setEmbeddedObject, ExecutionInputBindings bindings)
      throws IOException {
    var resolvedEmbeddedObject =
        SourceBackedStructuredInputResolver.resolveEmbeddedObject(
            setEmbeddedObject.embeddedObject(), bindings);
    return sameReference(resolvedEmbeddedObject, setEmbeddedObject.embeddedObject())
        ? setEmbeddedObject
        : new DrawingMutationAction.SetEmbeddedObject(resolvedEmbeddedObject);
  }

  private static StructuredMutationAction resolveSetDataValidation(
      StructuredMutationAction.SetDataValidation setDataValidation, ExecutionInputBindings bindings)
      throws IOException {
    var resolvedValidation =
        SourceBackedStructuredInputResolver.resolveDataValidation(
            setDataValidation.validation(), bindings);
    return sameReference(resolvedValidation, setDataValidation.validation())
        ? setDataValidation
        : new StructuredMutationAction.SetDataValidation(resolvedValidation);
  }

  private static StructuredMutationAction resolveSetTable(
      StructuredMutationAction.SetTable setTable, ExecutionInputBindings bindings)
      throws IOException {
    var resolvedTable =
        SourceBackedStructuredInputResolver.resolveTable(setTable.table(), bindings);
    return sameReference(resolvedTable, setTable.table())
        ? setTable
        : new StructuredMutationAction.SetTable(resolvedTable);
  }

  private static StructuredMutationAction resolveImportCustomXmlMapping(
      StructuredMutationAction.ImportCustomXmlMapping importCustomXmlMapping,
      ExecutionInputBindings bindings)
      throws IOException {
    var resolvedImport =
        SourceBackedStructuredInputResolver.resolveCustomXmlImport(
            importCustomXmlMapping.mapping(), bindings);
    return sameReference(resolvedImport, importCustomXmlMapping.mapping())
        ? importCustomXmlMapping
        : new StructuredMutationAction.ImportCustomXmlMapping(resolvedImport);
  }

  private static WorkbookMutationAction resolveSetPrintLayout(
      WorkbookMutationAction.SetPrintLayout setPrintLayout, ExecutionInputBindings bindings)
      throws IOException {
    var resolvedPrintLayout =
        SourceBackedStructuredInputResolver.resolvePrintLayout(
            setPrintLayout.printLayout(), bindings);
    return sameReference(resolvedPrintLayout, setPrintLayout.printLayout())
        ? setPrintLayout
        : new WorkbookMutationAction.SetPrintLayout(resolvedPrintLayout);
  }
}
