package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.MutationAction;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;

/** Materializes only source-bearing action operands; source-free actions retain their identity. */
final class MutationResolution {
  private MutationResolution() {}

  static ResolvedMutationAction resolve(MutationAction action, ExecutionInputBindings bindings) {
    return switch (action) {
      case ResolvedMutationAction shared -> shared;
      case CellMutationAction.SetCell cell ->
          new ResolvedMutationAction.SetCell(
              WorkbookCommandCellValueSpecConverter.toResolvedCellWrite(cell.value()));
      case CellMutationAction.SetRange range ->
          new ResolvedMutationAction.SetRange(
              range.rows().toCellInputRows().stream()
                  .map(
                      row ->
                          row.stream()
                              .map(WorkbookCommandCellValueSpecConverter::toResolvedCellWrite)
                              .toList())
                  .toList());
      case CellMutationAction.AppendRow row ->
          new ResolvedMutationAction.AppendRow(
              row.values().toCellInputs().stream()
                  .map(WorkbookCommandCellValueSpecConverter::toResolvedCellWrite)
                  .toList());
      case CellMutationAction.SetArrayFormula formula ->
          new ResolvedMutationAction.SetArrayFormula(
              WorkbookCommandCellValueSpecConverter.toResolvedArrayFormula(formula.formula()));
      case CellMutationAction.SetComment comment ->
          new ResolvedMutationAction.SetComment(
              WorkbookCommandCellValueSpecConverter.toResolvedComment(comment.comment()));
      case DrawingMutationAction.SetChart chart ->
          new ResolvedMutationAction.SetChart(
              WorkbookCommandDrawingInputConverter.toResolvedChart(chart.chart()));
      case DrawingMutationAction.SetPicture picture ->
          new ResolvedMutationAction.SetPicture(
              WorkbookCommandDrawingInputConverter.toResolvedPicture(picture.picture()));
      case DrawingMutationAction.SetEmbeddedObject object ->
          new ResolvedMutationAction.SetEmbeddedObject(
              WorkbookCommandDrawingInputConverter.toResolvedEmbeddedObject(
                  object.embeddedObject()));
      case DrawingMutationAction.SetShape shape ->
          new ResolvedMutationAction.SetShape(
              WorkbookCommandDrawingInputConverter.toResolvedShape(shape.shape()));
      case DrawingMutationAction.SetSignatureLine signature ->
          new ResolvedMutationAction.SetSignatureLine(
              WorkbookCommandDrawingInputConverter.toResolvedSignatureLine(
                  signature.signatureLine()));
      case StructuredMutationAction.ImportCustomXmlMapping mapping ->
          new ResolvedMutationAction.ImportCustomXmlMapping(
              WorkbookCommandDrawingInputConverter.toResolvedCustomXmlImport(mapping.mapping()));
      case StructuredMutationAction.SetTable table ->
          new ResolvedMutationAction.SetTable(
              WorkbookCommandTabularInputConverter.toResolvedTable(table.table()));
      case StructuredMutationAction.SetDataValidation validation ->
          new ResolvedMutationAction.SetDataValidation(
              WorkbookCommandStructuredInputConverter.toResolvedDataValidation(
                  validation.validation()));
      case WorkbookMutationAction.SetPrintLayout layout ->
          new ResolvedMutationAction.SetPrintLayout(
              WorkbookCommandLayoutInputConverter.toResolvedPrintLayout(layout.printLayout()));
      case WorkbookMutationAction.SetSheetProtection protection ->
          new ResolvedMutationAction.SetSheetProtection(
              protection.protection(), protection.passwordRef().map(bindings::resolveSecret));
      case WorkbookMutationAction.SetWorkbookProtection protection ->
          new ResolvedMutationAction.SetWorkbookProtection(
              WorkbookCommandLayoutInputConverter.toResolvedWorkbookProtection(
                  protection.protection(), bindings));
    };
  }
}
