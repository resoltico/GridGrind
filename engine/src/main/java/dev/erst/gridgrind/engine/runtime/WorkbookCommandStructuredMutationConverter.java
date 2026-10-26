package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelRangeSelection;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookFormattingCommand;
import dev.erst.gridgrind.excel.WorkbookMetadataCommand;
import dev.erst.gridgrind.excel.WorkbookTabularCommand;

/** Converts structured workbook-feature mutation families into engine commands. */
final class WorkbookCommandStructuredMutationConverter {
  private WorkbookCommandStructuredMutationConverter() {}

  static WorkbookCommand toCommand(Selector target, StructuredMutationAction action) {
    return switch (action) {
      case StructuredMutationAction.ImportCustomXmlMapping importCustomXmlMapping ->
          new WorkbookMetadataCommand.ImportCustomXmlMapping(
              WorkbookCommandDrawingInputConverter.toResolvedCustomXmlImport(
                  importCustomXmlMapping.mapping()));
      case StructuredMutationAction.SetPivotTable setPivotTable -> {
        WorkbookCommandSelectorSupport.ensurePivotTableIdentity(target, setPivotTable);
        yield new WorkbookTabularCommand.SetPivotTable(setPivotTable.pivotTable());
      }
      case StructuredMutationAction.SetDataValidation setDataValidation -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookFormattingCommand.SetDataValidation(
            selector.sheetName().value(),
            selector.range().value(),
            WorkbookCommandStructuredInputConverter.toResolvedDataValidation(
                setDataValidation.validation()));
      }
      case StructuredMutationAction.ClearDataValidations _ -> {
        SelectorConverter.SheetLocalRangeSelection selection =
            SelectorConverter.toSheetLocalRangeSelection((RangeSelector) target);
        yield new WorkbookFormattingCommand.ClearDataValidations(
            selection.sheetName(), selection.selection());
      }
      case StructuredMutationAction.SetConditionalFormatting setConditionalFormatting -> {
        SelectorConverter.SheetLocalRangeSelection selection =
            SelectorConverter.toSheetLocalRangeSelection((RangeSelector) target);
        ExcelRangeSelection.Selected selected =
            switch (selection.selection()) {
              case ExcelRangeSelection.Selected explicit -> explicit;
              case ExcelRangeSelection.All _ ->
                  throw new IllegalArgumentException(
                      "SET_CONDITIONAL_FORMATTING requires explicit range targets");
            };
        yield new WorkbookFormattingCommand.SetConditionalFormatting(
            selection.sheetName(),
            WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingBlock(
                selected.ranges(), setConditionalFormatting.conditionalFormatting()));
      }
      case StructuredMutationAction.ClearConditionalFormatting _ -> {
        SelectorConverter.SheetLocalRangeSelection selection =
            SelectorConverter.toSheetLocalRangeSelection((RangeSelector) target);
        yield new WorkbookFormattingCommand.ClearConditionalFormatting(
            selection.sheetName(), selection.selection());
      }
      case StructuredMutationAction.SetAutofilter setAutofilter -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookTabularCommand.SetAutofilter(
            selector.sheetName().value(),
            selector.range().value(),
            setAutofilter.criteria().stream()
                .map(WorkbookCommandStructuredInputConverter::toExcelAutofilterFilterColumn)
                .toList(),
            setAutofilter
                .sortState()
                .flatMap(WorkbookCommandStructuredInputConverter::toExcelAutofilterSortState));
      }
      case StructuredMutationAction.ClearAutofilter _ ->
          new WorkbookTabularCommand.ClearAutofilter(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name().value());
      case StructuredMutationAction.SetTable setTable -> {
        WorkbookCommandSelectorSupport.ensureTableIdentity(target, setTable);
        yield new WorkbookTabularCommand.SetTable(
            WorkbookCommandTabularInputConverter.toResolvedTable(setTable.table()));
      }
      case StructuredMutationAction.DeleteTable _ -> {
        TableSelector.ByNameOnSheet selector =
            WorkbookCommandSelectorSupport.tableByNameOnSheet(target, action);
        yield new WorkbookTabularCommand.DeleteTable(
            selector.name().value(), selector.sheetName().value());
      }
      case StructuredMutationAction.DeletePivotTable _ -> {
        PivotTableSelector.ByNameOnSheet selector =
            WorkbookCommandSelectorSupport.pivotTableByNameOnSheet(target, action);
        yield new WorkbookTabularCommand.DeletePivotTable(
            selector.name(), selector.sheetName().value());
      }
      case StructuredMutationAction.SetNamedRange setNamedRange -> {
        WorkbookCommandSelectorSupport.ensureNamedRangeIdentity(target, setNamedRange);
        yield new WorkbookMetadataCommand.SetNamedRange(
            new ExcelNamedRangeDefinition(
                setNamedRange.name().value(),
                WorkbookCommandLayoutInputConverter.toExcelNamedRangeScope(setNamedRange.scope()),
                WorkbookCommandLayoutInputConverter.toExcelNamedRangeTarget(
                    setNamedRange.target())));
      }
      case StructuredMutationAction.DeleteNamedRange _ -> {
        NamedRangeSelector.ScopedExact selector =
            WorkbookCommandSelectorSupport.namedRangeScopedExact(target, action);
        yield new WorkbookMetadataCommand.DeleteNamedRange(
            WorkbookCommandLayoutInputConverter.toExcelNamedRangeName(selector),
            WorkbookCommandLayoutInputConverter.toExcelNamedRangeScope(selector));
      }
    };
  }
}
