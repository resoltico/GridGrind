package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
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
              WorkbookCommandConverter.toExcelCustomXmlImportDefinition(
                  importCustomXmlMapping.mapping()));
      case StructuredMutationAction.SetPivotTable setPivotTable -> {
        WorkbookCommandSelectorSupport.ensurePivotTableIdentity(target, setPivotTable);
        yield new WorkbookTabularCommand.SetPivotTable(
            WorkbookCommandConverter.toExcelPivotTableDefinition(setPivotTable.pivotTable()));
      }
      case StructuredMutationAction.SetDataValidation setDataValidation -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookFormattingCommand.SetDataValidation(
            selector.sheetName(),
            selector.range(),
            WorkbookCommandConverter.toExcelDataValidationDefinition(
                setDataValidation.validation()));
      }
      case StructuredMutationAction.ClearDataValidations _ -> {
        SelectorConverter.SheetLocalRangeSelection selection =
            SelectorConverter.toSheetLocalRangeSelection((RangeSelector) target);
        yield new WorkbookFormattingCommand.ClearDataValidations(
            selection.sheetName(), selection.selection());
      }
      case StructuredMutationAction.SetConditionalFormatting setConditionalFormatting ->
          new WorkbookFormattingCommand.SetConditionalFormatting(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandConverter.toExcelConditionalFormattingBlock(
                  setConditionalFormatting.conditionalFormatting()));
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
            selector.sheetName(),
            selector.range(),
            setAutofilter.criteria().stream()
                .map(WorkbookCommandStructuredInputConverter::toExcelAutofilterFilterColumn)
                .toList(),
            WorkbookCommandStructuredInputConverter.toExcelAutofilterSortState(
                    setAutofilter.sortState())
                .orElse(null));
      }
      case StructuredMutationAction.ClearAutofilter _ ->
          new WorkbookTabularCommand.ClearAutofilter(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name());
      case StructuredMutationAction.SetTable setTable -> {
        WorkbookCommandSelectorSupport.ensureTableIdentity(target, setTable);
        yield new WorkbookTabularCommand.SetTable(
            WorkbookCommandConverter.toExcelTableDefinition(setTable.table()));
      }
      case StructuredMutationAction.DeleteTable _ -> {
        TableSelector.ByNameOnSheet selector =
            WorkbookCommandSelectorSupport.tableByNameOnSheet(target, action);
        yield new WorkbookTabularCommand.DeleteTable(selector.name(), selector.sheetName());
      }
      case StructuredMutationAction.DeletePivotTable _ -> {
        PivotTableSelector.ByNameOnSheet selector =
            WorkbookCommandSelectorSupport.pivotTableByNameOnSheet(target, action);
        yield new WorkbookTabularCommand.DeletePivotTable(selector.name(), selector.sheetName());
      }
      case StructuredMutationAction.SetNamedRange setNamedRange -> {
        WorkbookCommandSelectorSupport.ensureNamedRangeIdentity(target, setNamedRange);
        yield new WorkbookMetadataCommand.SetNamedRange(
            new ExcelNamedRangeDefinition(
                setNamedRange.name(),
                WorkbookCommandConverter.toExcelNamedRangeScope(setNamedRange.scope()),
                WorkbookCommandConverter.toExcelNamedRangeTarget(setNamedRange.target())));
      }
      case StructuredMutationAction.DeleteNamedRange _ -> {
        NamedRangeSelector.ScopedExact selector =
            WorkbookCommandSelectorSupport.namedRangeScopedExact(target, action);
        yield new WorkbookMetadataCommand.DeleteNamedRange(
            WorkbookCommandConverter.toExcelNamedRangeName(selector),
            WorkbookCommandConverter.toExcelNamedRangeScope(selector));
      }
    };
  }
}
