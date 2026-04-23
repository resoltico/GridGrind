package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.WorkbookCommand;

/** Converts structured workbook-feature mutation families into engine commands. */
final class WorkbookCommandStructuredMutationConverter {
  private WorkbookCommandStructuredMutationConverter() {}

  static WorkbookCommand toCommand(Selector target, MutationAction action) {
    return switch (action) {
      case MutationAction.ImportCustomXmlMapping importCustomXmlMapping ->
          new WorkbookCommand.ImportCustomXmlMapping(
              WorkbookCommandConverter.toExcelCustomXmlImportDefinition(
                  importCustomXmlMapping.mapping()));
      case MutationAction.SetPivotTable setPivotTable -> {
        WorkbookCommandSelectorSupport.ensurePivotTableIdentity(target, setPivotTable);
        yield new WorkbookCommand.SetPivotTable(
            WorkbookCommandConverter.toExcelPivotTableDefinition(setPivotTable.pivotTable()));
      }
      case MutationAction.SetDataValidation setDataValidation -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookCommand.SetDataValidation(
            selector.sheetName(),
            selector.range(),
            WorkbookCommandConverter.toExcelDataValidationDefinition(
                setDataValidation.validation()));
      }
      case MutationAction.ClearDataValidations _ -> {
        SelectorConverter.SheetLocalRangeSelection selection =
            SelectorConverter.toSheetLocalRangeSelection((RangeSelector) target);
        yield new WorkbookCommand.ClearDataValidations(
            selection.sheetName(), selection.selection());
      }
      case MutationAction.SetConditionalFormatting setConditionalFormatting ->
          new WorkbookCommand.SetConditionalFormatting(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandConverter.toExcelConditionalFormattingBlock(
                  setConditionalFormatting.conditionalFormatting()));
      case MutationAction.ClearConditionalFormatting _ -> {
        SelectorConverter.SheetLocalRangeSelection selection =
            SelectorConverter.toSheetLocalRangeSelection((RangeSelector) target);
        yield new WorkbookCommand.ClearConditionalFormatting(
            selection.sheetName(), selection.selection());
      }
      case MutationAction.SetAutofilter setAutofilter -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookCommand.SetAutofilter(
            selector.sheetName(),
            selector.range(),
            setAutofilter.criteria().stream()
                .map(WorkbookCommandStructuredInputConverter::toExcelAutofilterFilterColumn)
                .toList(),
            WorkbookCommandStructuredInputConverter.toExcelAutofilterSortState(
                setAutofilter.sortState()));
      }
      case MutationAction.ClearAutofilter _ ->
          new WorkbookCommand.ClearAutofilter(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name());
      case MutationAction.SetTable setTable -> {
        WorkbookCommandSelectorSupport.ensureTableIdentity(target, setTable);
        yield new WorkbookCommand.SetTable(
            WorkbookCommandConverter.toExcelTableDefinition(setTable.table()));
      }
      case MutationAction.DeleteTable _ -> {
        TableSelector.ByNameOnSheet selector =
            WorkbookCommandSelectorSupport.tableByNameOnSheet(target, action);
        yield new WorkbookCommand.DeleteTable(selector.name(), selector.sheetName());
      }
      case MutationAction.DeletePivotTable _ -> {
        PivotTableSelector.ByNameOnSheet selector =
            WorkbookCommandSelectorSupport.pivotTableByNameOnSheet(target, action);
        yield new WorkbookCommand.DeletePivotTable(selector.name(), selector.sheetName());
      }
      case MutationAction.SetNamedRange setNamedRange -> {
        WorkbookCommandSelectorSupport.ensureNamedRangeIdentity(target, setNamedRange);
        yield new WorkbookCommand.SetNamedRange(
            new ExcelNamedRangeDefinition(
                setNamedRange.name(),
                WorkbookCommandConverter.toExcelNamedRangeScope(setNamedRange.scope()),
                WorkbookCommandConverter.toExcelNamedRangeTarget(setNamedRange.target())));
      }
      case MutationAction.DeleteNamedRange _ -> {
        NamedRangeSelector.ScopedExact selector =
            WorkbookCommandSelectorSupport.namedRangeScopedExact(target, action);
        yield new WorkbookCommand.DeleteNamedRange(
            WorkbookCommandConverter.toExcelNamedRangeName(selector),
            WorkbookCommandConverter.toExcelNamedRangeScope(selector));
      }
      default -> null;
    };
  }
}
