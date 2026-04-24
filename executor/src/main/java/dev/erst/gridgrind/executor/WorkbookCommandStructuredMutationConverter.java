package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.WorkbookCommand;
import java.util.Optional;

/** Converts structured workbook-feature mutation families into engine commands. */
final class WorkbookCommandStructuredMutationConverter {
  private WorkbookCommandStructuredMutationConverter() {}

  static Optional<WorkbookCommand> toCommand(Selector target, MutationAction action) {
    return switch (action) {
      case MutationAction.ImportCustomXmlMapping importCustomXmlMapping ->
          Optional.of(
              new WorkbookCommand.ImportCustomXmlMapping(
                  WorkbookCommandConverter.toExcelCustomXmlImportDefinition(
                      importCustomXmlMapping.mapping())));
      case MutationAction.SetPivotTable setPivotTable -> {
        WorkbookCommandSelectorSupport.ensurePivotTableIdentity(target, setPivotTable);
        yield Optional.of(
            new WorkbookCommand.SetPivotTable(
                WorkbookCommandConverter.toExcelPivotTableDefinition(setPivotTable.pivotTable())));
      }
      case MutationAction.SetDataValidation setDataValidation -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield Optional.of(
            new WorkbookCommand.SetDataValidation(
                selector.sheetName(),
                selector.range(),
                WorkbookCommandConverter.toExcelDataValidationDefinition(
                    setDataValidation.validation())));
      }
      case MutationAction.ClearDataValidations _ -> {
        SelectorConverter.SheetLocalRangeSelection selection =
            SelectorConverter.toSheetLocalRangeSelection((RangeSelector) target);
        yield Optional.of(
            new WorkbookCommand.ClearDataValidations(selection.sheetName(), selection.selection()));
      }
      case MutationAction.SetConditionalFormatting setConditionalFormatting ->
          Optional.of(
              new WorkbookCommand.SetConditionalFormatting(
                  WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
                  WorkbookCommandConverter.toExcelConditionalFormattingBlock(
                      setConditionalFormatting.conditionalFormatting())));
      case MutationAction.ClearConditionalFormatting _ -> {
        SelectorConverter.SheetLocalRangeSelection selection =
            SelectorConverter.toSheetLocalRangeSelection((RangeSelector) target);
        yield Optional.of(
            new WorkbookCommand.ClearConditionalFormatting(
                selection.sheetName(), selection.selection()));
      }
      case MutationAction.SetAutofilter setAutofilter -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield Optional.of(
            new WorkbookCommand.SetAutofilter(
                selector.sheetName(),
                selector.range(),
                setAutofilter.criteria().stream()
                    .map(WorkbookCommandStructuredInputConverter::toExcelAutofilterFilterColumn)
                    .toList(),
                WorkbookCommandStructuredInputConverter.toExcelAutofilterSortState(
                        setAutofilter.sortState())
                    .orElse(null)));
      }
      case MutationAction.ClearAutofilter _ ->
          Optional.of(
              new WorkbookCommand.ClearAutofilter(
                  WorkbookCommandSelectorSupport.sheetByName(target, action).name()));
      case MutationAction.SetTable setTable -> {
        WorkbookCommandSelectorSupport.ensureTableIdentity(target, setTable);
        yield Optional.of(
            new WorkbookCommand.SetTable(
                WorkbookCommandConverter.toExcelTableDefinition(setTable.table())));
      }
      case MutationAction.DeleteTable _ -> {
        TableSelector.ByNameOnSheet selector =
            WorkbookCommandSelectorSupport.tableByNameOnSheet(target, action);
        yield Optional.of(new WorkbookCommand.DeleteTable(selector.name(), selector.sheetName()));
      }
      case MutationAction.DeletePivotTable _ -> {
        PivotTableSelector.ByNameOnSheet selector =
            WorkbookCommandSelectorSupport.pivotTableByNameOnSheet(target, action);
        yield Optional.of(
            new WorkbookCommand.DeletePivotTable(selector.name(), selector.sheetName()));
      }
      case MutationAction.SetNamedRange setNamedRange -> {
        WorkbookCommandSelectorSupport.ensureNamedRangeIdentity(target, setNamedRange);
        yield Optional.of(
            new WorkbookCommand.SetNamedRange(
                new ExcelNamedRangeDefinition(
                    setNamedRange.name(),
                    WorkbookCommandConverter.toExcelNamedRangeScope(setNamedRange.scope()),
                    WorkbookCommandConverter.toExcelNamedRangeTarget(setNamedRange.target()))));
      }
      case MutationAction.DeleteNamedRange _ -> {
        NamedRangeSelector.ScopedExact selector =
            WorkbookCommandSelectorSupport.namedRangeScopedExact(target, action);
        yield Optional.of(
            new WorkbookCommand.DeleteNamedRange(
                WorkbookCommandConverter.toExcelNamedRangeName(selector),
                WorkbookCommandConverter.toExcelNamedRangeScope(selector)));
      }
      default -> Optional.empty();
    };
  }
}
