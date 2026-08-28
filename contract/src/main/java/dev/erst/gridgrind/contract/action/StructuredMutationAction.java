package dev.erst.gridgrind.contract.action;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadata;
import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingDefinitionInput;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Mutation family for workbook-scoped structured Excel features. */
public sealed interface StructuredMutationAction extends MutationAction {
  /** Imports one XML document into one existing workbook custom-XML mapping. */
  @ProtocolTypeMetadata(
      id = "IMPORT_CUSTOM_XML_MAPPING",
      summary = "Import one XML document into one existing workbook custom-XML mapping.",
      targetSelectors = {WorkbookSelector.class})
  record ImportCustomXmlMapping(CustomXmlImportInput mapping) implements StructuredMutationAction {
    public ImportCustomXmlMapping {
      Objects.requireNonNull(mapping, "mapping must not be null");
    }
  }

  /** Creates or replaces one workbook-global pivot-table definition. */
  @ProtocolTypeMetadata(
      id = "SET_PIVOT_TABLE",
      summary = "Create or replace one workbook-global pivot-table definition.",
      targetSelectors = {PivotTableSelector.ByNameOnSheet.class})
  record SetPivotTable(PivotTableInput pivotTable) implements StructuredMutationAction {
    public SetPivotTable {
      Objects.requireNonNull(pivotTable, "pivotTable must not be null");
    }
  }

  /** Creates or replaces one data-validation rule over the requested sheet range. */
  @ProtocolTypeMetadata(
      id = "SET_DATA_VALIDATION",
      summary = "Create or replace one data-validation rule over the requested sheet range.",
      targetSelectors = {RangeSelector.ByRange.class})
  record SetDataValidation(DataValidationInput validation) implements StructuredMutationAction {
    public SetDataValidation {
      Objects.requireNonNull(validation, "validation must not be null");
    }
  }

  /** Removes data-validation structures on the sheet that match the provided range selection. */
  @ProtocolTypeMetadata(
      id = "CLEAR_DATA_VALIDATIONS",
      summary = "Remove data-validation structures on the sheet that match the range selection.",
      targetSelectors = {RangeSelector.class})
  record ClearDataValidations() implements StructuredMutationAction {
    public ClearDataValidations {}
  }

  /** Creates or replaces one logical conditional-formatting block over explicit sheet ranges. */
  @ProtocolTypeMetadata(
      id = "SET_CONDITIONAL_FORMATTING",
      summary =
          "Create or replace one logical conditional-formatting block over explicit sheet ranges.",
      targetSelectors = {RangeSelector.ByRange.class, RangeSelector.ByRanges.class})
  record SetConditionalFormatting(ConditionalFormattingDefinitionInput conditionalFormatting)
      implements StructuredMutationAction {
    public SetConditionalFormatting {
      Objects.requireNonNull(conditionalFormatting, "conditionalFormatting must not be null");
    }
  }

  /** Removes conditional-formatting blocks on the sheet that match the provided range selection. */
  @ProtocolTypeMetadata(
      id = "CLEAR_CONDITIONAL_FORMATTING",
      summary = "Remove conditional-formatting blocks on the sheet that match the range selection.",
      targetSelectors = {RangeSelector.class})
  record ClearConditionalFormatting() implements StructuredMutationAction {
    public ClearConditionalFormatting {}
  }

  /** Creates or replaces one sheet-level autofilter range. */
  @ProtocolTypeMetadata(
      id = "SET_AUTOFILTER",
      summary = "Create or replace one sheet-level autofilter range.",
      targetSelectors = {RangeSelector.ByRange.class})
  record SetAutofilter(
      List<AutofilterFilterColumnInput> criteria,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<AutofilterSortStateInput> sortState)
      implements StructuredMutationAction {
    /** Creates a plain sheet-level autofilter without criteria or explicit sort state. */
    public SetAutofilter() {
      this(List.of(), Optional.empty());
    }

    /** Creates one sheet-level autofilter with an optional explicit sort state. */
    public SetAutofilter(
        List<AutofilterFilterColumnInput> criteria, AutofilterSortStateInput sortState) {
      this(criteria, Optional.ofNullable(sortState));
    }

    public SetAutofilter {
      Objects.requireNonNull(criteria, "criteria must not be null");
      Objects.requireNonNull(sortState, "sortState must not be null");
      List<AutofilterFilterColumnInput> copiedCriteria = new ArrayList<>(criteria.size());
      for (AutofilterFilterColumnInput criterion : criteria) {
        Objects.requireNonNull(criterion, "criteria must not contain null values");
        copiedCriteria.add(criterion);
      }
      criteria = List.copyOf(copiedCriteria);
    }
  }

  /** Clears the sheet-level autofilter range on one sheet. */
  @ProtocolTypeMetadata(
      id = "CLEAR_AUTOFILTER",
      summary = "Clear the sheet-level autofilter range on one sheet.",
      targetSelectors = {SheetSelector.ByName.class})
  record ClearAutofilter() implements StructuredMutationAction {
    public ClearAutofilter {}
  }

  /** Creates or replaces one workbook-global table definition. */
  @ProtocolTypeMetadata(
      id = "SET_TABLE",
      summary = "Create or replace one workbook-global table definition.",
      targetSelectors = {TableSelector.ByNameOnSheet.class})
  record SetTable(TableInput table) implements StructuredMutationAction {
    public SetTable {
      Objects.requireNonNull(table, "table must not be null");
    }
  }

  /** Deletes one existing table by workbook-global name and expected sheet. */
  @ProtocolTypeMetadata(
      id = "DELETE_TABLE",
      summary = "Delete one existing table by workbook-global name and expected sheet.",
      targetSelectors = {TableSelector.ByNameOnSheet.class})
  record DeleteTable() implements StructuredMutationAction {
    public DeleteTable {}
  }

  /** Deletes one existing pivot table by workbook-global name and expected sheet. */
  @ProtocolTypeMetadata(
      id = "DELETE_PIVOT_TABLE",
      summary = "Delete one existing pivot table by workbook-global name and expected sheet.",
      targetSelectors = {PivotTableSelector.ByNameOnSheet.class})
  record DeletePivotTable() implements StructuredMutationAction {
    public DeletePivotTable {}
  }

  /** Creates or replaces one typed named range in workbook or sheet scope. */
  @ProtocolTypeMetadata(
      id = "SET_NAMED_RANGE",
      summary = "Create or replace one typed named range in workbook or sheet scope.",
      targetSelectors = {
        NamedRangeSelector.WorkbookScope.class,
        NamedRangeSelector.SheetScope.class
      })
  record SetNamedRange(String name, NamedRangeScope scope, NamedRangeTarget target)
      implements StructuredMutationAction {
    public SetNamedRange {
      Objects.requireNonNull(scope, "scope must not be null");
      Objects.requireNonNull(target, "target must not be null");
      MutationAction.Validation.requireNamedRangeName(name);
    }
  }

  /** Deletes one existing named range from workbook or sheet scope. */
  @ProtocolTypeMetadata(
      id = "DELETE_NAMED_RANGE",
      summary = "Delete one existing named range from workbook or sheet scope.",
      targetSelectors = {
        NamedRangeSelector.WorkbookScope.class,
        NamedRangeSelector.SheetScope.class
      })
  record DeleteNamedRange() implements StructuredMutationAction {
    public DeleteNamedRange {}
  }
}
