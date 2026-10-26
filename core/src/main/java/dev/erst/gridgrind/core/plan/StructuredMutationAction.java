package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.customxml.CustomXmlImportSpec;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateSpec;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationSpec;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Mutation family for workbook-scoped structured Excel features. */
public sealed interface StructuredMutationAction extends MutationAction {
  /** Imports one XML document into one existing workbook custom-XML mapping. */
  @Variant(
      id = "IMPORT_CUSTOM_XML_MAPPING",
      summary = "Import one XML document into one existing workbook custom-XML mapping.",
      targetSelectors = {WorkbookSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.STRUCTURED_MUTATION)
  record ImportCustomXmlMapping(CustomXmlImportSpec mapping) implements StructuredMutationAction {
    public ImportCustomXmlMapping {
      Objects.requireNonNull(mapping, "mapping must not be null");
    }
  }

  /** Creates or replaces one workbook-global pivot-table definition. */
  @Variant(
      id = "SET_PIVOT_TABLE",
      summary = "Create or replace one workbook-global pivot-table definition.",
      targetSelectors = {PivotTableSelector.ByNameOnSheet.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.STRUCTURED_MUTATION)
  record SetPivotTable(PivotTableSpec pivotTable) implements StructuredMutationAction {
    public SetPivotTable {
      Objects.requireNonNull(pivotTable, "pivotTable must not be null");
    }
  }

  /** Creates or replaces one data-validation rule over the requested sheet range. */
  @Variant(
      id = "SET_DATA_VALIDATION",
      summary = "Create or replace one data-validation rule over the requested sheet range.",
      targetSelectors = {RangeSelector.ByRange.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.STRUCTURED_MUTATION)
  record SetDataValidation(DataValidationSpec validation) implements StructuredMutationAction {
    public SetDataValidation {
      Objects.requireNonNull(validation, "validation must not be null");
    }
  }

  /** Removes data-validation structures on the sheet that match the provided range selection. */
  @Variant(
      id = "CLEAR_DATA_VALIDATIONS",
      summary = "Remove data-validation structures on the sheet that match the range selection.",
      targetSelectors = {RangeSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.STRUCTURED_MUTATION)
  record ClearDataValidations() implements StructuredMutationAction {
    public ClearDataValidations {}
  }

  /** Creates or replaces one logical conditional-formatting block over explicit sheet ranges. */
  @Variant(
      id = "SET_CONDITIONAL_FORMATTING",
      summary =
          "Create or replace one logical conditional-formatting block over explicit sheet ranges.",
      targetSelectors = {RangeSelector.ByRange.class, RangeSelector.ByRanges.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.STRUCTURED_MUTATION)
  record SetConditionalFormatting(ConditionalFormattingSpec conditionalFormatting)
      implements StructuredMutationAction {
    public SetConditionalFormatting {
      Objects.requireNonNull(conditionalFormatting, "conditionalFormatting must not be null");
    }
  }

  /** Removes conditional-formatting blocks on the sheet that match the provided range selection. */
  @Variant(
      id = "CLEAR_CONDITIONAL_FORMATTING",
      summary = "Remove conditional-formatting blocks on the sheet that match the range selection.",
      targetSelectors = {RangeSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.STRUCTURED_MUTATION)
  record ClearConditionalFormatting() implements StructuredMutationAction {
    public ClearConditionalFormatting {}
  }

  /** Creates or replaces one sheet-level autofilter range. */
  @Variant(
      id = "SET_AUTOFILTER",
      summary = "Create or replace one sheet-level autofilter range.",
      targetSelectors = {RangeSelector.ByRange.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.STRUCTURED_MUTATION)
  record SetAutofilter(
      List<AutofilterFilterColumnSpec> criteria, Optional<AutofilterSortStateSpec> sortState)
      implements StructuredMutationAction {
    /** Creates a plain sheet-level autofilter without criteria or explicit sort state. */
    public SetAutofilter() {
      this(List.of(), Optional.empty());
    }

    /** Creates one sheet-level autofilter with an optional explicit sort state. */
    public SetAutofilter(
        List<AutofilterFilterColumnSpec> criteria, AutofilterSortStateSpec sortState) {
      this(criteria, Optional.ofNullable(sortState));
    }

    public SetAutofilter {
      Objects.requireNonNull(criteria, "criteria must not be null");
      Objects.requireNonNull(sortState, "sortState must not be null");
      List<AutofilterFilterColumnSpec> copiedCriteria = new ArrayList<>(criteria.size());
      for (AutofilterFilterColumnSpec criterion : criteria) {
        Objects.requireNonNull(criterion, "criteria must not contain null values");
        copiedCriteria.add(criterion);
      }
      criteria = List.copyOf(copiedCriteria);
    }
  }

  /** Clears the sheet-level autofilter range on one sheet. */
  @Variant(
      id = "CLEAR_AUTOFILTER",
      summary = "Clear the sheet-level autofilter range on one sheet.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.STRUCTURED_MUTATION)
  record ClearAutofilter() implements StructuredMutationAction {
    public ClearAutofilter {}
  }

  /** Creates or replaces one workbook-global table definition. */
  @Variant(
      id = "SET_TABLE",
      summary = "Create or replace one workbook-global table definition.",
      targetSelectors = {TableSelector.ByNameOnSheet.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.STRUCTURED_MUTATION)
  record SetTable(TableSpec table) implements StructuredMutationAction {
    public SetTable {
      Objects.requireNonNull(table, "table must not be null");
    }
  }

  /** Deletes one existing table by workbook-global name and expected sheet. */
  @Variant(
      id = "DELETE_TABLE",
      summary = "Delete one existing table by workbook-global name and expected sheet.",
      targetSelectors = {TableSelector.ByNameOnSheet.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.STRUCTURED_MUTATION)
  record DeleteTable() implements StructuredMutationAction {
    public DeleteTable {}
  }

  /** Deletes one existing pivot table by workbook-global name and expected sheet. */
  @Variant(
      id = "DELETE_PIVOT_TABLE",
      summary = "Delete one existing pivot table by workbook-global name and expected sheet.",
      targetSelectors = {PivotTableSelector.ByNameOnSheet.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.STRUCTURED_MUTATION)
  record DeletePivotTable() implements StructuredMutationAction {
    public DeletePivotTable {}
  }

  /** Creates or replaces one typed named range in workbook or sheet scope. */
  @Variant(
      id = "SET_NAMED_RANGE",
      summary = "Create or replace one typed named range in workbook or sheet scope.",
      targetSelectors = {
        NamedRangeSelector.WorkbookScope.class,
        NamedRangeSelector.SheetScope.class
      },
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.STRUCTURED_MUTATION)
  record SetNamedRange(
      @dev.erst.gridgrind.core.metadata.FieldRule(
              kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.LENGTH_RANGE,
              minimum = 1,
              maximum =
                  dev.erst.gridgrind.core.vocabulary.ProtocolConstraintValues
                      .DEFINED_NAME_MAX_CODE_POINTS)
          @dev.erst.gridgrind.core.metadata.FieldRule(
              kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.STRING_PATTERN,
              text =
                  dev.erst.gridgrind.core.vocabulary.ProtocolConstraintValues.DEFINED_NAME_PATTERN)
          DefinedName name,
      NamedRangeScope scope,
      NamedRangeTarget target)
      implements StructuredMutationAction {
    public SetNamedRange {
      Objects.requireNonNull(scope, "scope must not be null");
      Objects.requireNonNull(target, "target must not be null");
      java.util.Objects.requireNonNull(name, "name must not be null");
    }
  }

  /** Deletes one existing named range from workbook or sheet scope. */
  @Variant(
      id = "DELETE_NAMED_RANGE",
      summary = "Delete one existing named range from workbook or sheet scope.",
      targetSelectors = {
        NamedRangeSelector.WorkbookScope.class,
        NamedRangeSelector.SheetScope.class
      },
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.STRUCTURED_MUTATION)
  record DeleteNamedRange() implements StructuredMutationAction {
    public DeleteNamedRange {}
  }
}
