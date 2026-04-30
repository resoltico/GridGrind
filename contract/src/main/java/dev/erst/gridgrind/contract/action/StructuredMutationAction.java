package dev.erst.gridgrind.contract.action;

import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Mutation family for workbook-scoped structured Excel features. */
public sealed interface StructuredMutationAction extends MutationAction {
  /** Imports one XML document into one existing workbook custom-XML mapping. */
  record ImportCustomXmlMapping(CustomXmlImportInput mapping) implements StructuredMutationAction {
    public ImportCustomXmlMapping {
      Objects.requireNonNull(mapping, "mapping must not be null");
    }
  }

  /** Creates or replaces one workbook-global pivot-table definition. */
  record SetPivotTable(PivotTableInput pivotTable) implements StructuredMutationAction {
    public SetPivotTable {
      Objects.requireNonNull(pivotTable, "pivotTable must not be null");
    }
  }

  /** Creates or replaces one data-validation rule over the requested sheet range. */
  record SetDataValidation(DataValidationInput validation) implements StructuredMutationAction {
    public SetDataValidation {
      Objects.requireNonNull(validation, "validation must not be null");
    }
  }

  /** Removes data-validation structures on the sheet that match the provided range selection. */
  record ClearDataValidations() implements StructuredMutationAction {
    public ClearDataValidations {}
  }

  /** Creates or replaces one logical conditional-formatting block over the requested ranges. */
  record SetConditionalFormatting(ConditionalFormattingBlockInput conditionalFormatting)
      implements StructuredMutationAction {
    public SetConditionalFormatting {
      Objects.requireNonNull(conditionalFormatting, "conditionalFormatting must not be null");
    }
  }

  /** Removes conditional-formatting blocks on the sheet that match the provided range selection. */
  record ClearConditionalFormatting() implements StructuredMutationAction {
    public ClearConditionalFormatting {}
  }

  /** Creates or replaces one sheet-level autofilter range. */
  record SetAutofilter(
      List<AutofilterFilterColumnInput> criteria, AutofilterSortStateInput sortState)
      implements StructuredMutationAction {
    /** Creates a plain sheet-level autofilter without criteria or explicit sort state. */
    public SetAutofilter() {
      this(List.of(), null);
    }

    public SetAutofilter {
      Objects.requireNonNull(criteria, "criteria must not be null");
      List<AutofilterFilterColumnInput> copiedCriteria = new ArrayList<>(criteria.size());
      for (AutofilterFilterColumnInput criterion : criteria) {
        Objects.requireNonNull(criterion, "criteria must not contain null values");
        copiedCriteria.add(criterion);
      }
      criteria = List.copyOf(copiedCriteria);
    }
  }

  /** Clears the sheet-level autofilter range on one sheet. */
  record ClearAutofilter() implements StructuredMutationAction {
    public ClearAutofilter {}
  }

  /** Creates or replaces one workbook-global table definition. */
  record SetTable(TableInput table) implements StructuredMutationAction {
    public SetTable {
      Objects.requireNonNull(table, "table must not be null");
    }
  }

  /** Deletes one existing table by workbook-global name and expected sheet. */
  record DeleteTable() implements StructuredMutationAction {
    public DeleteTable {}
  }

  /** Deletes one existing pivot table by workbook-global name and expected sheet. */
  record DeletePivotTable() implements StructuredMutationAction {
    public DeletePivotTable {}
  }

  /** Creates or replaces one typed named range in workbook or sheet scope. */
  record SetNamedRange(String name, NamedRangeScope scope, NamedRangeTarget target)
      implements StructuredMutationAction {
    public SetNamedRange {
      Objects.requireNonNull(scope, "scope must not be null");
      Objects.requireNonNull(target, "target must not be null");
      MutationAction.Validation.requireNamedRangeName(name);
    }
  }

  /** Deletes one existing named range from workbook or sheet scope. */
  record DeleteNamedRange() implements StructuredMutationAction {
    public DeleteNamedRange {}
  }
}
