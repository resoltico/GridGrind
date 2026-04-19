package dev.erst.gridgrind.contract.selector;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.List;

/** Selects one or more named ranges by scope-aware identity. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = NamedRangeSelector.All.class, name = "ALL"),
  @JsonSubTypes.Type(value = NamedRangeSelector.AnyOf.class, name = "ANY_OF"),
  @JsonSubTypes.Type(value = NamedRangeSelector.ByName.class, name = "BY_NAME"),
  @JsonSubTypes.Type(value = NamedRangeSelector.ByNames.class, name = "BY_NAMES"),
  @JsonSubTypes.Type(value = NamedRangeSelector.WorkbookScope.class, name = "WORKBOOK_SCOPE"),
  @JsonSubTypes.Type(value = NamedRangeSelector.SheetScope.class, name = "SHEET_SCOPE")
})
public sealed interface NamedRangeSelector extends Selector
    permits NamedRangeSelector.All,
        NamedRangeSelector.AnyOf,
        NamedRangeSelector.ByName,
        NamedRangeSelector.ByNames,
        NamedRangeSelector.ScopedExact {

  /** Selects every user-facing named range. */
  record All() implements NamedRangeSelector {
    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects the union of one or more explicit named-range references. */
  record AnyOf(List<Ref> selectors) implements NamedRangeSelector {
    public AnyOf {
      selectors = SelectorSupport.copyDistinctNamedRangeRefs(selectors, "selectors");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }

  /** Selects every named range whose name matches across all scopes. */
  record ByName(String name) implements NamedRangeSelector, Ref {
    public ByName {
      name = SelectorSupport.requireDefinedName(name, "name");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }

  /** Selects every named range whose name appears in the provided name set across all scopes. */
  record ByNames(List<String> names) implements NamedRangeSelector {
    public ByNames {
      names = SelectorSupport.copyDistinctDefinedNames(names, "names");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }

  /** Selects one workbook-scoped named range by exact name. */
  record WorkbookScope(String name) implements ScopedExact, Ref {
    public WorkbookScope {
      name = SelectorSupport.requireDefinedName(name, "name");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  /** Selects one sheet-scoped named range by exact name and exact sheet. */
  record SheetScope(String name, String sheetName) implements ScopedExact, Ref {
    public SheetScope {
      name = SelectorSupport.requireDefinedName(name, "name");
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  /** Marker for one explicit named-range reference used inside {@link AnyOf}. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = ByName.class, name = "BY_NAME"),
    @JsonSubTypes.Type(value = WorkbookScope.class, name = "WORKBOOK_SCOPE"),
    @JsonSubTypes.Type(value = SheetScope.class, name = "SHEET_SCOPE")
  })
  sealed interface Ref permits ByName, WorkbookScope, SheetScope {}

  /** Marker for one exact named-range identity that can be mutated or deleted authoritatively. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = WorkbookScope.class, name = "WORKBOOK_SCOPE"),
    @JsonSubTypes.Type(value = SheetScope.class, name = "SHEET_SCOPE")
  })
  sealed interface ScopedExact extends NamedRangeSelector permits WorkbookScope, SheetScope {}
}
