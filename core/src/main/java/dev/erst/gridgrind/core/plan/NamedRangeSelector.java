package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.util.List;

/** Selects one or more named ranges by scope-aware identity. */
@Union(discriminator = "type", catalog = true, displayOrder = 36)
public sealed interface NamedRangeSelector extends Selector
    permits NamedRangeSelector.All,
        NamedRangeSelector.AnyOf,
        NamedRangeSelector.ByNames,
        NamedRangeSelector.ByName,
        NamedRangeSelector.ScopedExact {

  /** Selects every user-facing named range. */
  @Variant(id = "NAMED_RANGE_ALL", summary = "Select every user-facing named range.")
  record All() implements NamedRangeSelector {
    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects the union of one or more explicit named-range references. */
  @Variant(
      id = "NAMED_RANGE_ANY_OF",
      summary = "Select the union of one or more explicit named-range references.")
  record AnyOf(List<Ref> selectors) implements NamedRangeSelector {
    public AnyOf {
      selectors = SelectorListValidation.copyDistinctNamedRangeRefs(selectors, "selectors");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }

  /** Selects every named range whose name matches across all scopes. */
  @Variant(
      id = "NAMED_RANGE_BY_NAME",
      summary = "Match a named range reference across all scopes by exact name.")
  record ByName(DefinedName name) implements NamedRangeSelector, Ref {
    public ByName {
      java.util.Objects.requireNonNull(name, "name must not be null");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }

  /** Selects every named range whose name appears in the provided name set across all scopes. */
  @Variant(
      id = "NAMED_RANGE_BY_NAMES",
      summary = "Match named ranges across all scopes by exact name set.")
  record ByNames(List<DefinedName> names) implements NamedRangeSelector {
    public ByNames {
      names = SelectorListValidation.copyDistinctNamedRangeNames(names, "names");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }

  /** Selects one workbook-scoped named range by exact name. */
  @Variant(
      id = "NAMED_RANGE_WORKBOOK_SCOPE",
      summary = "Match one workbook-scoped named range reference by exact name.")
  record WorkbookScope(DefinedName name) implements ScopedExact, Ref {
    public WorkbookScope {
      java.util.Objects.requireNonNull(name, "name must not be null");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  /** Selects one sheet-scoped named range by exact name and exact sheet. */
  @Variant(
      id = "NAMED_RANGE_SHEET_SCOPE",
      summary = "Match one sheet-scoped named range reference on one sheet.")
  record SheetScope(DefinedName name, SheetName sheetName) implements ScopedExact, Ref {
    public SheetScope {
      java.util.Objects.requireNonNull(name, "name must not be null");
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  /** Marker for one explicit named-range reference used inside {@link AnyOf}. */
  @Union(discriminator = "type", catalog = true, displayOrder = 34)
  sealed interface Ref
      permits dev.erst.gridgrind.core.plan.NamedRangeSelector.ByName,
          dev.erst.gridgrind.core.plan.NamedRangeSelector.WorkbookScope,
          dev.erst.gridgrind.core.plan.NamedRangeSelector.SheetScope {}

  /** Marker for one exact named-range identity that can be mutated or deleted authoritatively. */
  @Union(discriminator = "type")
  sealed interface ScopedExact extends NamedRangeSelector permits WorkbookScope, SheetScope {}
}
