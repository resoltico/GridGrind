package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Variant;

/** Presence and absence assertions over sheets, named ranges, tables, pivots, and charts. */
public sealed interface PresenceAssertion extends Assertion
    permits PresenceAssertion.SheetPresent,
        PresenceAssertion.SheetAbsent,
        PresenceAssertion.NamedRangePresent,
        PresenceAssertion.NamedRangeAbsent,
        PresenceAssertion.TablePresent,
        PresenceAssertion.TableAbsent,
        PresenceAssertion.PivotTablePresent,
        PresenceAssertion.PivotTableAbsent,
        PresenceAssertion.ChartPresent,
        PresenceAssertion.ChartAbsent {

  @Variant(
      id = "EXPECT_SHEET_PRESENT",
      summary = "Require the selected sheet selector to resolve to one or more sheets.",
      targetSelectors = {SheetSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record SheetPresent() implements PresenceAssertion {}

  @Variant(
      id = "EXPECT_SHEET_ABSENT",
      summary = "Require the selected sheet selector to resolve to no sheets.",
      targetSelectors = {SheetSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record SheetAbsent() implements PresenceAssertion {}

  @Variant(
      id = "EXPECT_NAMED_RANGE_PRESENT",
      summary = "Require the selected named-range selector to resolve to one or more named ranges.",
      targetSelectors = {NamedRangeSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record NamedRangePresent() implements PresenceAssertion {}

  @Variant(
      id = "EXPECT_NAMED_RANGE_ABSENT",
      summary = "Require the selected named-range selector to resolve to no named ranges.",
      targetSelectors = {NamedRangeSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record NamedRangeAbsent() implements PresenceAssertion {}

  @Variant(
      id = "EXPECT_TABLE_PRESENT",
      summary = "Require the selected table selector to resolve to one or more tables.",
      targetSelectors = {TableSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record TablePresent() implements PresenceAssertion {}

  @Variant(
      id = "EXPECT_TABLE_ABSENT",
      summary = "Require the selected table selector to resolve to no tables.",
      targetSelectors = {TableSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record TableAbsent() implements PresenceAssertion {}

  @Variant(
      id = "EXPECT_PIVOT_TABLE_PRESENT",
      summary = "Require the selected pivot-table selector to resolve to one or more pivot tables.",
      targetSelectors = {PivotTableSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record PivotTablePresent() implements PresenceAssertion {}

  @Variant(
      id = "EXPECT_PIVOT_TABLE_ABSENT",
      summary = "Require the selected pivot-table selector to resolve to no pivot tables.",
      targetSelectors = {PivotTableSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record PivotTableAbsent() implements PresenceAssertion {}

  @Variant(
      id = "EXPECT_CHART_PRESENT",
      summary = "Require the selected chart selector to resolve to one or more charts.",
      targetSelectors = {ChartSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record ChartPresent() implements PresenceAssertion {}

  @Variant(
      id = "EXPECT_CHART_ABSENT",
      summary = "Require the selected chart selector to resolve to no charts.",
      targetSelectors = {ChartSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record ChartAbsent() implements PresenceAssertion {}
}
