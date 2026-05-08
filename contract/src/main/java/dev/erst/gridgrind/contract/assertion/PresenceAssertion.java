package dev.erst.gridgrind.contract.assertion;

import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadata;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;

/** Presence and absence assertions over named ranges, tables, pivots, and charts. */
public sealed interface PresenceAssertion extends Assertion
    permits PresenceAssertion.NamedRangePresent,
        PresenceAssertion.NamedRangeAbsent,
        PresenceAssertion.TablePresent,
        PresenceAssertion.TableAbsent,
        PresenceAssertion.PivotTablePresent,
        PresenceAssertion.PivotTableAbsent,
        PresenceAssertion.ChartPresent,
        PresenceAssertion.ChartAbsent {

  @ProtocolTypeMetadata(
      id = "EXPECT_NAMED_RANGE_PRESENT",
      summary = "Require the selected named-range selector to resolve to one or more named ranges.",
      targetSelectors = {NamedRangeSelector.class})
  record NamedRangePresent() implements PresenceAssertion {}

  @ProtocolTypeMetadata(
      id = "EXPECT_NAMED_RANGE_ABSENT",
      summary = "Require the selected named-range selector to resolve to no named ranges.",
      targetSelectors = {NamedRangeSelector.class})
  record NamedRangeAbsent() implements PresenceAssertion {}

  @ProtocolTypeMetadata(
      id = "EXPECT_TABLE_PRESENT",
      summary = "Require the selected table selector to resolve to one or more tables.",
      targetSelectors = {TableSelector.class})
  record TablePresent() implements PresenceAssertion {}

  @ProtocolTypeMetadata(
      id = "EXPECT_TABLE_ABSENT",
      summary = "Require the selected table selector to resolve to no tables.",
      targetSelectors = {TableSelector.class})
  record TableAbsent() implements PresenceAssertion {}

  @ProtocolTypeMetadata(
      id = "EXPECT_PIVOT_TABLE_PRESENT",
      summary = "Require the selected pivot-table selector to resolve to one or more pivot tables.",
      targetSelectors = {PivotTableSelector.class})
  record PivotTablePresent() implements PresenceAssertion {}

  @ProtocolTypeMetadata(
      id = "EXPECT_PIVOT_TABLE_ABSENT",
      summary = "Require the selected pivot-table selector to resolve to no pivot tables.",
      targetSelectors = {PivotTableSelector.class})
  record PivotTableAbsent() implements PresenceAssertion {}

  @ProtocolTypeMetadata(
      id = "EXPECT_CHART_PRESENT",
      summary = "Require the selected chart selector to resolve to one or more charts.",
      targetSelectors = {ChartSelector.class})
  record ChartPresent() implements PresenceAssertion {}

  @ProtocolTypeMetadata(
      id = "EXPECT_CHART_ABSENT",
      summary = "Require the selected chart selector to resolve to no charts.",
      targetSelectors = {ChartSelector.class})
  record ChartAbsent() implements PresenceAssertion {}
}
