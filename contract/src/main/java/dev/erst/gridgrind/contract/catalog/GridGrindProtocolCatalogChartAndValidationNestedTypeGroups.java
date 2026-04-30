package dev.erst.gridgrind.contract.catalog;

import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.descriptor;
import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.nestedTypeGroup;

import dev.erst.gridgrind.contract.dto.AutofilterFilterCriterionInput;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.ChartDataSourceInput;
import dev.erst.gridgrind.contract.dto.ChartLegendInput;
import dev.erst.gridgrind.contract.dto.ChartPlotInput;
import dev.erst.gridgrind.contract.dto.ChartTitleInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.FontHeightInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.PrintAreaInput;
import dev.erst.gridgrind.contract.dto.PrintScalingInput;
import dev.erst.gridgrind.contract.dto.PrintTitleColumnsInput;
import dev.erst.gridgrind.contract.dto.PrintTitleRowsInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import java.util.List;

/** Owns one focused subset of nested-type group descriptors for the protocol catalog. */
final class GridGrindProtocolCatalogChartAndValidationNestedTypeGroups {
  private GridGrindProtocolCatalogChartAndValidationNestedTypeGroups() {}

  static final List<CatalogNestedTypeDescriptor> CHART_AND_VALIDATION_GROUPS =
      List.of(
          nestedTypeGroup(
              "drawingAnchorInputTypes",
              DrawingAnchorInput.class,
              List.of(
                  descriptor(
                      DrawingAnchorInput.TwoCell.class,
                      "TWO_CELL",
                      "Authored two-cell drawing anchor with explicit start and end markers."
                          + " behavior is explicit on the wire; the Java authoring surface"
                          + " exposes DrawingAnchorInput.TwoCell.moveAndResize(...)."))),
          nestedTypeGroup(
              "chartPlotInputTypes",
              ChartPlotInput.class,
              List.of(
                  descriptor(
                      ChartPlotInput.Area.class,
                      "AREA",
                      "Authored area-chart plot."
                          + " Supply explicit grouping and axes, or use the Java convenience"
                          + " constructor for the standard category/value axis pair.",
                      "varyColors"),
                  descriptor(
                      ChartPlotInput.Area3D.class,
                      "AREA_3D",
                      "Authored 3D area-chart plot."
                          + " Supply explicit grouping and axes, or use the Java convenience"
                          + " constructor for the standard category/value axis pair.",
                      "varyColors",
                      "gapDepth"),
                  descriptor(
                      ChartPlotInput.Bar.class,
                      "BAR",
                      "Authored bar-chart plot."
                          + " Supply explicit barDirection, grouping, and axes."
                          + " The Java authoring surface exposes one convenience constructor for"
                          + " the standard category/value axis pair.",
                      "varyColors",
                      "gapWidth",
                      "overlap"),
                  descriptor(
                      ChartPlotInput.Bar3D.class,
                      "BAR_3D",
                      "Authored 3D bar-chart plot."
                          + " Supply explicit barDirection, grouping, and axes."
                          + " The Java authoring surface exposes one convenience constructor for"
                          + " the standard category/value axis pair.",
                      "varyColors",
                      "gapDepth",
                      "gapWidth",
                      "shape"),
                  descriptor(
                      ChartPlotInput.Doughnut.class,
                      "DOUGHNUT",
                      "Authored doughnut-chart plot."
                          + " varyColors, firstSliceAngle, and holeSize are optional.",
                      "varyColors",
                      "firstSliceAngle",
                      "holeSize"),
                  descriptor(
                      ChartPlotInput.Line.class,
                      "LINE",
                      "Authored line-chart plot."
                          + " Supply explicit grouping and axes, or use the Java convenience"
                          + " constructor for the standard category/value axis pair.",
                      "varyColors"),
                  descriptor(
                      ChartPlotInput.Line3D.class,
                      "LINE_3D",
                      "Authored 3D line-chart plot."
                          + " Supply explicit grouping and axes, or use the Java convenience"
                          + " constructor for the standard category/value axis pair.",
                      "varyColors",
                      "gapDepth"),
                  descriptor(
                      ChartPlotInput.Pie.class,
                      "PIE",
                      "Authored pie-chart plot." + " varyColors and firstSliceAngle are optional.",
                      "varyColors",
                      "firstSliceAngle"),
                  descriptor(
                      ChartPlotInput.Pie3D.class,
                      "PIE_3D",
                      "Authored 3D pie-chart plot.",
                      "varyColors"),
                  descriptor(
                      ChartPlotInput.Radar.class,
                      "RADAR",
                      "Authored radar-chart plot."
                          + " Supply explicit style and axes, or use the Java convenience"
                          + " constructor for the standard category/value axis pair.",
                      "varyColors"),
                  descriptor(
                      ChartPlotInput.Scatter.class,
                      "SCATTER",
                      "Authored scatter-chart plot."
                          + " Supply explicit style and axes, or use the Java convenience"
                          + " constructor for the standard X/Y axis pair.",
                      "varyColors"),
                  descriptor(
                      ChartPlotInput.Surface.class,
                      "SURFACE",
                      "Authored surface-chart plot."
                          + " Supply explicit axes, or use the Java convenience constructor for"
                          + " the standard category/value/series axis set.",
                      "varyColors",
                      "wireframe"),
                  descriptor(
                      ChartPlotInput.Surface3D.class,
                      "SURFACE_3D",
                      "Authored 3D surface-chart plot."
                          + " Supply explicit axes, or use the Java convenience constructor for"
                          + " the standard category/value/series axis set.",
                      "varyColors",
                      "wireframe"))),
          nestedTypeGroup(
              "chartTitleInputTypes",
              ChartTitleInput.class,
              List.of(
                  descriptor(
                      ChartTitleInput.None.class, "NONE", "Remove any chart or series title."),
                  descriptor(ChartTitleInput.Text.class, "TEXT", "Use one explicit static title."),
                  descriptor(
                      ChartTitleInput.Formula.class,
                      "FORMULA",
                      "Bind the chart or series title to one workbook formula that resolves"
                          + " to one cell."))),
          nestedTypeGroup(
              "chartLegendInputTypes",
              ChartLegendInput.class,
              List.of(
                  descriptor(ChartLegendInput.Hidden.class, "HIDDEN", "Hide the legend entirely."),
                  descriptor(
                      ChartLegendInput.Visible.class,
                      "VISIBLE",
                      "Show the legend at one explicit position."))),
          nestedTypeGroup(
              "chartDataSourceInputTypes",
              ChartDataSourceInput.class,
              List.of(
                  descriptor(
                      ChartDataSourceInput.Reference.class,
                      "REFERENCE",
                      "Workbook-backed chart source formula or defined name."),
                  descriptor(
                      ChartDataSourceInput.StringLiteral.class,
                      "STRING_LITERAL",
                      "Literal string chart source stored directly in the chart part."),
                  descriptor(
                      ChartDataSourceInput.NumericLiteral.class,
                      "NUMERIC_LITERAL",
                      "Literal numeric chart source stored directly in the chart part."))),
          nestedTypeGroup(
              "pivotTableSourceTypes",
              PivotTableInput.Source.class,
              List.of(
                  descriptor(
                      PivotTableInput.Source.Range.class,
                      "RANGE",
                      "Use one explicit contiguous sheet range with the header row in the first"
                          + " row."),
                  descriptor(
                      PivotTableInput.Source.NamedRange.class,
                      "NAMED_RANGE",
                      "Use one existing workbook- or sheet-scoped named range as the pivot"
                          + " source."),
                  descriptor(
                      PivotTableInput.Source.Table.class,
                      "TABLE",
                      "Use one existing workbook-global table as the pivot source."))),
          nestedTypeGroup(
              "fontHeightTypes",
              FontHeightInput.class,
              List.of(
                  descriptor(
                      FontHeightInput.Points.class,
                      "POINTS",
                      "Specify font height in point units."
                          + " Write format: {\"type\":\"POINTS\",\"points\":13}."
                          + " Read-back (GET_CELLS, GET_WINDOW): style.fontHeight is"
                          + " {\"twips\":260,\"points\":13} with both fields present,"
                          + " not this discriminated type format."),
                  descriptor(
                      FontHeightInput.Twips.class,
                      "TWIPS",
                      "Specify font height in exact twips (20 twips = 1 point)."
                          + " Write format: {\"type\":\"TWIPS\",\"twips\":260}."
                          + " Read-back returns the same plain object shape as POINTS."))),
          GridGrindProtocolCatalogStyleTypeGroups.COLOR_INPUT_TYPES,
          GridGrindProtocolCatalogStyleTypeGroups.CELL_GRADIENT_FILL_INPUT_TYPES,
          GridGrindProtocolCatalogStyleTypeGroups.CELL_FILL_INPUT_TYPES,
          GridGrindProtocolCatalogStyleTypeGroups.CELL_COLOR_REPORT_TYPES,
          GridGrindProtocolCatalogStyleTypeGroups.CELL_GRADIENT_FILL_REPORT_TYPES,
          GridGrindProtocolCatalogStyleTypeGroups.CELL_FILL_REPORT_TYPES,
          nestedTypeGroup(
              "dataValidationRuleTypes",
              DataValidationRuleInput.class,
              List.of(
                  descriptor(
                      DataValidationRuleInput.ExplicitList.class,
                      "EXPLICIT_LIST",
                      "Allow only one of the supplied explicit values."
                          + " An empty values array preserves Excel's explicit-empty-list state."),
                  descriptor(
                      DataValidationRuleInput.FormulaList.class,
                      "FORMULA_LIST",
                      "Allow values from a formula-driven list expression."),
                  descriptor(
                      DataValidationRuleInput.WholeNumber.class,
                      "WHOLE_NUMBER",
                      "Apply a whole-number comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.DecimalNumber.class,
                      "DECIMAL_NUMBER",
                      "Apply a decimal-number comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.DateRule.class,
                      "DATE",
                      "Apply a date comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.TimeRule.class,
                      "TIME",
                      "Apply a time comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.TextLength.class,
                      "TEXT_LENGTH",
                      "Apply a text-length comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.CustomFormula.class,
                      "CUSTOM_FORMULA",
                      "Allow values that satisfy a custom formula."))),
          nestedTypeGroup(
              "autofilterFilterCriterionTypes",
              AutofilterFilterCriterionInput.class,
              List.of(
                  descriptor(
                      AutofilterFilterCriterionInput.Values.class,
                      "VALUES",
                      "Retain rows whose cell values match one or more explicit values."),
                  descriptor(
                      AutofilterFilterCriterionInput.Custom.class,
                      "CUSTOM",
                      "Retain rows that satisfy one or two comparator-based custom conditions."),
                  descriptor(
                      AutofilterFilterCriterionInput.Dynamic.class,
                      "DYNAMIC",
                      "Retain rows using one dynamic-date or moving-window autofilter rule.",
                      "value",
                      "maxValue"),
                  descriptor(
                      AutofilterFilterCriterionInput.Top10.class,
                      "TOP10",
                      "Retain top or bottom N or percent values."),
                  descriptor(
                      AutofilterFilterCriterionInput.Color.class,
                      "COLOR",
                      "Retain rows matching one cell color or font color criterion."),
                  descriptor(
                      AutofilterFilterCriterionInput.Icon.class,
                      "ICON",
                      "Retain rows matching one icon-set member."))),
          nestedTypeGroup(
              "conditionalFormattingRuleTypes",
              ConditionalFormattingRuleInput.class,
              List.of(
                  descriptor(
                      ConditionalFormattingRuleInput.FormulaRule.class,
                      "FORMULA_RULE",
                      "Apply one formula-driven conditional-formatting rule."
                          + " Supply one differential style."),
                  descriptor(
                      ConditionalFormattingRuleInput.CellValueRule.class,
                      "CELL_VALUE_RULE",
                      "Apply one cell-value comparison conditional-formatting rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN."
                          + " Supply one differential style.",
                      "formula2"),
                  descriptor(
                      ConditionalFormattingRuleInput.ColorScaleRule.class,
                      "COLOR_SCALE_RULE",
                      "Apply one color-scale conditional-formatting rule with ordered thresholds"
                          + " and colors."),
                  descriptor(
                      ConditionalFormattingRuleInput.DataBarRule.class,
                      "DATA_BAR_RULE",
                      "Apply one data-bar conditional-formatting rule with explicit thresholds"
                          + " and widths."),
                  descriptor(
                      ConditionalFormattingRuleInput.IconSetRule.class,
                      "ICON_SET_RULE",
                      "Apply one icon-set conditional-formatting rule with authored thresholds."),
                  descriptor(
                      ConditionalFormattingRuleInput.Top10Rule.class,
                      "TOP10_RULE",
                      "Apply one top/bottom-N conditional-formatting rule with a differential"
                          + " style."))),
          nestedTypeGroup(
              "printAreaTypes",
              PrintAreaInput.class,
              List.of(
                  descriptor(
                      PrintAreaInput.None.class, "NONE", "Sheet has no explicit print area."),
                  descriptor(
                      PrintAreaInput.Range.class,
                      "RANGE",
                      "Sheet prints the provided rectangular A1-style range."))),
          nestedTypeGroup(
              "printScalingTypes",
              PrintScalingInput.class,
              List.of(
                  descriptor(
                      PrintScalingInput.Automatic.class,
                      "AUTOMATIC",
                      "Sheet uses Excel's default scaling instead of fit-to-page counts."),
                  descriptor(
                      PrintScalingInput.Fit.class,
                      "FIT",
                      "Sheet fits printed content into the provided page counts."
                          + " A value of 0 on one axis keeps that axis unconstrained."))),
          nestedTypeGroup(
              "printTitleRowsTypes",
              PrintTitleRowsInput.class,
              List.of(
                  descriptor(
                      PrintTitleRowsInput.None.class,
                      "NONE",
                      "Sheet has no repeating print-title rows."),
                  descriptor(
                      PrintTitleRowsInput.Band.class,
                      "BAND",
                      "Sheet repeats the provided inclusive zero-based row band on every printed page."))),
          nestedTypeGroup(
              "printTitleColumnsTypes",
              PrintTitleColumnsInput.class,
              List.of(
                  descriptor(
                      PrintTitleColumnsInput.None.class,
                      "NONE",
                      "Sheet has no repeating print-title columns."),
                  descriptor(
                      PrintTitleColumnsInput.Band.class,
                      "BAND",
                      "Sheet repeats the provided inclusive zero-based column band on every printed page."))),
          nestedTypeGroup(
              "tableStyleTypes",
              TableStyleInput.class,
              List.of(
                  descriptor(TableStyleInput.None.class, "NONE", "Clear table style metadata."),
                  descriptor(
                      TableStyleInput.Named.class,
                      "NAMED",
                      "Apply one named workbook table style with explicit stripe and emphasis"
                          + " flags."))),
          nestedTypeGroup(
              "calculationStrategyTypes",
              CalculationStrategyInput.class,
              List.of(
                  descriptor(
                      CalculationStrategyInput.DoNotCalculate.class,
                      "DO_NOT_CALCULATE",
                      "Skip immediate server-side formula calculation."),
                  descriptor(
                      CalculationStrategyInput.EvaluateAll.class,
                      "EVALUATE_ALL",
                      "Preflight and evaluate every formula cell after mutation steps complete."),
                  descriptor(
                      CalculationStrategyInput.EvaluateTargets.class,
                      "EVALUATE_TARGETS",
                      "Preflight and evaluate the explicit formula-cell target list only.",
                      "cells"),
                  descriptor(
                      CalculationStrategyInput.ClearCachesOnly.class,
                      "CLEAR_CACHES_ONLY",
                      "Strip persisted formula caches without running immediate evaluation."))));
}
