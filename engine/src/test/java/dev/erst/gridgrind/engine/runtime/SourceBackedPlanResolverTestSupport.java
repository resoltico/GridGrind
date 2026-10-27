package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.mutate;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.mutations;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.request;

import dev.erst.gridgrind.core.model.chart.ChartDataSourceSpec;
import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ChartSeriesSpec;
import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextSpec;
import dev.erst.gridgrind.core.model.print.PrintLayoutSpec;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.MutationAction;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.util.Base64;
import java.util.List;

/** Shared helpers for source-backed plan resolver coverage slices. */
class SourceBackedPlanResolverTestSupport {

  final DrawingAnchorSpec.TwoCell twoCellAnchor() {
    return new DrawingAnchorSpec.TwoCell(
        new DrawingMarkerSpec(0, 0, 0, 0),
        new DrawingMarkerSpec(2, 3, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  final PrintLayoutSpec printLayoutWithHeader(HeaderFooterTextSpec header) {
    PrintLayoutSpec defaults = PrintLayoutSpec.defaults();
    return new PrintLayoutSpec(
        defaults.printArea(),
        defaults.orientation(),
        defaults.scaling(),
        defaults.repeatingRows(),
        defaults.repeatingColumns(),
        header,
        defaults.footer(),
        defaults.setup());
  }

  final PrintLayoutSpec printLayoutWithFooter(HeaderFooterTextSpec footer) {
    PrintLayoutSpec defaults = PrintLayoutSpec.defaults();
    return new PrintLayoutSpec(
        defaults.printArea(),
        defaults.orientation(),
        defaults.scaling(),
        defaults.repeatingRows(),
        defaults.repeatingColumns(),
        defaults.header(),
        footer,
        defaults.setup());
  }

  final PrintLayoutSpec printLayoutWithHeaderAndFooter(
      HeaderFooterTextSpec header, HeaderFooterTextSpec footer) {
    PrintLayoutSpec defaults = PrintLayoutSpec.defaults();
    return new PrintLayoutSpec(
        defaults.printArea(),
        defaults.orientation(),
        defaults.scaling(),
        defaults.repeatingRows(),
        defaults.repeatingColumns(),
        header,
        footer,
        defaults.setup());
  }

  final boolean requiresStandardInputFor(MutationAction action) {
    Object target =
        switch (action) {
          case CellMutationAction.SetCell _ ->
              new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1"));
          case CellMutationAction.SetRange _ ->
              new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B2"));
          case CellMutationAction.SetComment _ ->
              new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1"));
          case StructuredMutationAction.SetDataValidation _ ->
              new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1"));
          case StructuredMutationAction.SetTable setTable ->
              new dev.erst.gridgrind.core.plan.TableSelector.ByNameOnSheet(
                  setTable.table().name(), setTable.table().sheetName());
          default -> new SheetSelector.ByName(new SheetName("Budget"));
        };
    WorkbookPlan plan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(mutate((dev.erst.gridgrind.core.plan.Selector) target, action)));
    return SourceBackedPlanResolver.requiresStandardInput(plan);
  }

  final PictureDataSpec inlinePictureData() {
    return new PictureDataSpec(
        ExcelPictureFormat.PNG,
        new BinarySourceSpec.InlineBase64(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC"));
  }

  final ChartSeriesSpec chartSeries() {
    return new ChartSeriesSpec(
        new ChartTitleSpec.None(),
        new ChartDataSourceSpec.Reference("Budget!$A$2:$A$3"),
        new ChartDataSourceSpec.Reference("Budget!$B$2:$B$3"),
        java.util.Optional.empty(),
        java.util.Optional.empty(),
        java.util.Optional.empty(),
        java.util.Optional.empty());
  }

  final ChartSpec chartInput(
      String name,
      DrawingAnchorSpec.TwoCell anchor,
      ChartTitleSpec title,
      ChartLegendSpec legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      Boolean plotOnlyVisibleCells,
      ChartPlotSpec plot) {
    return new ChartSpec(
        name,
        anchor,
        title == null ? new ChartTitleSpec.None() : title,
        legend == null
            ? new ChartLegendSpec.Visible(
                dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition.RIGHT)
            : legend,
        displayBlanksAs == null ? ExcelChartDisplayBlanksAs.GAP : displayBlanksAs,
        plotOnlyVisibleCells == null ? Boolean.TRUE : plotOnlyVisibleCells,
        List.of(plot));
  }

  final byte[] pngBytes() {
    return Base64.getDecoder()
        .decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC");
  }
}
