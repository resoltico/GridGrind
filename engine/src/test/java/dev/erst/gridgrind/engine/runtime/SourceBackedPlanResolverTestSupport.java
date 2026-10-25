package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.mutate;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.mutations;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.request;

import dev.erst.gridgrind.core.model.chart.ChartDataSourceInput;
import dev.erst.gridgrind.core.model.chart.ChartInput;
import dev.erst.gridgrind.core.model.chart.ChartLegendInput;
import dev.erst.gridgrind.core.model.chart.ChartPlotInput;
import dev.erst.gridgrind.core.model.chart.ChartSeriesInput;
import dev.erst.gridgrind.core.model.chart.ChartTitleInput;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorInput;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerInput;
import dev.erst.gridgrind.core.model.drawing.PictureDataInput;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextInput;
import dev.erst.gridgrind.core.model.print.PrintLayoutInput;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.MutationAction;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.source.BinarySourceInput;
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

  final DrawingAnchorInput.TwoCell twoCellAnchor() {
    return new DrawingAnchorInput.TwoCell(
        new DrawingMarkerInput(0, 0, 0, 0),
        new DrawingMarkerInput(2, 3, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  final PrintLayoutInput printLayoutWithHeader(HeaderFooterTextInput header) {
    PrintLayoutInput defaults = PrintLayoutInput.defaults();
    return new PrintLayoutInput(
        defaults.printArea(),
        defaults.orientation(),
        defaults.scaling(),
        defaults.repeatingRows(),
        defaults.repeatingColumns(),
        header,
        defaults.footer(),
        defaults.setup());
  }

  final PrintLayoutInput printLayoutWithFooter(HeaderFooterTextInput footer) {
    PrintLayoutInput defaults = PrintLayoutInput.defaults();
    return new PrintLayoutInput(
        defaults.printArea(),
        defaults.orientation(),
        defaults.scaling(),
        defaults.repeatingRows(),
        defaults.repeatingColumns(),
        defaults.header(),
        footer,
        defaults.setup());
  }

  final PrintLayoutInput printLayoutWithHeaderAndFooter(
      HeaderFooterTextInput header, HeaderFooterTextInput footer) {
    PrintLayoutInput defaults = PrintLayoutInput.defaults();
    return new PrintLayoutInput(
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

  final PictureDataInput inlinePictureData() {
    return new PictureDataInput(
        ExcelPictureFormat.PNG,
        new BinarySourceInput.InlineBase64(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC"));
  }

  final ChartSeriesInput chartSeries() {
    return new ChartSeriesInput(
        new ChartTitleInput.None(),
        new ChartDataSourceInput.Reference("Budget!$A$2:$A$3"),
        new ChartDataSourceInput.Reference("Budget!$B$2:$B$3"),
        java.util.Optional.empty(),
        java.util.Optional.empty(),
        java.util.Optional.empty(),
        java.util.Optional.empty());
  }

  final ChartInput chartInput(
      String name,
      DrawingAnchorInput.TwoCell anchor,
      ChartTitleInput title,
      ChartLegendInput legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      Boolean plotOnlyVisibleCells,
      ChartPlotInput plot) {
    return new ChartInput(
        name,
        anchor,
        title == null ? new ChartTitleInput.None() : title,
        legend == null
            ? new ChartLegendInput.Visible(
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
