package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.mutate;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.mutations;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.request;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.dto.ChartDataSourceInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ChartLegendInput;
import dev.erst.gridgrind.contract.dto.ChartPlotInput;
import dev.erst.gridgrind.contract.dto.ChartSeriesInput;
import dev.erst.gridgrind.contract.dto.ChartTitleInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.DrawingMarkerInput;
import dev.erst.gridgrind.contract.dto.HeaderFooterTextInput;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
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
          case CellMutationAction.SetCell _ -> new CellSelector.ByAddress("Budget", "A1");
          case CellMutationAction.SetRange _ -> new RangeSelector.ByRange("Budget", "A1:B2");
          case CellMutationAction.SetComment _ -> new CellSelector.ByAddress("Budget", "A1");
          case StructuredMutationAction.SetDataValidation _ ->
              new RangeSelector.ByRange("Budget", "A1");
          case StructuredMutationAction.SetTable setTable ->
              new dev.erst.gridgrind.contract.selector.TableSelector.ByNameOnSheet(
                  setTable.table().name(), setTable.table().sheetName());
          default -> new SheetSelector.ByName("Budget");
        };
    WorkbookPlan plan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(mutate((dev.erst.gridgrind.contract.selector.Selector) target, action)));
    return SourceBackedPlanResolver.requiresStandardInput(plan);
  }

  final PictureDataInput inlinePictureData() {
    return new PictureDataInput(
        ExcelPictureFormat.PNG,
        new BinarySourceInput.InlineBase64(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII="));
  }

  final ChartSeriesInput chartSeries() {
    return new ChartSeriesInput(
        new ChartTitleInput.None(),
        new ChartDataSourceInput.Reference("Budget!$A$2:$A$3"),
        new ChartDataSourceInput.Reference("Budget!$B$2:$B$3"),
        null,
        null,
        null,
        null);
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
                dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition.RIGHT)
            : legend,
        displayBlanksAs == null ? ExcelChartDisplayBlanksAs.GAP : displayBlanksAs,
        plotOnlyVisibleCells == null ? Boolean.TRUE : plotOnlyVisibleCells,
        List.of(plot));
  }

  final byte[] pngBytes() {
    return Base64.getDecoder()
        .decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");
  }
}
