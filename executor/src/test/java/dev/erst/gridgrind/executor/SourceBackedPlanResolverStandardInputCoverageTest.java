package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.mutate;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.mutations;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.request;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ChartPlotInput;
import dev.erst.gridgrind.contract.dto.ChartTitleInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.HeaderFooterTextInput;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Source-backed standard-input requirement coverage. */
class SourceBackedPlanResolverStandardInputCoverageTest
    extends SourceBackedPlanResolverTestSupport {
  @Test
  void requiresStandardInputCoversAllMutationFamiliesAndNonMutationSteps() {
    CellMutationAction.SetRange fullyInlineRange =
        new CellMutationAction.SetRange(
            List.of(
                List.of(
                    new CellInput.Blank(),
                    new CellInput.RichText(
                        List.of(new RichTextRunInput(TextSourceInput.inline("Ada"), null))),
                    new CellInput.Numeric(1.0d),
                    new CellInput.BooleanValue(true),
                    new CellInput.Date(LocalDate.of(2026, 4, 18)),
                    new CellInput.DateTime(LocalDateTime.of(2026, 4, 18, 12, 30)),
                    new CellInput.Formula(TextSourceInput.inline("SUM(A1:A2)")))));

    assertFalse(
        SourceBackedPlanResolver.requiresStandardInput(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
                dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
                List.of(
                    new AssertionStep(
                        "assert-owner",
                        new CellSelector.ByAddress("Budget", "A1"),
                        new Assertion.CellValue(new ExpectedCellValue.Text("Owner")))))));
    assertFalse(
        SourceBackedPlanResolver.requiresStandardInput(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
                dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
                List.of(
                    new InspectionStep(
                        "inspect-cells",
                        new CellSelector.ByAddress("Budget", "A1"),
                        new InspectionQuery.GetCells())))));

    assertFalse(
        SourceBackedPlanResolver.requiresStandardInput(
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                mutations(
                    mutate(new RangeSelector.ByRange("Budget", "A1:G1"), fullyInlineRange)))));
    assertFalse(requiresStandardInputFor(new WorkbookMutationAction.EnsureSheet()));
    assertFalse(
        requiresStandardInputFor(new CellMutationAction.SetCell(new CellInput.Numeric(1.0d))));
    assertFalse(
        requiresStandardInputFor(new CellMutationAction.SetCell(new CellInput.BooleanValue(true))));
    assertFalse(
        requiresStandardInputFor(
            new CellMutationAction.SetCell(new CellInput.Date(LocalDate.of(2026, 4, 18)))));
    assertFalse(
        requiresStandardInputFor(
            new CellMutationAction.SetCell(
                new CellInput.DateTime(LocalDateTime.of(2026, 4, 18, 12, 30)))));

    assertTrue(
        requiresStandardInputFor(
            new CellMutationAction.SetCell(new CellInput.Text(TextSourceInput.standardInput()))));
    assertTrue(
        requiresStandardInputFor(
            new CellMutationAction.SetRange(
                List.of(
                    List.of(
                        new CellInput.RichText(
                            List.of(
                                new RichTextRunInput(TextSourceInput.standardInput(), null))))))));
    assertTrue(
        requiresStandardInputFor(
            new CellMutationAction.SetComment(
                new CommentInput(
                    TextSourceInput.standardInput(),
                    "Ada",
                    false,
                    java.util.Optional.empty(),
                    java.util.Optional.empty()))));
    assertFalse(
        requiresStandardInputFor(
            new CellMutationAction.SetComment(
                new CommentInput(
                    TextSourceInput.inline("Ada"),
                    "Ada",
                    false,
                    java.util.Optional.empty(),
                    java.util.Optional.empty()))));
    assertTrue(
        requiresStandardInputFor(
            new CellMutationAction.SetComment(
                new CommentInput(
                    TextSourceInput.inline("Ada"),
                    "Ada",
                    false,
                    java.util.Optional.of(
                        List.of(new RichTextRunInput(TextSourceInput.standardInput(), null))),
                    java.util.Optional.empty()))));
    assertTrue(
        requiresStandardInputFor(
            new DrawingMutationAction.SetPicture(
                new PictureInput(
                    "Logo",
                    inlinePictureData(),
                    twoCellAnchor(),
                    TextSourceInput.standardInput()))));
    assertTrue(
        requiresStandardInputFor(
            new DrawingMutationAction.SetPicture(
                new PictureInput(
                    "Logo",
                    new PictureDataInput(ExcelPictureFormat.PNG, BinarySourceInput.standardInput()),
                    twoCellAnchor(),
                    null))));
    assertTrue(
        requiresStandardInputFor(
            new DrawingMutationAction.SetChart(
                chartInput(
                    "Bar",
                    twoCellAnchor(),
                    new ChartTitleInput.Text(TextSourceInput.standardInput()),
                    null,
                    ExcelChartDisplayBlanksAs.GAP,
                    true,
                    new ChartPlotInput.Bar(
                        false,
                        dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection.COLUMN,
                        dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping.CLUSTERED,
                        null,
                        null,
                        List.of(chartSeries()))))));
    assertFalse(
        requiresStandardInputFor(
            new DrawingMutationAction.SetChart(
                chartInput(
                    "Line",
                    twoCellAnchor(),
                    new ChartTitleInput.Formula("Budget!$A$1"),
                    null,
                    ExcelChartDisplayBlanksAs.ZERO,
                    true,
                    new ChartPlotInput.Line(
                        false,
                        dev.erst.gridgrind.excel.foundation.ExcelChartGrouping.STANDARD,
                        List.of(chartSeries()))))));
    assertFalse(
        requiresStandardInputFor(
            new DrawingMutationAction.SetChart(
                chartInput(
                    "Pie",
                    twoCellAnchor(),
                    new ChartTitleInput.None(),
                    null,
                    ExcelChartDisplayBlanksAs.GAP,
                    true,
                    new ChartPlotInput.Pie(false, 0, List.of(chartSeries()))))));
    assertTrue(
        requiresStandardInputFor(
            new DrawingMutationAction.SetShape(
                new ShapeInput(
                    "Shape",
                    ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                    twoCellAnchor(),
                    "roundRect",
                    TextSourceInput.standardInput()))));
    assertFalse(
        requiresStandardInputFor(
            new DrawingMutationAction.SetShape(
                new ShapeInput(
                    "Shape",
                    ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                    twoCellAnchor(),
                    "roundRect",
                    null))));
    assertTrue(
        requiresStandardInputFor(
            new DrawingMutationAction.SetEmbeddedObject(
                new EmbeddedObjectInput(
                    "Payload",
                    "Payload",
                    "payload.bin",
                    "open",
                    BinarySourceInput.standardInput(),
                    inlinePictureData(),
                    twoCellAnchor()))));
    assertTrue(
        requiresStandardInputFor(
            new DrawingMutationAction.SetEmbeddedObject(
                new EmbeddedObjectInput(
                    "Payload",
                    "Payload",
                    "payload.bin",
                    "open",
                    BinarySourceInput.inlineBase64("cGF5bG9hZA=="),
                    new PictureDataInput(ExcelPictureFormat.PNG, BinarySourceInput.standardInput()),
                    twoCellAnchor()))));
    assertFalse(
        requiresStandardInputFor(
            new DrawingMutationAction.SetEmbeddedObject(
                new EmbeddedObjectInput(
                    "Payload",
                    "Payload",
                    "payload.bin",
                    "open",
                    BinarySourceInput.inlineBase64("cGF5bG9hZA=="),
                    inlinePictureData(),
                    twoCellAnchor()))));
    assertTrue(
        requiresStandardInputFor(
            new StructuredMutationAction.SetDataValidation(
                new DataValidationInput(
                    new DataValidationRuleInput.ExplicitList(List.of("Open")),
                    false,
                    false,
                    new DataValidationPromptInput(
                        TextSourceInput.standardInput(), TextSourceInput.inline("Pick one"), false),
                    null))));
    assertTrue(
        requiresStandardInputFor(
            new StructuredMutationAction.SetDataValidation(
                new DataValidationInput(
                    new DataValidationRuleInput.ExplicitList(List.of("Open")),
                    false,
                    false,
                    null,
                    new DataValidationErrorAlertInput(
                        ExcelDataValidationErrorStyle.STOP,
                        TextSourceInput.inline("Invalid"),
                        TextSourceInput.standardInput(),
                        false)))));
    assertFalse(
        requiresStandardInputFor(
            new StructuredMutationAction.SetDataValidation(
                new DataValidationInput(
                    new DataValidationRuleInput.ExplicitList(List.of("Open")),
                    false,
                    false,
                    null,
                    null))));
    assertTrue(
        requiresStandardInputFor(
            new StructuredMutationAction.SetTable(
                new TableInput(
                    "BudgetTable",
                    "Budget",
                    "A1:B2",
                    false,
                    true,
                    new TableStyleInput.Named("TableStyleMedium2", false, false, true, false),
                    TextSourceInput.standardInput(),
                    false,
                    false,
                    false,
                    "",
                    "",
                    "",
                    List.of()))));
    assertTrue(
        requiresStandardInputFor(
            new CellMutationAction.AppendRow(
                List.of(new CellInput.Formula(TextSourceInput.standardInput())))));
    assertTrue(
        requiresStandardInputFor(
            new WorkbookMutationAction.SetPrintLayout(
                printLayoutWithHeader(
                    new HeaderFooterTextInput(
                        TextSourceInput.standardInput(),
                        TextSourceInput.inline(""),
                        TextSourceInput.inline(""))))));
    assertTrue(
        requiresStandardInputFor(
            new WorkbookMutationAction.SetPrintLayout(
                printLayoutWithHeader(
                    new HeaderFooterTextInput(
                        TextSourceInput.inline("left"),
                        TextSourceInput.standardInput(),
                        TextSourceInput.inline("right"))))));
    assertTrue(
        requiresStandardInputFor(
            new WorkbookMutationAction.SetPrintLayout(
                printLayoutWithHeader(
                    new HeaderFooterTextInput(
                        TextSourceInput.inline("left"),
                        TextSourceInput.inline("center"),
                        TextSourceInput.standardInput())))));
    assertTrue(
        requiresStandardInputFor(
            new WorkbookMutationAction.SetPrintLayout(
                printLayoutWithFooter(
                    new HeaderFooterTextInput(
                        TextSourceInput.standardInput(),
                        TextSourceInput.inline(""),
                        TextSourceInput.inline(""))))));
  }

  @Test
  void requiresStandardInputTracksSemanticTableSelectors() {
    TableCellSelector.ByColumnName standardInputTarget =
        new TableCellSelector.ByColumnName(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName("BudgetTable"),
                "Item",
                new CellInput.Text(new TextSourceInput.StandardInput())),
            "Amount");
    TableCellSelector.ByColumnName inlineTarget =
        new TableCellSelector.ByColumnName(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName("BudgetTable"),
                "Item",
                new CellInput.Text(TextSourceInput.inline("Hosting"))),
            "Amount");

    WorkbookPlan standardInputPlan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
            dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
            List.of(
                new InspectionStep(
                    "inspect-amount", standardInputTarget, new InspectionQuery.GetCells())));
    WorkbookPlan inlinePlan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
            dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
            List.of(
                new InspectionStep(
                    "inspect-amount", inlineTarget, new InspectionQuery.GetCells())));

    assertTrue(SourceBackedPlanResolver.requiresStandardInput(standardInputPlan));
    assertFalse(SourceBackedPlanResolver.requiresStandardInput(inlinePlan));
  }
}
