package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.drawing.EmbeddedObjectSpec;
import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextSpec;
import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Source-backed standard-input requirement coverage. */
class SourceBackedPlanResolverStdinCoverageTest extends SourceBackedPlanResolverTestSupport {
  @Test
  void requiresStandardInputCoversAllMutationFamiliesAndNonMutationSteps() {
    CellMutationAction.SetRange fullyInlineRange =
        new CellMutationAction.SetRange(
            new dev.erst.gridgrind.core.model.value.CellGridSpec.Typed(
                List.of(
                    List.of(
                        new CellValueSpec.Blank(),
                        new CellValueSpec.RichText(
                            List.of(
                                new RichTextRunSpec(
                                    TextSourceSpec.inline("Ada"), java.util.Optional.empty()))),
                        new CellValueSpec.NumberValue(1.0d),
                        new CellValueSpec.BooleanValue(true),
                        new CellValueSpec.Date(LocalDate.of(2026, 4, 18)),
                        new CellValueSpec.DateTime(LocalDateTime.of(2026, 4, 18, 12, 30)),
                        new CellValueSpec.Formula(TextSourceSpec.inline("SUM(A1:A2)"))))));

    assertFalse(
        SourceBackedPlanResolver.requiresStandardInput(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                dev.erst.gridgrind.core.plan.ExecutionPolicySpec.defaults(),
                dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec.empty(),
                List.of(
                    new AssertionStep(
                        new StepId("assert-owner"),
                        new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                        new CellAssertion.CellValue(
                            new dev.erst.gridgrind.core.model.value.CellScalarValue.Text(
                                "Owner")))))));
    assertFalse(
        SourceBackedPlanResolver.requiresStandardInput(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                dev.erst.gridgrind.core.plan.ExecutionPolicySpec.defaults(),
                dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec.empty(),
                List.of(
                    new InspectionStep(
                        new StepId("inspect-cells"),
                        new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                        new SheetIntrospectionQuery.GetCells())))));

    assertFalse(
        SourceBackedPlanResolver.requiresStandardInput(
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                mutations(
                    mutate(
                        new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:G1")),
                        fullyInlineRange)))));
    assertFalse(requiresStandardInputFor(new WorkbookMutationAction.EnsureSheet()));
    assertFalse(
        requiresStandardInputFor(
            new CellMutationAction.SetCell(new CellValueSpec.NumberValue(1.0d))));
    assertFalse(
        requiresStandardInputFor(
            new CellMutationAction.SetCell(new CellValueSpec.BooleanValue(true))));
    assertFalse(
        requiresStandardInputFor(
            new CellMutationAction.SetCell(new CellValueSpec.Date(LocalDate.of(2026, 4, 18)))));
    assertFalse(
        requiresStandardInputFor(
            new CellMutationAction.SetCell(
                new CellValueSpec.DateTime(LocalDateTime.of(2026, 4, 18, 12, 30)))));

    assertTrue(
        requiresStandardInputFor(
            new CellMutationAction.SetCell(
                new CellValueSpec.Text(TextSourceSpec.standardInput()))));
    assertTrue(
        requiresStandardInputFor(
            new CellMutationAction.SetRange(
                new dev.erst.gridgrind.core.model.value.CellGridSpec.Typed(
                    List.of(
                        List.of(
                            new CellValueSpec.RichText(
                                List.of(
                                    new RichTextRunSpec(
                                        TextSourceSpec.standardInput(),
                                        java.util.Optional.empty())))))))));
    assertTrue(
        requiresStandardInputFor(
            new CellMutationAction.SetComment(
                new CommentSpec(
                    TextSourceSpec.standardInput(),
                    "Ada",
                    false,
                    java.util.Optional.empty(),
                    java.util.Optional.empty()))));
    assertFalse(
        requiresStandardInputFor(
            new CellMutationAction.SetComment(
                new CommentSpec(
                    TextSourceSpec.inline("Ada"),
                    "Ada",
                    false,
                    java.util.Optional.empty(),
                    java.util.Optional.empty()))));
    assertTrue(
        requiresStandardInputFor(
            new CellMutationAction.SetComment(
                new CommentSpec(
                    TextSourceSpec.inline("Ada"),
                    "Ada",
                    false,
                    java.util.Optional.of(
                        List.of(richTextRun(TextSourceSpec.standardInput(), null))),
                    java.util.Optional.empty()))));
    assertTrue(
        requiresStandardInputFor(
            new DrawingMutationAction.SetPicture(
                picture(
                    "Logo",
                    inlinePictureData(),
                    twoCellAnchor(),
                    TextSourceSpec.standardInput()))));
    assertTrue(
        requiresStandardInputFor(
            new DrawingMutationAction.SetPicture(
                picture(
                    "Logo",
                    new PictureDataSpec(ExcelPictureFormat.PNG, BinarySourceSpec.standardInput()),
                    twoCellAnchor(),
                    null))));
    assertTrue(
        requiresStandardInputFor(
            new DrawingMutationAction.SetChart(
                chartInput(
                    "Bar",
                    twoCellAnchor(),
                    new ChartTitleSpec.Text(TextSourceSpec.standardInput()),
                    null,
                    ExcelChartDisplayBlanksAs.GAP,
                    true,
                    new ChartPlotSpec.Bar(
                        false,
                        dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection.COLUMN,
                        dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping.CLUSTERED,
                        Optional.empty(),
                        Optional.empty(),
                        List.of(chartSeries()))))));
    assertFalse(
        requiresStandardInputFor(
            new DrawingMutationAction.SetChart(
                chartInput(
                    "Line",
                    twoCellAnchor(),
                    new ChartTitleSpec.Formula("Budget!$A$1"),
                    null,
                    ExcelChartDisplayBlanksAs.ZERO,
                    true,
                    new ChartPlotSpec.Line(
                        false,
                        dev.erst.gridgrind.core.vocabulary.ExcelChartGrouping.STANDARD,
                        List.of(chartSeries()))))));
    assertFalse(
        requiresStandardInputFor(
            new DrawingMutationAction.SetChart(
                chartInput(
                    "Pie",
                    twoCellAnchor(),
                    new ChartTitleSpec.None(),
                    null,
                    ExcelChartDisplayBlanksAs.GAP,
                    true,
                    new ChartPlotSpec.Pie(false, Optional.of(0), List.of(chartSeries()))))));
    assertTrue(
        requiresStandardInputFor(
            new DrawingMutationAction.SetShape(
                simpleShape(
                    "Shape", twoCellAnchor(), "roundRect", TextSourceSpec.standardInput()))));
    assertFalse(
        requiresStandardInputFor(
            new DrawingMutationAction.SetShape(
                simpleShape("Shape", twoCellAnchor(), "roundRect", null))));
    assertTrue(
        requiresStandardInputFor(
            new DrawingMutationAction.SetEmbeddedObject(
                new EmbeddedObjectSpec(
                    "Payload",
                    "Payload",
                    "payload.bin",
                    "open",
                    BinarySourceSpec.standardInput(),
                    inlinePictureData(),
                    twoCellAnchor()))));
    assertTrue(
        requiresStandardInputFor(
            new DrawingMutationAction.SetEmbeddedObject(
                new EmbeddedObjectSpec(
                    "Payload",
                    "Payload",
                    "payload.bin",
                    "open",
                    BinarySourceSpec.inlineBase64("cGF5bG9hZA=="),
                    new PictureDataSpec(ExcelPictureFormat.PNG, BinarySourceSpec.standardInput()),
                    twoCellAnchor()))));
    assertFalse(
        requiresStandardInputFor(
            new DrawingMutationAction.SetEmbeddedObject(
                new EmbeddedObjectSpec(
                    "Payload",
                    "Payload",
                    "payload.bin",
                    "open",
                    BinarySourceSpec.inlineBase64("cGF5bG9hZA=="),
                    inlinePictureData(),
                    twoCellAnchor()))));
    assertTrue(
        requiresStandardInputFor(
            new StructuredMutationAction.SetDataValidation(
                dataValidation(
                    new DataValidationRuleSpec.ExplicitList(List.of("Open")),
                    false,
                    false,
                    new DataValidationPromptSpec(
                        TextSourceSpec.standardInput(), TextSourceSpec.inline("Pick one"), false),
                    null))));
    assertTrue(
        requiresStandardInputFor(
            new StructuredMutationAction.SetDataValidation(
                dataValidation(
                    new DataValidationRuleSpec.ExplicitList(List.of("Open")),
                    false,
                    false,
                    null,
                    new DataValidationErrorAlertSpec(
                        ExcelDataValidationErrorStyle.STOP,
                        TextSourceSpec.inline("Invalid"),
                        TextSourceSpec.standardInput(),
                        false)))));
    assertFalse(
        requiresStandardInputFor(
            new StructuredMutationAction.SetDataValidation(
                dataValidation(
                    new DataValidationRuleSpec.ExplicitList(List.of("Open")),
                    false,
                    false,
                    null,
                    null))));
    assertTrue(
        requiresStandardInputFor(
            new StructuredMutationAction.SetTable(
                new TableSpec(
                    new TableName("BudgetTable"),
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                    new RangeRef("A1:B2"),
                    false,
                    true,
                    new TableStyleSpec.Named("TableStyleMedium2", false, false, true, false),
                    TextSourceSpec.standardInput(),
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
                new dev.erst.gridgrind.core.model.value.CellRowSpec.Typed(
                    List.of(new CellValueSpec.Formula(TextSourceSpec.standardInput()))))));
    assertTrue(
        requiresStandardInputFor(
            new WorkbookMutationAction.SetPrintLayout(
                printLayoutWithHeader(
                    new HeaderFooterTextSpec(
                        TextSourceSpec.standardInput(),
                        TextSourceSpec.inline(""),
                        TextSourceSpec.inline(""))))));
    assertTrue(
        requiresStandardInputFor(
            new WorkbookMutationAction.SetPrintLayout(
                printLayoutWithHeader(
                    new HeaderFooterTextSpec(
                        TextSourceSpec.inline("left"),
                        TextSourceSpec.standardInput(),
                        TextSourceSpec.inline("right"))))));
    assertTrue(
        requiresStandardInputFor(
            new WorkbookMutationAction.SetPrintLayout(
                printLayoutWithHeader(
                    new HeaderFooterTextSpec(
                        TextSourceSpec.inline("left"),
                        TextSourceSpec.inline("center"),
                        TextSourceSpec.standardInput())))));
    assertTrue(
        requiresStandardInputFor(
            new WorkbookMutationAction.SetPrintLayout(
                printLayoutWithFooter(
                    new HeaderFooterTextSpec(
                        TextSourceSpec.standardInput(),
                        TextSourceSpec.inline(""),
                        TextSourceSpec.inline(""))))));
  }

  @Test
  void requiresStandardInputTracksSemanticTableSelectors() {
    TableCellSelector.ByColumnName standardInputTarget =
        new TableCellSelector.ByColumnName(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName(new TableName("BudgetTable")),
                "Item",
                new CellValueSpec.Text(new TextSourceSpec.Stdin())),
            "Amount");
    TableCellSelector.ByColumnName inlineTarget =
        new TableCellSelector.ByColumnName(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName(new TableName("BudgetTable")),
                "Item",
                new CellValueSpec.Text(TextSourceSpec.inline("Hosting"))),
            "Amount");

    WorkbookPlan standardInputPlan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            dev.erst.gridgrind.core.plan.ExecutionPolicySpec.defaults(),
            dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec.empty(),
            List.of(
                new InspectionStep(
                    new StepId("inspect-amount"),
                    standardInputTarget,
                    new SheetIntrospectionQuery.GetCells())));
    WorkbookPlan inlinePlan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            dev.erst.gridgrind.core.plan.ExecutionPolicySpec.defaults(),
            dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec.empty(),
            List.of(
                new InspectionStep(
                    new StepId("inspect-amount"),
                    inlineTarget,
                    new SheetIntrospectionQuery.GetCells())));

    assertTrue(SourceBackedPlanResolver.requiresStandardInput(standardInputPlan));
    assertFalse(SourceBackedPlanResolver.requiresStandardInput(inlinePlan));
  }
}
