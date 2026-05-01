package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
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
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Source-backed resolver mutation-family inlining coverage. */
class SourceBackedPlanResolverMutationInliningCoverageTest
    extends SourceBackedPlanResolverTestSupport {
  @Test
  void resolveCoversRemainingMutationFamiliesAndPreservesNonMutationSteps() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-source-backed-phase7-");
    Files.writeString(
        workingDirectory.resolve("comment.txt"), "Ada Lovelace", StandardCharsets.UTF_8);
    Files.writeString(workingDirectory.resolve("run1.txt"), "Ada", StandardCharsets.UTF_8);
    Files.writeString(workingDirectory.resolve("run2.txt"), " Lovelace", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("picture-description.txt"),
        "Queue preview",
        StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("chart-title.txt"), "Quarterly Trend", StandardCharsets.UTF_8);
    Files.writeString(workingDirectory.resolve("shape.txt"), "Queue shape", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("prompt-title.txt"), "Choose", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("prompt-text.txt"), "Pick one value", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("error-title.txt"), "Invalid", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("error-text.txt"), "Try again", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("table-comment.txt"),
        "Quarterly budget table",
        StandardCharsets.UTF_8);
    Files.writeString(workingDirectory.resolve("header-left.txt"), "Left", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("header-center.txt"), "Center", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("header-right.txt"), "Right", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("formula.txt"), "=SUM(B2:B3)", StandardCharsets.UTF_8);
    Files.write(
        workingDirectory.resolve("payload.bin"), "payload".getBytes(StandardCharsets.UTF_8));
    Files.write(workingDirectory.resolve("preview.bin"), pngBytes());

    AssertionStep assertion =
        new AssertionStep(
            "assert-budget",
            new CellSelector.ByAddress("Budget", "A1"),
            new Assertion.CellValue(new ExpectedCellValue.Text("Quarterly Budget")));
    InspectionStep inspection =
        new InspectionStep(
            "inspect-budget",
            new CellSelector.ByAddress("Budget", "A1"),
            new InspectionQuery.GetCells());

    WorkbookPlan plan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
            dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    "step-01-set-comment",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new CellMutationAction.SetComment(
                        new CommentInput(
                            TextSourceInput.utf8File("comment.txt"),
                            "Ada",
                            false,
                            java.util.Optional.of(
                                List.of(
                                    new RichTextRunInput(
                                        TextSourceInput.utf8File("run1.txt"), null),
                                    new RichTextRunInput(
                                        TextSourceInput.utf8File("run2.txt"), null))),
                            java.util.Optional.empty()))),
                new MutationStep(
                    "step-02-set-picture",
                    new SheetSelector.ByName("Budget"),
                    new DrawingMutationAction.SetPicture(
                        new PictureInput(
                            "Logo",
                            new PictureDataInput(
                                ExcelPictureFormat.PNG, BinarySourceInput.file("preview.bin")),
                            twoCellAnchor(),
                            TextSourceInput.utf8File("picture-description.txt")))),
                new MutationStep(
                    "step-03-set-bar-chart",
                    new SheetSelector.ByName("Budget"),
                    new DrawingMutationAction.SetChart(
                        chartInput(
                            "Bar",
                            twoCellAnchor(),
                            new ChartTitleInput.Text(TextSourceInput.utf8File("chart-title.txt")),
                            null,
                            ExcelChartDisplayBlanksAs.GAP,
                            true,
                            new ChartPlotInput.Bar(
                                false,
                                dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection.COLUMN,
                                dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping.CLUSTERED,
                                null,
                                null,
                                List.of(chartSeries()))))),
                new MutationStep(
                    "step-04-set-line-chart",
                    new SheetSelector.ByName("Budget"),
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
                                List.of(chartSeries()))))),
                new MutationStep(
                    "step-05-set-pie-chart",
                    new SheetSelector.ByName("Budget"),
                    new DrawingMutationAction.SetChart(
                        chartInput(
                            "Pie",
                            twoCellAnchor(),
                            new ChartTitleInput.None(),
                            null,
                            ExcelChartDisplayBlanksAs.GAP,
                            true,
                            new ChartPlotInput.Pie(false, 0, List.of(chartSeries()))))),
                new MutationStep(
                    "step-06-set-shape",
                    new SheetSelector.ByName("Budget"),
                    new DrawingMutationAction.SetShape(
                        new ShapeInput(
                            "Shape",
                            ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                            twoCellAnchor(),
                            "roundRect",
                            TextSourceInput.utf8File("shape.txt")))),
                new MutationStep(
                    "step-07-set-embedded",
                    new SheetSelector.ByName("Budget"),
                    new DrawingMutationAction.SetEmbeddedObject(
                        new EmbeddedObjectInput(
                            "Payload",
                            "Payload",
                            "payload.bin",
                            "open",
                            BinarySourceInput.file("payload.bin"),
                            new PictureDataInput(
                                ExcelPictureFormat.PNG, BinarySourceInput.file("preview.bin")),
                            twoCellAnchor()))),
                new MutationStep(
                    "step-08-set-validation",
                    new RangeSelector.ByRange("Budget", "A1"),
                    new StructuredMutationAction.SetDataValidation(
                        new DataValidationInput(
                            new DataValidationRuleInput.WholeNumber(
                                ExcelComparisonOperator.BETWEEN, "1", "10"),
                            false,
                            false,
                            new DataValidationPromptInput(
                                TextSourceInput.utf8File("prompt-title.txt"),
                                TextSourceInput.utf8File("prompt-text.txt"),
                                false),
                            new DataValidationErrorAlertInput(
                                ExcelDataValidationErrorStyle.STOP,
                                TextSourceInput.utf8File("error-title.txt"),
                                TextSourceInput.utf8File("error-text.txt"),
                                false)))),
                new MutationStep(
                    "step-09-set-table",
                    new dev.erst.gridgrind.contract.selector.TableSelector.ByNameOnSheet(
                        "BudgetTable", "Budget"),
                    new StructuredMutationAction.SetTable(
                        new TableInput(
                            "BudgetTable",
                            "Budget",
                            "A1:B2",
                            false,
                            true,
                            new TableStyleInput.Named(
                                "TableStyleMedium2", false, false, true, false),
                            TextSourceInput.utf8File("table-comment.txt"),
                            false,
                            false,
                            false,
                            "",
                            "",
                            "",
                            List.of()))),
                new MutationStep(
                    "step-10-append-row",
                    new SheetSelector.ByName("Budget"),
                    new CellMutationAction.AppendRow(
                        List.of(
                            new CellInput.Text(TextSourceInput.utf8File("chart-title.txt")),
                            new CellInput.RichText(
                                List.of(
                                    new RichTextRunInput(
                                        TextSourceInput.utf8File("run1.txt"), null),
                                    new RichTextRunInput(
                                        TextSourceInput.utf8File("run2.txt"), null))),
                            new CellInput.Formula(TextSourceInput.utf8File("formula.txt"))))),
                new MutationStep(
                    "step-11-print-layout",
                    new SheetSelector.ByName("Budget"),
                    new WorkbookMutationAction.SetPrintLayout(
                        printLayoutWithHeader(
                            new HeaderFooterTextInput(
                                TextSourceInput.utf8File("header-left.txt"),
                                TextSourceInput.utf8File("header-center.txt"),
                                TextSourceInput.utf8File("header-right.txt"))))),
                assertion,
                inspection));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            plan,
            new ExecutionInputBindings(
                workingDirectory, "stdin-text".getBytes(StandardCharsets.UTF_8)));

    CellMutationAction.SetComment commentAction =
        assertInstanceOf(
            CellMutationAction.SetComment.class, ((MutationStep) resolved.steps().get(0)).action());
    assertEquals("Ada Lovelace", ((TextSourceInput.Inline) commentAction.comment().text()).text());
    assertEquals(
        "Ada",
        ((TextSourceInput.Inline) commentAction.comment().runs().orElseThrow().getFirst().source())
            .text());

    DrawingMutationAction.SetPicture pictureAction =
        assertInstanceOf(
            DrawingMutationAction.SetPicture.class,
            ((MutationStep) resolved.steps().get(1)).action());
    assertEquals(
        Base64.getEncoder().encodeToString(pngBytes()),
        ((BinarySourceInput.InlineBase64) pictureAction.picture().image().source()).base64Data());
    assertEquals(
        "Queue preview", ((TextSourceInput.Inline) pictureAction.picture().description()).text());

    DrawingMutationAction.SetChart barChart =
        assertInstanceOf(
            DrawingMutationAction.SetChart.class,
            ((MutationStep) resolved.steps().get(2)).action());
    ChartInput resolvedBarChart = barChart.chart();
    assertInstanceOf(ChartPlotInput.Bar.class, resolvedBarChart.plots().getFirst());
    assertEquals(
        "Quarterly Trend",
        ((TextSourceInput.Inline)
                assertInstanceOf(ChartTitleInput.Text.class, resolvedBarChart.title()).source())
            .text());

    DrawingMutationAction.SetChart lineChart =
        assertInstanceOf(
            DrawingMutationAction.SetChart.class,
            ((MutationStep) resolved.steps().get(3)).action());
    assertInstanceOf(ChartPlotInput.Line.class, lineChart.chart().plots().getFirst());
    assertInstanceOf(ChartTitleInput.Formula.class, lineChart.chart().title());

    DrawingMutationAction.SetChart pieChart =
        assertInstanceOf(
            DrawingMutationAction.SetChart.class,
            ((MutationStep) resolved.steps().get(4)).action());
    assertInstanceOf(ChartPlotInput.Pie.class, pieChart.chart().plots().getFirst());
    assertInstanceOf(ChartTitleInput.None.class, pieChart.chart().title());

    DrawingMutationAction.SetShape shapeAction =
        assertInstanceOf(
            DrawingMutationAction.SetShape.class,
            ((MutationStep) resolved.steps().get(5)).action());
    assertEquals("Queue shape", ((TextSourceInput.Inline) shapeAction.shape().text()).text());

    DrawingMutationAction.SetEmbeddedObject embeddedAction =
        assertInstanceOf(
            DrawingMutationAction.SetEmbeddedObject.class,
            ((MutationStep) resolved.steps().get(6)).action());
    assertEquals(
        Base64.getEncoder().encodeToString("payload".getBytes(StandardCharsets.UTF_8)),
        ((BinarySourceInput.InlineBase64) embeddedAction.embeddedObject().payload()).base64Data());

    StructuredMutationAction.SetDataValidation validationAction =
        assertInstanceOf(
            StructuredMutationAction.SetDataValidation.class,
            ((MutationStep) resolved.steps().get(7)).action());
    assertEquals(
        "Choose", ((TextSourceInput.Inline) validationAction.validation().prompt().title()).text());
    assertEquals(
        "Try again",
        ((TextSourceInput.Inline) validationAction.validation().errorAlert().text()).text());

    StructuredMutationAction.SetTable tableAction =
        assertInstanceOf(
            StructuredMutationAction.SetTable.class,
            ((MutationStep) resolved.steps().get(8)).action());
    assertEquals(
        "Quarterly budget table", ((TextSourceInput.Inline) tableAction.table().comment()).text());

    CellMutationAction.AppendRow appendRow =
        assertInstanceOf(
            CellMutationAction.AppendRow.class, ((MutationStep) resolved.steps().get(9)).action());
    assertEquals(
        "SUM(B2:B3)",
        ((TextSourceInput.Inline)
                assertInstanceOf(CellInput.Formula.class, appendRow.values().get(2)).source())
            .text());

    WorkbookMutationAction.SetPrintLayout printLayout =
        assertInstanceOf(
            WorkbookMutationAction.SetPrintLayout.class,
            ((MutationStep) resolved.steps().get(10)).action());
    assertEquals(
        "Left", ((TextSourceInput.Inline) printLayout.printLayout().header().left()).text());
    assertEquals(
        "Center", ((TextSourceInput.Inline) printLayout.printLayout().header().center()).text());
    assertEquals(
        "Right", ((TextSourceInput.Inline) printLayout.printLayout().header().right()).text());

    assertSame(assertion, resolved.steps().get(11));
    assertSame(inspection, resolved.steps().get(12));
  }
}
