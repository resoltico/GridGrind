package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.mutate;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.mutations;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.request;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ChartPlotInput;
import dev.erst.gridgrind.contract.dto.ChartTitleInput;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.SignatureLineInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Source-backed resolver identity and simple inlining coverage. */
class SourceBackedPlanResolverIdentityCoverageTest extends SourceBackedPlanResolverTestSupport {
  @Test
  void sameOptionalReferenceDistinguishesPresenceMismatches() throws Exception {
    assertTrue(
        SourceBackedResolutionIdentitySupport.sameOptionalReference(
            Optional.empty(), Optional.empty()));
    assertFalse(
        SourceBackedResolutionIdentitySupport.sameOptionalReference(
            Optional.of("left"), Optional.empty()));
  }

  @Test
  void resolveRewritesTextFormulaAndBinarySourcesIntoInlineValues() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-source-backed-resolve-");
    Path titleFile = workingDirectory.resolve("title.txt");
    Path formulaFile = workingDirectory.resolve("formula.txt");
    Path payloadFile = workingDirectory.resolve("payload.bin");
    Files.writeString(titleFile, "Quarterly Budget", StandardCharsets.UTF_8);
    Files.writeString(formulaFile, "=SUM(B2:B3)", StandardCharsets.UTF_8);
    byte[] payloadBytes = "payload".getBytes(StandardCharsets.UTF_8);
    Files.write(payloadFile, payloadBytes);

    WorkbookPlan plan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Budget", "A1"),
                    new CellMutationAction.SetCell(
                        new CellInput.Text(new TextSourceInput.Utf8File("title.txt")))),
                mutate(
                    new CellSelector.ByAddress("Budget", "A2"),
                    new CellMutationAction.SetCell(
                        new CellInput.Formula(new TextSourceInput.Utf8File("formula.txt")))),
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new DrawingMutationAction.SetShape(
                        new ShapeInput(
                            "QueueShape",
                            ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                            twoCellAnchor(),
                            "roundRect",
                            new TextSourceInput.StandardInput()))),
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new DrawingMutationAction.SetEmbeddedObject(
                        new EmbeddedObjectInput(
                            "OpsPayload",
                            "Ops payload",
                            "ops-payload.txt",
                            "open",
                            new BinarySourceInput.File("payload.bin"),
                            inlinePictureData(),
                            twoCellAnchor()))),
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new DrawingMutationAction.SetSignatureLine(
                        new SignatureLineInput(
                            "OpsSignature",
                            twoCellAnchor(),
                            false,
                            java.util.Optional.of("Review before signing."),
                            java.util.Optional.of("Ada Lovelace"),
                            java.util.Optional.of("Finance"),
                            java.util.Optional.of("ada@example.com"),
                            java.util.Optional.empty(),
                            java.util.Optional.of("invalid"),
                            java.util.Optional.of(
                                new PictureDataInput(
                                    ExcelPictureFormat.PNG,
                                    new BinarySourceInput.File("payload.bin"))))))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            plan,
            new ExecutionInputBindings(workingDirectory, "Queue".getBytes(StandardCharsets.UTF_8)));

    CellMutationAction.SetCell titleAction =
        assertInstanceOf(
            CellMutationAction.SetCell.class, ((MutationStep) resolved.steps().get(0)).action());
    CellInput.Text titleValue = assertInstanceOf(CellInput.Text.class, titleAction.value());
    assertEquals("Quarterly Budget", ((TextSourceInput.Inline) titleValue.source()).text());

    CellMutationAction.SetCell formulaAction =
        assertInstanceOf(
            CellMutationAction.SetCell.class, ((MutationStep) resolved.steps().get(1)).action());
    CellInput.Formula formulaValue =
        assertInstanceOf(CellInput.Formula.class, formulaAction.value());
    assertEquals("SUM(B2:B3)", ((TextSourceInput.Inline) formulaValue.source()).text());

    DrawingMutationAction.SetShape shapeAction =
        assertInstanceOf(
            DrawingMutationAction.SetShape.class,
            ((MutationStep) resolved.steps().get(2)).action());
    assertEquals("Queue", ((TextSourceInput.Inline) shapeAction.shape().text()).text());

    DrawingMutationAction.SetEmbeddedObject embeddedObjectAction =
        assertInstanceOf(
            DrawingMutationAction.SetEmbeddedObject.class,
            ((MutationStep) resolved.steps().get(3)).action());
    assertEquals(
        Base64.getEncoder().encodeToString(payloadBytes),
        ((BinarySourceInput.InlineBase64) embeddedObjectAction.embeddedObject().payload())
            .base64Data());

    DrawingMutationAction.SetSignatureLine signatureLineAction =
        assertInstanceOf(
            DrawingMutationAction.SetSignatureLine.class,
            ((MutationStep) resolved.steps().get(4)).action());
    assertEquals(
        Base64.getEncoder().encodeToString(payloadBytes),
        ((BinarySourceInput.InlineBase64)
                signatureLineAction.signatureLine().plainSignature().orElseThrow().source())
            .base64Data());
  }

  @Test
  void requiresStandardInputTracksSourceBackedStandardInputFields() {
    WorkbookPlan standardInputPlan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Budget", "A1"),
                    new CellMutationAction.SetCell(
                        new CellInput.Text(new TextSourceInput.StandardInput()))),
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new DrawingMutationAction.SetSignatureLine(
                        new SignatureLineInput(
                            "OpsSignature",
                            twoCellAnchor(),
                            false,
                            java.util.Optional.of("Review before signing."),
                            java.util.Optional.of("Ada Lovelace"),
                            java.util.Optional.of("Finance"),
                            java.util.Optional.of("ada@example.com"),
                            java.util.Optional.empty(),
                            java.util.Optional.of("invalid"),
                            java.util.Optional.of(
                                new PictureDataInput(
                                    ExcelPictureFormat.PNG,
                                    new BinarySourceInput.StandardInput())))))));
    WorkbookPlan fileBackedPlan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Budget", "A1"),
                    new CellMutationAction.SetCell(
                        new CellInput.Text(new TextSourceInput.Utf8File("title.txt"))))));

    assertTrue(SourceBackedPlanResolver.requiresStandardInput(standardInputPlan));
    assertFalse(SourceBackedPlanResolver.requiresStandardInput(fileBackedPlan));
  }

  @Test
  void resolveRewritesCustomXmlImportSourcesIntoInlineValues() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-custom-xml-resolve-");
    Files.writeString(
        workingDirectory.resolve("mapping.xml"),
        "<CORSO><NOME>Grid</NOME><DOCENTE>Grind</DOCENTE></CORSO>",
        StandardCharsets.UTF_8);

    WorkbookPlan plan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new WorkbookSelector.Current(),
                    new StructuredMutationAction.ImportCustomXmlMapping(
                        new CustomXmlImportInput(
                            new CustomXmlMappingLocator(1L, "CORSO_mapping"),
                            new TextSourceInput.Utf8File("mapping.xml"))))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(plan, new ExecutionInputBindings(workingDirectory));

    StructuredMutationAction.ImportCustomXmlMapping action =
        assertInstanceOf(
            StructuredMutationAction.ImportCustomXmlMapping.class,
            assertInstanceOf(MutationStep.class, resolved.steps().getFirst()).action());
    assertEquals(new CustomXmlMappingLocator(1L, "CORSO_mapping"), action.mapping().locator());
    assertEquals(
        "<CORSO><NOME>Grid</NOME><DOCENTE>Grind</DOCENTE></CORSO>",
        assertInstanceOf(TextSourceInput.Inline.class, action.mapping().xml()).text());
  }

  @Test
  void resolveRewritesSourceBackedTableKeySelectorsIntoInlineValues() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-source-backed-selector-");
    Files.writeString(workingDirectory.resolve("item.txt"), "Hosting", StandardCharsets.UTF_8);

    TableCellSelector.ByColumnName selector =
        new TableCellSelector.ByColumnName(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName("BudgetTable"),
                "Item",
                new CellInput.Text(new TextSourceInput.Utf8File("item.txt"))),
            "Amount");
    WorkbookPlan plan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
            dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
            List.of(
                new InspectionStep("inspect-amount", selector, new InspectionQuery.GetCells()),
                new AssertionStep(
                    "assert-amount",
                    selector,
                    new Assertion.CellValue(new ExpectedCellValue.NumericValue(125.0)))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(plan, new ExecutionInputBindings(workingDirectory));

    InspectionStep resolvedInspection =
        assertInstanceOf(InspectionStep.class, resolved.steps().get(0));
    TableCellSelector.ByColumnName resolvedSelector =
        assertInstanceOf(TableCellSelector.ByColumnName.class, resolvedInspection.target());
    TableRowSelector.ByKeyCell resolvedRow =
        assertInstanceOf(TableRowSelector.ByKeyCell.class, resolvedSelector.row());
    CellInput.Text resolvedValue =
        assertInstanceOf(CellInput.Text.class, resolvedRow.expectedValue());
    assertEquals("Hosting", ((TextSourceInput.Inline) resolvedValue.source()).text());

    AssertionStep resolvedAssertion =
        assertInstanceOf(AssertionStep.class, resolved.steps().get(1));
    TableCellSelector.ByColumnName assertedSelector =
        assertInstanceOf(TableCellSelector.ByColumnName.class, resolvedAssertion.target());
    TableRowSelector.ByKeyCell assertedRow =
        assertInstanceOf(TableRowSelector.ByKeyCell.class, assertedSelector.row());
    CellInput.Text assertedValue =
        assertInstanceOf(CellInput.Text.class, assertedRow.expectedValue());
    assertEquals("Hosting", ((TextSourceInput.Inline) assertedValue.source()).text());
  }

  @Test
  void resolveRewritesChangedRangeRowsAndTextTitledCharts() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-source-backed-rewrites-");
    Files.writeString(
        workingDirectory.resolve("range-cell.txt"), "Quarterly Budget", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("line-title.txt"), "Quarterly Trend", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("pie-title.txt"), "Category Mix", StandardCharsets.UTF_8);

    MutationStep rangeStep =
        new MutationStep(
            "set-range-file-backed",
            new RangeSelector.ByRange("Budget", "A1:B1"),
            new CellMutationAction.SetRange(
                List.of(
                    List.of(
                        new CellInput.Text(TextSourceInput.utf8File("range-cell.txt")),
                        new CellInput.Numeric(42.0d)))));
    MutationStep lineChartStep =
        new MutationStep(
            "set-line-chart-file-backed",
            new SheetSelector.ByName("Budget"),
            new DrawingMutationAction.SetChart(
                chartInput(
                    "Line",
                    twoCellAnchor(),
                    new ChartTitleInput.Text(TextSourceInput.utf8File("line-title.txt")),
                    null,
                    ExcelChartDisplayBlanksAs.ZERO,
                    true,
                    new ChartPlotInput.Line(
                        false,
                        dev.erst.gridgrind.excel.foundation.ExcelChartGrouping.STANDARD,
                        List.of(chartSeries())))));
    MutationStep pieChartStep =
        new MutationStep(
            "set-pie-chart-file-backed",
            new SheetSelector.ByName("Budget"),
            new DrawingMutationAction.SetChart(
                chartInput(
                    "Pie",
                    twoCellAnchor(),
                    new ChartTitleInput.Text(TextSourceInput.utf8File("pie-title.txt")),
                    null,
                    ExcelChartDisplayBlanksAs.GAP,
                    true,
                    new ChartPlotInput.Pie(false, 15, List.of(chartSeries())))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
                dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
                List.of(rangeStep, lineChartStep, pieChartStep)),
            new ExecutionInputBindings(workingDirectory));

    MutationStep resolvedRangeStep = assertInstanceOf(MutationStep.class, resolved.steps().get(0));
    assertNotSame(rangeStep, resolvedRangeStep);
    CellMutationAction.SetRange resolvedRange =
        assertInstanceOf(CellMutationAction.SetRange.class, resolvedRangeStep.action());
    CellInput.Text resolvedRangeCell =
        assertInstanceOf(CellInput.Text.class, resolvedRange.rows().getFirst().getFirst());
    assertEquals("Quarterly Budget", ((TextSourceInput.Inline) resolvedRangeCell.source()).text());

    MutationStep resolvedLineStep = assertInstanceOf(MutationStep.class, resolved.steps().get(1));
    assertNotSame(lineChartStep, resolvedLineStep);
    ChartInput resolvedLineChart =
        assertInstanceOf(DrawingMutationAction.SetChart.class, resolvedLineStep.action()).chart();
    assertInstanceOf(ChartPlotInput.Line.class, resolvedLineChart.plots().getFirst());
    ChartTitleInput.Text lineTitle =
        assertInstanceOf(ChartTitleInput.Text.class, resolvedLineChart.title());
    assertEquals("Quarterly Trend", ((TextSourceInput.Inline) lineTitle.source()).text());

    MutationStep resolvedPieStep = assertInstanceOf(MutationStep.class, resolved.steps().get(2));
    assertNotSame(pieChartStep, resolvedPieStep);
    ChartInput resolvedPieChart =
        assertInstanceOf(DrawingMutationAction.SetChart.class, resolvedPieStep.action()).chart();
    assertInstanceOf(ChartPlotInput.Pie.class, resolvedPieChart.plots().getFirst());
    ChartTitleInput.Text pieTitle =
        assertInstanceOf(ChartTitleInput.Text.class, resolvedPieChart.title());
    assertEquals("Category Mix", ((TextSourceInput.Inline) pieTitle.source()).text());
  }
}
