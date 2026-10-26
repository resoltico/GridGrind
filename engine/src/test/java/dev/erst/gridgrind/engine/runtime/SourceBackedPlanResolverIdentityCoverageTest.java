package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.customxml.CustomXmlImportSpec;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.model.drawing.EmbeddedObjectSpec;
import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.drawing.ShapeSpec;
import dev.erst.gridgrind.core.model.drawing.SignatureLineSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TableName;
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
    Files.writeString(formulaFile, "SUM(B2:B3)", StandardCharsets.UTF_8);
    byte[] payloadBytes = "payload".getBytes(StandardCharsets.UTF_8);
    Files.write(payloadFile, payloadBytes);

    WorkbookPlan plan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                    new CellMutationAction.SetCell(
                        new CellValueSpec.Text(new TextSourceSpec.Utf8File("title.txt")))),
                mutate(
                    new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A2")),
                    new CellMutationAction.SetCell(
                        new CellValueSpec.Formula(new TextSourceSpec.Utf8File("formula.txt")))),
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new DrawingMutationAction.SetShape(
                        simpleShape(
                            "QueueShape",
                            twoCellAnchor(),
                            "roundRect",
                            new TextSourceSpec.Stdin()))),
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new DrawingMutationAction.SetEmbeddedObject(
                        new EmbeddedObjectSpec(
                            "OpsPayload",
                            "Ops payload",
                            "ops-payload.txt",
                            "open",
                            new BinarySourceSpec.File("payload.bin"),
                            inlinePictureData(),
                            twoCellAnchor()))),
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new DrawingMutationAction.SetSignatureLine(
                        new SignatureLineSpec(
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
                                new PictureDataSpec(
                                    ExcelPictureFormat.PNG,
                                    new BinarySourceSpec.File("payload.bin"))))))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            plan,
            ExecutionInputBindingsFixtureSupport.bindings(
                workingDirectory, plan, "Queue".getBytes(StandardCharsets.UTF_8)));

    CellMutationAction.SetCell titleAction =
        assertInstanceOf(
            CellMutationAction.SetCell.class, ((MutationStep) resolved.steps().get(0)).action());
    CellValueSpec.Text titleValue = assertInstanceOf(CellValueSpec.Text.class, titleAction.value());
    assertEquals(
        "Quarterly Budget",
        assertInstanceOf(TextSourceSpec.Inline.class, titleValue.value()).text());

    CellMutationAction.SetCell formulaAction =
        assertInstanceOf(
            CellMutationAction.SetCell.class, ((MutationStep) resolved.steps().get(1)).action());
    CellValueSpec.Formula formulaValue =
        assertInstanceOf(CellValueSpec.Formula.class, formulaAction.value());
    assertEquals(
        "SUM(B2:B3)", assertInstanceOf(TextSourceSpec.Inline.class, formulaValue.value()).text());

    DrawingMutationAction.SetShape shapeAction =
        assertInstanceOf(
            DrawingMutationAction.SetShape.class,
            ((MutationStep) resolved.steps().get(2)).action());
    assertEquals(
        "Queue",
        ((TextSourceSpec.Inline)
                assertInstanceOf(ShapeSpec.SimpleShape.class, shapeAction.shape())
                    .text()
                    .orElseThrow())
            .text());

    DrawingMutationAction.SetEmbeddedObject embeddedObjectAction =
        assertInstanceOf(
            DrawingMutationAction.SetEmbeddedObject.class,
            ((MutationStep) resolved.steps().get(3)).action());
    assertEquals(
        Base64.getEncoder().encodeToString(payloadBytes),
        ((BinarySourceSpec.InlineBase64) embeddedObjectAction.embeddedObject().payload())
            .base64Data());

    DrawingMutationAction.SetSignatureLine signatureLineAction =
        assertInstanceOf(
            DrawingMutationAction.SetSignatureLine.class,
            ((MutationStep) resolved.steps().get(4)).action());
    assertEquals(
        Base64.getEncoder().encodeToString(payloadBytes),
        ((BinarySourceSpec.InlineBase64)
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
                    new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                    new CellMutationAction.SetCell(
                        new CellValueSpec.Text(new TextSourceSpec.Stdin()))),
                mutate(
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new DrawingMutationAction.SetSignatureLine(
                        new SignatureLineSpec(
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
                                new PictureDataSpec(
                                    ExcelPictureFormat.PNG, new BinarySourceSpec.Stdin())))))));
    WorkbookPlan fileBackedPlan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                    new CellMutationAction.SetCell(
                        new CellValueSpec.Text(new TextSourceSpec.Utf8File("title.txt"))))));

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
                        new CustomXmlImportSpec(
                            new CustomXmlMappingLocator(1L, "CORSO_mapping"),
                            new TextSourceSpec.Utf8File("mapping.xml"))))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            plan, ExecutionInputBindingsFixtureSupport.bindings(workingDirectory, plan));

    StructuredMutationAction.ImportCustomXmlMapping action =
        assertInstanceOf(
            StructuredMutationAction.ImportCustomXmlMapping.class,
            assertInstanceOf(MutationStep.class, resolved.steps().getFirst()).action());
    assertEquals(new CustomXmlMappingLocator(1L, "CORSO_mapping"), action.mapping().locator());
    assertEquals(
        "<CORSO><NOME>Grid</NOME><DOCENTE>Grind</DOCENTE></CORSO>",
        assertInstanceOf(TextSourceSpec.Inline.class, action.mapping().xml()).text());
  }

  @Test
  void resolveRewritesSourceBackedTableKeySelectorsIntoInlineValues() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-source-backed-selector-");
    Files.writeString(workingDirectory.resolve("item.txt"), "Hosting", StandardCharsets.UTF_8);

    TableCellSelector.ByColumnName selector =
        new TableCellSelector.ByColumnName(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName(new TableName("BudgetTable")),
                "Item",
                new CellValueSpec.Text(new TextSourceSpec.Utf8File("item.txt"))),
            "Amount");
    WorkbookPlan plan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            dev.erst.gridgrind.core.plan.ExecutionPolicySpec.defaults(),
            dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec.empty(),
            List.of(
                new InspectionStep(
                    new StepId("inspect-amount"), selector, new SheetIntrospectionQuery.GetCells()),
                new AssertionStep(
                    new StepId("assert-amount"),
                    selector,
                    new CellAssertion.Value(
                        new dev.erst.gridgrind.core.model.value.CellValue.NumberValue(125.0)))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            plan,
            ExecutionInputBindingsFixtureSupport.bindings(
                workingDirectory, List.of(workingDirectory.resolve("item.txt"))));

    InspectionStep resolvedInspection =
        assertInstanceOf(InspectionStep.class, resolved.steps().get(0));
    TableCellSelector.ByColumnName resolvedSelector =
        assertInstanceOf(TableCellSelector.ByColumnName.class, resolvedInspection.target());
    TableRowSelector.ByKeyCell resolvedRow =
        assertInstanceOf(TableRowSelector.ByKeyCell.class, resolvedSelector.row());
    CellValueSpec.Text resolvedValue =
        assertInstanceOf(CellValueSpec.Text.class, resolvedRow.expectedValue());
    assertEquals(
        "Hosting", assertInstanceOf(TextSourceSpec.Inline.class, resolvedValue.value()).text());

    AssertionStep resolvedAssertion =
        assertInstanceOf(AssertionStep.class, resolved.steps().get(1));
    TableCellSelector.ByColumnName assertedSelector =
        assertInstanceOf(TableCellSelector.ByColumnName.class, resolvedAssertion.target());
    TableRowSelector.ByKeyCell assertedRow =
        assertInstanceOf(TableRowSelector.ByKeyCell.class, assertedSelector.row());
    CellValueSpec.Text assertedValue =
        assertInstanceOf(CellValueSpec.Text.class, assertedRow.expectedValue());
    assertEquals(
        "Hosting", assertInstanceOf(TextSourceSpec.Inline.class, assertedValue.value()).text());
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
            new StepId("set-range-file-backed"),
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B1")),
            new CellMutationAction.SetRange(
                new dev.erst.gridgrind.core.model.value.CellGridSpec.Typed(
                    List.of(
                        List.of(
                            new CellValueSpec.Text(TextSourceSpec.utf8File("range-cell.txt")),
                            new CellValueSpec.NumberValue(42.0d))))));
    MutationStep lineChartStep =
        new MutationStep(
            new StepId("set-line-chart-file-backed"),
            new SheetSelector.ByName(new SheetName("Budget")),
            new DrawingMutationAction.SetChart(
                chartInput(
                    "Line",
                    twoCellAnchor(),
                    new ChartTitleSpec.Text(TextSourceSpec.utf8File("line-title.txt")),
                    null,
                    ExcelChartDisplayBlanksAs.ZERO,
                    true,
                    new ChartPlotSpec.Line(
                        false,
                        dev.erst.gridgrind.core.vocabulary.ExcelChartGrouping.STANDARD,
                        List.of(chartSeries())))));
    MutationStep pieChartStep =
        new MutationStep(
            new StepId("set-pie-chart-file-backed"),
            new SheetSelector.ByName(new SheetName("Budget")),
            new DrawingMutationAction.SetChart(
                chartInput(
                    "Pie",
                    twoCellAnchor(),
                    new ChartTitleSpec.Text(TextSourceSpec.utf8File("pie-title.txt")),
                    null,
                    ExcelChartDisplayBlanksAs.GAP,
                    true,
                    new ChartPlotSpec.Pie(false, Optional.of(15), List.of(chartSeries())))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                dev.erst.gridgrind.core.plan.ExecutionPolicySpec.defaults(),
                dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec.empty(),
                List.of(rangeStep, lineChartStep, pieChartStep)),
            ExecutionInputBindingsFixtureSupport.bindings(
                workingDirectory,
                List.of(
                    workingDirectory.resolve("range-cell.txt"),
                    workingDirectory.resolve("line-title.txt"),
                    workingDirectory.resolve("pie-title.txt"))));

    MutationStep resolvedRangeStep = assertInstanceOf(MutationStep.class, resolved.steps().get(0));
    assertNotSame(rangeStep, resolvedRangeStep);
    CellMutationAction.SetRange resolvedRange =
        assertInstanceOf(CellMutationAction.SetRange.class, resolvedRangeStep.action());
    CellValueSpec.Text resolvedRangeCell =
        assertInstanceOf(
            CellValueSpec.Text.class, resolvedRange.rows().toCellInputRows().getFirst().getFirst());
    assertEquals(
        "Quarterly Budget",
        assertInstanceOf(TextSourceSpec.Inline.class, resolvedRangeCell.value()).text());

    MutationStep resolvedLineStep = assertInstanceOf(MutationStep.class, resolved.steps().get(1));
    assertNotSame(lineChartStep, resolvedLineStep);
    ChartSpec resolvedLineChart =
        assertInstanceOf(DrawingMutationAction.SetChart.class, resolvedLineStep.action()).chart();
    assertInstanceOf(ChartPlotSpec.Line.class, resolvedLineChart.plots().getFirst());
    ChartTitleSpec.Text lineTitle =
        assertInstanceOf(ChartTitleSpec.Text.class, resolvedLineChart.title());
    assertEquals("Quarterly Trend", ((TextSourceSpec.Inline) lineTitle.source()).text());

    MutationStep resolvedPieStep = assertInstanceOf(MutationStep.class, resolved.steps().get(2));
    assertNotSame(pieChartStep, resolvedPieStep);
    ChartSpec resolvedPieChart =
        assertInstanceOf(DrawingMutationAction.SetChart.class, resolvedPieStep.action()).chart();
    assertInstanceOf(ChartPlotSpec.Pie.class, resolvedPieChart.plots().getFirst());
    ChartTitleSpec.Text pieTitle =
        assertInstanceOf(ChartTitleSpec.Text.class, resolvedPieChart.title());
    assertEquals("Category Mix", ((TextSourceSpec.Inline) pieTitle.source()).text());
  }
}
