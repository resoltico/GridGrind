package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ProtocolStyleTestAccess.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlDataBindingReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlExportReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlLinkedCellReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingReport;
import dev.erst.gridgrind.core.model.drawing.CommentAnchorReport;
import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerReport;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectPayloadReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.formula.ArrayFormulaReport;
import dev.erst.gridgrind.core.model.layout.SheetDefaultsReport;
import dev.erst.gridgrind.core.model.layout.SheetDisplayReport;
import dev.erst.gridgrind.core.model.layout.SheetOutlineSummaryReport;
import dev.erst.gridgrind.core.model.layout.SheetPresentationReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.print.*;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.PrintLayoutReport;
import dev.erst.gridgrind.core.model.print.PrintMarginsReport;
import dev.erst.gridgrind.core.model.print.PrintSetupReport;
import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionReport;
import dev.erst.gridgrind.core.model.security.OoxmlEncryptionReport;
import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.core.model.security.OoxmlSignatureReport;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellGradientFillReport;
import dev.erst.gridgrind.core.model.style.CellGradientStopReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.DifferentialStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateReport;
import dev.erst.gridgrind.core.model.table.TableColumnReport;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorReport;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.plan.SheetInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookAnalysisResult;
import dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookInspectionResult;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingShapeKind;
import dev.erst.gridgrind.core.vocabulary.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelIgnoredErrorType;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlHashAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.DrawingGeometryFixtures;
import java.math.BigDecimal;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for converting advanced engine read results into protocol response shapes. */
class InspectionResultConverterTest extends DefaultGridGrindRequestExecutorTestSupport {
  @Test
  void convertsPackageSecurityReadResultsIntoProtocolShapes() {
    WorkbookInspectionResult.PackageSecurityResult packageSecurity =
        assertInstanceOf(
            WorkbookInspectionResult.PackageSecurityResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookCoreResult.PackageSecurityResult(
                    new StepId("security"),
                    new OoxmlPackageSecurityReport(
                        new OoxmlEncryptionReport.Encrypted(
                            ExcelOoxmlEncryptionMode.AGILE,
                            ExcelOoxmlCipherAlgorithm.AES_256,
                            ExcelOoxmlHashAlgorithm.SHA_512,
                            ExcelOoxmlChainingMode.CBC,
                            256,
                            16,
                            100_000),
                        List.of(
                            new OoxmlSignatureReport(
                                "/_xmlsignatures/sig1.xml",
                                Optional.of("CN=GridGrind Signing Test"),
                                Optional.of("CN=GridGrind Signing Test"),
                                Optional.of("01AB"),
                                ExcelOoxmlSignatureState.VALID))))));

    assertInstanceOf(
        OoxmlEncryptionReport.Encrypted.class, packageSecurity.security().encryption());
    assertEquals(
        ExcelOoxmlSignatureState.VALID, packageSecurity.security().signatures().getFirst().state());
    assertEquals(
        "CN=GridGrind Signing Test",
        packageSecurity.security().signatures().getFirst().signer().orElseThrow().subject());
  }

  @Test
  void convertsWorkbookCustomXmlReadResultsIntoProtocolShapes() {
    WorkbookInspectionResult.CustomXmlMappingsResult mappings =
        assertInstanceOf(
            WorkbookInspectionResult.CustomXmlMappingsResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookCoreResult.CustomXmlMappingsResult(
                    new StepId("custom-xml-mappings"),
                    List.of(
                        customXmlMappingSnapshot(
                            1L,
                            "CORSO_mapping",
                            "CORSO",
                            "Schema1",
                            settings(false, true, false, true, true),
                            schema(
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.of("<xsd:schema/>")),
                            Optional.of(
                                new CustomXmlDataBindingReport(
                                    "binding", true, 5L, "binding.xml", 2L)),
                            List.of(
                                new CustomXmlLinkedCellReport(
                                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Foglio1"),
                                    new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
                                    "/CORSO/NOME",
                                    "string")),
                            List.of())))));
    WorkbookInspectionResult.CustomXmlExportResult exported =
        assertInstanceOf(
            WorkbookInspectionResult.CustomXmlExportResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookCoreResult.CustomXmlExportResult(
                    new StepId("custom-xml-export"),
                    new CustomXmlExportReport(
                        customXmlMappingSnapshot(
                            1L,
                            "CORSO_mapping",
                            "CORSO",
                            "Schema1",
                            settings(false, true, false, true, true),
                            schema(
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.of("<xsd:schema/>")),
                            Optional.empty(),
                            List.of(
                                new CustomXmlLinkedCellReport(
                                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Foglio1"),
                                    new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
                                    "/CORSO/NOME",
                                    "string")),
                            List.of()),
                        "UTF-8",
                        true,
                        "<CORSO><NOME>Grid</NOME></CORSO>"))));

    assertEquals("custom-xml-mappings", mappings.stepId().value());
    assertEquals("CORSO_mapping", mappings.mappings().getFirst().name());
    assertEquals(2L, mappings.mappings().getFirst().dataBinding().loadMode());
    assertEquals("A1", mappings.mappings().getFirst().linkedCells().getFirst().address().value());
    assertEquals("custom-xml-export", exported.stepId().value());
    assertEquals("UTF-8", exported.export().encoding());
    assertTrue(exported.export().xml().contains("<NOME>Grid</NOME>"));
  }

  @Test
  void convertsPlainCommentFactsDirectly() {
    assertTrue(InspectionResultCellReportSupport.toCommentReport((ResolvedComment) null).isEmpty());
    CommentReport plainComment =
        InspectionResultCellReportSupport.toCommentReport(
                new ResolvedComment("Review", "GridGrind", false))
            .orElseThrow();
    assertEquals("Review", plainComment.text());
    assertEquals("GridGrind", plainComment.author());
    assertEquals(java.util.Optional.empty(), plainComment.runs());
  }

  @Test
  void convertsAdvancedReadResultsIntoProtocolShapes() {
    assertWorkbookProtectionResult();
    assertCellsResult();
    assertCommentsResult();
    assertSheetAndPrintLayoutResults();
    assertAutofiltersResult();
    assertTablesResult();
    assertConditionalFormattingResult();
    assertDrawingResults();
  }

  private static void assertWorkbookProtectionResult() {
    WorkbookInspectionResult.WorkbookProtectionResult protection =
        assertInstanceOf(
            WorkbookInspectionResult.WorkbookProtectionResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookProtectionResult(
                    new StepId("workbook-protection"),
                    new WorkbookProtectionReport(true, false, true, true, false))));
    assertTrue(protection.protection().structureLocked());
  }

  private static void assertCellsResult() {
    SheetInspectionResult.CellsResult cells =
        assertInstanceOf(
            SheetInspectionResult.CellsResult.class,
            InspectionResultConverter.toReadResult(
                excelCellsResult("cells", "Budget", List.of(advancedCell()))));
    dev.erst.gridgrind.core.model.value.CellReport.TextReport cell =
        assertInstanceOf(
            dev.erst.gridgrind.core.model.value.CellReport.TextReport.class,
            cells.cells().getFirst());
    assertEquals(CellColorReport.theme(2, 0.25d), style(cell).font().fontColor());
    assertEquals(
        CellColorReport.indexed(12),
        assertInstanceOf(CellBorderSideReport.Colored.class, style(cell).border().bottom())
            .color());
    assertEquals(
        2,
        assertInstanceOf(CellGradientFillReport.Linear.class, fillGradient(style(cell).fill()))
            .stops()
            .size());
    assertEquals(
        "https://example.com/tasks/42",
        ((HyperlinkTarget.Url) cell.hyperlink().orElseThrow()).target());
    assertNotNull(cell.comment().orElseThrow());
    assertEquals(2, cell.comment().orElseThrow().runs().orElseThrow().size());
    assertEquals(1, cell.comment().orElseThrow().anchor().orElseThrow().firstColumn());
  }

  private static void assertCommentsResult() {
    SheetInspectionResult.CommentsResult comments =
        assertInstanceOf(
            SheetInspectionResult.CommentsResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookSheetResult.CommentsResult(
                    new StepId("comments"),
                    "Budget",
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookSheetResult.CellComment(
                            "C3", richComment())))));
    assertEquals(2, comments.comments().getFirst().comment().runs().orElseThrow().size());
    assertEquals(6, comments.comments().getFirst().comment().anchor().orElseThrow().lastRow());
  }

  private static void assertSheetAndPrintLayoutResults() {
    SheetInspectionResult.PrintLayoutResult printLayout =
        assertInstanceOf(
            SheetInspectionResult.PrintLayoutResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookSheetResult.PrintLayoutResult(
                    new StepId("print-layout"), "Budget", advancedPrintLayout())));
    SheetInspectionResult.SheetLayoutResult sheetLayout =
        assertInstanceOf(
            SheetInspectionResult.SheetLayoutResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayoutResult(
                    new StepId("sheet-layout"), advancedSheetLayout())));
    assertFalse(sheetLayout.layout().presentation().display().displayGridlines());
    assertEquals(
        "#112233",
        assertInstanceOf(
                CellColorReport.Rgb.class,
                sheetLayout.layout().presentation().tabColor().orElseThrow())
            .rgb()
            .value());
    assertEquals(12, sheetLayout.layout().presentation().sheetDefaults().defaultColumnWidth());
    assertEquals(
        "B2:B12", sheetLayout.layout().presentation().ignoredErrors().getFirst().range().value());
    assertEquals(
        ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT,
        sheetLayout.layout().presentation().ignoredErrors().getFirst().errorTypes().getFirst());
    assertTrue(printLayout.layout().setup().printGridlines());
    assertEquals(List.of(10, 20), printLayout.layout().setup().rowBreaks());
    assertEquals(9, printLayout.layout().setup().paperSize());
  }

  private static void assertAutofiltersResult() {
    SheetInspectionResult.AutofiltersResult autofilters =
        assertInstanceOf(
            SheetInspectionResult.AutofiltersResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookRuleResult.AutofiltersResult(
                    new StepId("autofilters"), "Budget", advancedAutofilters())));
    AutofilterEntryReport.SheetOwned sheetOwned =
        assertInstanceOf(
            AutofilterEntryReport.SheetOwned.class, autofilters.autofilters().getFirst());
    assertEquals(6, sheetOwned.filterColumns().size());
    assertInstanceOf(
        AutofilterFilterCriterionReport.Custom.class,
        sheetOwned.filterColumns().get(1).criterion());
    assertInstanceOf(
        AutofilterFilterCriterionReport.Dynamic.class,
        sheetOwned.filterColumns().get(2).criterion());
    assertInstanceOf(
        AutofilterFilterCriterionReport.Top10.class, sheetOwned.filterColumns().get(3).criterion());
    assertInstanceOf(
        AutofilterFilterCriterionReport.Color.class, sheetOwned.filterColumns().get(4).criterion());
    assertInstanceOf(
        AutofilterFilterCriterionReport.Icon.class, sheetOwned.filterColumns().get(5).criterion());
    assertEquals("A1:F5", sheetOwned.sortState().orElseThrow().range().value());
    assertEquals(3, sheetOwned.sortState().orElseThrow().conditions().size());
    assertInstanceOf(
        AutofilterSortConditionReport.CellColor.class,
        sheetOwned.sortState().orElseThrow().conditions().get(0));
    assertInstanceOf(
        AutofilterSortConditionReport.FontColor.class,
        sheetOwned.sortState().orElseThrow().conditions().get(1));
    assertInstanceOf(
        AutofilterSortConditionReport.Icon.class,
        sheetOwned.sortState().orElseThrow().conditions().get(2));
    assertInstanceOf(AutofilterEntryReport.TableOwned.class, autofilters.autofilters().get(1));
  }

  private static void assertTablesResult() {
    WorkbookAssetInspectionResult.TablesResult tables =
        assertInstanceOf(
            WorkbookAssetInspectionResult.TablesResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookRuleResult.TablesResult(
                    new StepId("tables"), List.of(advancedTable(), normalizedOptionalTable()))));
    assertEquals(
        Optional.of("HeaderStyle"), tables.tables().getFirst().presentation().headerRowCellStyle());
    assertEquals(
        Optional.of("Total"),
        tables.tables().getFirst().structure().columns().get(1).totalsRowLabel());
    assertEquals(Optional.empty(), tables.tables().get(1).presentation().comment());
    assertEquals(Optional.empty(), tables.tables().get(1).presentation().headerRowCellStyle());
    assertEquals(Optional.empty(), tables.tables().get(1).presentation().dataCellStyle());
    assertEquals(Optional.empty(), tables.tables().get(1).presentation().totalsRowCellStyle());
    assertEquals(
        Optional.empty(), tables.tables().get(1).structure().columns().getFirst().uniqueName());
    assertEquals(
        Optional.empty(), tables.tables().get(1).structure().columns().getFirst().totalsRowLabel());
  }

  private static void assertConditionalFormattingResult() {
    SheetInspectionResult.ConditionalFormattingResult conditionalFormatting =
        assertInstanceOf(
            SheetInspectionResult.ConditionalFormattingResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookRuleResult.ConditionalFormattingResult(
                    new StepId("conditional-formatting"),
                    "Budget",
                    List.of(
                        new ConditionalFormattingEntryReport(
                            List.of("A1:A5").stream()
                                .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                                .toList(),
                            List.of(
                                new ConditionalFormattingRuleReport.Top10Rule(
                                    1,
                                    false,
                                    10,
                                    true,
                                    false,
                                    Optional.ofNullable(differentialStyle()))))))));
    assertInstanceOf(
        ConditionalFormattingRuleReport.Top10Rule.class,
        conditionalFormatting.conditionalFormattingBlocks().getFirst().rules().getFirst());
  }

  private static void assertDrawingResults() {
    var signatureSetup =
        new DrawingObjectReport.SignatureSetup(
            Optional.of("{ABC}"),
            Optional.of(false),
            Optional.of("Review before signing."),
            Optional.of("Ada Lovelace"),
            Optional.of("Finance"),
            Optional.of("ada@example.com"));
    var signaturePreview =
        new DrawingObjectReport.SignaturePreview(
            ExcelPictureFormat.PNG,
            "image/png",
            42L,
            Optional.of("sig123"),
            Optional.of(400),
            Optional.of(150));
    var signatureDrawingObject =
        new DrawingObjectReport.SignatureLine(
            "OpsSignature",
            DrawingGeometryFixtures.report(
                new DrawingAnchorSpec.TwoCell(
                    new DrawingMarkerSpec(7, 3, 0, 0),
                    new DrawingMarkerSpec(10, 9, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE)),
            Optional.of(signatureSetup),
            Optional.of(signaturePreview));
    WorkbookAssetInspectionResult.DrawingObjectsResult drawingObjects =
        assertInstanceOf(
            WorkbookAssetInspectionResult.DrawingObjectsResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookDrawingResult.DrawingObjectsResult(
                    new StepId("drawing-objects"),
                    "Budget",
                    List.of(
                        new DrawingObjectReport.Picture(
                            "OpsPicture",
                            DrawingGeometryFixtures.report(
                                new DrawingAnchorSpec.TwoCell(
                                    new DrawingMarkerSpec(1, 2, 3, 4),
                                    new DrawingMarkerSpec(4, 6, 7, 8),
                                    ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE)),
                            ExcelPictureFormat.PNG,
                            "image/png",
                            68L,
                            "abc123",
                            null,
                            null,
                            "Queue preview"),
                        new DrawingObjectReport.Shape(
                            "OpsShape",
                            new DrawingAnchorReport.OneCell(
                                new DrawingMarkerReport(5, 6, 0, 0), 10L, 20L, null),
                            ExcelDrawingShapeKind.SIMPLE_SHAPE,
                            "rect",
                            "Queue",
                            0),
                        new DrawingObjectReport.Chart(
                            "OpsChart",
                            DrawingGeometryFixtures.report(
                                new DrawingAnchorSpec.TwoCell(
                                    new DrawingMarkerSpec(1, 2, 0, 0),
                                    new DrawingMarkerSpec(6, 12, 0, 0),
                                    ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE)),
                            true,
                            List.of("BAR"),
                            "Roadmap"),
                        new DrawingObjectReport.EmbeddedObject(
                            "OpsEmbed",
                            new DrawingAnchorReport.Absolute(1L, 2L, 10L, 20L, null),
                            ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                            "Payload",
                            "payload.txt",
                            "payload.txt",
                            "application/octet-stream",
                            7L,
                            "def456",
                            null,
                            null,
                            null),
                        signatureDrawingObject))));
    WorkbookAssetInspectionResult.DrawingObjectPayloadResult drawingPayload =
        assertInstanceOf(
            WorkbookAssetInspectionResult.DrawingObjectPayloadResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookDrawingResult.DrawingObjectPayloadResult(
                    new StepId("drawing-payload"),
                    "Budget",
                    new DrawingObjectPayloadReport.EmbeddedObject(
                        "OpsEmbed",
                        ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                        "application/octet-stream",
                        "payload.txt",
                        "def456",
                        java.util.Base64.getEncoder()
                            .encodeToString(
                                new BinaryContent(
                                        "payload".getBytes(java.nio.charset.StandardCharsets.UTF_8))
                                    .bytes()),
                        "Payload",
                        "payload.txt"))));
    assertEquals(5, drawingObjects.drawingObjects().size());
    DrawingObjectReport.Picture picture =
        assertInstanceOf(DrawingObjectReport.Picture.class, drawingObjects.drawingObjects().get(0));
    assertEquals(ExcelPictureFormat.PNG, picture.format());
    assertEquals(
        ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE,
        assertInstanceOf(DrawingAnchorReport.TwoCell.class, picture.anchor()).behavior());
    assertInstanceOf(DrawingObjectReport.Shape.class, drawingObjects.drawingObjects().get(1));
    DrawingObjectReport.Chart chartObject =
        assertInstanceOf(DrawingObjectReport.Chart.class, drawingObjects.drawingObjects().get(2));
    assertTrue(chartObject.supported());
    assertEquals(List.of("BAR"), chartObject.plotTypeTokens());
    assertInstanceOf(
        DrawingObjectReport.EmbeddedObject.class, drawingObjects.drawingObjects().get(3));
    DrawingObjectReport.SignatureLine signatureLine =
        assertInstanceOf(
            DrawingObjectReport.SignatureLine.class, drawingObjects.drawingObjects().get(4));
    assertEquals("{ABC}", signatureLine.setup().orElseThrow().setupId().orElseThrow());
    assertFalse(signatureLine.setup().orElseThrow().allowComments().orElseThrow());
    assertEquals(400, signatureLine.preview().orElseThrow().widthPixels().orElseThrow());
    assertEquals("cGF5bG9hZA==", drawingPayload.payload().base64Data());
  }

  @Test
  void convertsPivotReadResultsIntoProtocolShapes() {
    WorkbookAssetInspectionResult.PivotTablesResult pivotTables =
        assertInstanceOf(
            WorkbookAssetInspectionResult.PivotTablesResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookDrawingResult.PivotTablesResult(
                    new StepId("pivots"),
                    List.of(
                        new PivotTableReport.Supported(
                            "Sales Pivot 2026",
                            new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
                            new PivotTableReport.Anchor(
                                new dev.erst.gridgrind.core.vocabulary.CellAddress("C5"),
                                new dev.erst.gridgrind.core.vocabulary.RangeLabel("C5:G9")),
                            new PivotTableReport.PivotSourceReport.Table(
                                new dev.erst.gridgrind.core.vocabulary.TableName("SalesTable"),
                                new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:D5")),
                            List.of(new PivotTableReport.Field(0, "Region")),
                            List.of(new PivotTableReport.Field(1, "Stage")),
                            List.of(new PivotTableReport.Field(2, "Owner")),
                            List.of(
                                new PivotTableReport.DataField(
                                    3,
                                    "Amount",
                                    ExcelPivotDataConsolidateFunction.SUM,
                                    "Total Amount",
                                    Optional.of("#,##0.00"))),
                            true),
                        new PivotTableReport.Unsupported(
                            "Broken Pivot",
                            new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
                            new PivotTableReport.Anchor(
                                new dev.erst.gridgrind.core.vocabulary.CellAddress("A3"),
                                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A3:C8")),
                            "Pivot cache source no longer resolves cleanly.")))));
    WorkbookAnalysisResult.PivotTableHealthResult pivotTableHealth =
        assertInstanceOf(
            WorkbookAnalysisResult.PivotTableHealthResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookAnalysisResult.PivotTableHealthResult(
                    new StepId("pivot-health"),
                    new WorkbookAnalysis.PivotTableHealth(
                        1,
                        new WorkbookAnalysis.AnalysisSummary(1, 0, 1, 0),
                        List.of(
                            new WorkbookAnalysis.AnalysisFinding(
                                AnalysisFindingCode.PIVOT_TABLE_MISSING_NAME,
                                AnalysisSeverity.WARNING,
                                "Pivot table name is missing",
                                "GridGrind assigned a synthetic identifier for readback.",
                                new WorkbookAnalysis.AnalysisLocation.Sheet("Report"),
                                List.of("_GG_PIVOT_Report_A3")))))));

    PivotTableReport.Supported supported =
        assertInstanceOf(PivotTableReport.Supported.class, pivotTables.pivotTables().getFirst());
    PivotTableReport.Unsupported unsupported =
        assertInstanceOf(PivotTableReport.Unsupported.class, pivotTables.pivotTables().get(1));

    assertEquals(
        "SalesTable",
        ((PivotTableReport.PivotSourceReport.Table) supported.source()).name().value());
    assertEquals("Amount", supported.dataFields().getFirst().sourceColumnName());
    assertTrue(supported.valuesAxisOnColumns());
    assertEquals("Pivot cache source no longer resolves cleanly.", unsupported.detail());
    assertEquals(1, pivotTableHealth.analysis().checkedPivotTableCount());
    assertEquals(
        AnalysisFindingCode.PIVOT_TABLE_MISSING_NAME,
        pivotTableHealth.analysis().findings().getFirst().code());
  }

  @Test
  void convertsArrayFormulaReadResultsIntoProtocolShapes() {
    SheetInspectionResult.ArrayFormulasResult arrayFormulas =
        assertInstanceOf(
            SheetInspectionResult.ArrayFormulasResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookSheetResult.ArrayFormulasResult(
                    new StepId("array-formulas"),
                    List.of(
                        new ArrayFormulaReport(
                            new dev.erst.gridgrind.core.vocabulary.SheetLabel("Calc"),
                            new dev.erst.gridgrind.core.vocabulary.RangeLabel("D2:D4"),
                            new dev.erst.gridgrind.core.vocabulary.CellAddress("D2"),
                            "B2:B4*C2:C4",
                            false),
                        new ArrayFormulaReport(
                            new dev.erst.gridgrind.core.vocabulary.SheetLabel("Calc"),
                            new dev.erst.gridgrind.core.vocabulary.RangeLabel("F2"),
                            new dev.erst.gridgrind.core.vocabulary.CellAddress("F2"),
                            "SUM(B2:C2)",
                            true)))));

    assertEquals("array-formulas", arrayFormulas.stepId().value());
    assertEquals(2, arrayFormulas.arrayFormulas().size());
    assertEquals("D2:D4", arrayFormulas.arrayFormulas().getFirst().range().value());
    assertTrue(arrayFormulas.arrayFormulas().get(1).singleCell());
  }

  @Test
  void convertsPivotSourceVariantsIntoProtocolShapes() {
    WorkbookAssetInspectionResult.PivotTablesResult pivotTables =
        assertInstanceOf(
            WorkbookAssetInspectionResult.PivotTablesResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookDrawingResult.PivotTablesResult(
                    new StepId("pivots"),
                    List.of(
                        new PivotTableReport.Supported(
                            "Range Pivot",
                            new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
                            new PivotTableReport.Anchor(
                                new dev.erst.gridgrind.core.vocabulary.CellAddress("C5"),
                                new dev.erst.gridgrind.core.vocabulary.RangeLabel("C5:G9")),
                            new PivotTableReport.PivotSourceReport.Range(
                                new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:D5")),
                            List.of(),
                            List.of(),
                            List.of(),
                            List.of(
                                new PivotTableReport.DataField(
                                    3,
                                    "Amount",
                                    ExcelPivotDataConsolidateFunction.SUM,
                                    "Total Amount",
                                    Optional.empty())),
                            false),
                        new PivotTableReport.Supported(
                            "Named Range Pivot",
                            new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
                            new PivotTableReport.Anchor(
                                new dev.erst.gridgrind.core.vocabulary.CellAddress("A3"),
                                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A3:E8")),
                            new PivotTableReport.PivotSourceReport.NamedRange(
                                new dev.erst.gridgrind.core.vocabulary.DefinedName("PivotSource"),
                                new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:D5")),
                            List.of(),
                            List.of(),
                            List.of(),
                            List.of(
                                new PivotTableReport.DataField(
                                    3,
                                    "Amount",
                                    ExcelPivotDataConsolidateFunction.SUM,
                                    "Total Amount",
                                    Optional.empty())),
                            false)))));

    PivotTableReport.Supported rangePivot =
        assertInstanceOf(PivotTableReport.Supported.class, pivotTables.pivotTables().getFirst());
    PivotTableReport.Supported namedRangePivot =
        assertInstanceOf(PivotTableReport.Supported.class, pivotTables.pivotTables().get(1));

    assertEquals(
        "Data",
        ((PivotTableReport.PivotSourceReport.Range) rangePivot.source()).sheetName().value());
    assertEquals(
        "PivotSource",
        ((PivotTableReport.PivotSourceReport.NamedRange) namedRangePivot.source()).name().value());
  }

  @Test
  void convertsChartReadResultsIntoProtocolShapes() {
    WorkbookAssetInspectionResult.ChartsResult charts =
        assertInstanceOf(
            WorkbookAssetInspectionResult.ChartsResult.class,
            InspectionResultConverter.toReadResult(baseChartsResult()));
    ChartReport chartReport = charts.charts().getFirst();
    ChartReport.Bar chartPlot =
        assertInstanceOf(ChartReport.Bar.class, chartReport.plots().getFirst());
    assertEquals("Roadmap", ((ChartReport.Title.Text) chartReport.title()).text());
    assertEquals(
        "ChartValues",
        ((ChartReport.DataSource.NumericReference) chartPlot.series().getFirst().values())
            .formula());

    WorkbookAssetInspectionResult.ChartsResult advancedCharts =
        assertInstanceOf(
            WorkbookAssetInspectionResult.ChartsResult.class,
            InspectionResultConverter.toReadResult(advancedChartsResult()));
    ChartReport lineReport = advancedCharts.charts().get(0);
    ChartReport.Line linePlot =
        assertInstanceOf(ChartReport.Line.class, lineReport.plots().getFirst());
    assertTrue(lineReport.title() instanceof ChartReport.Title.None);
    assertEquals(
        List.of("Jan", "Feb"),
        ((ChartReport.DataSource.StringLiteral) linePlot.series().getFirst().categories())
            .values());
    assertEquals(
        List.of("10", "18"),
        ((ChartReport.DataSource.NumericLiteral) linePlot.series().getFirst().values()).values());
    ChartReport pieReport = advancedCharts.charts().get(1);
    ChartReport.Pie piePlot = assertInstanceOf(ChartReport.Pie.class, pieReport.plots().getFirst());
    assertEquals(Optional.of(120), piePlot.firstSliceAngle());
    assertEquals("Actual", ((ChartReport.Title.Formula) pieReport.title()).cachedText());
    ChartReport unsupportedChart = advancedCharts.charts().get(2);
    ChartReport.Unsupported unsupportedPlot =
        assertInstanceOf(ChartReport.Unsupported.class, unsupportedChart.plots().getFirst());
    assertEquals("AREA", unsupportedPlot.plotTypeToken());
  }

  private static dev.erst.gridgrind.excel.WorkbookDrawingResult.ChartsResult baseChartsResult() {
    return new dev.erst.gridgrind.excel.WorkbookDrawingResult.ChartsResult(
        new StepId("charts"),
        "Budget",
        List.of(
            chartSnapshot(
                "OpsChart",
                new DrawingAnchorSpec.TwoCell(
                    new DrawingMarkerSpec(1, 2, 0, 0),
                    new DrawingMarkerSpec(6, 12, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE),
                new ChartReport.Title.Text("Roadmap"),
                new ChartReport.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
                ExcelChartDisplayBlanksAs.SPAN,
                false,
                new ChartReport.Bar(
                    true,
                    ExcelChartBarDirection.COLUMN,
                    ExcelChartBarGrouping.CLUSTERED,
                    Optional.empty(),
                    Optional.empty(),
                    List.of(
                        new ChartReport.Axis(
                            ExcelChartAxisKind.CATEGORY,
                            ExcelChartAxisPosition.BOTTOM,
                            ExcelChartAxisCrosses.AUTO_ZERO,
                            true),
                        new ChartReport.Axis(
                            ExcelChartAxisKind.VALUE,
                            ExcelChartAxisPosition.LEFT,
                            ExcelChartAxisCrosses.AUTO_ZERO,
                            true)),
                    List.of(
                        chartSnapshotSeries(
                            new ChartReport.Title.Formula("Chart!$B$1", "Plan"),
                            new ChartReport.DataSource.StringReference(
                                "ChartCategories", List.of("Jan", "Feb", "Mar")),
                            new ChartReport.DataSource.NumericReference(
                                "ChartValues",
                                Optional.empty(),
                                List.of("10.0", "18.0", "15.0"))))))));
  }

  private static dev.erst.gridgrind.excel.WorkbookDrawingResult.ChartsResult
      advancedChartsResult() {
    return new dev.erst.gridgrind.excel.WorkbookDrawingResult.ChartsResult(
        new StepId("charts-advanced"),
        "Budget",
        List.of(
            chartSnapshot(
                "TrendChart",
                new DrawingAnchorSpec.TwoCell(
                    new DrawingMarkerSpec(2, 3, 0, 0),
                    new DrawingMarkerSpec(7, 14, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                new ChartReport.Title.None(),
                new ChartReport.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.GAP,
                true,
                new ChartReport.Line(
                    false,
                    ExcelChartGrouping.STANDARD,
                    List.of(
                        new ChartReport.Axis(
                            ExcelChartAxisKind.CATEGORY,
                            ExcelChartAxisPosition.BOTTOM,
                            ExcelChartAxisCrosses.AUTO_ZERO,
                            true),
                        new ChartReport.Axis(
                            ExcelChartAxisKind.VALUE,
                            ExcelChartAxisPosition.LEFT,
                            ExcelChartAxisCrosses.MIN,
                            true)),
                    List.of(
                        chartSnapshotSeries(
                            new ChartReport.Title.None(),
                            new ChartReport.DataSource.StringLiteral(List.of("Jan", "Feb")),
                            new ChartReport.DataSource.NumericLiteral(
                                Optional.of("0.0"), List.of("10", "18")))))),
            chartSnapshot(
                "ShareChart",
                new DrawingAnchorSpec.TwoCell(
                    new DrawingMarkerSpec(8, 3, 0, 0),
                    new DrawingMarkerSpec(13, 14, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                new ChartReport.Title.Formula("Budget!$C$1", "Actual"),
                new ChartReport.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.ZERO,
                false,
                new ChartReport.Pie(
                    true,
                    Optional.of(120),
                    List.of(
                        chartSnapshotSeries(
                            new ChartReport.Title.Text("Actual"),
                            new ChartReport.DataSource.StringReference(
                                "ChartCategories", List.of("Jan", "Feb")),
                            new ChartReport.DataSource.NumericReference(
                                "ChartActual", Optional.of("0.0"), List.of("12", "16")))))),
            chartSnapshot(
                "AreaOnly",
                new DrawingAnchorSpec.TwoCell(
                    new DrawingMarkerSpec(14, 3, 0, 0),
                    new DrawingMarkerSpec(19, 14, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                new ChartReport.Title.None(),
                new ChartReport.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.GAP,
                true,
                new ChartReport.Unsupported(
                    "AREA",
                    "Chart plot family is outside the current modeled simple-chart contract."))));
  }

  private static ChartReport chartSnapshot(
      String name,
      DrawingAnchorSpec.TwoCell anchor,
      ChartReport.Title title,
      ChartReport.Legend legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      boolean plotOnlyVisibleCells,
      ChartReport.Plot plot) {
    return new ChartReport(
        name,
        DrawingGeometryFixtures.report(anchor),
        title,
        legend,
        displayBlanksAs,
        plotOnlyVisibleCells,
        List.of(plot));
  }

  private static ChartReport.Series chartSnapshotSeries(
      ChartReport.Title title, ChartReport.DataSource categories, ChartReport.DataSource values) {
    return new ChartReport.Series(
        title,
        categories,
        values,
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }

  private static ExcelCellSnapshot advancedCell() {
    List<RichTextRunReport> richText =
        List.copyOf(List.of(new RichTextRunReport("Styled", advancedFont())));
    return new ExcelCellSnapshot.TextSnapshot(
        "C3",
        "Styled",
        advancedStyle(),
        ExcelCellMetadataSnapshot.of(
            Optional.of(new HyperlinkTarget.Url("https://example.com/tasks/42")),
            Optional.of(richComment())),
        "Styled",
        richText);
  }

  private static CommentReport richComment() {
    return new CommentReport(
        "Hi there",
        "Ada",
        true,
        Optional.of(
            List.copyOf(
                List.of(
                    new RichTextRunReport("Hi ", advancedFont()),
                    new RichTextRunReport("there", accentFont())))),
        Optional.of(new CommentAnchorReport(1, 2, 4, 6)));
  }

  private static CellStyleReport advancedStyle() {
    return new CellStyleReport(
        "0.00",
        new CellAlignmentReport(
            false, ExcelHorizontalAlignment.CENTER, ExcelVerticalAlignment.CENTER, 0, 0),
        advancedFont(),
        CellFillReport.gradient(
            CellGradientFillReport.linear(
                45.0d,
                List.of(
                    new CellGradientStopReport(0.0d, CellColorReport.rgb(new RgbHex("#112233"))),
                    new CellGradientStopReport(1.0d, CellColorReport.theme(4, 0.45d))))),
        new CellBorderReport(
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None(),
            new CellBorderSideReport.Colored(ExcelBorderStyle.THICK, CellColorReport.indexed(12)),
            new CellBorderSideReport.None()),
        new CellProtectionReport(true, false));
  }

  private static CellFontReport advancedFont() {
    return new CellFontReport(
        true,
        false,
        "Aptos",
        FontHeightReport.fromPoints(new BigDecimal("11")),
        CellColorReport.theme(2, 0.25d),
        false,
        false);
  }

  private static CellFontReport accentFont() {
    return new CellFontReport(
        false,
        false,
        "Aptos",
        FontHeightReport.fromPoints(new BigDecimal("11")),
        CellColorReport.rgb(new RgbHex("#AABBCC")),
        false,
        false);
  }

  private static PrintLayoutReport advancedPrintLayout() {
    return new PrintLayoutReport(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
        new PrintAreaReport.Range(new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:F20")),
        ExcelPrintOrientation.LANDSCAPE,
        new PrintScalingReport.Fit(1, 0),
        new PrintTitleRowsReport.Band(0, 0),
        new PrintTitleColumnsReport.None(),
        new HeaderFooterTextReport("Ops", "Queue", ""),
        new HeaderFooterTextReport("", "Internal", "Page &P"),
        new PrintSetupReport(
            new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
            true,
            true,
            false,
            9,
            false,
            true,
            2,
            true,
            3,
            List.of(10, 20),
            List.of(2, 4)));
  }

  private static dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayout advancedSheetLayout() {
    return new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayout(
        "Budget",
        new ExcelSheetPane.Frozen(1, 1, 1, 1),
        125,
        new SheetPresentationReport(
            new SheetDisplayReport(false, false, true, true, true),
            Optional.of(CellColorReport.rgb(new RgbHex("#112233"))),
            new SheetOutlineSummaryReport(false, false),
            new SheetDefaultsReport(12, 18.5d),
            List.of(
                new IgnoredErrorReport(
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel("B2:B12"),
                    List.of(
                        ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT,
                        ExcelIgnoredErrorType.FORMULA)))),
        List.of(
            new dev.erst.gridgrind.excel.WorkbookSheetResult.ColumnLayout(
                0, 16.0d, false, 0, false)),
        List.of(
            new dev.erst.gridgrind.excel.WorkbookSheetResult.RowLayout(0, 20.0d, false, 0, false)));
  }

  private static List<AutofilterEntryReport> advancedAutofilters() {
    AutofilterSortStateReport sortState =
        new AutofilterSortStateReport(
            new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:F5"),
            true,
            false,
            java.util.Optional.empty(),
            List.of(
                new AutofilterSortConditionReport.CellColor(
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel("A2:A5"),
                    true,
                    CellColorReport.rgb(new RgbHex("#AABBCC"))),
                new AutofilterSortConditionReport.FontColor(
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel("B2:B5"),
                    false,
                    CellColorReport.theme(4, 0.15d)),
                new AutofilterSortConditionReport.Icon(
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel("C2:C5"), false, 2)));
    List<AutofilterFilterColumnReport> filterColumns =
        List.of(
            new AutofilterFilterColumnReport(
                0L, false, new AutofilterFilterCriterionReport.Values(List.of("Queued"), true)),
            new AutofilterFilterColumnReport(
                1L,
                true,
                new AutofilterFilterCriterionReport.Custom(
                    true,
                    List.of(
                        new AutofilterFilterCriterionReport.CustomConditionReport(
                            "equal", "Ada")))),
            new AutofilterFilterColumnReport(
                2L, true, new AutofilterFilterCriterionReport.Dynamic("TODAY", 1.0d, 2.0d)),
            new AutofilterFilterColumnReport(
                3L, true, new AutofilterFilterCriterionReport.Top10(true, false, 10.0d, 8.0d)),
            new AutofilterFilterColumnReport(
                4L,
                true,
                new AutofilterFilterCriterionReport.Color(false, CellColorReport.theme(4, 0.45d))),
            new AutofilterFilterColumnReport(
                5L, true, new AutofilterFilterCriterionReport.Icon("3TrafficLights1", 2)));
    return List.of(
        new AutofilterEntryReport.SheetOwned(
            new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:F5"),
            filterColumns,
            java.util.Optional.of(sortState)),
        new AutofilterEntryReport.TableOwned(
            new dev.erst.gridgrind.core.vocabulary.RangeLabel("H1:I5"),
            "QueueTable",
            List.of(),
            java.util.Optional.of(sortState)));
  }

  private static TableEntryReport advancedTable() {
    return new TableEntryReport(
        new dev.erst.gridgrind.core.vocabulary.TableLabel("QueueTable"),
        new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
        new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B5"),
        new TableEntryReport.Structure(
            1,
            1,
            List.of("Item", "Amount"),
            List.of(
                new TableColumnReport(
                    1L,
                    "Item",
                    Optional.ofNullable(""),
                    Optional.ofNullable(""),
                    Optional.ofNullable(""),
                    Optional.ofNullable("")),
                new TableColumnReport(
                    2L,
                    "Amount",
                    Optional.ofNullable("UniqueAmount"),
                    Optional.ofNullable("Total"),
                    Optional.ofNullable("sum"),
                    Optional.ofNullable("[@Amount]*2")))),
        new TableStyleReport.Named("TableStyleMedium2", false, false, true, false),
        new TableEntryReport.Behavior(true, true, true, false),
        new TableEntryReport.Presentation(
            Optional.of("Queue comment"),
            Optional.of("HeaderStyle"),
            Optional.of("DataStyle"),
            Optional.of("TotalsStyle")));
  }

  private static TableEntryReport normalizedOptionalTable() {
    return new TableEntryReport(
        new dev.erst.gridgrind.core.vocabulary.TableLabel("OptionalTable"),
        new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
        new dev.erst.gridgrind.core.vocabulary.RangeLabel("D1:E2"),
        new TableEntryReport.Structure(
            1,
            0,
            List.of("Code"),
            List.of(
                new TableColumnReport(
                    3L,
                    "Code",
                    Optional.empty(),
                    Optional.ofNullable(" "),
                    Optional.ofNullable(""),
                    Optional.ofNullable(" ")))),
        new TableStyleReport.None(),
        new TableEntryReport.Behavior(false, false, false, false),
        new TableEntryReport.Presentation(
            Optional.of(" "), Optional.of(""), Optional.of(" "), Optional.empty()));
  }

  private static CustomXmlMappingReport customXmlMappingSnapshot(
      long mapId,
      String name,
      String rootElement,
      String schemaId,
      CustomXmlMappingReport.Settings settings,
      CustomXmlMappingReport.Schema schema,
      Optional<CustomXmlDataBindingReport> dataBinding,
      List<CustomXmlLinkedCellReport> linkedCells,
      List<dev.erst.gridgrind.core.model.customxml.CustomXmlLinkedTableReport> linkedTables) {
    return new CustomXmlMappingReport(
        mapId,
        name,
        rootElement,
        schemaId,
        settings,
        schema,
        dataBinding == null ? null : dataBinding.orElse(null),
        linkedCells,
        linkedTables);
  }

  private static CustomXmlMappingReport.Settings settings(
      boolean showImportExportValidationErrors,
      boolean autoFit,
      boolean append,
      boolean preserveSortAfLayout,
      boolean preserveFormat) {
    return new CustomXmlMappingReport.Settings(
        showImportExportValidationErrors, autoFit, append, preserveSortAfLayout, preserveFormat);
  }

  private static CustomXmlMappingReport.Schema schema(
      Optional<String> namespace,
      Optional<String> language,
      Optional<String> reference,
      Optional<String> xml) {
    return new CustomXmlMappingReport.Schema(
        nullable(namespace), nullable(language), nullable(reference), nullable(xml));
  }

  private static DifferentialStyleReport differentialStyle() {
    return new DifferentialStyleReport(
        "0.00",
        true,
        null,
        null,
        Optional.ofNullable(
            dev.erst.gridgrind.excel.DifferentialStyleFixtures.color(
                ColorSpec.rgb(new RgbHex("#AABBCC")))),
        null,
        null,
        Optional.empty(),
        Optional.empty(),
        List.of());
  }

  private static <T> T nullable(Optional<T> value) {
    return value == null ? null : value.orElse(null);
  }
}
