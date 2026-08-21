package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ProtocolStyleTestAccess.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.SheetInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookAnalysisResult;
import dev.erst.gridgrind.contract.query.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookInspectionResult;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlDataBindingSnapshot;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlExportSnapshot;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlLinkedCellSnapshot;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlMappingSettings;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlMappingSnapshot;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlSchemaSnapshot;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectPayload;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectSnapshot;
import dev.erst.gridgrind.excel.drawing.ExcelSignatureLineSnapshot;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlHashAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlEncryptionSnapshot;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageSecuritySnapshot;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlSignatureSnapshot;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableSnapshot;
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
                    "security",
                    new ExcelOoxmlPackageSecuritySnapshot(
                        new ExcelOoxmlEncryptionSnapshot.Encrypted(
                            ExcelOoxmlEncryptionMode.AGILE,
                            ExcelOoxmlCipherAlgorithm.AES_256,
                            ExcelOoxmlHashAlgorithm.SHA_512,
                            ExcelOoxmlChainingMode.CBC,
                            256,
                            16,
                            100_000),
                        List.of(
                            new ExcelOoxmlSignatureSnapshot(
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
                    "custom-xml-mappings",
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
                                new ExcelCustomXmlDataBindingSnapshot(
                                    Optional.of("binding"),
                                    Optional.of(true),
                                    Optional.of(5L),
                                    Optional.of("binding.xml"),
                                    2L)),
                            List.of(
                                new ExcelCustomXmlLinkedCellSnapshot(
                                    "Foglio1", "A1", "/CORSO/NOME", "string")),
                            List.of())))));
    WorkbookInspectionResult.CustomXmlExportResult exported =
        assertInstanceOf(
            WorkbookInspectionResult.CustomXmlExportResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookCoreResult.CustomXmlExportResult(
                    "custom-xml-export",
                    new ExcelCustomXmlExportSnapshot(
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
                                new ExcelCustomXmlLinkedCellSnapshot(
                                    "Foglio1", "A1", "/CORSO/NOME", "string")),
                            List.of()),
                        "UTF-8",
                        true,
                        "<CORSO><NOME>Grid</NOME></CORSO>"))));

    assertEquals("custom-xml-mappings", mappings.stepId());
    assertEquals("CORSO_mapping", mappings.mappings().getFirst().name());
    assertEquals(2L, mappings.mappings().getFirst().dataBinding().loadMode());
    assertEquals("A1", mappings.mappings().getFirst().linkedCells().getFirst().address());
    assertEquals("custom-xml-export", exported.stepId());
    assertEquals("UTF-8", exported.export().encoding());
    assertTrue(exported.export().xml().contains("<NOME>Grid</NOME>"));
  }

  @Test
  void convertsPlainCommentAndWorkbookProtectionFactsDirectly() {
    assertTrue(InspectionResultCellReportSupport.toCommentReport((ExcelComment) null).isEmpty());
    CommentReport plainComment =
        InspectionResultCellReportSupport.toCommentReport(
                new ExcelComment("Review", "GridGrind", false))
            .orElseThrow();
    WorkbookProtectionReport protection =
        InspectionResultWorkbookCoreReportSupport.toWorkbookProtectionReport(
            new ExcelWorkbookProtectionSnapshot(true, false, true, true, false));

    assertEquals("Review", plainComment.text());
    assertEquals("GridGrind", plainComment.author());
    assertEquals(java.util.Optional.empty(), plainComment.runs());
    assertTrue(protection.structureLocked());
    assertTrue(protection.revisionsLocked());
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
                    "workbook-protection",
                    new ExcelWorkbookProtectionSnapshot(true, false, true, true, false))));
    assertTrue(protection.protection().structureLocked());
  }

  private static void assertCellsResult() {
    SheetInspectionResult.CellsResult cells =
        assertInstanceOf(
            SheetInspectionResult.CellsResult.class,
            InspectionResultConverter.toReadResult(
                excelCellsResult("cells", "Budget", List.of(advancedCell()))));
    dev.erst.gridgrind.contract.dto.CellReport.TextReport cell =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.CellReport.TextReport.class, cells.cells().getFirst());
    assertEquals(CellColorReport.theme(2, 0.25d), style(cell).font().fontColor());
    assertEquals(CellColorReport.indexed(12), style(cell).border().bottom().color());
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
                    "comments",
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
                    "print-layout", "Budget", advancedPrintLayout())));
    SheetInspectionResult.SheetLayoutResult sheetLayout =
        assertInstanceOf(
            SheetInspectionResult.SheetLayoutResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayoutResult(
                    "sheet-layout", advancedSheetLayout())));
    assertFalse(sheetLayout.layout().presentation().display().displayGridlines());
    assertEquals(
        "#112233",
        assertInstanceOf(
                CellColorReport.Rgb.class,
                sheetLayout.layout().presentation().tabColor().orElseThrow())
            .rgb());
    assertEquals(12, sheetLayout.layout().presentation().sheetDefaults().defaultColumnWidth());
    assertEquals("B2:B12", sheetLayout.layout().presentation().ignoredErrors().getFirst().range());
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
                    "autofilters", "Budget", advancedAutofilters())));
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
    assertEquals("A1:F5", sheetOwned.sortState().orElseThrow().range());
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
                    "tables", List.of(advancedTable(), normalizedOptionalTable()))));
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
                    "conditional-formatting",
                    "Budget",
                    List.of(
                        new ExcelConditionalFormattingBlockSnapshot(
                            List.of("A1:A5"),
                            List.of(
                                new ExcelConditionalFormattingRuleSnapshot.Top10Rule(
                                    1, false, 10, true, false, differentialStyle())))))));
    assertInstanceOf(
        ConditionalFormattingRuleReport.Top10Rule.class,
        conditionalFormatting.conditionalFormattingBlocks().getFirst().rules().getFirst());
  }

  private static void assertDrawingResults() {
    var signatureSetup =
        new ExcelSignatureLineSnapshot.Setup(
            Optional.of("{ABC}"),
            Optional.of(false),
            Optional.of("Review before signing."),
            Optional.of("Ada Lovelace"),
            Optional.of("Finance"),
            Optional.of("ada@example.com"));
    var signaturePreview =
        new ExcelSignatureLineSnapshot.Preview(
            ExcelPictureFormat.PNG,
            "image/png",
            42L,
            Optional.of("sig123"),
            Optional.of(400),
            Optional.of(150));
    var signatureDrawingObject =
        new ExcelDrawingObjectSnapshot.SignatureLine(
            "OpsSignature",
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(7, 3, 0, 0),
                new ExcelDrawingMarker(10, 9, 0, 0),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
            Optional.of(signatureSetup),
            Optional.of(signaturePreview));
    WorkbookAssetInspectionResult.DrawingObjectsResult drawingObjects =
        assertInstanceOf(
            WorkbookAssetInspectionResult.DrawingObjectsResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookDrawingResult.DrawingObjectsResult(
                    "drawing-objects",
                    "Budget",
                    List.of(
                        new ExcelDrawingObjectSnapshot.Picture(
                            "OpsPicture",
                            new ExcelDrawingAnchor.TwoCell(
                                new ExcelDrawingMarker(1, 2, 3, 4),
                                new ExcelDrawingMarker(4, 6, 7, 8),
                                ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE),
                            ExcelPictureFormat.PNG,
                            "image/png",
                            68L,
                            "abc123",
                            null,
                            null,
                            "Queue preview"),
                        new ExcelDrawingObjectSnapshot.Shape(
                            "OpsShape",
                            new ExcelDrawingAnchor.OneCell(
                                new ExcelDrawingMarker(5, 6, 0, 0), 10L, 20L, null),
                            ExcelDrawingShapeKind.SIMPLE_SHAPE,
                            "rect",
                            "Queue",
                            0),
                        new ExcelDrawingObjectSnapshot.Chart(
                            "OpsChart",
                            new ExcelDrawingAnchor.TwoCell(
                                new ExcelDrawingMarker(1, 2, 0, 0),
                                new ExcelDrawingMarker(6, 12, 0, 0),
                                ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE),
                            true,
                            List.of("BAR"),
                            "Roadmap"),
                        new ExcelDrawingObjectSnapshot.EmbeddedObject(
                            "OpsEmbed",
                            new ExcelDrawingAnchor.Absolute(1L, 2L, 10L, 20L, null),
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
                    "drawing-payload",
                    "Budget",
                    new ExcelDrawingObjectPayload.EmbeddedObject(
                        "OpsEmbed",
                        ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                        "application/octet-stream",
                        "payload.txt",
                        "def456",
                        new ExcelBinaryData(
                            "payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
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
    DrawingObjectPayloadReport.Picture picturePayload =
        assertInstanceOf(
            DrawingObjectPayloadReport.Picture.class,
            InspectionResultDrawingReportSupport.toDrawingObjectPayloadReport(
                new ExcelDrawingObjectPayload.Picture(
                    "OpsPicture",
                    ExcelPictureFormat.PNG,
                    "image/png",
                    "OpsPicture.png",
                    "abc123",
                    new ExcelBinaryData(
                        "picture".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
                    "Queue preview")));
    assertEquals("OpsPicture.png", picturePayload.fileName());
    assertEquals("cGljdHVyZQ==", picturePayload.base64Data());
  }

  @Test
  void convertsPivotReadResultsIntoProtocolShapes() {
    WorkbookAssetInspectionResult.PivotTablesResult pivotTables =
        assertInstanceOf(
            WorkbookAssetInspectionResult.PivotTablesResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookDrawingResult.PivotTablesResult(
                    "pivots",
                    List.of(
                        new ExcelPivotTableSnapshot.Supported(
                            "Sales Pivot 2026",
                            "Report",
                            new ExcelPivotTableSnapshot.Anchor("C5", "C5:G9"),
                            new ExcelPivotTableSnapshot.Source.Table("SalesTable", "Data", "A1:D5"),
                            List.of(new ExcelPivotTableSnapshot.Field(0, "Region")),
                            List.of(new ExcelPivotTableSnapshot.Field(1, "Stage")),
                            List.of(new ExcelPivotTableSnapshot.Field(2, "Owner")),
                            List.of(
                                new ExcelPivotTableSnapshot.DataField(
                                    3,
                                    "Amount",
                                    ExcelPivotDataConsolidateFunction.SUM,
                                    "Total Amount",
                                    Optional.of("#,##0.00"))),
                            true),
                        new ExcelPivotTableSnapshot.Unsupported(
                            "Broken Pivot",
                            "Report",
                            new ExcelPivotTableSnapshot.Anchor("A3", "A3:C8"),
                            "Pivot cache source no longer resolves cleanly.")))));
    WorkbookAnalysisResult.PivotTableHealthResult pivotTableHealth =
        assertInstanceOf(
            WorkbookAnalysisResult.PivotTableHealthResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookAnalysisResult.PivotTableHealthResult(
                    "pivot-health",
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

    assertEquals("SalesTable", ((PivotTableReport.Source.Table) supported.source()).name());
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
                    "array-formulas",
                    List.of(
                        new ExcelArrayFormulaSnapshot("Calc", "D2:D4", "D2", "B2:B4*C2:C4", false),
                        new ExcelArrayFormulaSnapshot("Calc", "F2", "F2", "SUM(B2:C2)", true)))));

    assertEquals("array-formulas", arrayFormulas.stepId());
    assertEquals(2, arrayFormulas.arrayFormulas().size());
    assertEquals("D2:D4", arrayFormulas.arrayFormulas().getFirst().range());
    assertTrue(arrayFormulas.arrayFormulas().get(1).singleCell());
  }

  @Test
  void convertsPivotSourceVariantsIntoProtocolShapes() {
    WorkbookAssetInspectionResult.PivotTablesResult pivotTables =
        assertInstanceOf(
            WorkbookAssetInspectionResult.PivotTablesResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookDrawingResult.PivotTablesResult(
                    "pivots",
                    List.of(
                        new ExcelPivotTableSnapshot.Supported(
                            "Range Pivot",
                            "Report",
                            new ExcelPivotTableSnapshot.Anchor("C5", "C5:G9"),
                            new ExcelPivotTableSnapshot.Source.Range("Data", "A1:D5"),
                            List.of(),
                            List.of(),
                            List.of(),
                            List.of(
                                new ExcelPivotTableSnapshot.DataField(
                                    3,
                                    "Amount",
                                    ExcelPivotDataConsolidateFunction.SUM,
                                    "Total Amount",
                                    Optional.empty())),
                            false),
                        new ExcelPivotTableSnapshot.Supported(
                            "Named Range Pivot",
                            "Report",
                            new ExcelPivotTableSnapshot.Anchor("A3", "A3:E8"),
                            new ExcelPivotTableSnapshot.Source.NamedRange(
                                "PivotSource", "Data", "A1:D5"),
                            List.of(),
                            List.of(),
                            List.of(),
                            List.of(
                                new ExcelPivotTableSnapshot.DataField(
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

    assertEquals("Data", ((PivotTableReport.Source.Range) rangePivot.source()).sheetName());
    assertEquals(
        "PivotSource", ((PivotTableReport.Source.NamedRange) namedRangePivot.source()).name());
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

  @Test
  void directChartReportConversionCoversStandaloneSwitchBranches() {
    ExcelChartSnapshot lineSnapshot =
        chartSnapshot(
            "OpsLine",
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(2, 3, 0, 0),
                new ExcelDrawingMarker(7, 14, 0, 0),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
            new ExcelChartSnapshot.Title.None(),
            new ExcelChartSnapshot.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            new ExcelChartSnapshot.Line(
                false,
                ExcelChartGrouping.STANDARD,
                List.of(
                    new ExcelChartSnapshot.Axis(
                        ExcelChartAxisKind.CATEGORY,
                        ExcelChartAxisPosition.TOP,
                        ExcelChartAxisCrosses.AUTO_ZERO,
                        true),
                    new ExcelChartSnapshot.Axis(
                        ExcelChartAxisKind.VALUE,
                        ExcelChartAxisPosition.RIGHT,
                        ExcelChartAxisCrosses.MAX,
                        false)),
                List.of(
                    chartSnapshotSeries(
                        new ExcelChartSnapshot.Title.None(),
                        new ExcelChartSnapshot.DataSource.StringLiteral(List.of("Jan", "Feb")),
                        new ExcelChartSnapshot.DataSource.NumericLiteral(
                            Optional.of("0.0"), List.of("10", "18"))))));
    ExcelChartSnapshot pieSnapshot =
        chartSnapshot(
            "OpsPie",
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(8, 3, 0, 0),
                new ExcelDrawingMarker(13, 14, 0, 0),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
            new ExcelChartSnapshot.Title.Formula("Budget!$C$1", "Actual"),
            new ExcelChartSnapshot.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            false,
            new ExcelChartSnapshot.Pie(
                true,
                Optional.of(120),
                List.of(
                    chartSnapshotSeries(
                        new ExcelChartSnapshot.Title.Text("Actual"),
                        new ExcelChartSnapshot.DataSource.StringReference(
                            "ChartCategories", List.of("Jan", "Feb")),
                        new ExcelChartSnapshot.DataSource.NumericReference(
                            "ChartActual", Optional.of("0.0"), List.of("12", "16"))))));
    ExcelChartSnapshot unsupportedSnapshot =
        chartSnapshot(
            "AreaOnly",
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(14, 3, 0, 0),
                new ExcelDrawingMarker(19, 14, 0, 0),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
            new ExcelChartSnapshot.Title.None(),
            new ExcelChartSnapshot.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            new ExcelChartSnapshot.Unsupported(
                "AREA", "Only simple single-plot charts are modeled."));

    ChartReport lineReport =
        InspectionResultDrawingReportSupport.toChartReport((ExcelChartSnapshot) lineSnapshot);
    ChartReport.Line linePlot =
        assertInstanceOf(ChartReport.Line.class, lineReport.plots().getFirst());
    ChartReport pieReport =
        InspectionResultDrawingReportSupport.toChartReport((ExcelChartSnapshot) pieSnapshot);
    ChartReport.Pie piePlot = assertInstanceOf(ChartReport.Pie.class, pieReport.plots().getFirst());
    ChartReport unsupportedReport =
        InspectionResultDrawingReportSupport.toChartReport(
            (ExcelChartSnapshot) unsupportedSnapshot);
    ChartReport.Unsupported unsupportedPlot =
        assertInstanceOf(ChartReport.Unsupported.class, unsupportedReport.plots().getFirst());

    assertTrue(lineReport.title() instanceof ChartReport.Title.None);
    assertTrue(lineReport.legend() instanceof ChartReport.Legend.Hidden);
    assertEquals(
        List.of("Jan", "Feb"),
        assertInstanceOf(
                ChartReport.DataSource.StringLiteral.class,
                linePlot.series().getFirst().categories())
            .values());
    assertEquals(
        Optional.of("0.0"),
        assertInstanceOf(
                ChartReport.DataSource.NumericLiteral.class, linePlot.series().getFirst().values())
            .formatCode());
    assertEquals(
        "Actual",
        assertInstanceOf(ChartReport.Title.Formula.class, pieReport.title()).cachedText());
    assertEquals(
        List.of("12", "16"),
        assertInstanceOf(
                ChartReport.DataSource.NumericReference.class, piePlot.series().getFirst().values())
            .cachedValues());
    assertEquals("AREA", unsupportedPlot.plotTypeToken());
  }

  private static dev.erst.gridgrind.excel.WorkbookDrawingResult.ChartsResult baseChartsResult() {
    return new dev.erst.gridgrind.excel.WorkbookDrawingResult.ChartsResult(
        "charts",
        "Budget",
        List.of(
            chartSnapshot(
                "OpsChart",
                new ExcelDrawingAnchor.TwoCell(
                    new ExcelDrawingMarker(1, 2, 0, 0),
                    new ExcelDrawingMarker(6, 12, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE),
                new ExcelChartSnapshot.Title.Text("Roadmap"),
                new ExcelChartSnapshot.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
                ExcelChartDisplayBlanksAs.SPAN,
                false,
                new ExcelChartSnapshot.Bar(
                    true,
                    ExcelChartBarDirection.COLUMN,
                    ExcelChartBarGrouping.CLUSTERED,
                    Optional.empty(),
                    Optional.empty(),
                    List.of(
                        new ExcelChartSnapshot.Axis(
                            ExcelChartAxisKind.CATEGORY,
                            ExcelChartAxisPosition.BOTTOM,
                            ExcelChartAxisCrosses.AUTO_ZERO,
                            true),
                        new ExcelChartSnapshot.Axis(
                            ExcelChartAxisKind.VALUE,
                            ExcelChartAxisPosition.LEFT,
                            ExcelChartAxisCrosses.AUTO_ZERO,
                            true)),
                    List.of(
                        chartSnapshotSeries(
                            new ExcelChartSnapshot.Title.Formula("Chart!$B$1", "Plan"),
                            new ExcelChartSnapshot.DataSource.StringReference(
                                "ChartCategories", List.of("Jan", "Feb", "Mar")),
                            new ExcelChartSnapshot.DataSource.NumericReference(
                                "ChartValues",
                                Optional.empty(),
                                List.of("10.0", "18.0", "15.0"))))))));
  }

  private static dev.erst.gridgrind.excel.WorkbookDrawingResult.ChartsResult
      advancedChartsResult() {
    return new dev.erst.gridgrind.excel.WorkbookDrawingResult.ChartsResult(
        "charts-advanced",
        "Budget",
        List.of(
            chartSnapshot(
                "TrendChart",
                new ExcelDrawingAnchor.TwoCell(
                    new ExcelDrawingMarker(2, 3, 0, 0),
                    new ExcelDrawingMarker(7, 14, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                new ExcelChartSnapshot.Title.None(),
                new ExcelChartSnapshot.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.GAP,
                true,
                new ExcelChartSnapshot.Line(
                    false,
                    ExcelChartGrouping.STANDARD,
                    List.of(
                        new ExcelChartSnapshot.Axis(
                            ExcelChartAxisKind.CATEGORY,
                            ExcelChartAxisPosition.BOTTOM,
                            ExcelChartAxisCrosses.AUTO_ZERO,
                            true),
                        new ExcelChartSnapshot.Axis(
                            ExcelChartAxisKind.VALUE,
                            ExcelChartAxisPosition.LEFT,
                            ExcelChartAxisCrosses.MIN,
                            true)),
                    List.of(
                        chartSnapshotSeries(
                            new ExcelChartSnapshot.Title.None(),
                            new ExcelChartSnapshot.DataSource.StringLiteral(List.of("Jan", "Feb")),
                            new ExcelChartSnapshot.DataSource.NumericLiteral(
                                Optional.of("0.0"), List.of("10", "18")))))),
            chartSnapshot(
                "ShareChart",
                new ExcelDrawingAnchor.TwoCell(
                    new ExcelDrawingMarker(8, 3, 0, 0),
                    new ExcelDrawingMarker(13, 14, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                new ExcelChartSnapshot.Title.Formula("Budget!$C$1", "Actual"),
                new ExcelChartSnapshot.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.ZERO,
                false,
                new ExcelChartSnapshot.Pie(
                    true,
                    Optional.of(120),
                    List.of(
                        chartSnapshotSeries(
                            new ExcelChartSnapshot.Title.Text("Actual"),
                            new ExcelChartSnapshot.DataSource.StringReference(
                                "ChartCategories", List.of("Jan", "Feb")),
                            new ExcelChartSnapshot.DataSource.NumericReference(
                                "ChartActual", Optional.of("0.0"), List.of("12", "16")))))),
            chartSnapshot(
                "AreaOnly",
                new ExcelDrawingAnchor.TwoCell(
                    new ExcelDrawingMarker(14, 3, 0, 0),
                    new ExcelDrawingMarker(19, 14, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                new ExcelChartSnapshot.Title.None(),
                new ExcelChartSnapshot.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.GAP,
                true,
                new ExcelChartSnapshot.Unsupported(
                    "AREA",
                    "Chart plot family is outside the current modeled simple-chart contract."))));
  }

  private static ExcelChartSnapshot chartSnapshot(
      String name,
      ExcelDrawingAnchor anchor,
      ExcelChartSnapshot.Title title,
      ExcelChartSnapshot.Legend legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      boolean plotOnlyVisibleCells,
      ExcelChartSnapshot.Plot plot) {
    return new ExcelChartSnapshot(
        name, anchor, title, legend, displayBlanksAs, plotOnlyVisibleCells, List.of(plot));
  }

  private static ExcelChartSnapshot.Series chartSnapshotSeries(
      ExcelChartSnapshot.Title title,
      ExcelChartSnapshot.DataSource categories,
      ExcelChartSnapshot.DataSource values) {
    return new ExcelChartSnapshot.Series(
        title,
        categories,
        values,
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }

  private static ExcelCellSnapshot advancedCell() {
    ExcelRichTextSnapshot richText =
        new ExcelRichTextSnapshot(List.of(new ExcelRichTextRunSnapshot("Styled", advancedFont())));
    return new ExcelCellSnapshot.TextSnapshot(
        "C3",
        "Styled",
        advancedStyle(),
        ExcelCellMetadataSnapshot.of(
            Optional.of(new ExcelHyperlink.Url("https://example.com/tasks/42")),
            Optional.of(richComment())),
        "Styled",
        richText);
  }

  private static ExcelCommentSnapshot richComment() {
    return new ExcelCommentSnapshot(
        "Hi there",
        "Ada",
        true,
        Optional.of(
            new ExcelRichTextSnapshot(
                List.of(
                    new ExcelRichTextRunSnapshot("Hi ", advancedFont()),
                    new ExcelRichTextRunSnapshot("there", accentFont())))),
        Optional.of(new ExcelCommentAnchorSnapshot(1, 2, 4, 6)));
  }

  private static ExcelCellStyleSnapshot advancedStyle() {
    return new ExcelCellStyleSnapshot(
        "0.00",
        new ExcelCellAlignmentSnapshot(
            false, ExcelHorizontalAlignment.CENTER, ExcelVerticalAlignment.CENTER, 0, 0),
        advancedFont(),
        ExcelCellFillSnapshot.gradient(
            ExcelGradientFillSnapshot.linear(
                45.0d,
                List.of(
                    new ExcelGradientStopSnapshot(0.0d, ExcelColorSnapshot.rgb("#112233")),
                    new ExcelGradientStopSnapshot(1.0d, ExcelColorSnapshot.theme(4, 0.45d))))),
        new ExcelBorderSnapshot(
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.THICK, ExcelColorSnapshot.indexed(12)),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null)),
        new ExcelCellProtectionSnapshot(true, false));
  }

  private static ExcelCellFontSnapshot advancedFont() {
    return new ExcelCellFontSnapshot(
        true,
        false,
        "Aptos",
        ExcelFontHeight.fromPoints(new BigDecimal("11")),
        ExcelColorSnapshot.theme(2, 0.25d),
        false,
        false);
  }

  private static ExcelCellFontSnapshot accentFont() {
    return new ExcelCellFontSnapshot(
        false,
        false,
        "Aptos",
        ExcelFontHeight.fromPoints(new BigDecimal("11")),
        ExcelColorSnapshot.rgb("#AABBCC"),
        false,
        false);
  }

  private static ExcelPrintLayoutSnapshot advancedPrintLayout() {
    return new ExcelPrintLayoutSnapshot(
        new ExcelPrintLayout(
            new ExcelPrintLayout.Area.Range("A1:F20"),
            ExcelPrintOrientation.LANDSCAPE,
            new ExcelPrintLayout.Scaling.Fit(1, 0),
            new ExcelPrintLayout.TitleRows.Band(0, 0),
            new ExcelPrintLayout.TitleColumns.None(),
            new ExcelHeaderFooterText("Ops", "Queue", ""),
            new ExcelHeaderFooterText("", "Internal", "Page &P")),
        new ExcelPrintSetupSnapshot(
            new ExcelPrintMarginsSnapshot(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
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
        new ExcelSheetPresentationSnapshot(
            new ExcelSheetDisplay(false, false, true, true, true),
            Optional.of(ExcelColorSnapshot.rgb("#112233")),
            new ExcelSheetOutlineSummary(false, false),
            new ExcelSheetDefaults(12, 18.5d),
            List.of(
                new ExcelIgnoredError(
                    "B2:B12",
                    List.of(
                        ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT,
                        ExcelIgnoredErrorType.FORMULA)))),
        List.of(
            new dev.erst.gridgrind.excel.WorkbookSheetResult.ColumnLayout(
                0, 16.0d, false, 0, false)),
        List.of(
            new dev.erst.gridgrind.excel.WorkbookSheetResult.RowLayout(0, 20.0d, false, 0, false)));
  }

  private static List<ExcelAutofilterSnapshot> advancedAutofilters() {
    ExcelAutofilterSortStateSnapshot sortState =
        new ExcelAutofilterSortStateSnapshot(
            "A1:F5",
            true,
            false,
            java.util.Optional.empty(),
            List.of(
                new ExcelAutofilterSortConditionSnapshot.CellColor(
                    "A2:A5", true, ExcelColorSnapshot.rgb("#AABBCC")),
                new ExcelAutofilterSortConditionSnapshot.FontColor(
                    "B2:B5", false, ExcelColorSnapshot.theme(4, 0.15d)),
                new ExcelAutofilterSortConditionSnapshot.Icon("C2:C5", false, 2)));
    List<ExcelAutofilterFilterColumnSnapshot> filterColumns =
        List.of(
            new ExcelAutofilterFilterColumnSnapshot(
                0L,
                false,
                new ExcelAutofilterFilterCriterionSnapshot.Values(List.of("Queued"), true)),
            new ExcelAutofilterFilterColumnSnapshot(
                1L,
                true,
                new ExcelAutofilterFilterCriterionSnapshot.Custom(
                    true,
                    List.of(
                        new ExcelAutofilterFilterCriterionSnapshot.CustomCondition(
                            "equal", "Ada")))),
            new ExcelAutofilterFilterColumnSnapshot(
                2L, true, new ExcelAutofilterFilterCriterionSnapshot.Dynamic("TODAY", 1.0d, 2.0d)),
            new ExcelAutofilterFilterColumnSnapshot(
                3L,
                true,
                new ExcelAutofilterFilterCriterionSnapshot.Top10(true, false, 10.0d, 8.0d)),
            new ExcelAutofilterFilterColumnSnapshot(
                4L,
                true,
                new ExcelAutofilterFilterCriterionSnapshot.Color(
                    false, ExcelColorSnapshot.theme(4, 0.45d))),
            new ExcelAutofilterFilterColumnSnapshot(
                5L, true, new ExcelAutofilterFilterCriterionSnapshot.Icon("3TrafficLights1", 2)));
    return List.of(
        new ExcelAutofilterSnapshot.SheetOwned(
            "A1:F5", filterColumns, java.util.Optional.of(sortState)),
        new ExcelAutofilterSnapshot.TableOwned(
            "H1:I5", "QueueTable", List.of(), java.util.Optional.of(sortState)));
  }

  private static ExcelTableSnapshot advancedTable() {
    return new ExcelTableSnapshot(
        "QueueTable",
        "Budget",
        "A1:B5",
        new ExcelTableSnapshot.Structure(
            1,
            1,
            List.of("Item", "Amount"),
            List.of(
                new ExcelTableColumnSnapshot(1L, "Item", "", "", "", ""),
                new ExcelTableColumnSnapshot(
                    2L, "Amount", "UniqueAmount", "Total", "sum", "[@Amount]*2"))),
        new ExcelTableStyleSnapshot.Named("TableStyleMedium2", false, false, true, false),
        new ExcelTableSnapshot.Behavior(true, true, true, false),
        new ExcelTableSnapshot.Presentation(
            Optional.of("Queue comment"),
            Optional.of("HeaderStyle"),
            Optional.of("DataStyle"),
            Optional.of("TotalsStyle")));
  }

  private static ExcelTableSnapshot normalizedOptionalTable() {
    return new ExcelTableSnapshot(
        "OptionalTable",
        "Budget",
        "D1:E2",
        new ExcelTableSnapshot.Structure(
            1,
            0,
            List.of("Code"),
            List.of(new ExcelTableColumnSnapshot(3L, "Code", null, " ", "", " "))),
        new ExcelTableStyleSnapshot.None(),
        new ExcelTableSnapshot.Behavior(false, false, false, false),
        new ExcelTableSnapshot.Presentation(
            Optional.of(" "), Optional.of(""), Optional.of(" "), Optional.empty()));
  }

  private static ExcelCustomXmlMappingSnapshot customXmlMappingSnapshot(
      long mapId,
      String name,
      String rootElement,
      String schemaId,
      ExcelCustomXmlMappingSettings settings,
      ExcelCustomXmlSchemaSnapshot schema,
      Optional<ExcelCustomXmlDataBindingSnapshot> dataBinding,
      List<ExcelCustomXmlLinkedCellSnapshot> linkedCells,
      List<dev.erst.gridgrind.excel.customxml.ExcelCustomXmlLinkedTableSnapshot> linkedTables) {
    return new ExcelCustomXmlMappingSnapshot(
        mapId,
        name,
        rootElement,
        schemaId,
        settings,
        schema,
        dataBinding,
        linkedCells,
        linkedTables);
  }

  private static ExcelCustomXmlMappingSettings settings(
      boolean showImportExportValidationErrors,
      boolean autoFit,
      boolean append,
      boolean preserveSortAfLayout,
      boolean preserveFormat) {
    return new ExcelCustomXmlMappingSettings(
        showImportExportValidationErrors, autoFit, append, preserveSortAfLayout, preserveFormat);
  }

  private static ExcelCustomXmlSchemaSnapshot schema(
      Optional<String> namespace,
      Optional<String> language,
      Optional<String> reference,
      Optional<String> xml) {
    return new ExcelCustomXmlSchemaSnapshot(namespace, language, reference, xml);
  }

  private static ExcelDifferentialStyleSnapshot differentialStyle() {
    return new ExcelDifferentialStyleSnapshot(
        "0.00", true, null, null, ExcelColor.rgb("#AABBCC"), null, null, null, null, List.of());
  }
}
