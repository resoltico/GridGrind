package dev.erst.gridgrind.core.model.security;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartDataSourceSpec;
import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ChartSeriesSpec;
import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.customxml.CustomXmlImportSpec;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.model.drawing.CommentAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.EmbeddedObjectSpec;
import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.drawing.PictureSpec;
import dev.erst.gridgrind.core.model.drawing.ShapeSpec;
import dev.erst.gridgrind.core.model.drawing.SignatureLineSpec;
import dev.erst.gridgrind.core.model.formula.ArrayFormulaSpec;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.layout.SheetDefaultsSpec;
import dev.erst.gridgrind.core.model.layout.SheetDisplaySpec;
import dev.erst.gridgrind.core.model.layout.SheetOutlineSummarySpec;
import dev.erst.gridgrind.core.model.layout.SheetPresentationSpec;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextSpec;
import dev.erst.gridgrind.core.model.print.PrintAreaSpec;
import dev.erst.gridgrind.core.model.print.PrintLayoutSpec;
import dev.erst.gridgrind.core.model.print.PrintMarginsSpec;
import dev.erst.gridgrind.core.model.print.PrintScalingSpec;
import dev.erst.gridgrind.core.model.print.PrintSetupSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionSpec;
import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.CellFillSpec;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellGradientFillSpec;
import dev.erst.gridgrind.core.model.style.CellGradientStopSpec;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnSpec;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateSpec;
import dev.erst.gridgrind.core.model.table.TableColumnSpec;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdSpec;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorSpec;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.SecretReference;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelIgnoredErrorType;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteHash;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.ProtocolCellAddressValidation;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct tests for advanced mutation-oriented protocol DTO families. */
class AdvancedMutationProtocolTypesTest {
  @Test
  @SuppressWarnings("NullOptional") // The factory deliberately normalizes absent JSON optionals.
  void ooxmlSecurityInputsNormalizeAndValidate() {
    OoxmlOpenSecuritySpec openSecurity = new OoxmlOpenSecuritySpec(Optional.of(ref("source-pass")));
    OoxmlEncryptionSpec encryption = OoxmlEncryptionSpec.strong(ref("persist-pass"));
    OoxmlSignatureSpec signature =
        new OoxmlSignatureSpec(
            "tmp/signing-material.p12",
            ref("keystore-pass"),
            Optional.of(ref("key-pass")),
            Optional.empty(),
            ExcelOoxmlSignatureDigestAlgorithm.SHA256,
            Optional.empty());
    OoxmlPersistenceSecuritySpec persistence =
        new OoxmlPersistenceSecuritySpec(
            new OoxmlPersistenceEncryptionSpec.Encrypt(encryption),
            new OoxmlPersistenceSignatureSpec.Sign(signature));

    assertEquals(Optional.of(ref("source-pass")), openSecurity.passwordRef());
    assertEquals(ExcelOoxmlWriteCipher.AES_256, encryption.cipher());
    assertEquals(ExcelOoxmlWriteHash.SHA_512, encryption.hash());
    assertEquals(Optional.of(ref("key-pass")), signature.keyPasswordRef());
    assertEquals(ExcelOoxmlSignatureDigestAlgorithm.SHA256, signature.digestAlgorithm());
    assertTrue(signature.alias().isEmpty());
    assertTrue(signature.description().isEmpty());
    assertEquals(
        encryption,
        assertInstanceOf(OoxmlPersistenceEncryptionSpec.Encrypt.class, persistence.encryption())
            .encryption());
    assertEquals(
        signature,
        assertInstanceOf(OoxmlPersistenceSignatureSpec.Sign.class, persistence.signature())
            .signature());

    assertThrows(IllegalArgumentException.class, () -> new SecretReference(" "));
    assertThrows(
        NullPointerException.class,
        () -> new OoxmlEncryptionSpec(ref("persist-pass"), null, ExcelOoxmlWriteHash.SHA_512));
    assertThrows(
        NullPointerException.class,
        () -> new OoxmlEncryptionSpec(ref("persist-pass"), ExcelOoxmlWriteCipher.AES_256, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new OoxmlSignatureSpec(
                "tmp/signing-material.p12",
                ref("keystore-pass"),
                Optional.empty(),
                Optional.of(" "),
                ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new OoxmlSignatureSpec(
                " ",
                ref("keystore-pass"),
                Optional.empty(),
                Optional.empty(),
                ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                Optional.empty()));

    assertThrows(NullPointerException.class, () -> new OoxmlPersistenceSecuritySpec(null, null));
  }

  @Test
  void ooxmlSecurityHelpersCollapseEmptyRequestState() {
    OoxmlOpenSecuritySpec openSecurity = new OoxmlOpenSecuritySpec(Optional.empty());
    OoxmlEncryptionSpec encryption =
        new OoxmlEncryptionSpec(
            ref("persist-pass"), ExcelOoxmlWriteCipher.AES_192, ExcelOoxmlWriteHash.SHA_384);
    OoxmlPersistenceSecuritySpec encryptionAndUnsigned =
        new OoxmlPersistenceSecuritySpec(
            new OoxmlPersistenceEncryptionSpec.Encrypt(encryption),
            new OoxmlPersistenceSignatureSpec.None());
    OoxmlPersistenceSecuritySpec plaintextAndSigned =
        new OoxmlPersistenceSecuritySpec(
            new OoxmlPersistenceEncryptionSpec.None(),
            new OoxmlPersistenceSignatureSpec.Sign(
                new OoxmlSignatureSpec(
                    "tmp/signing-material.p12",
                    ref("keystore-pass"),
                    Optional.of(ref("key-pass")),
                    Optional.empty(),
                    ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                    Optional.empty())));
    WorkbookPlan.WorkbookSource.ExistingFile source =
        new WorkbookPlan.WorkbookSource.ExistingFile("budget.xlsx", openSecurity);
    WorkbookPlan.WorkbookPersistence.Overwrite unsecuredOverwrite =
        new WorkbookPlan.WorkbookPersistence.Overwrite(Optional.empty());
    WorkbookPlan.WorkbookPersistence.Overwrite securedOverwrite =
        new WorkbookPlan.WorkbookPersistence.Overwrite(encryptionAndUnsigned);

    assertEquals(Optional.empty(), openSecurity.passwordRef());
    assertTrue(source.security().isEmpty());
    assertTrue(unsecuredOverwrite.security().isEmpty());
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookPlan.WorkbookPersistence.Overwrite((OoxmlPersistenceSecuritySpec) null));
    assertEquals(encryptionAndUnsigned, securedOverwrite.security().orElseThrow());
    assertEquals(
        plaintextAndSigned,
        new WorkbookPlan.WorkbookPersistence.SaveAs(
                "secured.xlsx",
                WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
                plaintextAndSigned)
            .security()
            .orElseThrow());
  }

  @Test
  void workbookProtectionNamedRangeAndCommentInputsNormalizeAndValidate() {
    WorkbookProtectionSpec protection =
        new WorkbookProtectionSpec(
            false, true, false, Optional.of(ref("book-secret")), Optional.of(ref("review-secret")));

    assertFalse(protection.structureLocked());
    assertTrue(protection.windowsLocked());
    assertFalse(protection.revisionsLocked());
    assertEquals(ref("book-secret"), protection.workbookPasswordRef().orElseThrow());
    assertEquals(ref("review-secret"), protection.revisionsPasswordRef().orElseThrow());

    NamedRangeTarget.Range explicit =
        NamedRangeTarget.range(new SheetName("Budget"), new RangeRef("C3:A1"));
    NamedRangeTarget.Formula formula = NamedRangeTarget.formula("SUM(Budget!A1:A3)");
    assertEquals("Budget", explicit.sheetName().value());
    assertEquals("C3:A1", explicit.range().value());
    assertEquals("SUM(Budget!A1:A3)", formula.formula());
    assertThrows(IllegalArgumentException.class, () -> NamedRangeTarget.formula(" "));
    assertThrows(
        IllegalArgumentException.class,
        () -> NamedRangeTarget.range(new SheetName("Budget"), new RangeRef(" ")));
    assertThrows(
        NullPointerException.class, () -> new NamedRangeTarget.Range(null, new RangeRef("A1")));

    CommentAnchorSpec anchor = new CommentAnchorSpec(1, 2, 4, 6);
    assertEquals(1, anchor.firstColumn());
    assertEquals(2, anchor.firstRow());
    assertEquals(4, anchor.lastColumn());
    assertEquals(6, anchor.lastRow());
    assertThrows(IllegalArgumentException.class, () -> new CommentAnchorSpec(-1, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new CommentAnchorSpec(1, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new CommentAnchorSpec(0, 2, 1, 1));

    CommentSpec richComment =
        new CommentSpec(
            text("Ada Lovelace"),
            "GridGrind",
            true,
            Optional.of(
                List.of(
                    new RichTextRunSpec(text("Ada"), Optional.empty()),
                    new RichTextRunSpec(text(" Lovelace"), Optional.empty()))),
            Optional.of(anchor));
    assertTrue(richComment.visible());
    assertEquals(anchor, richComment.anchor().orElseThrow());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CommentSpec(
                text("Ada"), "GridGrind", true, Optional.of(List.of()), Optional.empty()));
    assertThrows(
        NullPointerException.class,
        () ->
            new CommentSpec(
                text("Ada"),
                "GridGrind",
                true,
                Optional.of(List.of((RichTextRunSpec) null)),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CommentSpec(
                text("Ada Lovelace"),
                "GridGrind",
                true,
                Optional.of(List.of(new RichTextRunSpec(text("Ada"), Optional.empty()))),
                Optional.empty()));
  }

  private static SecretReference ref(String id) {
    return new SecretReference(id);
  }

  @Test
  void drawingInputsNormalizeAndValidate() {
    DrawingMarkerSpec from = new DrawingMarkerSpec(1, 2, 3, 4);
    DrawingMarkerSpec to = new DrawingMarkerSpec(4, 6);
    DrawingAnchorSpec.TwoCell anchor =
        new DrawingAnchorSpec.TwoCell(from, to, ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
    DrawingAnchorSpec.TwoCell defaultAnchor = DrawingAnchorSpec.TwoCell.moveAndResize(from, to);
    PictureDataSpec pictureData =
        new PictureDataSpec(
            ExcelPictureFormat.PNG,
            binary(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC"));
    PictureSpec picture =
        new PictureSpec("OpsPicture", pictureData, anchor, Optional.of(text("Queue preview")));
    ShapeSpec.SimpleShape shape =
        new ShapeSpec.SimpleShape("OpsShape", anchor, " rect ", Optional.of(text("Queue")));
    ShapeSpec.Connector connector = new ShapeSpec.Connector("OpsConnector", anchor);
    EmbeddedObjectSpec embeddedObject =
        new EmbeddedObjectSpec(
            "OpsEmbed",
            "Payload",
            "payload.txt",
            "payload.txt",
            binary("cGF5bG9hZA=="),
            pictureData,
            anchor);
    SignatureLineSpec signatureLine =
        new SignatureLineSpec(
            "OpsSignature",
            anchor,
            true,
            java.util.Optional.of("Review before signing."),
            java.util.Optional.of("Ada Lovelace"),
            java.util.Optional.of("Finance"),
            java.util.Optional.of("ada@example.com"),
            java.util.Optional.empty(),
            java.util.Optional.of("invalid"),
            java.util.Optional.of(pictureData));

    DrawingAnchorSpec.TwoCell sameRowAnchor =
        new DrawingAnchorSpec.TwoCell(
            new DrawingMarkerSpec(1, 2, 0, 3),
            new DrawingMarkerSpec(2, 2, 0, 4),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    DrawingAnchorSpec.TwoCell sameColAnchor =
        new DrawingAnchorSpec.TwoCell(
            new DrawingMarkerSpec(1, 2, 3, 0),
            new DrawingMarkerSpec(1, 3, 4, 0),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);

    assertEquals(4, to.columnIndex());
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE, anchor.behavior());
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE, defaultAnchor.behavior());
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE, sameRowAnchor.behavior());
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE, sameColAnchor.behavior());
    assertEquals(ExcelPictureFormat.PNG, picture.image().format());
    assertEquals(text("Queue preview"), picture.description().orElseThrow());
    assertEquals("rect", shape.presetGeometryToken());
    assertEquals(ExcelAuthoredDrawingShapeKind.CONNECTOR, connector.kind());
    assertEquals("payload.txt", embeddedObject.fileName());
    assertTrue(signatureLine.allowComments());
    assertEquals("Ada Lovelace", signatureLine.suggestedSigner().orElseThrow());
    assertThrows(
        IllegalArgumentException.class,
        () -> new ShapeSpec.SimpleShape("DefaultShape", anchor, " ", Optional.empty()));
    assertThrows(
        NullPointerException.class,
        () -> new ShapeSpec.SimpleShape("NullPreset", anchor, null, Optional.empty()));
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerSpec(-1, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerSpec(0, -1, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerSpec(0, 0, -1, 0));
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerSpec(0, 0, 0, -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorSpec.TwoCell(
                new DrawingMarkerSpec(1, 2, 0, 0),
                new DrawingMarkerSpec(1, 1, 0, 0),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorSpec.TwoCell(
                new DrawingMarkerSpec(2, 2, 0, 0),
                new DrawingMarkerSpec(1, 2, 0, 0),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorSpec.TwoCell(
                new DrawingMarkerSpec(1, 2, 0, 4),
                new DrawingMarkerSpec(2, 2, 0, 3),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorSpec.TwoCell(
                new DrawingMarkerSpec(1, 2, 4, 0),
                new DrawingMarkerSpec(1, 3, 3, 0),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertThrows(NullPointerException.class, () -> new DrawingAnchorSpec.TwoCell(from, to, null));
    assertThrows(
        NullPointerException.class, () -> new PictureDataSpec(null, binary("cGF5bG9hZA==")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PictureDataSpec(ExcelPictureFormat.PNG, binary(" ")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PictureDataSpec(ExcelPictureFormat.PNG, binary("not-base64")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PictureSpec(" ", pictureData, anchor, Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PictureSpec("OpsPicture", pictureData, anchor, Optional.of(text(" "))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ShapeSpec.SimpleShape(" ", anchor, "rect", Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ShapeSpec.SimpleShape("OpsShape", anchor, "rect", Optional.of(text(" "))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new EmbeddedObjectSpec(
                "OpsEmbed",
                "Payload",
                "payload.txt",
                "payload.txt",
                binary("not-base64"),
                pictureData,
                anchor));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new EmbeddedObjectSpec(
                "OpsEmbed",
                "Payload",
                "payload.txt",
                " ",
                binary("cGF5bG9hZA=="),
                pictureData,
                anchor));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new SignatureLineSpec(
                "OpsSignature",
                anchor,
                false,
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new SignatureLineSpec(
                "OpsSignature",
                anchor,
                false,
                java.util.Optional.of("Review before signing."),
                java.util.Optional.of("Ada Lovelace"),
                java.util.Optional.of("Finance"),
                java.util.Optional.of("ada@example.com"),
                java.util.Optional.of("line1\nline2\nline3\nline4"),
                java.util.Optional.empty(),
                java.util.Optional.of(pictureData)));
  }

  @Test
  void arrayFormulaInputsNormalizeInlineWrappers() {
    ArrayFormulaSpec wrapped = new ArrayFormulaSpec(text("{=SUM(B2:B4*C2:C4)}"));
    ArrayFormulaSpec bare = new ArrayFormulaSpec(text("SUM(B2:B4*C2:C4)"));
    ArrayFormulaSpec prefixed = new ArrayFormulaSpec(text("=SUM(B2:B4*C2:C4)"));

    assertEquals(
        "SUM(B2:B4*C2:C4)", assertInstanceOf(TextSourceSpec.Inline.class, wrapped.source()).text());
    assertEquals(
        "SUM(B2:B4*C2:C4)", assertInstanceOf(TextSourceSpec.Inline.class, bare.source()).text());
    assertEquals(
        "SUM(B2:B4*C2:C4)",
        assertInstanceOf(TextSourceSpec.Inline.class, prefixed.source()).text());

    assertThrows(IllegalArgumentException.class, () -> new ArrayFormulaSpec(text(" ")));
    assertThrows(IllegalArgumentException.class, () -> new ArrayFormulaSpec(text("{=}")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ArrayFormulaSpec(text("DDE(\"cmd\",\"/C calc\",\"\")")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ArrayFormulaSpec(text("=dde(\"cmd\",\"/C calc\",\"\")")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ArrayFormulaSpec(text("{=DDE(\"cmd\",\"/C calc\",\"\")}")));
  }

  @Test
  void customXmlInputsNormalizeAndValidate() {
    CustomXmlMappingLocator byId =
        new CustomXmlMappingLocator(java.util.Optional.of(1L), java.util.Optional.empty());
    CustomXmlMappingLocator byName =
        new CustomXmlMappingLocator(
            java.util.Optional.empty(), java.util.Optional.of("CORSO_mapping"));
    CustomXmlImportSpec inlineImport =
        new CustomXmlImportSpec(byName, text("<CORSO><NOME>Grid</NOME></CORSO>"));

    assertEquals(1L, byId.mapId().orElse(null));
    assertEquals("CORSO_mapping", byName.name().orElse(null));
    assertEquals(byName, inlineImport.locator());
    assertEquals(
        "<CORSO><NOME>Grid</NOME></CORSO>",
        assertInstanceOf(TextSourceSpec.Inline.class, inlineImport.xml()).text());

    assertThrows(
        IllegalArgumentException.class,
        () -> new CustomXmlMappingLocator(java.util.Optional.empty(), java.util.Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CustomXmlMappingLocator(java.util.Optional.of(0L), java.util.Optional.empty()));
    assertThrows(IllegalArgumentException.class, () -> new CustomXmlImportSpec(byId, text(" ")));
  }

  @Test
  void chartInputsNormalizeAndValidate() {
    DrawingAnchorSpec.TwoCell anchor =
        DrawingAnchorSpec.TwoCell.moveAndResize(
            new DrawingMarkerSpec(1, 2, 0, 0), new DrawingMarkerSpec(6, 12, 0, 0));
    ChartSeriesSpec firstSeries = chartSeries(new ChartTitleSpec.Formula("B1"), "A2:A4", "B2:B4");
    ChartSeriesSpec secondSeries = chartSeries(null, "ChartCategories", "ChartActual");
    ChartSpec bar =
        chartInput(
            "OpsChart",
            anchor,
            null,
            null,
            null,
            null,
            new ChartPlotSpec.Bar(
                false,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.CLUSTERED,
                Optional.empty(),
                Optional.empty(),
                List.of(firstSeries)));
    ChartSpec line =
        chartInput(
            "TrendChart",
            anchor,
            new ChartTitleSpec.Text(text("Trend")),
            new ChartLegendSpec.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            false,
            new ChartPlotSpec.Line(true, ExcelChartGrouping.STANDARD, List.of(secondSeries)));
    ChartSpec defaultLine =
        chartInput(
            "DefaultTrend",
            anchor,
            null,
            null,
            null,
            null,
            new ChartPlotSpec.Line(false, ExcelChartGrouping.STANDARD, List.of(secondSeries)));
    ChartSpec pie =
        chartInput(
            "ShareChart",
            anchor,
            null,
            null,
            null,
            null,
            new ChartPlotSpec.Pie(false, Optional.of(180), List.of(secondSeries)));
    ChartSpec defaultPie =
        chartInput(
            "DefaultShare",
            anchor,
            null,
            null,
            null,
            null,
            new ChartPlotSpec.Pie(false, Optional.empty(), List.of(secondSeries)));
    ChartSpec explicitPie =
        chartInput(
            "ExplicitShare",
            anchor,
            new ChartTitleSpec.Text(text("Share")),
            new ChartLegendSpec.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            false,
            new ChartPlotSpec.Pie(true, Optional.of(90), List.of(secondSeries)));

    ChartPlotSpec.Bar barPlot = assertInstanceOf(ChartPlotSpec.Bar.class, bar.plots().getFirst());
    ChartPlotSpec.Line linePlot =
        assertInstanceOf(ChartPlotSpec.Line.class, line.plots().getFirst());
    ChartPlotSpec.Line defaultLinePlot =
        assertInstanceOf(ChartPlotSpec.Line.class, defaultLine.plots().getFirst());
    ChartPlotSpec.Pie piePlot = assertInstanceOf(ChartPlotSpec.Pie.class, pie.plots().getFirst());
    ChartPlotSpec.Pie defaultPiePlot =
        assertInstanceOf(ChartPlotSpec.Pie.class, defaultPie.plots().getFirst());
    ChartPlotSpec.Pie explicitPiePlot =
        assertInstanceOf(ChartPlotSpec.Pie.class, explicitPie.plots().getFirst());

    assertTrue(bar.title() instanceof ChartTitleSpec.None);
    assertEquals(new ChartLegendSpec.Visible(ExcelChartLegendPosition.RIGHT), bar.legend());
    assertEquals(ExcelChartDisplayBlanksAs.GAP, bar.displayBlanksAs());
    assertTrue(bar.plotOnlyVisibleCells());
    assertFalse(barPlot.varyColors());
    assertEquals(ExcelChartBarDirection.COLUMN, barPlot.barDirection());
    assertEquals("B1", ((ChartTitleSpec.Formula) barPlot.series().getFirst().title()).formula());
    assertTrue(secondSeries.title() instanceof ChartTitleSpec.None);
    assertEquals(ExcelChartDisplayBlanksAs.ZERO, line.displayBlanksAs());
    assertFalse(line.plotOnlyVisibleCells());
    assertTrue(linePlot.varyColors());
    assertEquals(Optional.of(180), piePlot.firstSliceAngle());
    assertTrue(defaultLine.title() instanceof ChartTitleSpec.None);
    assertEquals(new ChartLegendSpec.Visible(ExcelChartLegendPosition.RIGHT), defaultLine.legend());
    assertEquals(ExcelChartDisplayBlanksAs.GAP, defaultLine.displayBlanksAs());
    assertTrue(defaultLine.plotOnlyVisibleCells());
    assertFalse(defaultLinePlot.varyColors());
    assertTrue(defaultPie.title() instanceof ChartTitleSpec.None);
    assertEquals(new ChartLegendSpec.Visible(ExcelChartLegendPosition.RIGHT), defaultPie.legend());
    assertEquals(ExcelChartDisplayBlanksAs.GAP, defaultPie.displayBlanksAs());
    assertTrue(defaultPie.plotOnlyVisibleCells());
    assertFalse(defaultPiePlot.varyColors());
    assertEquals(Optional.empty(), defaultPiePlot.firstSliceAngle());
    assertEquals(new ChartTitleSpec.Text(text("Share")), explicitPie.title());
    assertEquals(new ChartLegendSpec.Hidden(), explicitPie.legend());
    assertEquals(ExcelChartDisplayBlanksAs.ZERO, explicitPie.displayBlanksAs());
    assertFalse(explicitPie.plotOnlyVisibleCells());
    assertTrue(explicitPiePlot.varyColors());
    assertEquals(Optional.of(90), explicitPiePlot.firstSliceAngle());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            chartInput(
                " ",
                anchor,
                null,
                null,
                null,
                null,
                new ChartPlotSpec.Bar(
                    false,
                    ExcelChartBarDirection.COLUMN,
                    ExcelChartBarGrouping.CLUSTERED,
                    Optional.empty(),
                    Optional.empty(),
                    List.of(firstSeries))));
    assertThrows(
        NullPointerException.class,
        () ->
            chartInput(
                "OpsChart",
                null,
                null,
                null,
                null,
                null,
                new ChartPlotSpec.Line(false, ExcelChartGrouping.STANDARD, List.of(firstSeries))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartPlotSpec.Pie(false, Optional.of(361), List.of(firstSeries)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartPlotSpec.Pie(false, Optional.of(-1), List.of(firstSeries)));
    assertThrows(IllegalArgumentException.class, () -> new ChartTitleSpec.Text(text(" ")));
    assertThrows(IllegalArgumentException.class, () -> new ChartTitleSpec.Formula(" "));
    assertThrows(NullPointerException.class, () -> new ChartLegendSpec.Visible(null));
    assertThrows(IllegalArgumentException.class, () -> new ChartDataSourceSpec.Reference(" "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartDataSourceSpec.Reference("DDE(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartDataSourceSpec.Reference("=dde(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartPlotSpec.Bar(
                false,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.CLUSTERED,
                Optional.empty(),
                Optional.empty(),
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ChartPlotSpec.Line(false, ExcelChartGrouping.STANDARD, List.of(firstSeries, null)));
  }

  @Test
  void autofilterInputsNormalizeAndValidateAcrossAllCriterionFamilies() {
    AutofilterSortConditionSpec colorSort =
        new AutofilterSortConditionSpec.CellColor(
            new RangeRef("B2:B9"), true, ColorSpec.rgb(new RgbHex("#aabbcc")));
    AutofilterSortConditionSpec iconSort =
        new AutofilterSortConditionSpec.Icon(new RangeRef("C2:C9"), false, 2);
    AutofilterSortStateSpec sortState =
        AutofilterSortStateSpec.withoutSortMethod(
            new RangeRef("A1:F9"), false, true, List.of(colorSort, iconSort));
    AutofilterSortStateSpec sortStateWithMethod =
        new AutofilterSortStateSpec(
            new RangeRef("A1:F9"),
            true,
            false,
            Optional.of(dev.erst.gridgrind.core.vocabulary.ExcelAutofilterSortMethod.PINYIN),
            List.of(colorSort));

    AutofilterSortConditionSpec.CellColor typedColorSort =
        assertInstanceOf(AutofilterSortConditionSpec.CellColor.class, colorSort);
    AutofilterSortConditionSpec.Icon typedIconSort =
        assertInstanceOf(AutofilterSortConditionSpec.Icon.class, iconSort);
    assertEquals(
        "#AABBCC", assertInstanceOf(ColorSpec.Rgb.class, typedColorSort.color()).rgb().value());
    assertEquals(2, typedIconSort.iconId());
    assertFalse(sortState.caseSensitive());
    assertTrue(sortState.columnSort());
    assertEquals(Optional.empty(), sortState.sortMethod());
    assertEquals(List.of(colorSort, iconSort), sortState.conditions());
    assertEquals(
        Optional.of(dev.erst.gridgrind.core.vocabulary.ExcelAutofilterSortMethod.PINYIN),
        sortStateWithMethod.sortMethod());
    assertEquals(
        new AutofilterFilterCriterionSpec.Dynamic("TODAY", Optional.empty(), Optional.empty()),
        new AutofilterFilterCriterionSpec.Dynamic("TODAY", Optional.empty(), Optional.empty()));
    assertThrows(
        NullPointerException.class, () -> new AutofilterSortConditionSpec.Value(null, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            AutofilterSortStateSpec.withoutSortMethod(
                new RangeRef(" "), false, false, List.of(colorSort)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterSortConditionSpec.Value(new RangeRef(" "), false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterSortConditionSpec.Icon(new RangeRef("A1:A2"), false, -1));
    assertThrows(
        NullPointerException.class,
        () ->
            AutofilterSortStateSpec.withoutSortMethod(
                new RangeRef("A1:F9"), false, false, (List<AutofilterSortConditionSpec>) null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            AutofilterSortStateSpec.withoutSortMethod(
                new RangeRef("A1:F9"), false, false, List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            AutofilterSortStateSpec.withoutSortMethod(
                new RangeRef("A1:F9"), false, false, List.of((AutofilterSortConditionSpec) null)));

    AutofilterFilterCriterionSpec.Values values =
        new AutofilterFilterCriterionSpec.Values(List.of("Queued", "Ready"), true);
    AutofilterFilterCriterionSpec.Custom custom =
        new AutofilterFilterCriterionSpec.Custom(
            true, List.of(new AutofilterFilterCriterionSpec.CustomConditionSpec("equal", "Ada")));
    AutofilterFilterCriterionSpec.Dynamic dynamic =
        new AutofilterFilterCriterionSpec.Dynamic("TODAY", Optional.of(1.0d), Optional.of(2.0d));
    AutofilterFilterCriterionSpec.Top10 top10 =
        new AutofilterFilterCriterionSpec.Top10(10, true, false);
    AutofilterFilterCriterionSpec.Color color =
        new AutofilterFilterCriterionSpec.Color(false, ColorSpec.theme(3, 0.25d));
    AutofilterFilterCriterionSpec.Icon icon =
        new AutofilterFilterCriterionSpec.Icon("3TrafficLights1", 2);
    AutofilterFilterColumnSpec column = AutofilterFilterColumnSpec.visibleButton(2L, values);

    assertEquals(List.of("Queued", "Ready"), values.values());
    assertTrue(custom.and());
    assertEquals("TODAY", dynamic.dynamicType());
    assertEquals(Optional.of(1.0d), dynamic.value());
    assertEquals(Optional.of(2.0d), dynamic.maxValue());
    assertEquals(10, top10.value());
    assertEquals(3, assertInstanceOf(ColorSpec.Theme.class, color.color()).theme());
    assertEquals("3TrafficLights1", icon.iconSet());
    assertTrue(column.showButton());
    assertEquals(values, column.criterion());

    assertThrows(
        NullPointerException.class, () -> new AutofilterFilterCriterionSpec.Values(null, false));
    assertThrows(
        NullPointerException.class,
        () -> new AutofilterFilterCriterionSpec.Values(List.of("Queued", null), false));
    assertThrows(
        NullPointerException.class, () -> new AutofilterFilterCriterionSpec.Custom(true, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionSpec.Custom(true, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionSpec.CustomConditionSpec(" ", "Ada"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionSpec.CustomConditionSpec("equal", " "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionSpec.Dynamic(" ", Optional.of(1.0d), Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new AutofilterFilterCriterionSpec.Dynamic(
                "TODAY", Optional.of(Double.POSITIVE_INFINITY), Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new AutofilterFilterCriterionSpec.Dynamic(
                "TODAY", Optional.of(1.0d), Optional.of(Double.NaN)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionSpec.Top10(0, true, false));
    assertThrows(
        NullPointerException.class, () -> new AutofilterFilterCriterionSpec.Color(true, null));
    assertThrows(
        IllegalArgumentException.class, () -> new AutofilterFilterCriterionSpec.Icon(" ", 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionSpec.Icon("3TrafficLights1", -1));
    assertThrows(
        IllegalArgumentException.class, () -> new AutofilterFilterColumnSpec(-1L, true, values));
    assertThrows(NullPointerException.class, () -> new AutofilterFilterColumnSpec(0L, false, null));
  }

  @Test
  void colorAndThresholdInputsNormalizeAndValidate() {
    ColorSpec.Rgb rgb = ColorSpec.rgb(new RgbHex("#aabbcc"));
    ColorSpec.Rgb tintedRgb = ColorSpec.rgb(new RgbHex("#112233"), 0.25d);
    ColorSpec.Theme themed = ColorSpec.theme(4, 0.5d);
    ColorSpec.Indexed indexed = ColorSpec.indexed(64);
    assertEquals("#AABBCC", rgb.rgb().value());
    assertEquals("#112233", tintedRgb.rgb().value());
    assertEquals(Optional.of(0.25d), tintedRgb.tint());
    assertEquals(4, themed.theme());
    assertEquals(Optional.of(0.5d), themed.tint());
    assertEquals(64, indexed.indexed());
    assertThrows(NullPointerException.class, () -> ColorSpec.rgb(null));
    assertThrows(IllegalArgumentException.class, () -> ColorSpec.rgb(new RgbHex(" ")));
    assertThrows(IllegalArgumentException.class, () -> ColorSpec.rgb(new RgbHex("123456")));
    assertThrows(IllegalArgumentException.class, () -> ColorSpec.theme(-1));
    assertThrows(IllegalArgumentException.class, () -> ColorSpec.indexed(-1));
    assertThrows(
        IllegalArgumentException.class,
        () -> ColorSpec.rgb(new RgbHex("#AABBCC"), Double.NEGATIVE_INFINITY));

    ConditionalFormattingThresholdSpec.Formula threshold =
        new ConditionalFormattingThresholdSpec.Formula("A1*2");
    assertEquals("A1*2", threshold.formula());
    assertEquals(90.0d, new ConditionalFormattingThresholdSpec.Percentile(90.0d).value());
    assertEquals(0.0d, new ConditionalFormattingThresholdSpec.Percent(0.0d).value());
    assertEquals(100.0d, new ConditionalFormattingThresholdSpec.Percent(100.0d).value());
    assertEquals(0.0d, new ConditionalFormattingThresholdSpec.Percentile(0.0d).value());
    assertEquals(100.0d, new ConditionalFormattingThresholdSpec.Percentile(100.0d).value());
    assertThrows(
        IllegalArgumentException.class, () -> new ConditionalFormattingThresholdSpec.Formula(null));
    assertThrows(
        IllegalArgumentException.class, () -> new ConditionalFormattingThresholdSpec.Formula(" "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ConditionalFormattingThresholdSpec.Numeric(Double.NEGATIVE_INFINITY));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ConditionalFormattingThresholdSpec.Percent(Double.NaN));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ConditionalFormattingThresholdSpec.Percentile(Double.NaN));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ConditionalFormattingThresholdSpec.Percent(-0.1d));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ConditionalFormattingThresholdSpec.Percentile(100.1d));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ConditionalFormattingThresholdSpec.Formula("DDE(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleSpec.ColorScaleRule(
                false,
                List.of(new ConditionalFormattingThresholdSpec.Min()),
                List.of(ColorSpec.rgb(new RgbHex("#112233")))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleSpec.ColorScaleRule(
                false,
                List.of(
                    new ConditionalFormattingThresholdSpec.Min(),
                    new ConditionalFormattingThresholdSpec.Max()),
                List.of(ColorSpec.rgb(new RgbHex("#112233")))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleSpec.DataBarRule(
                false,
                ColorSpec.rgb(new RgbHex("#112233")),
                false,
                -1,
                90,
                new ConditionalFormattingThresholdSpec.Numeric(0.0d),
                new ConditionalFormattingThresholdSpec.Numeric(1.0d)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleSpec.DataBarRule(
                false,
                ColorSpec.rgb(new RgbHex("#112233")),
                false,
                0,
                -1,
                new ConditionalFormattingThresholdSpec.Numeric(0.0d),
                new ConditionalFormattingThresholdSpec.Numeric(1.0d)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleSpec.DataBarRule(
                false,
                ColorSpec.rgb(new RgbHex("#112233")),
                false,
                10,
                5,
                new ConditionalFormattingThresholdSpec.Numeric(0.0d),
                new ConditionalFormattingThresholdSpec.Numeric(1.0d)));
    assertThrows(
        NullPointerException.class,
        () -> new ConditionalFormattingRuleSpec.IconSetRule(false, null, false, false, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleSpec.IconSetRule(
                false,
                ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
                false,
                false,
                List.of(
                    new ConditionalFormattingThresholdSpec.Percent(0.0d),
                    new ConditionalFormattingThresholdSpec.Percent(50.0d))));
    assertThrows(
        NullPointerException.class,
        () -> new ConditionalFormattingRuleSpec.Top10Rule(false, 1, true, false, null));
  }

  @Test
  void gradientAndStructuredStyleInputsNormalizeAndValidate() {
    CellGradientStopSpec start =
        new CellGradientStopSpec(0.0d, ColorSpec.rgb(new RgbHex("#112233")));
    CellGradientStopSpec finish = new CellGradientStopSpec(1.0d, ColorSpec.theme(5, 0.2d));
    CellGradientFillSpec gradient =
        CellGradientFillSpec.path(
            Optional.of(0.1d),
            Optional.of(0.2d),
            Optional.of(0.3d),
            Optional.of(0.4d),
            List.of(start, finish));
    assertGradientInputsNormalizeAndValidate(start, finish, gradient);
    assertFontInputsNormalizeAndValidate();
    assertFillInputsNormalizeAndValidate(gradient);
    assertBorderSideInputsNormalizeAndValidate();
  }

  private void assertGradientInputsNormalizeAndValidate(
      CellGradientStopSpec start, CellGradientStopSpec finish, CellGradientFillSpec gradient) {
    CellGradientFillSpec.Path path = assertInstanceOf(CellGradientFillSpec.Path.class, gradient);
    assertEquals(Optional.of(0.1d), path.left());
    assertEquals(Optional.of(0.2d), path.right());
    assertEquals(Optional.of(0.3d), path.top());
    assertEquals(Optional.of(0.4d), path.bottom());
    assertEquals(List.of(start, finish), path.stops());
    assertInstanceOf(
        CellGradientFillSpec.Linear.class,
        CellGradientFillSpec.linear(Optional.empty(), List.of(start, finish)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientStopSpec(-0.1d, ColorSpec.rgb(new RgbHex("#112233"))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientStopSpec(Double.NaN, ColorSpec.rgb(new RgbHex("#112233"))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientStopSpec(1.5d, ColorSpec.rgb(new RgbHex("#112233"))));
    assertThrows(NullPointerException.class, () -> new CellGradientStopSpec(0.5d, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellGradientFillSpec.linear(
                Optional.of(Double.POSITIVE_INFINITY), List.of(start, finish)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellGradientFillSpec.path(
                Optional.of(Double.NaN),
                Optional.of(0.2d),
                Optional.of(0.3d),
                Optional.of(0.4d),
                List.of(start, finish)));
    assertThrows(
        IllegalArgumentException.class,
        () -> CellGradientFillSpec.linear(Optional.empty(), List.of(start)));
    assertThrows(
        NullPointerException.class,
        () -> CellGradientFillSpec.linear(Optional.empty(), List.of(start, null)));
  }

  private void assertFontInputsNormalizeAndValidate() {
    CellFontPatch themedFont =
        new CellFontPatch(
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.of(ColorSpec.theme(2, 0.4d)),
            Optional.of(true),
            Optional.empty());
    CellFontPatch namedFont =
        new CellFontPatch(
            Optional.empty(),
            Optional.empty(),
            Optional.of("Calibri"),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    CellFontPatch indexedFont =
        new CellFontPatch(
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.of(ColorSpec.indexed(64)),
            Optional.empty(),
            Optional.empty());
    assertThemeColor(themedFont.fontColor().orElseThrow(), 2, 0.4d);
    assertTrue(themedFont.underline().orElseThrow());
    assertEquals("Calibri", namedFont.fontName().orElseThrow());
    assertIndexedColor(indexedFont.fontColor().orElseThrow(), 64, null);
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFontPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.of(" "),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFontPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(IllegalArgumentException.class, () -> ColorSpec.theme(-1));
    assertThrows(IllegalArgumentException.class, () -> ColorSpec.indexed(-1));
    assertThrows(
        IllegalArgumentException.class,
        () -> ColorSpec.rgb(new RgbHex("#112233"), Double.POSITIVE_INFINITY));
  }

  private void assertFillInputsNormalizeAndValidate(CellGradientFillSpec gradient) {
    CellFillSpec.Gradient gradientFill =
        assertInstanceOf(CellFillSpec.Gradient.class, CellFillSpec.gradient(gradient));
    CellFillSpec.PatternForeground themedFill =
        assertInstanceOf(
            CellFillSpec.PatternForeground.class,
            CellFillSpec.patternForeground(ExcelFillPattern.SOLID, ColorSpec.theme(2, 0.3d)));
    CellFillSpec.PatternForeground fgIndexedFill =
        assertInstanceOf(
            CellFillSpec.PatternForeground.class,
            CellFillSpec.patternForeground(ExcelFillPattern.SOLID, ColorSpec.indexed(64)));
    CellFillSpec.PatternForeground fgIndexedTintedFill =
        assertInstanceOf(
            CellFillSpec.PatternForeground.class,
            CellFillSpec.patternForeground(ExcelFillPattern.SOLID, ColorSpec.indexed(64, 0.3d)));
    CellFillSpec.PatternBackground bgIndexedFill =
        assertInstanceOf(
            CellFillSpec.PatternBackground.class,
            CellFillSpec.patternBackground(ExcelFillPattern.FINE_DOTS, ColorSpec.indexed(64)));
    assertEquals(gradient, gradientFill.gradient());
    assertThemeColor(themedFill.foregroundColor(), 2, 0.3d);
    assertIndexedColor(fgIndexedFill.foregroundColor(), 64, null);
    assertIndexedColor(fgIndexedTintedFill.foregroundColor(), 64, 0.3d);
    assertIndexedColor(bgIndexedFill.backgroundColor(), 64, null);
    assertThrows(NullPointerException.class, () -> CellFillSpec.gradient(null));
    assertThrows(NullPointerException.class, () -> CellFillSpec.pattern(null));
    assertThrows(
        NullPointerException.class,
        () -> CellFillSpec.patternForeground(ExcelFillPattern.SOLID, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> CellFillSpec.patternForeground(ExcelFillPattern.NONE, ColorSpec.theme(2)));
    assertThrows(
        IllegalArgumentException.class,
        () -> CellFillSpec.patternBackground(ExcelFillPattern.SOLID, ColorSpec.indexed(64)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellFillSpec.patternColors(
                ExcelFillPattern.SOLID,
                ColorSpec.rgb(new RgbHex("#112233")),
                ColorSpec.indexed(64)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellFillSpec.patternColors(
                ExcelFillPattern.NONE,
                ColorSpec.rgb(new RgbHex("#112233")),
                ColorSpec.indexed(64)));
    assertThrows(
        NullPointerException.class,
        () -> CellFillSpec.patternBackground(ExcelFillPattern.FINE_DOTS, null));
    assertThrows(
        NullPointerException.class,
        () ->
            CellFillSpec.patternColors(
                ExcelFillPattern.FINE_DOTS, ColorSpec.rgb(new RgbHex("#112233")), null));
  }

  private void assertBorderSideInputsNormalizeAndValidate() {
    BorderSidePatch noneStyleOnly = new BorderSidePatch(ExcelBorderStyle.NONE);
    BorderSidePatch borderSide = new BorderSidePatch(null, ColorSpec.theme(1, 0.15d));
    BorderSidePatch indexedBorderSide = new BorderSidePatch(null, ColorSpec.indexed(64));
    assertEquals(Optional.of(ExcelBorderStyle.NONE), noneStyleOnly.style());
    assertThemeColor(borderSide.color().orElseThrow(), 1, 0.15d);
    assertIndexedColor(indexedBorderSide.color().orElseThrow(), 64, null);
    assertThrows(
        IllegalArgumentException.class,
        () -> new BorderSidePatch(Optional.empty(), Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new BorderSidePatch(ExcelBorderStyle.NONE, ColorSpec.theme(1)));
    assertThrows(IllegalArgumentException.class, () -> ColorSpec.theme(-1));
    assertThrows(IllegalArgumentException.class, () -> ColorSpec.indexed(-1));
    assertThrows(
        IllegalArgumentException.class, () -> ColorSpec.theme(1, Double.POSITIVE_INFINITY));
  }

  @Test
  void printAndTableInputsNormalizeAndValidate() {
    PrintMarginsSpec margins = new PrintMarginsSpec(0.5d, 0.6d, 0.7d, 0.8d, 0.3d, 0.2d);
    PrintSetupSpec explicitSetup =
        new PrintSetupSpec(
            margins, true, true, false, 9, true, false, 2, true, 4, List.of(6), List.of(3));
    PrintSetupSpec defaultLikeSetup =
        new PrintSetupSpec(
            new PrintMarginsSpec(0.7d, 0.7d, 0.75d, 0.75d, 0.3d, 0.3d),
            false,
            false,
            false,
            0,
            false,
            false,
            0,
            false,
            0,
            List.of(),
            List.of());

    assertEquals(margins, explicitSetup.margins());
    assertTrue(explicitSetup.printGridlines());
    assertTrue(explicitSetup.horizontallyCentered());
    assertFalse(explicitSetup.verticallyCentered());
    assertEquals(9, explicitSetup.paperSize());
    assertTrue(explicitSetup.draft());
    assertFalse(explicitSetup.blackAndWhite());
    assertEquals(2, explicitSetup.copies());
    assertTrue(explicitSetup.useFirstPageNumber());
    assertEquals(4, explicitSetup.firstPageNumber());
    assertEquals(List.of(6), explicitSetup.rowBreaks());
    assertEquals(List.of(3), explicitSetup.columnBreaks());
    assertFalse(defaultLikeSetup.printGridlines());
    assertEquals(
        new PrintMarginsSpec(0.7d, 0.7d, 0.75d, 0.75d, 0.3d, 0.3d), defaultLikeSetup.margins());
    assertEquals(0, defaultLikeSetup.paperSize());
    assertEquals(0, defaultLikeSetup.copies());
    assertEquals(0, defaultLikeSetup.firstPageNumber());
    assertEquals(PrintSetupSpec.defaults().margins(), PrintSetupSpec.defaults().margins());
    assertEquals(PrintSetupSpec.defaults(), PrintLayoutSpec.defaults().setup());
    assertEquals(
        explicitSetup,
        new PrintLayoutSpec(
                new PrintAreaSpec.None(),
                ExcelPrintOrientation.PORTRAIT,
                new PrintScalingSpec.Automatic(),
                new PrintTitleRowsSpec.None(),
                new PrintTitleColumnsSpec.None(),
                HeaderFooterTextSpec.blank(),
                HeaderFooterTextSpec.blank(),
                explicitSetup)
            .setup());
    assertThrows(
        IllegalArgumentException.class,
        () -> new PrintMarginsSpec(-0.1d, 0.6d, 0.7d, 0.8d, 0.3d, 0.2d));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PrintMarginsSpec(0.5d, 0.6d, 0.7d, 0.8d, Double.NaN, 0.2d));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupSpec(
                margins, false, false, false, -1, false, false, 1, false, 1, List.of(), List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupSpec(
                margins, false, false, false, 1, false, false, -1, false, 1, List.of(), List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupSpec(
                margins, false, false, false, 1, false, false, 1, false, -1, List.of(), List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new PrintSetupSpec(
                margins,
                false,
                false,
                false,
                1,
                false,
                false,
                1,
                false,
                1,
                List.of((Integer) null),
                List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupSpec(
                margins,
                false,
                false,
                false,
                1,
                false,
                false,
                1,
                false,
                1,
                List.of(-1),
                List.of()));

    TableColumnSpec tableColumn = new TableColumnSpec(0, "", "", " SUM ", "");
    TableColumnSpec blankTotalsFunction = new TableColumnSpec(0, "", "", " ", "");
    assertEquals("", tableColumn.uniqueName());
    assertEquals("", tableColumn.totalsRowLabel());
    assertEquals("sum", tableColumn.totalsRowFunction());
    assertEquals("", tableColumn.calculatedColumnFormula());
    assertEquals("", blankTotalsFunction.totalsRowFunction());
    assertThrows(IllegalArgumentException.class, () -> new TableColumnSpec(-1, "", "", "", ""));
    assertThrows(NullPointerException.class, () -> new TableColumnSpec(0, null, "", "", ""));

    TableSpec defaultedTable =
        new TableSpec(
            new TableName("BudgetTable"),
            new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
            new RangeRef("A1:C4"),
            false,
            true,
            new TableStyleSpec.None(),
            text(""),
            false,
            false,
            false,
            "",
            "",
            "",
            List.of());
    TableSpec explicitTable =
        new TableSpec(
            new TableName("BudgetTable"),
            new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
            new RangeRef("A1:C4"),
            false,
            false,
            new TableStyleSpec.None(),
            text("Table note"),
            true,
            true,
            true,
            "HeaderStyle",
            "DataStyle",
            "TotalsStyle",
            List.of());
    assertFalse(defaultedTable.showTotalsRow());
    assertTrue(defaultedTable.hasAutofilter());
    assertEquals(text(""), defaultedTable.comment());
    assertEquals(List.of(), defaultedTable.columns());
    assertEquals(text("Table note"), explicitTable.comment());
    assertEquals("HeaderStyle", explicitTable.headerRowCellStyle());
    assertEquals("DataStyle", explicitTable.dataCellStyle());
    assertEquals("TotalsStyle", explicitTable.totalsRowCellStyle());
    assertThrows(
        NullPointerException.class,
        () ->
            new TableSpec(
                new TableName("BudgetTable"),
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                new RangeRef("A1:C4"),
                false,
                true,
                new TableStyleSpec.None(),
                null,
                false,
                false,
                false,
                null,
                null,
                null,
                List.of((TableColumnSpec) null)));
  }

  @Test
  void sheetPresentationInputsNormalizeAndValidate() {
    SheetPresentationSpec explicitPresentation =
        new SheetPresentationSpec(
            new SheetDisplaySpec(false, false, true, true, true),
            java.util.Optional.of(ColorSpec.rgb(new RgbHex("#112233"))),
            new SheetOutlineSummarySpec(false, false),
            new SheetDefaultsSpec(12, 18.5d),
            List.of(
                new IgnoredErrorSpec(
                    new RangeRef("B2:B12"),
                    List.of(
                        ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT,
                        ExcelIgnoredErrorType.FORMULA))));
    SheetPresentationSpec defaultedPresentation = SheetPresentationSpec.defaults();

    assertEquals(
        new SheetDisplaySpec(false, false, true, true, true), explicitPresentation.display());
    assertEquals(
        java.util.Optional.of(ColorSpec.rgb(new RgbHex("#112233"))),
        explicitPresentation.tabColor());
    assertEquals(new SheetOutlineSummarySpec(false, false), explicitPresentation.outlineSummary());
    assertEquals(new SheetDefaultsSpec(12, 18.5d), explicitPresentation.sheetDefaults());
    assertEquals(SheetDisplaySpec.defaults(), defaultedPresentation.display());
    assertEquals(SheetOutlineSummarySpec.defaults(), defaultedPresentation.outlineSummary());
    assertEquals(SheetDefaultsSpec.defaults(), defaultedPresentation.sheetDefaults());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new IgnoredErrorSpec(
                new RangeRef("B2:B12"),
                List.of(
                    ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT,
                    ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new SheetPresentationSpec(
                SheetDisplaySpec.defaults(),
                java.util.Optional.empty(),
                SheetOutlineSummarySpec.defaults(),
                SheetDefaultsSpec.defaults(),
                List.of(
                    new IgnoredErrorSpec(
                        new RangeRef("B2:B12"),
                        List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT)),
                    new IgnoredErrorSpec(
                        new RangeRef("B2:B12"), List.of(ExcelIgnoredErrorType.FORMULA)))));
  }

  @Test
  void pivotInputsSelectionsAndCellAddressValidationNormalizeAndValidate() {
    PivotTableSpec input =
        new PivotTableSpec(
            "Sales Pivot 2026",
            new SheetName("Report"),
            new PivotTableSpec.PivotSourceSpec.Range(new SheetName("Data"), new RangeRef("A1:D5")),
            new PivotTableSpec.Anchor(new CellAddress("$C$5")),
            List.of("Region"),
            List.of("Stage"),
            List.of("Owner"),
            List.of(
                new PivotTableSpec.DataField(
                    "Amount",
                    ExcelPivotDataConsolidateFunction.SUM,
                    "Amount",
                    Optional.of("#,##0.00"))));
    PivotTableSpec tableSourceInput =
        new PivotTableSpec(
            "Sales Table Pivot",
            new SheetName("Report"),
            new PivotTableSpec.PivotSourceSpec.Table(new TableName("SalesTable2026")),
            new PivotTableSpec.Anchor(new CellAddress("D7")),
            List.of(),
            List.of(),
            List.of(),
            List.of(
                new PivotTableSpec.DataField(
                    "Amount",
                    ExcelPivotDataConsolidateFunction.SUM,
                    "Total Amount",
                    Optional.empty())));
    dev.erst.gridgrind.core.plan.PivotTableSelector.ByNames selection =
        new dev.erst.gridgrind.core.plan.PivotTableSelector.ByNames(
            List.of("Sales Pivot 2026", "Ops Pivot"));

    assertEquals("$C$5", input.anchor().topLeftAddress().value());
    assertEquals("Amount", input.dataFields().getFirst().displayName());
    assertEquals(List.of("Sales Pivot 2026", "Ops Pivot"), selection.names());
    assertEquals(List.of(), tableSourceInput.rowLabels());
    assertEquals(List.of(), tableSourceInput.columnLabels());
    assertEquals(List.of(), tableSourceInput.reportFilters());
    assertEquals(
        "SalesTable2026",
        ((PivotTableSpec.PivotSourceSpec.Table) tableSourceInput.source()).name().value());
    assertEquals("B12", ProtocolCellAddressValidation.validateAddress("B12"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolCellAddressValidation.validateAddress("A1048577"));
    assertEquals(
        "PivotSource",
        new PivotTableSpec.PivotSourceSpec.NamedRange(new DefinedName("PivotSource"))
            .name()
            .value());
    assertThrows(
        IllegalArgumentException.class,
        () -> new PivotTableSpec.PivotSourceSpec.NamedRange(new DefinedName("Sales Pivot 2026")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PivotTableSpec.PivotSourceSpec.Table(new TableName("Sales Table")));
    assertThrows(
        NullPointerException.class,
        () -> new PivotTableSpec.PivotSourceSpec.Range(new SheetName("Data"), null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PivotTableSpec.PivotSourceSpec.Range(new SheetName("Data"), new RangeRef(" ")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec(
                "Sales Pivot 2026",
                new SheetName("Report"),
                new PivotTableSpec.PivotSourceSpec.Range(
                    new SheetName("Data"), new RangeRef("A1:D5")),
                new PivotTableSpec.Anchor(new CellAddress("C5")),
                List.of("Region"),
                List.of("Stage"),
                List.of("Owner"),
                List.of(
                    new PivotTableSpec.DataField(
                        "Region",
                        ExcelPivotDataConsolidateFunction.SUM,
                        "Total",
                        Optional.empty()))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec(
                "Sales Pivot 2026",
                new SheetName("Report"),
                new PivotTableSpec.PivotSourceSpec.Range(
                    new SheetName("Data"), new RangeRef("A1:D5")),
                new PivotTableSpec.Anchor(new CellAddress("C5")),
                List.of("Region", "region"),
                List.of(),
                List.of(),
                List.of(
                    new PivotTableSpec.DataField(
                        "Amount",
                        ExcelPivotDataConsolidateFunction.SUM,
                        "Total",
                        Optional.empty()))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec(
                "Sales Pivot 2026",
                new SheetName("Report"),
                new PivotTableSpec.PivotSourceSpec.Range(
                    new SheetName("Data"), new RangeRef("A1:D5")),
                new PivotTableSpec.Anchor(new CellAddress("C5")),
                List.of("Region"),
                List.of("region"),
                List.of(),
                List.of(
                    new PivotTableSpec.DataField(
                        "Amount",
                        ExcelPivotDataConsolidateFunction.SUM,
                        "Total",
                        Optional.empty()))));
    assertThrows(
        NullPointerException.class,
        () ->
            new PivotTableSpec(
                "Sales Pivot 2026",
                new SheetName("Report"),
                new PivotTableSpec.PivotSourceSpec.Range(
                    new SheetName("Data"), new RangeRef("A1:D5")),
                new PivotTableSpec.Anchor(new CellAddress("C5")),
                List.of((String) null),
                List.of(),
                List.of(),
                List.of(
                    new PivotTableSpec.DataField(
                        "Amount",
                        ExcelPivotDataConsolidateFunction.SUM,
                        "Total",
                        Optional.empty()))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec(
                "Sales Pivot 2026",
                new SheetName("Report"),
                new PivotTableSpec.PivotSourceSpec.Range(
                    new SheetName("Data"), new RangeRef("A1:D5")),
                new PivotTableSpec.Anchor(new CellAddress("C5")),
                List.of(" "),
                List.of(),
                List.of(),
                List.of(
                    new PivotTableSpec.DataField(
                        "Amount",
                        ExcelPivotDataConsolidateFunction.SUM,
                        "Total",
                        Optional.empty()))));
    assertThrows(
        NullPointerException.class,
        () ->
            new PivotTableSpec(
                "Sales Pivot 2026",
                new SheetName("Report"),
                new PivotTableSpec.PivotSourceSpec.Range(
                    new SheetName("Data"), new RangeRef("A1:D5")),
                new PivotTableSpec.Anchor(new CellAddress("C5")),
                List.of(),
                List.of(),
                List.of(),
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec(
                "Sales Pivot 2026",
                new SheetName("Report"),
                new PivotTableSpec.PivotSourceSpec.Range(
                    new SheetName("Data"), new RangeRef("A1:D5")),
                new PivotTableSpec.Anchor(new CellAddress("C5")),
                List.of(),
                List.of(),
                List.of(),
                List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec.DataField(
                "Amount", ExcelPivotDataConsolidateFunction.SUM, " ", Optional.empty()));
    assertThrows(
        NullPointerException.class,
        () -> new PivotTableSpec.DataField("Amount", null, "Total", Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec.DataField(
                " ", ExcelPivotDataConsolidateFunction.SUM, "Total", Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec.DataField(
                "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total", Optional.of(" ")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PivotTableSpec.Anchor(new CellAddress("Sheet1!A1")));
    assertThrows(
        NullPointerException.class, () -> ProtocolCellAddressValidation.validateAddress(null));
    assertThrows(
        IllegalArgumentException.class, () -> ProtocolCellAddressValidation.validateAddress(" "));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolCellAddressValidation.validateAddress("XFE1"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolCellAddressValidation.validateAddress("Sheet1!A1"));
    assertThrows(
        NullPointerException.class,
        () -> new dev.erst.gridgrind.core.plan.PivotTableSelector.ByNames(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new dev.erst.gridgrind.core.plan.PivotTableSelector.ByNames(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new dev.erst.gridgrind.core.plan.PivotTableSelector.ByNames(
                List.of("Sales Pivot 2026", "sales pivot 2026")));
  }

  private static TextSourceSpec text(String value) {
    return TextSourceSpec.inline(value);
  }

  private static BinarySourceSpec binary(String value) {
    return BinarySourceSpec.inlineBase64(value);
  }

  private static void assertThemeColor(ColorSpec color, int theme, Double tint) {
    ColorSpec.Theme themed = assertInstanceOf(ColorSpec.Theme.class, color);
    assertEquals(theme, themed.theme());
    assertEquals(Optional.ofNullable(tint), themed.tint());
  }

  private static void assertIndexedColor(ColorSpec color, int indexed, Double tint) {
    ColorSpec.Indexed indexedColor = assertInstanceOf(ColorSpec.Indexed.class, color);
    assertEquals(indexed, indexedColor.indexed());
    assertEquals(Optional.ofNullable(tint), indexedColor.tint());
  }

  private static ChartSpec chartInput(
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
        legend == null ? new ChartLegendSpec.Visible(ExcelChartLegendPosition.RIGHT) : legend,
        displayBlanksAs == null ? ExcelChartDisplayBlanksAs.GAP : displayBlanksAs,
        plotOnlyVisibleCells == null ? Boolean.TRUE : plotOnlyVisibleCells,
        List.of(plot));
  }

  private static ChartSeriesSpec chartSeries(
      ChartTitleSpec title, String categoriesFormula, String valuesFormula) {
    return new ChartSeriesSpec(
        title == null ? new ChartTitleSpec.None() : title,
        new ChartDataSourceSpec.Reference(categoriesFormula),
        new ChartDataSourceSpec.Reference(valuesFormula),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }
}
