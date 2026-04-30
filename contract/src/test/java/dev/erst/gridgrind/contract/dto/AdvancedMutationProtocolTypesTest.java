package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct tests for advanced mutation-oriented protocol DTO families. */
class AdvancedMutationProtocolTypesTest {
  @Test
  void ooxmlSecurityInputsNormalizeAndValidate() {
    OoxmlOpenSecurityInput openSecurity = new OoxmlOpenSecurityInput(Optional.of("source-pass"));
    OoxmlEncryptionInput encryption = OoxmlEncryptionInput.agile("persist-pass");
    OoxmlSignatureInput signature =
        new OoxmlSignatureInput(
            "tmp/signing-material.p12",
            "keystore-pass",
            "keystore-pass",
            Optional.empty(),
            ExcelOoxmlSignatureDigestAlgorithm.SHA256,
            Optional.empty());
    OoxmlPersistenceSecurityInput persistence =
        new OoxmlPersistenceSecurityInput(encryption, signature);

    assertEquals(Optional.of("source-pass"), openSecurity.password());
    assertEquals(ExcelOoxmlEncryptionMode.AGILE, encryption.mode());
    assertEquals("keystore-pass", signature.keyPassword());
    assertEquals(ExcelOoxmlSignatureDigestAlgorithm.SHA256, signature.digestAlgorithm());
    assertTrue(signature.alias().isEmpty());
    assertTrue(signature.description().isEmpty());
    assertEquals(encryption, persistence.encryption());
    assertEquals(signature, persistence.signature());

    assertThrows(
        IllegalArgumentException.class, () -> new OoxmlOpenSecurityInput(Optional.of(" ")));
    assertThrows(IllegalArgumentException.class, () -> OoxmlEncryptionInput.agile(" "));
    assertThrows(NullPointerException.class, () -> new OoxmlEncryptionInput("persist-pass", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new OoxmlSignatureInput(
                "tmp/signing-material.p12",
                "keystore-pass",
                " ",
                Optional.empty(),
                ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new OoxmlSignatureInput(
                "tmp/signing-material.p12",
                "keystore-pass",
                "keystore-pass",
                Optional.of(" "),
                ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class, () -> new OoxmlPersistenceSecurityInput(null, null));
  }

  @Test
  void ooxmlSecurityHelpersCollapseEmptyRequestState() {
    OoxmlOpenSecurityInput openSecurity = new OoxmlOpenSecurityInput(Optional.empty());
    OoxmlPersistenceSecurityInput encryptionOnly =
        new OoxmlPersistenceSecurityInput(
            new OoxmlEncryptionInput("persist-pass", ExcelOoxmlEncryptionMode.STANDARD), null);
    OoxmlPersistenceSecurityInput signatureOnly =
        new OoxmlPersistenceSecurityInput(
            null,
            new OoxmlSignatureInput(
                "tmp/signing-material.p12",
                "keystore-pass",
                "key-pass",
                Optional.empty(),
                ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                Optional.empty()));
    WorkbookPlan.WorkbookSource.ExistingFile source =
        new WorkbookPlan.WorkbookSource.ExistingFile("budget.xlsx", openSecurity);
    WorkbookPlan.WorkbookPersistence.OverwriteSource unsecuredOverwrite =
        new WorkbookPlan.WorkbookPersistence.OverwriteSource();
    WorkbookPlan.WorkbookPersistence.OverwriteSource securedOverwrite =
        new WorkbookPlan.WorkbookPersistence.OverwriteSource(encryptionOnly);

    assertEquals(Optional.empty(), openSecurity.password());
    assertTrue(openSecurity.isEmpty());
    assertTrue(source.security().isEmpty());
    assertTrue(unsecuredOverwrite.security().isEmpty());
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookPlan.WorkbookPersistence.OverwriteSource(
                (OoxmlPersistenceSecurityInput) null));
    assertEquals(encryptionOnly, securedOverwrite.security().orElseThrow());
    assertEquals(
        signatureOnly,
        new WorkbookPlan.WorkbookPersistence.SaveAs("secured.xlsx", signatureOnly)
            .security()
            .orElseThrow());
  }

  @Test
  void workbookProtectionNamedRangeAndCommentInputsNormalizeAndValidate() {
    WorkbookProtectionInput protection =
        new WorkbookProtectionInput(false, true, false, "book-secret", "review-secret");

    assertFalse(protection.structureLocked());
    assertTrue(protection.windowsLocked());
    assertFalse(protection.revisionsLocked());
    assertEquals("book-secret", protection.workbookPassword());
    assertEquals("review-secret", protection.revisionsPassword());
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookProtectionInput(true, false, false, " ", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookProtectionInput(true, false, false, null, " "));

    NamedRangeTarget explicit = new NamedRangeTarget("Budget", "C3:A1");
    NamedRangeTarget formula = new NamedRangeTarget("SUM(Budget!A1:A3)");
    assertEquals("Budget", explicit.sheetName());
    assertEquals("C3:A1", explicit.range());
    assertNull(explicit.formula());
    assertNull(formula.sheetName());
    assertNull(formula.range());
    assertEquals("SUM(Budget!A1:A3)", formula.formula());
    assertThrows(IllegalArgumentException.class, () -> new NamedRangeTarget(" "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new NamedRangeTarget("Budget", "A1", "SUM(Budget!A1)"));
    assertThrows(
        IllegalArgumentException.class, () -> new NamedRangeTarget(null, "A1", "SUM(Budget!A1)"));

    CommentAnchorInput anchor = new CommentAnchorInput(1, 2, 4, 6);
    assertEquals(1, anchor.firstColumn());
    assertEquals(2, anchor.firstRow());
    assertEquals(4, anchor.lastColumn());
    assertEquals(6, anchor.lastRow());
    assertThrows(IllegalArgumentException.class, () -> new CommentAnchorInput(-1, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new CommentAnchorInput(1, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new CommentAnchorInput(0, 2, 1, 1));

    CommentInput richComment =
        new CommentInput(
            text("Ada Lovelace"),
            "GridGrind",
            true,
            Optional.of(
                List.of(
                    new RichTextRunInput(text("Ada"), null),
                    new RichTextRunInput(text(" Lovelace"), null))),
            Optional.of(anchor));
    assertTrue(richComment.visible());
    assertEquals(anchor, richComment.anchor().orElseThrow());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CommentInput(
                text("Ada"), "GridGrind", true, Optional.of(List.of()), Optional.empty()));
    assertThrows(
        NullPointerException.class,
        () ->
            new CommentInput(
                text("Ada"),
                "GridGrind",
                true,
                Optional.of(List.of((RichTextRunInput) null)),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CommentInput(
                text("Ada Lovelace"),
                "GridGrind",
                true,
                Optional.of(List.of(new RichTextRunInput(text("Ada"), null))),
                Optional.empty()));
  }

  @Test
  void drawingInputsNormalizeAndValidate() {
    DrawingMarkerInput from = new DrawingMarkerInput(1, 2, 3, 4);
    DrawingMarkerInput to = new DrawingMarkerInput(4, 6);
    DrawingAnchorInput.TwoCell anchor =
        new DrawingAnchorInput.TwoCell(from, to, ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
    DrawingAnchorInput.TwoCell defaultAnchor = DrawingAnchorInput.TwoCell.moveAndResize(from, to);
    PictureDataInput pictureData =
        new PictureDataInput(
            ExcelPictureFormat.PNG,
            binary(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII="));
    PictureInput picture =
        new PictureInput("OpsPicture", pictureData, anchor, text("Queue preview"));
    ShapeInput shape =
        new ShapeInput(
            "OpsShape",
            ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
            anchor,
            " rect ",
            text("Queue"));
    ShapeInput connector =
        new ShapeInput("OpsConnector", ExcelAuthoredDrawingShapeKind.CONNECTOR, anchor, null, null);
    EmbeddedObjectInput embeddedObject =
        new EmbeddedObjectInput(
            "OpsEmbed",
            "Payload",
            "payload.txt",
            "payload.txt",
            binary("cGF5bG9hZA=="),
            pictureData,
            anchor);
    SignatureLineInput signatureLine =
        new SignatureLineInput(
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

    DrawingAnchorInput.TwoCell sameRowAnchor =
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(1, 2, 0, 3),
            new DrawingMarkerInput(2, 2, 0, 4),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    DrawingAnchorInput.TwoCell sameColAnchor =
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(1, 2, 3, 0),
            new DrawingMarkerInput(1, 3, 4, 0),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);

    assertEquals(4, to.columnIndex());
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE, anchor.behavior());
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE, defaultAnchor.behavior());
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE, sameRowAnchor.behavior());
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE, sameColAnchor.behavior());
    assertEquals(ExcelPictureFormat.PNG, picture.image().format());
    assertEquals(text("Queue preview"), picture.description());
    assertEquals("rect", shape.presetGeometryToken());
    assertEquals(ExcelAuthoredDrawingShapeKind.CONNECTOR, connector.kind());
    assertEquals("payload.txt", embeddedObject.fileName());
    assertTrue(signatureLine.allowComments());
    assertEquals("Ada Lovelace", signatureLine.suggestedSigner().orElseThrow());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ShapeInput(
                "DefaultShape", ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, anchor, " ", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ShapeInput(
                "NullPreset", ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, anchor, null, null));
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerInput(-1, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerInput(0, -1, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerInput(0, 0, -1, 0));
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerInput(0, 0, 0, -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorInput.TwoCell(
                new DrawingMarkerInput(1, 2, 0, 0),
                new DrawingMarkerInput(1, 1, 0, 0),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorInput.TwoCell(
                new DrawingMarkerInput(2, 2, 0, 0),
                new DrawingMarkerInput(1, 2, 0, 0),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorInput.TwoCell(
                new DrawingMarkerInput(1, 2, 0, 4),
                new DrawingMarkerInput(2, 2, 0, 3),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorInput.TwoCell(
                new DrawingMarkerInput(1, 2, 4, 0),
                new DrawingMarkerInput(1, 3, 3, 0),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertThrows(NullPointerException.class, () -> new DrawingAnchorInput.TwoCell(from, to, null));
    assertThrows(
        NullPointerException.class, () -> new PictureDataInput(null, binary("cGF5bG9hZA==")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PictureDataInput(ExcelPictureFormat.PNG, binary(" ")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PictureDataInput(ExcelPictureFormat.PNG, binary("not-base64")));
    assertThrows(
        IllegalArgumentException.class, () -> new PictureInput(" ", pictureData, anchor, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PictureInput("OpsPicture", pictureData, anchor, text(" ")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ShapeInput(" ", ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, anchor, "rect", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ShapeInput(
                "OpsShape", ExcelAuthoredDrawingShapeKind.CONNECTOR, anchor, "rect", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ShapeInput(
                "OpsShape",
                ExcelAuthoredDrawingShapeKind.CONNECTOR,
                anchor,
                null,
                text("Connector")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ShapeInput(
                "OpsShape", ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, anchor, "rect", text(" ")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new EmbeddedObjectInput(
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
            new EmbeddedObjectInput(
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
            new SignatureLineInput(
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
            new SignatureLineInput(
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
    ArrayFormulaInput wrapped = new ArrayFormulaInput(text("{=SUM(B2:B4*C2:C4)}"));
    ArrayFormulaInput bare = new ArrayFormulaInput(text("SUM(B2:B4*C2:C4)"));
    ArrayFormulaInput prefixed = new ArrayFormulaInput(text("=SUM(B2:B4*C2:C4)"));

    assertEquals(
        "SUM(B2:B4*C2:C4)",
        assertInstanceOf(TextSourceInput.Inline.class, wrapped.source()).text());
    assertEquals(
        "SUM(B2:B4*C2:C4)", assertInstanceOf(TextSourceInput.Inline.class, bare.source()).text());
    assertEquals(
        "SUM(B2:B4*C2:C4)",
        assertInstanceOf(TextSourceInput.Inline.class, prefixed.source()).text());

    assertThrows(IllegalArgumentException.class, () -> new ArrayFormulaInput(text(" ")));
    assertThrows(IllegalArgumentException.class, () -> new ArrayFormulaInput(text("{=}")));
  }

  @Test
  void customXmlInputsNormalizeAndValidate() {
    CustomXmlMappingLocator byId = new CustomXmlMappingLocator(1L, null);
    CustomXmlMappingLocator byName = new CustomXmlMappingLocator(null, "CORSO_mapping");
    CustomXmlImportInput inlineImport =
        new CustomXmlImportInput(byName, text("<CORSO><NOME>Grid</NOME></CORSO>"));

    assertEquals(1L, byId.mapId());
    assertEquals("CORSO_mapping", byName.name());
    assertEquals(byName, inlineImport.locator());
    assertEquals(
        "<CORSO><NOME>Grid</NOME></CORSO>",
        assertInstanceOf(TextSourceInput.Inline.class, inlineImport.xml()).text());

    assertThrows(IllegalArgumentException.class, () -> new CustomXmlMappingLocator(null, null));
    assertThrows(IllegalArgumentException.class, () -> new CustomXmlMappingLocator(0L, null));
    assertThrows(IllegalArgumentException.class, () -> new CustomXmlImportInput(byId, text(" ")));
  }

  @Test
  void chartInputsNormalizeAndValidate() {
    DrawingAnchorInput.TwoCell anchor =
        DrawingAnchorInput.TwoCell.moveAndResize(
            new DrawingMarkerInput(1, 2, 0, 0), new DrawingMarkerInput(6, 12, 0, 0));
    ChartSeriesInput firstSeries = chartSeries(new ChartTitleInput.Formula("B1"), "A2:A4", "B2:B4");
    ChartSeriesInput secondSeries = chartSeries(null, "ChartCategories", "ChartActual");
    ChartInput bar =
        chartInput(
            "OpsChart",
            anchor,
            null,
            null,
            null,
            null,
            new ChartPlotInput.Bar(
                false,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.CLUSTERED,
                null,
                null,
                List.of(firstSeries)));
    ChartInput line =
        chartInput(
            "TrendChart",
            anchor,
            new ChartTitleInput.Text(text("Trend")),
            new ChartLegendInput.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            false,
            new ChartPlotInput.Line(true, ExcelChartGrouping.STANDARD, List.of(secondSeries)));
    ChartInput defaultLine =
        chartInput(
            "DefaultTrend",
            anchor,
            null,
            null,
            null,
            null,
            new ChartPlotInput.Line(false, ExcelChartGrouping.STANDARD, List.of(secondSeries)));
    ChartInput pie =
        chartInput(
            "ShareChart",
            anchor,
            null,
            null,
            null,
            null,
            new ChartPlotInput.Pie(false, 180, List.of(secondSeries)));
    ChartInput defaultPie =
        chartInput(
            "DefaultShare",
            anchor,
            null,
            null,
            null,
            null,
            new ChartPlotInput.Pie(false, null, List.of(secondSeries)));
    ChartInput explicitPie =
        chartInput(
            "ExplicitShare",
            anchor,
            new ChartTitleInput.Text(text("Share")),
            new ChartLegendInput.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            false,
            new ChartPlotInput.Pie(true, 90, List.of(secondSeries)));

    ChartPlotInput.Bar barPlot = assertInstanceOf(ChartPlotInput.Bar.class, bar.plots().getFirst());
    ChartPlotInput.Line linePlot =
        assertInstanceOf(ChartPlotInput.Line.class, line.plots().getFirst());
    ChartPlotInput.Line defaultLinePlot =
        assertInstanceOf(ChartPlotInput.Line.class, defaultLine.plots().getFirst());
    ChartPlotInput.Pie piePlot = assertInstanceOf(ChartPlotInput.Pie.class, pie.plots().getFirst());
    ChartPlotInput.Pie defaultPiePlot =
        assertInstanceOf(ChartPlotInput.Pie.class, defaultPie.plots().getFirst());
    ChartPlotInput.Pie explicitPiePlot =
        assertInstanceOf(ChartPlotInput.Pie.class, explicitPie.plots().getFirst());

    assertTrue(bar.title() instanceof ChartTitleInput.None);
    assertEquals(new ChartLegendInput.Visible(ExcelChartLegendPosition.RIGHT), bar.legend());
    assertEquals(ExcelChartDisplayBlanksAs.GAP, bar.displayBlanksAs());
    assertTrue(bar.plotOnlyVisibleCells());
    assertFalse(barPlot.varyColors());
    assertEquals(ExcelChartBarDirection.COLUMN, barPlot.barDirection());
    assertEquals("B1", ((ChartTitleInput.Formula) barPlot.series().getFirst().title()).formula());
    assertTrue(secondSeries.title() instanceof ChartTitleInput.None);
    assertEquals(ExcelChartDisplayBlanksAs.ZERO, line.displayBlanksAs());
    assertFalse(line.plotOnlyVisibleCells());
    assertTrue(linePlot.varyColors());
    assertEquals(180, piePlot.firstSliceAngle());
    assertTrue(defaultLine.title() instanceof ChartTitleInput.None);
    assertEquals(
        new ChartLegendInput.Visible(ExcelChartLegendPosition.RIGHT), defaultLine.legend());
    assertEquals(ExcelChartDisplayBlanksAs.GAP, defaultLine.displayBlanksAs());
    assertTrue(defaultLine.plotOnlyVisibleCells());
    assertFalse(defaultLinePlot.varyColors());
    assertTrue(defaultPie.title() instanceof ChartTitleInput.None);
    assertEquals(new ChartLegendInput.Visible(ExcelChartLegendPosition.RIGHT), defaultPie.legend());
    assertEquals(ExcelChartDisplayBlanksAs.GAP, defaultPie.displayBlanksAs());
    assertTrue(defaultPie.plotOnlyVisibleCells());
    assertFalse(defaultPiePlot.varyColors());
    assertNull(defaultPiePlot.firstSliceAngle());
    assertEquals(new ChartTitleInput.Text(text("Share")), explicitPie.title());
    assertEquals(new ChartLegendInput.Hidden(), explicitPie.legend());
    assertEquals(ExcelChartDisplayBlanksAs.ZERO, explicitPie.displayBlanksAs());
    assertFalse(explicitPie.plotOnlyVisibleCells());
    assertTrue(explicitPiePlot.varyColors());
    assertEquals(90, explicitPiePlot.firstSliceAngle());

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
                new ChartPlotInput.Bar(
                    false,
                    ExcelChartBarDirection.COLUMN,
                    ExcelChartBarGrouping.CLUSTERED,
                    null,
                    null,
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
                new ChartPlotInput.Line(false, ExcelChartGrouping.STANDARD, List.of(firstSeries))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartPlotInput.Pie(false, 361, List.of(firstSeries)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartPlotInput.Pie(false, -1, List.of(firstSeries)));
    assertThrows(IllegalArgumentException.class, () -> new ChartTitleInput.Text(text(" ")));
    assertThrows(IllegalArgumentException.class, () -> new ChartTitleInput.Formula(" "));
    assertThrows(NullPointerException.class, () -> new ChartLegendInput.Visible(null));
    assertThrows(IllegalArgumentException.class, () -> new ChartDataSourceInput.Reference(" "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartPlotInput.Bar(
                false,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.CLUSTERED,
                null,
                null,
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ChartPlotInput.Line(
                false, ExcelChartGrouping.STANDARD, List.of(firstSeries, null)));
  }

  @Test
  void autofilterInputsNormalizeAndValidateAcrossAllCriterionFamilies() {
    AutofilterSortConditionInput colorSort =
        new AutofilterSortConditionInput.CellColor("B2:B9", true, ColorInput.rgb("#aabbcc"));
    AutofilterSortConditionInput iconSort =
        new AutofilterSortConditionInput.Icon("C2:C9", false, 2);
    AutofilterSortStateInput sortState =
        AutofilterSortStateInput.withoutSortMethod(
            "A1:F9", false, true, List.of(colorSort, iconSort));
    AutofilterSortStateInput sortStateWithMethod =
        new AutofilterSortStateInput(
            "A1:F9",
            true,
            false,
            Optional.of(dev.erst.gridgrind.excel.foundation.ExcelAutofilterSortMethod.PINYIN),
            List.of(colorSort));

    AutofilterSortConditionInput.CellColor typedColorSort =
        assertInstanceOf(AutofilterSortConditionInput.CellColor.class, colorSort);
    AutofilterSortConditionInput.Icon typedIconSort =
        assertInstanceOf(AutofilterSortConditionInput.Icon.class, iconSort);
    assertEquals("#AABBCC", assertInstanceOf(ColorInput.Rgb.class, typedColorSort.color()).rgb());
    assertEquals(2, typedIconSort.iconId());
    assertFalse(sortState.caseSensitive());
    assertTrue(sortState.columnSort());
    assertEquals(Optional.empty(), sortState.sortMethod());
    assertEquals(List.of(colorSort, iconSort), sortState.conditions());
    assertEquals(
        Optional.of(dev.erst.gridgrind.excel.foundation.ExcelAutofilterSortMethod.PINYIN),
        sortStateWithMethod.sortMethod());
    assertEquals(
        new AutofilterFilterCriterionInput.Dynamic("TODAY", Optional.empty(), Optional.empty()),
        new AutofilterFilterCriterionInput.Dynamic("TODAY", Optional.empty(), Optional.empty()));
    assertThrows(
        NullPointerException.class, () -> new AutofilterSortConditionInput.Value(null, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> AutofilterSortStateInput.withoutSortMethod(" ", false, false, List.of(colorSort)));
    assertThrows(
        IllegalArgumentException.class, () -> new AutofilterSortConditionInput.Value(" ", false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterSortConditionInput.Icon("A1:A2", false, -1));
    assertThrows(
        NullPointerException.class,
        () ->
            AutofilterSortStateInput.withoutSortMethod(
                "A1:F9", false, false, (List<AutofilterSortConditionInput>) null));
    assertThrows(
        IllegalArgumentException.class,
        () -> AutofilterSortStateInput.withoutSortMethod("A1:F9", false, false, List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            AutofilterSortStateInput.withoutSortMethod(
                "A1:F9", false, false, List.of((AutofilterSortConditionInput) null)));

    AutofilterFilterCriterionInput.Values values =
        new AutofilterFilterCriterionInput.Values(List.of("Queued", "Ready"), true);
    AutofilterFilterCriterionInput.Custom custom =
        new AutofilterFilterCriterionInput.Custom(
            true, List.of(new AutofilterFilterCriterionInput.CustomConditionInput("equal", "Ada")));
    AutofilterFilterCriterionInput.Dynamic dynamic =
        new AutofilterFilterCriterionInput.Dynamic("TODAY", Optional.of(1.0d), Optional.of(2.0d));
    AutofilterFilterCriterionInput.Top10 top10 =
        new AutofilterFilterCriterionInput.Top10(10, true, false);
    AutofilterFilterCriterionInput.Color color =
        new AutofilterFilterCriterionInput.Color(false, ColorInput.theme(3, 0.25d));
    AutofilterFilterCriterionInput.Icon icon =
        new AutofilterFilterCriterionInput.Icon("3TrafficLights1", 2);
    AutofilterFilterColumnInput column = AutofilterFilterColumnInput.visibleButton(2L, values);

    assertEquals(List.of("Queued", "Ready"), values.values());
    assertTrue(custom.and());
    assertEquals("TODAY", dynamic.type());
    assertEquals(Optional.of(1.0d), dynamic.value());
    assertEquals(Optional.of(2.0d), dynamic.maxValue());
    assertEquals(10, top10.value());
    assertEquals(3, assertInstanceOf(ColorInput.Theme.class, color.color()).theme());
    assertEquals("3TrafficLights1", icon.iconSet());
    assertTrue(column.showButton());
    assertEquals(values, column.criterion());

    assertThrows(
        NullPointerException.class, () -> new AutofilterFilterCriterionInput.Values(null, false));
    assertThrows(
        NullPointerException.class,
        () -> new AutofilterFilterCriterionInput.Values(List.of("Queued", null), false));
    assertThrows(
        NullPointerException.class, () -> new AutofilterFilterCriterionInput.Custom(true, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionInput.Custom(true, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionInput.CustomConditionInput(" ", "Ada"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionInput.CustomConditionInput("equal", " "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionInput.Dynamic(" ", Optional.of(1.0d), Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new AutofilterFilterCriterionInput.Dynamic(
                "TODAY", Optional.of(Double.POSITIVE_INFINITY), Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new AutofilterFilterCriterionInput.Dynamic(
                "TODAY", Optional.of(1.0d), Optional.of(Double.NaN)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionInput.Top10(0, true, false));
    assertThrows(
        NullPointerException.class, () -> new AutofilterFilterCriterionInput.Color(true, null));
    assertThrows(
        IllegalArgumentException.class, () -> new AutofilterFilterCriterionInput.Icon(" ", 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionInput.Icon("3TrafficLights1", -1));
    assertThrows(
        IllegalArgumentException.class, () -> new AutofilterFilterColumnInput(-1L, true, values));
    assertThrows(
        NullPointerException.class, () -> new AutofilterFilterColumnInput(0L, false, null));
  }

  @Test
  void colorAndThresholdInputsNormalizeAndValidate() {
    ColorInput.Rgb rgb = ColorInput.rgb("#aabbcc");
    ColorInput.Rgb tintedRgb = ColorInput.rgb("#112233", 0.25d);
    ColorInput.Theme themed = ColorInput.theme(4, 0.5d);
    ColorInput.Indexed indexed = ColorInput.indexed(64);
    assertEquals("#AABBCC", rgb.rgb());
    assertEquals("#112233", tintedRgb.rgb());
    assertEquals(0.25d, tintedRgb.tint());
    assertEquals(4, themed.theme());
    assertEquals(0.5d, themed.tint());
    assertEquals(64, indexed.indexed());
    assertThrows(NullPointerException.class, () -> ColorInput.rgb(null));
    assertThrows(IllegalArgumentException.class, () -> ColorInput.rgb(" "));
    assertThrows(IllegalArgumentException.class, () -> ColorInput.rgb("123456"));
    assertThrows(IllegalArgumentException.class, () -> ColorInput.theme(-1));
    assertThrows(IllegalArgumentException.class, () -> ColorInput.indexed(-1));
    assertThrows(
        IllegalArgumentException.class, () -> ColorInput.rgb("#AABBCC", Double.NEGATIVE_INFINITY));

    ConditionalFormattingThresholdInput threshold =
        new ConditionalFormattingThresholdInput(
            ExcelConditionalFormattingThresholdType.PERCENTILE, "A1*2", 90.0d);
    assertEquals(ExcelConditionalFormattingThresholdType.PERCENTILE, threshold.type());
    assertEquals("A1*2", threshold.formula());
    assertEquals(90.0d, threshold.value());
    assertThrows(
        NullPointerException.class,
        () -> new ConditionalFormattingThresholdInput(null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingThresholdInput(
                ExcelConditionalFormattingThresholdType.NUMBER, " ", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingThresholdInput(
                ExcelConditionalFormattingThresholdType.NUMBER, null, Double.NEGATIVE_INFINITY));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.ColorScaleRule(
                false,
                List.of(
                    new ConditionalFormattingThresholdInput(
                        ExcelConditionalFormattingThresholdType.MIN, null, null)),
                List.of(ColorInput.rgb("#112233"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.ColorScaleRule(
                false,
                List.of(
                    new ConditionalFormattingThresholdInput(
                        ExcelConditionalFormattingThresholdType.MIN, null, null),
                    new ConditionalFormattingThresholdInput(
                        ExcelConditionalFormattingThresholdType.MAX, null, null)),
                List.of(ColorInput.rgb("#112233"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.DataBarRule(
                false,
                ColorInput.rgb("#112233"),
                false,
                -1,
                90,
                new ConditionalFormattingThresholdInput(
                    ExcelConditionalFormattingThresholdType.NUMBER, null, 0.0d),
                new ConditionalFormattingThresholdInput(
                    ExcelConditionalFormattingThresholdType.NUMBER, null, 1.0d)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.DataBarRule(
                false,
                ColorInput.rgb("#112233"),
                false,
                0,
                -1,
                new ConditionalFormattingThresholdInput(
                    ExcelConditionalFormattingThresholdType.NUMBER, null, 0.0d),
                new ConditionalFormattingThresholdInput(
                    ExcelConditionalFormattingThresholdType.NUMBER, null, 1.0d)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.DataBarRule(
                false,
                ColorInput.rgb("#112233"),
                false,
                10,
                5,
                new ConditionalFormattingThresholdInput(
                    ExcelConditionalFormattingThresholdType.NUMBER, null, 0.0d),
                new ConditionalFormattingThresholdInput(
                    ExcelConditionalFormattingThresholdType.NUMBER, null, 1.0d)));
    assertThrows(
        NullPointerException.class,
        () -> new ConditionalFormattingRuleInput.IconSetRule(false, null, false, false, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.IconSetRule(
                false,
                ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
                false,
                false,
                List.of(
                    new ConditionalFormattingThresholdInput(
                        ExcelConditionalFormattingThresholdType.PERCENT, null, 0.0d),
                    new ConditionalFormattingThresholdInput(
                        ExcelConditionalFormattingThresholdType.PERCENT, null, 50.0d))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ConditionalFormattingRuleInput.Top10Rule(false, 0, true, false, null));
  }

  @Test
  void gradientAndStructuredStyleInputsNormalizeAndValidate() {
    CellGradientStopInput start = new CellGradientStopInput(0.0d, ColorInput.rgb("#112233"));
    CellGradientStopInput finish = new CellGradientStopInput(1.0d, ColorInput.theme(5, 0.2d));
    CellGradientFillInput gradient =
        CellGradientFillInput.path(0.1d, 0.2d, 0.3d, 0.4d, List.of(start, finish));
    assertGradientInputsNormalizeAndValidate(start, finish, gradient);
    assertFontInputsNormalizeAndValidate();
    assertFillInputsNormalizeAndValidate(gradient);
    assertBorderSideInputsNormalizeAndValidate();
  }

  private void assertGradientInputsNormalizeAndValidate(
      CellGradientStopInput start, CellGradientStopInput finish, CellGradientFillInput gradient) {
    CellGradientFillInput.Path path = assertInstanceOf(CellGradientFillInput.Path.class, gradient);
    assertEquals(0.1d, path.left());
    assertEquals(0.2d, path.right());
    assertEquals(0.3d, path.top());
    assertEquals(0.4d, path.bottom());
    assertEquals(List.of(start, finish), path.stops());
    assertInstanceOf(
        CellGradientFillInput.Linear.class,
        CellGradientFillInput.linear(null, List.of(start, finish)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientStopInput(-0.1d, ColorInput.rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientStopInput(Double.NaN, ColorInput.rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientStopInput(1.5d, ColorInput.rgb("#112233")));
    assertThrows(NullPointerException.class, () -> new CellGradientStopInput(0.5d, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> CellGradientFillInput.linear(Double.POSITIVE_INFINITY, List.of(start, finish)));
    assertThrows(
        IllegalArgumentException.class,
        () -> CellGradientFillInput.path(Double.NaN, 0.2d, 0.3d, 0.4d, List.of(start, finish)));
    assertThrows(
        IllegalArgumentException.class, () -> CellGradientFillInput.linear(null, List.of(start)));
    assertThrows(
        NullPointerException.class, () -> CellGradientFillInput.linear(null, List.of(start, null)));
  }

  private void assertFontInputsNormalizeAndValidate() {
    CellFontInput themedFont =
        new CellFontInput(null, null, null, null, ColorInput.theme(2, 0.4d), true, null);
    CellFontInput namedFont = new CellFontInput(null, null, "Calibri", null, null, null, null);
    CellFontInput indexedFont =
        new CellFontInput(null, null, null, null, ColorInput.indexed(64), null, null);
    assertThemeColor(themedFont.fontColor(), 2, 0.4d);
    assertTrue(themedFont.underline());
    assertEquals("Calibri", namedFont.fontName());
    assertIndexedColor(indexedFont.fontColor(), 64, null);
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFontInput(null, null, " ", null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFontInput(null, null, null, null, null, null, null));
    assertThrows(IllegalArgumentException.class, () -> ColorInput.theme(-1));
    assertThrows(IllegalArgumentException.class, () -> ColorInput.indexed(-1));
    assertThrows(
        IllegalArgumentException.class, () -> ColorInput.rgb("#112233", Double.POSITIVE_INFINITY));
  }

  private void assertFillInputsNormalizeAndValidate(CellGradientFillInput gradient) {
    CellFillInput.Gradient gradientFill =
        assertInstanceOf(CellFillInput.Gradient.class, CellFillInput.gradient(gradient));
    CellFillInput.PatternForeground themedFill =
        assertInstanceOf(
            CellFillInput.PatternForeground.class,
            CellFillInput.patternForeground(ExcelFillPattern.SOLID, ColorInput.theme(2, 0.3d)));
    CellFillInput.PatternForeground fgIndexedFill =
        assertInstanceOf(
            CellFillInput.PatternForeground.class,
            CellFillInput.patternForeground(ExcelFillPattern.SOLID, ColorInput.indexed(64)));
    CellFillInput.PatternForeground fgIndexedTintedFill =
        assertInstanceOf(
            CellFillInput.PatternForeground.class,
            CellFillInput.patternForeground(ExcelFillPattern.SOLID, ColorInput.indexed(64, 0.3d)));
    CellFillInput.PatternBackground bgIndexedFill =
        assertInstanceOf(
            CellFillInput.PatternBackground.class,
            CellFillInput.patternBackground(ExcelFillPattern.FINE_DOTS, ColorInput.indexed(64)));
    assertEquals(gradient, gradientFill.gradient());
    assertThemeColor(themedFill.foregroundColor(), 2, 0.3d);
    assertIndexedColor(fgIndexedFill.foregroundColor(), 64, null);
    assertIndexedColor(fgIndexedTintedFill.foregroundColor(), 64, 0.3d);
    assertIndexedColor(bgIndexedFill.backgroundColor(), 64, null);
    assertThrows(NullPointerException.class, () -> CellFillInput.gradient(null));
    assertThrows(NullPointerException.class, () -> CellFillInput.pattern(null));
    assertThrows(
        NullPointerException.class,
        () -> CellFillInput.patternForeground(ExcelFillPattern.SOLID, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> CellFillInput.patternForeground(ExcelFillPattern.NONE, ColorInput.theme(2)));
    assertThrows(
        IllegalArgumentException.class,
        () -> CellFillInput.patternBackground(ExcelFillPattern.SOLID, ColorInput.indexed(64)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellFillInput.patternColors(
                ExcelFillPattern.SOLID, ColorInput.rgb("#112233"), ColorInput.indexed(64)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellFillInput.patternColors(
                ExcelFillPattern.NONE, ColorInput.rgb("#112233"), ColorInput.indexed(64)));
    assertThrows(
        NullPointerException.class,
        () -> CellFillInput.patternBackground(ExcelFillPattern.FINE_DOTS, null));
    assertThrows(
        NullPointerException.class,
        () ->
            CellFillInput.patternColors(
                ExcelFillPattern.FINE_DOTS, ColorInput.rgb("#112233"), null));
  }

  private void assertBorderSideInputsNormalizeAndValidate() {
    CellBorderSideInput noneStyleOnly = new CellBorderSideInput(ExcelBorderStyle.NONE);
    CellBorderSideInput borderSide = new CellBorderSideInput(null, ColorInput.theme(1, 0.15d));
    CellBorderSideInput indexedBorderSide = new CellBorderSideInput(null, ColorInput.indexed(64));
    assertEquals(ExcelBorderStyle.NONE, noneStyleOnly.style());
    assertThemeColor(borderSide.color(), 1, 0.15d);
    assertIndexedColor(indexedBorderSide.color(), 64, null);
    assertThrows(IllegalArgumentException.class, () -> new CellBorderSideInput(null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellBorderSideInput(ExcelBorderStyle.NONE, ColorInput.theme(1)));
    assertThrows(IllegalArgumentException.class, () -> ColorInput.theme(-1));
    assertThrows(IllegalArgumentException.class, () -> ColorInput.indexed(-1));
    assertThrows(
        IllegalArgumentException.class, () -> ColorInput.theme(1, Double.POSITIVE_INFINITY));
  }

  @Test
  void printAndTableInputsNormalizeAndValidate() {
    PrintMarginsInput margins = new PrintMarginsInput(0.5d, 0.6d, 0.7d, 0.8d, 0.3d, 0.2d);
    PrintSetupInput explicitSetup =
        new PrintSetupInput(
            margins, true, true, false, 9, true, false, 2, true, 4, List.of(6), List.of(3));
    PrintSetupInput defaultLikeSetup =
        new PrintSetupInput(
            new PrintMarginsInput(0.7d, 0.7d, 0.75d, 0.75d, 0.3d, 0.3d),
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
        new PrintMarginsInput(0.7d, 0.7d, 0.75d, 0.75d, 0.3d, 0.3d), defaultLikeSetup.margins());
    assertEquals(0, defaultLikeSetup.paperSize());
    assertEquals(0, defaultLikeSetup.copies());
    assertEquals(0, defaultLikeSetup.firstPageNumber());
    assertEquals(PrintSetupInput.defaults().margins(), PrintSetupInput.defaults().margins());
    assertEquals(PrintSetupInput.defaults(), PrintLayoutInput.defaults().setup());
    assertEquals(
        explicitSetup,
        new PrintLayoutInput(
                new PrintAreaInput.None(),
                ExcelPrintOrientation.PORTRAIT,
                new PrintScalingInput.Automatic(),
                new PrintTitleRowsInput.None(),
                new PrintTitleColumnsInput.None(),
                HeaderFooterTextInput.blank(),
                HeaderFooterTextInput.blank(),
                explicitSetup)
            .setup());
    assertThrows(
        IllegalArgumentException.class,
        () -> new PrintMarginsInput(-0.1d, 0.6d, 0.7d, 0.8d, 0.3d, 0.2d));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PrintMarginsInput(0.5d, 0.6d, 0.7d, 0.8d, Double.NaN, 0.2d));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupInput(
                margins, false, false, false, -1, false, false, 1, false, 1, List.of(), List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupInput(
                margins, false, false, false, 1, false, false, -1, false, 1, List.of(), List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupInput(
                margins, false, false, false, 1, false, false, 1, false, -1, List.of(), List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new PrintSetupInput(
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
            new PrintSetupInput(
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

    TableColumnInput tableColumn = new TableColumnInput(0, "", "", " SUM ", "");
    TableColumnInput blankTotalsFunction = new TableColumnInput(0, "", "", " ", "");
    assertEquals("", tableColumn.uniqueName());
    assertEquals("", tableColumn.totalsRowLabel());
    assertEquals("sum", tableColumn.totalsRowFunction());
    assertEquals("", tableColumn.calculatedColumnFormula());
    assertEquals("", blankTotalsFunction.totalsRowFunction());
    assertThrows(IllegalArgumentException.class, () -> new TableColumnInput(-1, "", "", "", ""));
    assertThrows(NullPointerException.class, () -> new TableColumnInput(0, null, "", "", ""));

    TableInput defaultedTable =
        new TableInput(
            "BudgetTable",
            "Budget",
            "A1:C4",
            false,
            true,
            new TableStyleInput.None(),
            text(""),
            false,
            false,
            false,
            "",
            "",
            "",
            List.of());
    TableInput explicitTable =
        new TableInput(
            "BudgetTable",
            "Budget",
            "A1:C4",
            false,
            false,
            new TableStyleInput.None(),
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
            new TableInput(
                "BudgetTable",
                "Budget",
                "A1:C4",
                false,
                true,
                new TableStyleInput.None(),
                null,
                false,
                false,
                false,
                null,
                null,
                null,
                List.of((TableColumnInput) null)));
  }

  @Test
  void sheetPresentationInputsNormalizeAndValidate() {
    SheetPresentationInput explicitPresentation =
        new SheetPresentationInput(
            new SheetDisplayInput(false, false, true, true, true),
            java.util.Optional.of(ColorInput.rgb("#112233")),
            new SheetOutlineSummaryInput(false, false),
            new SheetDefaultsInput(12, 18.5d),
            List.of(
                new IgnoredErrorInput(
                    "B2:B12",
                    List.of(
                        ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT,
                        ExcelIgnoredErrorType.FORMULA))));
    SheetPresentationInput defaultedPresentation = SheetPresentationInput.defaults();

    assertEquals(
        new SheetDisplayInput(false, false, true, true, true), explicitPresentation.display());
    assertEquals(java.util.Optional.of(ColorInput.rgb("#112233")), explicitPresentation.tabColor());
    assertEquals(new SheetOutlineSummaryInput(false, false), explicitPresentation.outlineSummary());
    assertEquals(new SheetDefaultsInput(12, 18.5d), explicitPresentation.sheetDefaults());
    assertEquals(SheetDisplayInput.defaults(), defaultedPresentation.display());
    assertEquals(SheetOutlineSummaryInput.defaults(), defaultedPresentation.outlineSummary());
    assertEquals(SheetDefaultsInput.defaults(), defaultedPresentation.sheetDefaults());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new IgnoredErrorInput(
                "B2:B12",
                List.of(
                    ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT,
                    ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            SheetPresentationInput.create(
                SheetDisplayInput.defaults(),
                java.util.Optional.empty(),
                SheetOutlineSummaryInput.defaults(),
                SheetDefaultsInput.defaults(),
                List.of(
                    new IgnoredErrorInput(
                        "B2:B12", List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT)),
                    new IgnoredErrorInput("B2:B12", List.of(ExcelIgnoredErrorType.FORMULA)))));
  }

  @Test
  void pivotInputsSelectionsAndCellAddressValidationNormalizeAndValidate() {
    PivotTableInput input =
        new PivotTableInput(
            "Sales Pivot 2026",
            "Report",
            new PivotTableInput.Source.Range("Data", "A1:D5"),
            new PivotTableInput.Anchor("$C$5"),
            List.of("Region"),
            List.of("Stage"),
            List.of("Owner"),
            List.of(
                new PivotTableInput.DataField(
                    "Amount", ExcelPivotDataConsolidateFunction.SUM, "Amount", "#,##0.00")));
    PivotTableInput tableSourceInput =
        new PivotTableInput(
            "Sales Table Pivot",
            "Report",
            new PivotTableInput.Source.Table("SalesTable2026"),
            new PivotTableInput.Anchor("D7"),
            List.of(),
            List.of(),
            List.of(),
            List.of(
                new PivotTableInput.DataField(
                    "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total Amount", null)));
    dev.erst.gridgrind.contract.selector.PivotTableSelector.ByNames selection =
        new dev.erst.gridgrind.contract.selector.PivotTableSelector.ByNames(
            List.of("Sales Pivot 2026", "Ops Pivot"));

    assertEquals("$C$5", input.anchor().topLeftAddress());
    assertEquals("Amount", input.dataFields().getFirst().displayName());
    assertEquals(List.of("Sales Pivot 2026", "Ops Pivot"), selection.names());
    assertEquals(List.of(), tableSourceInput.rowLabels());
    assertEquals(List.of(), tableSourceInput.columnLabels());
    assertEquals(List.of(), tableSourceInput.reportFilters());
    assertEquals(
        "SalesTable2026", ((PivotTableInput.Source.Table) tableSourceInput.source()).name());
    assertEquals("B12", ProtocolCellAddressValidation.validateAddress("B12"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolCellAddressValidation.validateAddress("A1048577"));
    assertEquals("PivotSource", new PivotTableInput.Source.NamedRange("PivotSource").name());
    assertThrows(
        IllegalArgumentException.class,
        () -> new PivotTableInput.Source.NamedRange("Sales Pivot 2026"));
    assertThrows(
        IllegalArgumentException.class, () -> new PivotTableInput.Source.Table("Sales Table"));
    assertThrows(NullPointerException.class, () -> new PivotTableInput.Source.Range("Data", null));
    assertThrows(
        IllegalArgumentException.class, () -> new PivotTableInput.Source.Range("Data", " "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableInput(
                "Sales Pivot 2026",
                "Report",
                new PivotTableInput.Source.Range("Data", "A1:D5"),
                new PivotTableInput.Anchor("C5"),
                List.of("Region"),
                List.of("Stage"),
                List.of("Owner"),
                List.of(
                    new PivotTableInput.DataField(
                        "Region", ExcelPivotDataConsolidateFunction.SUM, "Total", null))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableInput(
                "Sales Pivot 2026",
                "Report",
                new PivotTableInput.Source.Range("Data", "A1:D5"),
                new PivotTableInput.Anchor("C5"),
                List.of("Region", "region"),
                List.of(),
                List.of(),
                List.of(
                    new PivotTableInput.DataField(
                        "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total", null))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableInput(
                "Sales Pivot 2026",
                "Report",
                new PivotTableInput.Source.Range("Data", "A1:D5"),
                new PivotTableInput.Anchor("C5"),
                List.of("Region"),
                List.of("region"),
                List.of(),
                List.of(
                    new PivotTableInput.DataField(
                        "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total", null))));
    assertThrows(
        NullPointerException.class,
        () ->
            new PivotTableInput(
                "Sales Pivot 2026",
                "Report",
                new PivotTableInput.Source.Range("Data", "A1:D5"),
                new PivotTableInput.Anchor("C5"),
                List.of((String) null),
                List.of(),
                List.of(),
                List.of(
                    new PivotTableInput.DataField(
                        "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total", null))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableInput(
                "Sales Pivot 2026",
                "Report",
                new PivotTableInput.Source.Range("Data", "A1:D5"),
                new PivotTableInput.Anchor("C5"),
                List.of(" "),
                List.of(),
                List.of(),
                List.of(
                    new PivotTableInput.DataField(
                        "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total", null))));
    assertThrows(
        NullPointerException.class,
        () ->
            new PivotTableInput(
                "Sales Pivot 2026",
                "Report",
                new PivotTableInput.Source.Range("Data", "A1:D5"),
                new PivotTableInput.Anchor("C5"),
                List.of(),
                List.of(),
                List.of(),
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableInput(
                "Sales Pivot 2026",
                "Report",
                new PivotTableInput.Source.Range("Data", "A1:D5"),
                new PivotTableInput.Anchor("C5"),
                List.of(),
                List.of(),
                List.of(),
                List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableInput.DataField(
                "Amount", ExcelPivotDataConsolidateFunction.SUM, " ", null));
    assertThrows(
        NullPointerException.class,
        () -> new PivotTableInput.DataField("Amount", null, "Total", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableInput.DataField(
                " ", ExcelPivotDataConsolidateFunction.SUM, "Total", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableInput.DataField(
                "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total", " "));
    assertThrows(IllegalArgumentException.class, () -> new PivotTableInput.Anchor("Sheet1!A1"));
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
        () -> new dev.erst.gridgrind.contract.selector.PivotTableSelector.ByNames(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new dev.erst.gridgrind.contract.selector.PivotTableSelector.ByNames(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new dev.erst.gridgrind.contract.selector.PivotTableSelector.ByNames(
                List.of("Sales Pivot 2026", "sales pivot 2026")));
  }

  private static TextSourceInput text(String value) {
    return TextSourceInput.inline(value);
  }

  private static BinarySourceInput binary(String value) {
    return BinarySourceInput.inlineBase64(value);
  }

  private static void assertThemeColor(ColorInput color, int theme, Double tint) {
    ColorInput.Theme themed = assertInstanceOf(ColorInput.Theme.class, color);
    assertEquals(theme, themed.theme());
    assertEquals(tint, themed.tint());
  }

  private static void assertIndexedColor(ColorInput color, int indexed, Double tint) {
    ColorInput.Indexed indexedColor = assertInstanceOf(ColorInput.Indexed.class, color);
    assertEquals(indexed, indexedColor.indexed());
    assertEquals(tint, indexedColor.tint());
  }

  private static ChartInput chartInput(
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
        legend == null ? new ChartLegendInput.Visible(ExcelChartLegendPosition.RIGHT) : legend,
        displayBlanksAs == null ? ExcelChartDisplayBlanksAs.GAP : displayBlanksAs,
        plotOnlyVisibleCells == null ? Boolean.TRUE : plotOnlyVisibleCells,
        List.of(plot));
  }

  private static ChartSeriesInput chartSeries(
      ChartTitleInput title, String categoriesFormula, String valuesFormula) {
    return new ChartSeriesInput(
        title == null ? new ChartTitleInput.None() : title,
        new ChartDataSourceInput.Reference(categoriesFormula),
        new ChartDataSourceInput.Reference(valuesFormula),
        null,
        null,
        null,
        null);
  }
}
