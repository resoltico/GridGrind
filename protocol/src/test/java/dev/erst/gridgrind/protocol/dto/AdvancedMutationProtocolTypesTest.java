package dev.erst.gridgrind.protocol.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.excel.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.ExcelFillPattern;
import dev.erst.gridgrind.excel.ExcelIgnoredErrorType;
import dev.erst.gridgrind.excel.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.excel.ExcelPictureFormat;
import dev.erst.gridgrind.excel.ExcelPivotDataConsolidateFunction;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct tests for advanced mutation-oriented protocol DTO families. */
class AdvancedMutationProtocolTypesTest {
  @Test
  void ooxmlSecurityInputsNormalizeAndValidate() {
    OoxmlOpenSecurityInput openSecurity = new OoxmlOpenSecurityInput("source-pass");
    OoxmlEncryptionInput encryption = new OoxmlEncryptionInput("persist-pass", null);
    OoxmlSignatureInput signature =
        new OoxmlSignatureInput(
            "tmp/signing-material.p12", "keystore-pass", null, null, null, null);
    OoxmlPersistenceSecurityInput persistence =
        new OoxmlPersistenceSecurityInput(encryption, signature);

    assertEquals("source-pass", openSecurity.password());
    assertEquals(ExcelOoxmlEncryptionMode.AGILE, encryption.mode());
    assertEquals("keystore-pass", signature.keyPassword());
    assertEquals(ExcelOoxmlSignatureDigestAlgorithm.SHA256, signature.digestAlgorithm());
    assertNull(signature.alias());
    assertNull(signature.description());
    assertEquals(encryption, persistence.encryption());
    assertEquals(signature, persistence.signature());

    assertThrows(IllegalArgumentException.class, () -> new OoxmlOpenSecurityInput(" "));
    assertThrows(IllegalArgumentException.class, () -> new OoxmlEncryptionInput(" ", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new OoxmlSignatureInput(
                "tmp/signing-material.p12", "keystore-pass", " ", null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new OoxmlSignatureInput(
                "tmp/signing-material.p12", "keystore-pass", null, " ", null, null));
    assertThrows(
        IllegalArgumentException.class, () -> new OoxmlPersistenceSecurityInput(null, null));
  }

  @Test
  void ooxmlSecurityHelpersCollapseEmptyRequestState() {
    OoxmlOpenSecurityInput openSecurity = new OoxmlOpenSecurityInput(null);
    OoxmlPersistenceSecurityInput encryptionOnly =
        new OoxmlPersistenceSecurityInput(
            new OoxmlEncryptionInput("persist-pass", ExcelOoxmlEncryptionMode.STANDARD), null);
    OoxmlPersistenceSecurityInput signatureOnly =
        new OoxmlPersistenceSecurityInput(
            null,
            new OoxmlSignatureInput(
                "tmp/signing-material.p12", "keystore-pass", "key-pass", null, null, null));
    GridGrindRequest.WorkbookSource.ExistingFile source =
        new GridGrindRequest.WorkbookSource.ExistingFile("budget.xlsx", openSecurity);
    GridGrindRequest.WorkbookPersistence.OverwriteSource unsecuredOverwrite =
        new GridGrindRequest.WorkbookPersistence.OverwriteSource(
            (OoxmlPersistenceSecurityInput) null);
    GridGrindRequest.WorkbookPersistence.OverwriteSource securedOverwrite =
        new GridGrindRequest.WorkbookPersistence.OverwriteSource(encryptionOnly);

    assertNull(openSecurity.password());
    assertTrue(openSecurity.isEmpty());
    assertNull(source.security());
    assertNull(unsecuredOverwrite.security());
    assertEquals(encryptionOnly, securedOverwrite.security());
    assertEquals(
        signatureOnly,
        new GridGrindRequest.WorkbookPersistence.SaveAs("secured.xlsx", signatureOnly).security());
  }

  @Test
  void workbookProtectionNamedRangeAndCommentInputsNormalizeAndValidate() {
    WorkbookProtectionInput protection =
        new WorkbookProtectionInput(null, true, null, "book-secret", "review-secret");

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
            "Ada Lovelace",
            "GridGrind",
            true,
            List.of(new RichTextRunInput("Ada", null), new RichTextRunInput(" Lovelace", null)),
            anchor);
    assertTrue(richComment.visible());
    assertEquals(anchor, richComment.anchor());
    assertThrows(
        IllegalArgumentException.class,
        () -> new CommentInput("Ada", "GridGrind", true, List.of(), null));
    assertThrows(
        NullPointerException.class,
        () -> new CommentInput("Ada", "GridGrind", true, List.of((RichTextRunInput) null), null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CommentInput(
                "Ada Lovelace",
                "GridGrind",
                true,
                List.of(new RichTextRunInput("Ada", null)),
                null));
  }

  @Test
  void drawingInputsNormalizeAndValidate() {
    DrawingMarkerInput from = new DrawingMarkerInput(1, 2, 3, 4);
    DrawingMarkerInput to = new DrawingMarkerInput(4, 6);
    DrawingAnchorInput.TwoCell anchor =
        new DrawingAnchorInput.TwoCell(from, to, ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
    DrawingAnchorInput.TwoCell defaultAnchor = new DrawingAnchorInput.TwoCell(from, to, null);
    PictureDataInput pictureData =
        new PictureDataInput(
            ExcelPictureFormat.PNG,
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");
    PictureInput picture = new PictureInput("OpsPicture", pictureData, anchor, "Queue preview");
    ShapeInput shape =
        new ShapeInput(
            "OpsShape", ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, anchor, " rect ", "Queue");
    ShapeInput defaultShape =
        new ShapeInput(
            "DefaultShape", ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, anchor, " ", null);
    ShapeInput connector =
        new ShapeInput("OpsConnector", ExcelAuthoredDrawingShapeKind.CONNECTOR, anchor, null, null);
    EmbeddedObjectInput embeddedObject =
        new EmbeddedObjectInput(
            "OpsEmbed",
            "Payload",
            "payload.txt",
            "payload.txt",
            "cGF5bG9hZA==",
            pictureData,
            anchor);

    DrawingAnchorInput.TwoCell sameRowAnchor =
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(1, 2, 0, 3), new DrawingMarkerInput(2, 2, 0, 4), null);
    DrawingAnchorInput.TwoCell sameColAnchor =
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(1, 2, 3, 0), new DrawingMarkerInput(1, 3, 4, 0), null);

    assertEquals(4, to.columnIndex());
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE, anchor.behavior());
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE, defaultAnchor.behavior());
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE, sameRowAnchor.behavior());
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE, sameColAnchor.behavior());
    assertEquals(ExcelPictureFormat.PNG, picture.image().format());
    assertEquals("Queue preview", picture.description());
    ShapeInput nullPresetShape =
        new ShapeInput(
            "NullPreset", ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, anchor, null, null);
    assertEquals("rect", shape.presetGeometryToken());
    assertEquals("rect", defaultShape.presetGeometryToken());
    assertEquals("rect", nullPresetShape.presetGeometryToken());
    assertEquals(ExcelAuthoredDrawingShapeKind.CONNECTOR, connector.kind());
    assertEquals("payload.txt", embeddedObject.fileName());
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerInput(-1, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerInput(0, -1, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerInput(0, 0, -1, 0));
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerInput(0, 0, 0, -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorInput.TwoCell(
                new DrawingMarkerInput(1, 2, 0, 0), new DrawingMarkerInput(1, 1, 0, 0), null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorInput.TwoCell(
                new DrawingMarkerInput(2, 2, 0, 0), new DrawingMarkerInput(1, 2, 0, 0), null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorInput.TwoCell(
                new DrawingMarkerInput(1, 2, 0, 4), new DrawingMarkerInput(2, 2, 0, 3), null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorInput.TwoCell(
                new DrawingMarkerInput(1, 2, 4, 0), new DrawingMarkerInput(1, 3, 3, 0), null));
    assertThrows(NullPointerException.class, () -> new PictureDataInput(null, "cGF5bG9hZA=="));
    assertThrows(
        IllegalArgumentException.class, () -> new PictureDataInput(ExcelPictureFormat.PNG, " "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PictureDataInput(ExcelPictureFormat.PNG, "not-base64"));
    assertThrows(
        IllegalArgumentException.class, () -> new PictureInput(" ", pictureData, anchor, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PictureInput("OpsPicture", pictureData, anchor, " "));
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
                "OpsShape", ExcelAuthoredDrawingShapeKind.CONNECTOR, anchor, null, "Connector"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ShapeInput(
                "OpsShape", ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, anchor, "rect", " "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new EmbeddedObjectInput(
                "OpsEmbed",
                "Payload",
                "payload.txt",
                "payload.txt",
                "not-base64",
                pictureData,
                anchor));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new EmbeddedObjectInput(
                "OpsEmbed", "Payload", "payload.txt", " ", "cGF5bG9hZA==", pictureData, anchor));
  }

  @Test
  void chartInputsNormalizeAndValidate() {
    DrawingAnchorInput.TwoCell anchor =
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(1, 2, 0, 0), new DrawingMarkerInput(6, 12, 0, 0), null);
    ChartInput.Series firstSeries =
        new ChartInput.Series(
            new ChartInput.Title.Formula("B1"),
            new ChartInput.DataSource("A2:A4"),
            new ChartInput.DataSource("B2:B4"));
    ChartInput.Series secondSeries =
        new ChartInput.Series(
            null,
            new ChartInput.DataSource("ChartCategories"),
            new ChartInput.DataSource("ChartActual"));
    ChartInput.Bar bar =
        new ChartInput.Bar(
            "OpsChart", anchor, null, null, null, null, null, null, List.of(firstSeries));
    ChartInput.Line line =
        new ChartInput.Line(
            "TrendChart",
            anchor,
            new ChartInput.Title.Text("Trend"),
            new ChartInput.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            false,
            true,
            List.of(secondSeries));
    ChartInput.Line defaultLine =
        new ChartInput.Line(
            "DefaultTrend", anchor, null, null, null, null, null, List.of(secondSeries));
    ChartInput.Pie pie =
        new ChartInput.Pie(
            "ShareChart", anchor, null, null, null, null, null, 180, List.of(secondSeries));
    ChartInput.Pie defaultPie =
        new ChartInput.Pie(
            "DefaultShare", anchor, null, null, null, null, null, null, List.of(secondSeries));
    ChartInput.Pie explicitPie =
        new ChartInput.Pie(
            "ExplicitShare",
            anchor,
            new ChartInput.Title.Text("Share"),
            new ChartInput.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            false,
            true,
            90,
            List.of(secondSeries));

    assertTrue(bar.title() instanceof ChartInput.Title.None);
    assertEquals(new ChartInput.Legend.Visible(ExcelChartLegendPosition.RIGHT), bar.legend());
    assertEquals(ExcelChartDisplayBlanksAs.GAP, bar.displayBlanksAs());
    assertTrue(bar.plotOnlyVisibleCells());
    assertFalse(bar.varyColors());
    assertEquals(ExcelChartBarDirection.COLUMN, bar.barDirection());
    assertEquals("B1", ((ChartInput.Title.Formula) bar.series().getFirst().title()).formula());
    assertTrue(secondSeries.title() instanceof ChartInput.Title.None);
    assertEquals(ExcelChartDisplayBlanksAs.ZERO, line.displayBlanksAs());
    assertFalse(line.plotOnlyVisibleCells());
    assertTrue(line.varyColors());
    assertEquals(180, pie.firstSliceAngle());
    assertTrue(defaultLine.title() instanceof ChartInput.Title.None);
    assertEquals(
        new ChartInput.Legend.Visible(ExcelChartLegendPosition.RIGHT), defaultLine.legend());
    assertEquals(ExcelChartDisplayBlanksAs.GAP, defaultLine.displayBlanksAs());
    assertTrue(defaultLine.plotOnlyVisibleCells());
    assertFalse(defaultLine.varyColors());
    assertTrue(defaultPie.title() instanceof ChartInput.Title.None);
    assertEquals(
        new ChartInput.Legend.Visible(ExcelChartLegendPosition.RIGHT), defaultPie.legend());
    assertEquals(ExcelChartDisplayBlanksAs.GAP, defaultPie.displayBlanksAs());
    assertTrue(defaultPie.plotOnlyVisibleCells());
    assertFalse(defaultPie.varyColors());
    assertNull(defaultPie.firstSliceAngle());
    assertEquals(new ChartInput.Title.Text("Share"), explicitPie.title());
    assertEquals(new ChartInput.Legend.Hidden(), explicitPie.legend());
    assertEquals(ExcelChartDisplayBlanksAs.ZERO, explicitPie.displayBlanksAs());
    assertFalse(explicitPie.plotOnlyVisibleCells());
    assertTrue(explicitPie.varyColors());
    assertEquals(90, explicitPie.firstSliceAngle());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartInput.Bar(
                " ", anchor, null, null, null, null, null, null, List.of(firstSeries)));
    assertThrows(
        NullPointerException.class,
        () ->
            new ChartInput.Line(
                "OpsChart", null, null, null, null, null, null, List.of(firstSeries)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartInput.Pie(
                "OpsChart", anchor, null, null, null, null, null, 361, List.of(firstSeries)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartInput.Pie(
                "OpsChart", anchor, null, null, null, null, null, -1, List.of(firstSeries)));
    assertThrows(IllegalArgumentException.class, () -> new ChartInput.Title.Text(" "));
    assertThrows(IllegalArgumentException.class, () -> new ChartInput.Title.Formula(" "));
    assertThrows(NullPointerException.class, () -> new ChartInput.Legend.Visible(null));
    assertThrows(IllegalArgumentException.class, () -> new ChartInput.DataSource(" "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartInput.Bar("OpsChart", anchor, null, null, null, null, null, null, List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ChartInput.Line(
                "OpsChart", anchor, null, null, null, null, null, List.of(firstSeries, null)));
  }

  @Test
  void autofilterInputsNormalizeAndValidateAcrossAllCriterionFamilies() {
    AutofilterSortConditionInput colorSort =
        new AutofilterSortConditionInput("B2:B9", true, null, new ColorInput("#aabbcc"), null);
    AutofilterSortConditionInput iconSort =
        new AutofilterSortConditionInput("C2:C9", false, "ICON", null, 2);
    AutofilterSortStateInput sortState =
        new AutofilterSortStateInput("A1:F9", null, true, null, List.of(colorSort, iconSort));
    AutofilterSortStateInput sortStateWithMethod =
        new AutofilterSortStateInput("A1:F9", true, false, "PINYIN", List.of(colorSort));

    assertEquals("", colorSort.sortBy());
    assertEquals("#AABBCC", colorSort.color().rgb());
    assertEquals("ICON", iconSort.sortBy());
    assertEquals(2, iconSort.iconId());
    assertFalse(sortState.caseSensitive());
    assertTrue(sortState.columnSort());
    assertEquals("", sortState.sortMethod());
    assertEquals(List.of(colorSort, iconSort), sortState.conditions());
    assertEquals("PINYIN", sortStateWithMethod.sortMethod());
    assertEquals(
        new AutofilterFilterCriterionInput.Dynamic("TODAY", null, null),
        new AutofilterFilterCriterionInput.Dynamic("TODAY", null, null));
    assertThrows(
        NullPointerException.class,
        () -> new AutofilterSortConditionInput(null, false, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterSortStateInput(" ", false, false, null, List.of(colorSort)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterSortConditionInput(" ", false, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterSortConditionInput("A1:A2", false, null, null, -1));
    assertThrows(
        NullPointerException.class,
        () -> new AutofilterSortStateInput("A1:F9", false, false, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterSortStateInput("A1:F9", false, false, null, List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new AutofilterSortStateInput(
                "A1:F9", false, false, null, List.of((AutofilterSortConditionInput) null)));

    AutofilterFilterCriterionInput.Values values =
        new AutofilterFilterCriterionInput.Values(List.of("Queued", "Ready"), true);
    AutofilterFilterCriterionInput.Custom custom =
        new AutofilterFilterCriterionInput.Custom(
            true, List.of(new AutofilterFilterCriterionInput.CustomConditionInput("equal", "Ada")));
    AutofilterFilterCriterionInput.Dynamic dynamic =
        new AutofilterFilterCriterionInput.Dynamic("TODAY", 1.0d, 2.0d);
    AutofilterFilterCriterionInput.Top10 top10 =
        new AutofilterFilterCriterionInput.Top10(10, true, false);
    AutofilterFilterCriterionInput.Color color =
        new AutofilterFilterCriterionInput.Color(false, new ColorInput(null, 3, null, 0.25d));
    AutofilterFilterCriterionInput.Icon icon =
        new AutofilterFilterCriterionInput.Icon("3TrafficLights1", 2);
    AutofilterFilterColumnInput column = new AutofilterFilterColumnInput(2L, null, values);

    assertEquals(List.of("Queued", "Ready"), values.values());
    assertTrue(custom.and());
    assertEquals("TODAY", dynamic.type());
    assertEquals(10, top10.value());
    assertEquals(3, color.color().theme());
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
        () -> new AutofilterFilterCriterionInput.Dynamic(" ", 1.0d, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionInput.Dynamic("TODAY", Double.POSITIVE_INFINITY, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionInput.Dynamic("TODAY", 1.0d, Double.NaN));
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
    ColorInput rgb = new ColorInput("#aabbcc");
    ColorInput themed = new ColorInput(null, 4, null, 0.5d);
    ColorInput indexed = new ColorInput(null, null, 64, null);
    assertEquals("#AABBCC", rgb.rgb());
    assertEquals(4, themed.theme());
    assertEquals(0.5d, themed.tint());
    assertEquals(64, indexed.indexed());
    assertThrows(IllegalArgumentException.class, () -> new ColorInput(null, null, null, null));
    assertThrows(IllegalArgumentException.class, () -> new ColorInput(" ", null, null, null));
    assertThrows(IllegalArgumentException.class, () -> new ColorInput("123456", null, null, null));
    assertThrows(IllegalArgumentException.class, () -> new ColorInput(null, -1, null, null));
    assertThrows(IllegalArgumentException.class, () -> new ColorInput(null, null, -1, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ColorInput("#AABBCC", null, null, Double.NEGATIVE_INFINITY));

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
                List.of(new ColorInput("#112233"))));
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
                List.of(new ColorInput("#112233"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.DataBarRule(
                false,
                new ColorInput("#112233"),
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
                new ColorInput("#112233"),
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
                new ColorInput("#112233"),
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
    CellGradientStopInput start = new CellGradientStopInput(0.0d, new ColorInput("#112233"));
    CellGradientStopInput finish =
        new CellGradientStopInput(1.0d, new ColorInput(null, 5, null, 0.2d));
    CellGradientFillInput gradient =
        new CellGradientFillInput(" path ", null, 0.1d, 0.2d, 0.3d, 0.4d, List.of(start, finish));
    assertGradientInputsNormalizeAndValidate(start, finish, gradient);
    assertFontInputsNormalizeAndValidate();
    assertFillInputsNormalizeAndValidate(gradient);
    assertBorderSideInputsNormalizeAndValidate();
  }

  private void assertGradientInputsNormalizeAndValidate(
      CellGradientStopInput start, CellGradientStopInput finish, CellGradientFillInput gradient) {
    assertEquals("PATH", gradient.type());
    assertEquals(List.of(start, finish), gradient.stops());
    assertEquals(
        "LINEAR",
        new CellGradientFillInput(null, null, null, null, null, null, List.of(start, finish))
            .type());
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientStopInput(-0.1d, new ColorInput("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientStopInput(Double.NaN, new ColorInput("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientStopInput(1.5d, new ColorInput("#112233")));
    assertThrows(NullPointerException.class, () -> new CellGradientStopInput(0.5d, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientFillInput(" ", null, null, null, null, null, List.of(start, finish)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellGradientFillInput(
                "diagonal", null, null, null, null, null, List.of(start, finish)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellGradientFillInput(
                "LINEAR",
                Double.POSITIVE_INFINITY,
                null,
                null,
                null,
                null,
                List.of(start, finish)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellGradientFillInput(
                "LINEAR", 45.0d, 0.1d, null, null, null, List.of(start, finish)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellGradientFillInput(
                "PATH", 45.0d, 0.1d, 0.2d, 0.3d, 0.4d, List.of(start, finish)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientFillInput("LINEAR", null, null, null, null, null, List.of(start)));
    assertThrows(
        NullPointerException.class,
        () ->
            new CellGradientFillInput(
                "LINEAR", null, null, null, null, null, List.of(start, null)));
  }

  private void assertFontInputsNormalizeAndValidate() {
    CellFontInput themedFont =
        new CellFontInput(null, null, null, null, null, 2, null, 0.4d, true, null);
    CellFontInput namedFont =
        new CellFontInput(null, null, "Calibri", null, null, null, null, null, null, null);
    CellFontInput indexedFont =
        new CellFontInput(null, null, null, null, null, null, 64, null, null, null);
    assertEquals(2, themedFont.fontColorTheme());
    assertEquals(0.4d, themedFont.fontColorTint());
    assertTrue(themedFont.underline());
    assertEquals("Calibri", namedFont.fontName());
    assertEquals(64, indexedFont.fontColorIndexed());
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFontInput(null, null, " ", null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFontInput(null, null, null, null, null, null, null, 0.4d, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFontInput(null, null, null, null, null, -1, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFontInput(null, null, null, null, null, null, -1, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFontInput(
                null, null, null, null, null, null, null, Double.POSITIVE_INFINITY, null, null));
  }

  private void assertFillInputsNormalizeAndValidate(CellGradientFillInput gradient) {
    CellFillInput gradientFill =
        new CellFillInput(null, null, null, null, null, null, null, null, null, gradient);
    CellFillInput themedFill =
        new CellFillInput(null, null, 2, null, 0.3d, null, null, null, null, null);
    CellFillInput fgIndexedFill =
        new CellFillInput(null, null, null, 64, null, null, null, null, null, null);
    CellFillInput fgIndexedTintedFill =
        new CellFillInput(null, null, null, 64, 0.3d, null, null, null, null, null);
    CellFillInput bgIndexedFill =
        new CellFillInput(
            ExcelFillPattern.FINE_DOTS, null, null, null, null, null, null, 64, null, null);
    assertEquals(gradient, gradientFill.gradient());
    assertEquals(2, themedFill.foregroundColorTheme());
    assertEquals(0.3d, themedFill.foregroundColorTint());
    assertEquals(64, fgIndexedFill.foregroundColorIndexed());
    assertEquals(0.3d, fgIndexedTintedFill.foregroundColorTint());
    assertEquals(64, bgIndexedFill.backgroundColorIndexed());
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillInput(null, null, null, null, 0.3d, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillInput(null, null, -1, null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillInput(null, null, null, -1, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFillInput(
                null, null, null, null, Double.POSITIVE_INFINITY, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFillInput(
                ExcelFillPattern.SOLID, null, null, null, null, null, null, null, null, gradient));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFillInput(null, "#112233", null, null, null, null, null, null, null, gradient));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillInput(null, null, 2, null, null, null, null, null, null, gradient));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillInput(null, null, null, 64, null, null, null, null, null, gradient));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFillInput(null, null, null, null, null, "#AABBCC", null, null, null, gradient));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillInput(null, null, null, null, null, null, 2, null, null, gradient));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillInput(null, null, null, null, null, null, null, 64, null, gradient));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFillInput(
                ExcelFillPattern.NONE, null, 2, null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFillInput(
                ExcelFillPattern.NONE, null, null, 64, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFillInput(
                ExcelFillPattern.NONE, null, null, null, null, null, 2, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFillInput(
                ExcelFillPattern.NONE, null, null, null, null, null, null, 64, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillInput(null, null, null, null, null, null, 2, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillInput(null, null, null, null, null, null, null, 64, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFillInput(
                ExcelFillPattern.SOLID, "#112233", null, null, null, null, 2, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFillInput(
                ExcelFillPattern.SOLID, "#112233", null, null, null, null, null, 64, null, null));
  }

  private void assertBorderSideInputsNormalizeAndValidate() {
    CellBorderSideInput noneStyleOnly = new CellBorderSideInput(ExcelBorderStyle.NONE);
    CellBorderSideInput borderSide = new CellBorderSideInput(null, null, 1, null, 0.15d);
    CellBorderSideInput indexedBorderSide = new CellBorderSideInput(null, null, null, 64, null);
    assertEquals(ExcelBorderStyle.NONE, noneStyleOnly.style());
    assertEquals(1, borderSide.colorTheme());
    assertEquals(0.15d, borderSide.colorTint());
    assertEquals(64, indexedBorderSide.colorIndexed());
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellBorderSideInput(null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellBorderSideInput(null, null, null, null, 0.15d));
    assertThrows(
        IllegalArgumentException.class, () -> new CellBorderSideInput(null, null, -1, null, null));
    assertThrows(
        IllegalArgumentException.class, () -> new CellBorderSideInput(null, null, null, -1, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellBorderSideInput(null, null, null, null, Double.POSITIVE_INFINITY));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellBorderSideInput(ExcelBorderStyle.NONE, null, 1, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellBorderSideInput(ExcelBorderStyle.NONE, null, null, 64, null));
  }

  @Test
  void printAndTableInputsNormalizeAndValidate() {
    PrintMarginsInput margins = new PrintMarginsInput(0.5d, 0.6d, 0.7d, 0.8d, 0.3d, 0.2d);
    PrintSetupInput explicitSetup =
        new PrintSetupInput(
            margins, true, true, null, 9, true, null, 2, true, 4, List.of(6), List.of(3));
    PrintSetupInput defaultLikeSetup =
        new PrintSetupInput(null, null, null, null, null, null, null, null, null, null, null, null);

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
    assertEquals(
        PrintSetupInput.defaults(),
        new PrintLayoutInput(null, null, null, null, null, null, null, null).setup());
    assertEquals(
        explicitSetup,
        new PrintLayoutInput(null, null, null, null, null, null, null, explicitSetup).setup());
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
                margins, null, null, null, -1, null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupInput(
                margins, null, null, null, null, null, null, -1, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupInput(
                margins, null, null, null, null, null, null, null, null, -1, null, null));
    assertThrows(
        NullPointerException.class,
        () ->
            new PrintSetupInput(
                margins,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                List.of((Integer) null),
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupInput(
                margins, null, null, null, null, null, null, null, null, null, List.of(-1), null));

    TableColumnInput tableColumn = new TableColumnInput(0, null, null, " SUM ", null);
    TableColumnInput blankTotalsFunction = new TableColumnInput(0, null, null, " ", null);
    TableColumnInput nullTotalsFunction = new TableColumnInput(0, null, null, null, null);
    assertEquals("", tableColumn.uniqueName());
    assertEquals("", tableColumn.totalsRowLabel());
    assertEquals("sum", tableColumn.totalsRowFunction());
    assertEquals("", tableColumn.calculatedColumnFormula());
    assertEquals("", blankTotalsFunction.totalsRowFunction());
    assertEquals("", nullTotalsFunction.totalsRowFunction());
    assertThrows(
        IllegalArgumentException.class, () -> new TableColumnInput(-1, null, null, null, null));

    TableInput defaultedTable =
        new TableInput(
            "BudgetTable",
            "Budget",
            "A1:C4",
            null,
            null,
            new TableStyleInput.None(),
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
    TableInput explicitTable =
        new TableInput(
            "BudgetTable",
            "Budget",
            "A1:C4",
            false,
            false,
            new TableStyleInput.None(),
            "Table note",
            true,
            true,
            true,
            "HeaderStyle",
            "DataStyle",
            "TotalsStyle",
            List.of());
    assertFalse(defaultedTable.showTotalsRow());
    assertTrue(defaultedTable.hasAutofilter());
    assertEquals("", defaultedTable.comment());
    assertEquals(List.of(), defaultedTable.columns());
    assertEquals("Table note", explicitTable.comment());
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
                null,
                null,
                null,
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
            new ColorInput("#112233"),
            new SheetOutlineSummaryInput(false, false),
            new SheetDefaultsInput(12, 18.5d),
            List.of(
                new IgnoredErrorInput(
                    "B2:B12",
                    List.of(
                        ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT,
                        ExcelIgnoredErrorType.FORMULA))));
    SheetPresentationInput defaultedPresentation =
        new SheetPresentationInput(null, null, null, null, null);

    assertEquals(
        new SheetDisplayInput(false, false, true, true, true), explicitPresentation.display());
    assertEquals(new ColorInput("#112233"), explicitPresentation.tabColor());
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
            new SheetPresentationInput(
                null,
                null,
                null,
                null,
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
                    "Amount", ExcelPivotDataConsolidateFunction.SUM, null, "#,##0.00")));
    PivotTableInput tableSourceInput =
        new PivotTableInput(
            "Sales Table Pivot",
            "Report",
            new PivotTableInput.Source.Table("SalesTable2026"),
            new PivotTableInput.Anchor("D7"),
            null,
            null,
            null,
            List.of(
                new PivotTableInput.DataField(
                    "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total Amount", null)));
    PivotTableSelection.ByNames selection =
        new PivotTableSelection.ByNames(List.of("Sales Pivot 2026", "Ops Pivot"));

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
    PivotTableInput.DataField blankDisplayField =
        new PivotTableInput.DataField("Amount", ExcelPivotDataConsolidateFunction.SUM, " ", null);
    assertEquals("Amount", blankDisplayField.displayName());
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
    assertThrows(NullPointerException.class, () -> new PivotTableSelection.ByNames(null));
    assertThrows(IllegalArgumentException.class, () -> new PivotTableSelection.ByNames(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PivotTableSelection.ByNames(List.of("Sales Pivot 2026", "sales pivot 2026")));
  }
}
