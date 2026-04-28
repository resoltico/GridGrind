package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.ExcelStyleTestAccess.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Focused contract tests for advanced workbook-core value objects added during parity work. */
class ExcelAdvancedValueObjectTest {
  @Test
  void autofilterValueObjectsNormalizeAndValidate() {
    var customCondition = new ExcelAutofilterFilterCriterion.CustomCondition("equal", "Ada");
    var values =
        new ExcelAutofilterFilterCriterion.Values(new ArrayList<>(List.of("Queued", "")), true);
    var custom = new ExcelAutofilterFilterCriterion.Custom(true, List.of(customCondition));
    var dynamic = new ExcelAutofilterFilterCriterion.Dynamic("today", 1.0d, 2.0d);
    var top10 = new ExcelAutofilterFilterCriterion.Top10(10, false, true);
    var color = new ExcelAutofilterFilterCriterion.Color(false, ExcelColor.indexed(10, -0.25d));
    var icon = new ExcelAutofilterFilterCriterion.Icon("3TrafficLights1", 2);
    var column = new ExcelAutofilterFilterColumn(4L, false, color);
    var sortCondition =
        new ExcelAutofilterSortCondition("A2:A5", true, null, ExcelColor.rgb("#AABBCC"), 1);
    var sortState = new ExcelAutofilterSortState("A1:F5", true, true, "", List.of(sortCondition));

    assertEquals(List.of("Queued", ""), values.values());
    assertEquals(List.of(customCondition), custom.conditions());
    assertEquals("today", dynamic.type());
    assertEquals(10, top10.value());
    assertEquals(ExcelColor.indexed(10, -0.25d), color.color());
    assertEquals("3TrafficLights1", icon.iconSet());
    assertEquals(4L, column.columnId());
    assertEquals("", sortCondition.sortBy());
    assertEquals(List.of(sortCondition), sortState.conditions());

    List<String> valuesWithNull = new ArrayList<>(Arrays.asList("Queued", null));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelAutofilterFilterCriterion.Values(valuesWithNull, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterion.Custom(false, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterion.Dynamic(" ", null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterion.Dynamic("today", Double.NaN, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterion.Top10(0, true, false));
    assertThrows(
        NullPointerException.class, () -> new ExcelAutofilterFilterCriterion.Color(true, null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelAutofilterFilterCriterion.Icon(" ", 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterion.Icon("3TrafficLights1", -1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterion.CustomCondition(" ", "Ada"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterion.CustomCondition("equal", " "));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelAutofilterFilterColumn(-1L, true, values));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelAutofilterSortCondition(null, false, "", null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterSortCondition(" ", false, "", null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterSortCondition("A1", false, "", null, -1));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelAutofilterSortState("A1:B2", false, false, null, List.of(sortCondition)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterSortState(" ", false, false, "", List.of(sortCondition)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterSortState("A1:B2", false, false, "", List.of()));
  }

  @Test
  void styleValueObjectsNormalizeDefaultsAndValidate() {
    var themeColor = ExcelColor.theme(6, -0.35d);
    var indexedColor = ExcelColor.indexed(10);
    var gradient =
        ExcelGradientFill.path(
            0.1d,
            0.2d,
            0.3d,
            0.4d,
            List.of(
                new ExcelGradientStop(0.0d, ExcelColor.rgb("#112233")),
                new ExcelGradientStop(1.0d, indexedColor)));
    var gradientFill = ExcelCellFill.gradient(gradient);
    var printSetup =
        new ExcelPrintSetup(
            new ExcelPrintMargins(0.5d, 0.6d, 0.7d, 0.8d, 0.2d, 0.3d),
            false,
            true,
            true,
            8,
            true,
            true,
            2,
            true,
            4,
            List.of(1, 3),
            List.of(2));
    var printLayout =
        new ExcelPrintLayout(
            new ExcelPrintLayout.Area.Range("A1:C3"),
            ExcelPrintOrientation.LANDSCAPE,
            new ExcelPrintLayout.Scaling.Fit(1, 0),
            new ExcelPrintLayout.TitleRows.Band(0, 0),
            new ExcelPrintLayout.TitleColumns.Band(0, 1),
            new ExcelHeaderFooterText("Left", "Center", "Right"),
            new ExcelHeaderFooterText("Footer", "", ""));

    assertEquals("#AA00CC", ExcelColor.rgb("#aa00cc").rgb());
    assertEquals(6, themeColor.theme());
    assertEquals(10, indexedColor.indexed());
    assertEquals("PATH", gradientType(gradient));
    assertEquals(0.1d, gradientLeft(gradient));
    assertEquals(0.2d, gradientRight(gradient));
    assertEquals(0.3d, gradientTop(gradient));
    assertEquals(0.4d, gradientBottom(gradient));
    assertEquals(2, gradient.stops().size());
    assertSame(gradient, gradientFill.gradient());
    assertEquals(8, printSetup.paperSize());
    assertEquals(ExcelPrintSetup.defaults(), printLayout.setup());
    assertEquals(new ExcelPrintLayout.Area.None(), ExcelPrintLayout.defaults().printArea());

    assertThrows(IllegalArgumentException.class, () -> ExcelColor.rgb(" "));
    assertThrows(IllegalArgumentException.class, () -> ExcelColor.rgb("#12345"));
    assertThrows(IllegalArgumentException.class, () -> ExcelColor.theme(-1));
    assertThrows(IllegalArgumentException.class, () -> ExcelColor.indexed(-1));
    assertThrows(NullPointerException.class, () -> ExcelColor.rgb(null));
    assertThrows(
        IllegalArgumentException.class, () -> ExcelColor.theme(1, Double.POSITIVE_INFINITY));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelGradientFill.path(Double.NaN, null, null, null, List.of()));
    assertEquals(
        "LINEAR",
        gradientType(
            ExcelGradientFill.linear(
                null,
                List.of(
                    new ExcelGradientStop(0.0d, ExcelColor.rgb("#112233")),
                    new ExcelGradientStop(1.0d, ExcelColor.rgb("#445566"))))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ExcelGradientFill.linear(
                Double.POSITIVE_INFINITY,
                List.of(
                    new ExcelGradientStop(0.0d, ExcelColor.rgb("#112233")),
                    new ExcelGradientStop(1.0d, ExcelColor.rgb("#445566")))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ExcelGradientFill.path(
                Double.POSITIVE_INFINITY,
                null,
                null,
                null,
                List.of(
                    new ExcelGradientStop(0.0d, ExcelColor.rgb("#112233")),
                    new ExcelGradientStop(1.0d, ExcelColor.rgb("#445566")))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ExcelGradientFill.linear(
                Double.NaN,
                List.of(
                    new ExcelGradientStop(0.0d, ExcelColor.rgb("#112233")),
                    new ExcelGradientStop(1.0d, ExcelColor.rgb("#445566")))));
    assertThrows(NullPointerException.class, () -> ExcelGradientFill.linear(null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ExcelGradientFill.linear(
                null, List.of(new ExcelGradientStop(0.0d, ExcelColor.rgb("#112233")))));
    assertThrows(
        NullPointerException.class,
        () ->
            ExcelGradientFill.linear(
                null, Arrays.asList(new ExcelGradientStop(0.0d, ExcelColor.rgb("#112233")), null)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelGradientStop(-0.1d, ExcelColor.rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelGradientStop(1.5d, ExcelColor.rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelGradientStop(Double.NaN, ExcelColor.rgb("#112233")));
    assertThrows(NullPointerException.class, () -> ExcelCellFill.pattern(null));
    assertThrows(
        NullPointerException.class,
        () -> ExcelCellFill.patternBackground(null, ExcelColor.rgb("#112233")));
    assertThrows(NullPointerException.class, () -> ExcelCellFill.gradient(null));
    assertThrows(
        NullPointerException.class,
        () -> ExcelCellFill.patternForeground(null, ExcelColor.rgb("#112233")));
    assertThrows(
        NullPointerException.class,
        () -> ExcelCellFill.patternBackground(null, ExcelColor.rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelCellFill.patternForeground(ExcelFillPattern.NONE, ExcelColor.rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ExcelCellFill.patternColors(
                ExcelFillPattern.SOLID, ExcelColor.rgb("#112233"), ExcelColor.rgb("#445566")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPrintSetup(
                printSetup.margins(),
                false,
                false,
                false,
                -1,
                false,
                false,
                1,
                false,
                1,
                List.of(),
                List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPrintSetup(
                printSetup.margins(),
                false,
                false,
                false,
                1,
                false,
                false,
                -1,
                false,
                1,
                List.of(),
                List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPrintSetup(
                printSetup.margins(),
                false,
                false,
                false,
                1,
                false,
                false,
                1,
                false,
                -1,
                List.of(),
                List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPrintSetup(
                printSetup.margins(),
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
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelPrintMargins(0, 0, Double.NaN, 0, 0, 0));
  }

  @Test
  void drawingValueObjectsNormalizeDefaultsAndValidate() {
    ExcelBinaryData binary = new ExcelBinaryData(new byte[] {1, 2, 3});
    ExcelDrawingMarker from = new ExcelDrawingMarker(1, 2, 3, 4);
    ExcelDrawingMarker to = new ExcelDrawingMarker(4, 6);
    ExcelDrawingAnchor.TwoCell twoCell =
        new ExcelDrawingAnchor.TwoCell(from, to, ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
    ExcelPictureDefinition picture =
        new ExcelPictureDefinition("OpsPicture", binary, ExcelPictureFormat.PNG, twoCell, null);
    ExcelShapeDefinition shape =
        new ExcelShapeDefinition(
            "OpsShape", ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, twoCell, " rect ", "Queue");
    ExcelEmbeddedObjectDefinition embeddedObject =
        new ExcelEmbeddedObjectDefinition(
            "OpsEmbed",
            "Payload",
            "payload.txt",
            "payload.txt",
            new ExcelBinaryData("payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
            ExcelPictureFormat.PNG,
            binary,
            twoCell);
    ExcelDrawingObjectSnapshot.Shape groupSnapshot =
        new ExcelDrawingObjectSnapshot.Shape(
            "OpsGroup",
            new ExcelDrawingAnchor.Absolute(1L, 2L, 10L, 20L, null),
            ExcelDrawingShapeKind.GROUP,
            null,
            null,
            2);
    ExcelDrawingObjectPayload.EmbeddedObject payload =
        new ExcelDrawingObjectPayload.EmbeddedObject(
            "OpsEmbed",
            ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
            "application/octet-stream",
            "payload.txt",
            "abc123",
            new ExcelBinaryData("payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
            "Payload",
            "payload.txt");

    assertNotEquals(binary.bytes(), binary.bytes());
    assertEquals(3, binary.size());
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE, twoCell.behavior());
    assertEquals(ExcelPictureFormat.PNG, picture.format());
    assertEquals("rect", shape.presetGeometryToken());
    assertEquals("payload.txt", embeddedObject.fileName());
    assertEquals(2, groupSnapshot.childCount());
    assertEquals("Payload", payload.label());
    assertThrows(IllegalArgumentException.class, () -> new ExcelBinaryData(new byte[0]));
    assertThrows(IllegalArgumentException.class, () -> new ExcelDrawingMarker(-1, 0, 0, 0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(2, 2, 0, 0), new ExcelDrawingMarker(1, 2, 0, 0), null));
    assertDoesNotThrow(
        () ->
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(2, 1, 0, 3), new ExcelDrawingMarker(2, 3, 0, 4), null));
    assertDoesNotThrow(
        () ->
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(1, 2, 3, 0), new ExcelDrawingMarker(3, 2, 4, 0), null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelShapeDefinition(
                "OpsShape", ExcelAuthoredDrawingShapeKind.CONNECTOR, twoCell, "rect", null));
    ExcelShapeDefinition nullPresetShape =
        new ExcelShapeDefinition(
            "OpsNullPreset", ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, twoCell, null, null);
    assertEquals("rect", nullPresetShape.presetGeometryToken());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectPayload.EmbeddedObject(
                "OpsEmbed",
                ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                "application/octet-stream",
                " ",
                "abc123",
                binary,
                null,
                null));
  }

  @Test
  void namedRangeCommentAndTableContractsSupportAdvancedForms() {
    var authoringComment =
        new ExcelComment(
            "Quarterly Review",
            "GridGrind",
            true,
            new ExcelRichText(
                List.of(
                    new ExcelRichTextRun("Quarterly", null),
                    new ExcelRichTextRun(
                        " Review",
                        new ExcelCellFont(
                            true, false, "Aptos", null, ExcelColor.rgb("#112233"), null, null)))),
            new ExcelCommentAnchor(1, 2, 4, 6));
    var commentSnapshot =
        new ExcelCommentSnapshot(
            "Quarterly Review",
            "GridGrind",
            true,
            new ExcelRichTextSnapshot(
                List.of(
                    new ExcelRichTextRunSnapshot(
                        "Quarterly",
                        new ExcelCellFontSnapshot(
                            false,
                            false,
                            "Aptos",
                            ExcelFontHeight.fromPoints(java.math.BigDecimal.valueOf(11)),
                            null,
                            false,
                            false)),
                    new ExcelRichTextRunSnapshot(
                        " Review",
                        new ExcelCellFontSnapshot(
                            true,
                            false,
                            "Aptos",
                            ExcelFontHeight.fromPoints(java.math.BigDecimal.valueOf(11)),
                            ExcelColorSnapshot.rgb("#112233"),
                            false,
                            false)))),
            new ExcelCommentAnchorSnapshot(1, 2, 4, 6));
    var formulaTarget = new ExcelNamedRangeTarget("SUM(Ops!$B$2:$B$5)");
    var table =
        new ExcelTableDefinition(
            "Queue",
            "Ops",
            "A1:B4",
            true,
            true,
            new ExcelTableStyle.None(),
            null,
            true,
            true,
            false,
            null,
            null,
            null,
            List.of(new ExcelTableColumnDefinition(1, null, null, " SUM ", "B2*2")));
    var protection = new ExcelWorkbookProtectionSettings(true, false, true, "book", "review");
    ExcelComment copiedComment = commentSnapshot.toAuthoringComment();

    assertEquals("SUM(Ops!$B$2:$B$5)", formulaTarget.refersToFormula());
    assertEquals(authoringComment.text(), copiedComment.text());
    assertEquals(authoringComment.author(), copiedComment.author());
    assertEquals(authoringComment.visible(), copiedComment.visible());
    assertEquals(authoringComment.anchor(), copiedComment.anchor());
    assertEquals(authoringComment.runs().plainText(), copiedComment.runs().plainText());
    assertEquals(ExcelColor.rgb("#112233"), copiedComment.runs().runs().get(1).font().fontColor());
    assertEquals("", table.comment());
    assertEquals("", table.headerRowCellStyle());
    assertEquals("sum", table.columns().getFirst().totalsRowFunction());
    assertEquals("B2*2", table.columns().getFirst().calculatedColumnFormula());
    assertEquals("book", protection.workbookPassword());
    assertEquals("review", protection.revisionsPassword());
    assertEquals("A1", new ExcelNamedRangeTarget("Ops", "A1").range());
    assertEquals("Ops!$A$1", new ExcelNamedRangeTarget("Ops", "A1").refersToFormula());
    assertEquals("A1:B2", new ExcelNamedRangeTarget("Ops", "$A$1:$B$2").range());
    ExcelComment plainComment =
        new ExcelCommentSnapshot("Done", "GridGrind", false, null, null).toPlainComment();
    assertEquals("Done", plainComment.text());
    assertEquals("GridGrind", plainComment.author());
    ExcelComment plainAuthoringComment =
        new ExcelCommentSnapshot("Done", "GridGrind", false, null, null).toAuthoringComment();
    assertEquals("Done", plainAuthoringComment.text());
    assertNull(plainAuthoringComment.runs());
    assertNull(plainAuthoringComment.anchor());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelComment(
                "Quarterly Review",
                "GridGrind",
                true,
                new ExcelRichText(List.of(new ExcelRichTextRun("Mismatch", null))),
                null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelNamedRangeTarget(" "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelNamedRangeTarget("Ops", "A1", "SUM(Ops!$A$1)"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelWorkbookProtectionSettings(false, false, false, " ", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelWorkbookProtectionSettings(false, false, false, null, " "));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelTableColumnDefinition(-1, "", "", "", ""));
    ExcelTableColumnDefinition normalizedColumn =
        new ExcelTableColumnDefinition(0, null, null, null, null);
    assertEquals("", normalizedColumn.uniqueName());
    assertEquals("", normalizedColumn.totalsRowLabel());
    assertEquals("", normalizedColumn.totalsRowFunction());
    assertEquals("", normalizedColumn.calculatedColumnFormula());
    ExcelTableDefinition normalizedTable =
        new ExcelTableDefinition(
            "Queue",
            "Ops",
            "A1:B2",
            false,
            true,
            new ExcelTableStyle.None(),
            null,
            false,
            false,
            false,
            null,
            null,
            null,
            null);
    assertEquals("", normalizedTable.comment());
    assertEquals(List.of(), normalizedTable.columns());
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelTableDefinition(
                "Queue",
                "Ops",
                "A1:B2",
                false,
                true,
                new ExcelTableStyle.None(),
                "",
                false,
                false,
                false,
                "",
                "",
                "",
                Arrays.asList((ExcelTableColumnDefinition) null)));
  }

  @Test
  void conditionalFormattingAndSupportValueObjectsNormalizeAndValidate() throws Exception {
    assertConditionalFormattingSupportValueObjects();
    assertRgbColorSupportValueObjects();
    assertFormulaSheetRenameSupportValueObjects();
  }

  private static void assertConditionalFormattingSupportValueObjects() {
    for (ExcelConditionalFormattingIconSet iconSet : ExcelConditionalFormattingIconSet.values()) {
      assertEquals(
          iconSet,
          ExcelConditionalFormattingPoiBridge.fromPoi(
              ExcelConditionalFormattingPoiBridge.toPoi(iconSet)));
      assertEquals(
          ExcelConditionalFormattingPoiBridge.toPoi(iconSet).num, iconSet.thresholdCount());
    }
    for (ExcelConditionalFormattingThresholdType type :
        ExcelConditionalFormattingThresholdType.values()) {
      assertEquals(
          type,
          ExcelConditionalFormattingPoiBridge.fromPoi(
              ExcelConditionalFormattingPoiBridge.toPoi(type)));
    }

    var min =
        new ExcelConditionalFormattingThreshold(
            ExcelConditionalFormattingThresholdType.MIN, null, null);
    var percentile =
        new ExcelConditionalFormattingThreshold(
            ExcelConditionalFormattingThresholdType.PERCENTILE, null, 50.0d);
    var max =
        new ExcelConditionalFormattingThreshold(
            ExcelConditionalFormattingThresholdType.MAX, null, null);
    var colorScale =
        new ExcelConditionalFormattingRule.ColorScaleRule(
            List.of(min, percentile, max),
            List.of(
                ExcelColor.rgb("#112233"), ExcelColor.rgb("#445566"), ExcelColor.rgb("#778899")),
            false);
    var dataBar =
        new ExcelConditionalFormattingRule.DataBarRule(
            ExcelColor.rgb("#102030"), true, 10, 90, min, max, false);
    var iconSetRule =
        new ExcelConditionalFormattingRule.IconSetRule(
            ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
            true,
            true,
            List.of(
                new ExcelConditionalFormattingThreshold(
                    ExcelConditionalFormattingThresholdType.PERCENT, null, 0.0d),
                percentile,
                max),
            false);
    var top10 = new ExcelConditionalFormattingRule.Top10Rule(10, true, false, false, null);
    var anchor = new ExcelCommentAnchor(1, 2, 3, 4);

    assertEquals(3, colorScale.thresholds().size());
    assertEquals(90, dataBar.widthMax());
    assertTrue(iconSetRule.iconOnly());
    assertEquals(10, top10.rank());
    assertEquals(1, anchor.firstColumn());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingThreshold(
                ExcelConditionalFormattingThresholdType.NUMBER, " ", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingThreshold(
                ExcelConditionalFormattingThresholdType.NUMBER, null, Double.NaN));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRule.ColorScaleRule(
                List.of(min), List.of(ExcelColor.rgb("#112233")), false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRule.ColorScaleRule(
                List.of(min, max), List.of(ExcelColor.rgb("#112233")), false));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelConditionalFormattingRule.ColorScaleRule(
                null, List.of(ExcelColor.rgb("#112233")), false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRule.DataBarRule(
                ExcelColor.rgb("#102030"), false, 10, 9, min, max, false));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelConditionalFormattingRule.DataBarRule(null, false, 0, 0, min, max, false));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelConditionalFormattingRule.DataBarRule(
                ExcelColor.rgb("#102030"), false, 0, 0, null, max, false));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelConditionalFormattingRule.DataBarRule(
                ExcelColor.rgb("#102030"), false, 0, 0, min, null, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRule.DataBarRule(
                ExcelColor.rgb("#102030"), false, -1, 0, min, max, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRule.DataBarRule(
                ExcelColor.rgb("#102030"), false, 0, -1, min, max, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRule.IconSetRule(
                ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
                false,
                false,
                List.of(min, max),
                false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelConditionalFormattingRule.Top10Rule(0, false, false, false, null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelCommentAnchor(-1, 0, 1, 1));
    assertThrows(IllegalArgumentException.class, () -> new ExcelCommentAnchor(1, 0, 0, 1));
    assertThrows(IllegalArgumentException.class, () -> new ExcelCommentAnchor(1, 2, 1, 1));
    assertThrows(IllegalArgumentException.class, () -> new ExcelPrintMargins(-0.1d, 0, 0, 0, 0, 0));
  }

  private static void assertRgbColorSupportValueObjects() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      assertEquals(
          Optional.of("#AABBCC"), ExcelRgbColorSupport.normalizeRgbHex("#aabbcc", "color"));
      assertThrows(
          NullPointerException.class, () -> ExcelRgbColorSupport.normalizeRgbHex("#AABBCC", null));
      assertThrows(
          IllegalArgumentException.class, () -> ExcelRgbColorSupport.normalizeRgbHex(" ", "color"));
      assertEquals(Optional.empty(), ExcelRgbColorSupport.toRgbHex(null));
      assertEquals(
          Optional.of("#112233"),
          ExcelRgbColorSupport.toRgbHex(
              new XSSFColor(
                  new byte[] {0x11, 0x22, 0x33}, workbook.getStylesSource().getIndexedColors())));
      assertEquals(
          Optional.empty(),
          ExcelRgbColorSupport.toRgbHex(
              new XSSFColor(
                  new byte[] {0x11, 0x22}, workbook.getStylesSource().getIndexedColors())));
      assertEquals(
          Optional.of("#AABBCC"),
          ExcelRgbColorSupport.toRgbHex(ExcelRgbColorSupport.toXssfColor(workbook, "#AABBCC")));
      assertThrows(
          IllegalArgumentException.class, () -> ExcelRgbColorSupport.toXssfColor(workbook, null));
      assertThrows(
          NullPointerException.class, () -> ExcelRgbColorSupport.toXssfColor(null, "#AABBCC"));
      assertThrows(
          IllegalArgumentException.class,
          () -> ExcelRgbColorSupport.toXssfColor(workbook, "#ABCDE"));
    }
  }

  private static void assertFormulaSheetRenameSupportValueObjects() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Source");
      workbook.createSheet("Replica");
      workbook.createSheet("Ledger");
      assertEquals(
          "",
          ExcelFormulaSheetRenameSupport.renameSheet(
              workbook, "", org.apache.poi.ss.formula.FormulaType.CELL, 0, "Source", "Replica"));
      assertEquals(
          "SUM(Source!A1)",
          ExcelFormulaSheetRenameSupport.renameSheet(
              workbook,
              "SUM(Source!A1)",
              org.apache.poi.ss.formula.FormulaType.CELL,
              0,
              "Source",
              "Source"));
      assertEquals(
          "SUM(Replica!A1,Replica:Ledger!B2)",
          ExcelFormulaSheetRenameSupport.renameSheet(
              workbook,
              "SUM(Source!A1,Source:Ledger!B2)",
              org.apache.poi.ss.formula.FormulaType.CELL,
              0,
              "Source",
              "Replica"));
      assertEquals(
          "SUM(Replica!A1)",
          ExcelFormulaSheetRenameSupport.renameSheet(
              workbook,
              "SUM(Source!A1)",
              org.apache.poi.ss.formula.FormulaType.CELL,
              0,
              "Source",
              "Replica"));
      assertEquals(
          "SUM(Ledger:Replica!B2)",
          ExcelFormulaSheetRenameSupport.renameSheet(
              workbook,
              "SUM(Ledger:Source!B2)",
              org.apache.poi.ss.formula.FormulaType.CELL,
              0,
              "Source",
              "Replica"));
      assertEquals(
          "SUM(Replica!A1)",
          ExcelFormulaSheetRenameSupport.renameSheet(
              workbook,
              "SUM(Source!A1)",
              org.apache.poi.ss.formula.FormulaType.CONDFORMAT,
              0,
              "Source",
              "Replica"));

      var localName = workbook.createName();
      localName.setNameName("LocalTotal");
      localName.setSheetIndex(workbook.getSheetIndex("Source"));
      localName.setRefersToFormula("Source!$A$1");
      assertEquals(
          "Replica!LocalTotal+1",
          ExcelFormulaSheetRenameSupport.renameSheet(
              workbook,
              "Source!LocalTotal+1",
              org.apache.poi.ss.formula.FormulaType.CELL,
              0,
              "Source",
              "Replica"));
    }
  }

  @Test
  void formulaSheetRenameSupportPreservesExternalWorkbookReferences() throws Exception {
    try (XSSFWorkbook referencedWorkbook = new XSSFWorkbook();
        XSSFWorkbook workbook = new XSSFWorkbook()) {
      referencedWorkbook.createSheet("Source");
      workbook.linkExternalWorkbook("ext.xlsx", referencedWorkbook);
      workbook.createSheet("Source");
      workbook.createSheet("Replica");
      // External workbook refs produce Pxg with getExternalWorkbookNumber() >= 1;
      // renameSheet must leave them untouched.
      String result =
          ExcelFormulaSheetRenameSupport.renameSheet(
              workbook,
              "[ext.xlsx]Source!$A$1",
              org.apache.poi.ss.formula.FormulaType.CELL,
              0,
              "Source",
              "Replica");
      assertFalse(result.contains("Replica"));
    }
  }
}
