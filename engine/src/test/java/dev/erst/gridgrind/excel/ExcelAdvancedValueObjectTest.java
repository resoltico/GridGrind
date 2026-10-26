package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.ExcelStyleTestAccess.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.CommentAnchorReport;
import dev.erst.gridgrind.core.model.drawing.CommentAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerReport;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.drawing.ResolvedEmbeddedObject;
import dev.erst.gridgrind.core.model.drawing.ResolvedPicture;
import dev.erst.gridgrind.core.model.drawing.ResolvedShape;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.PrintAreaSpec;
import dev.erst.gridgrind.core.model.print.PrintMarginsSpec;
import dev.erst.gridgrind.core.model.print.PrintScalingSpec;
import dev.erst.gridgrind.core.model.print.PrintSetupSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec;
import dev.erst.gridgrind.core.model.print.ResolvedPrintLayout;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFillSpec;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellGradientFillSpec;
import dev.erst.gridgrind.core.model.style.CellGradientStopSpec;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.style.ResolvedRichText;
import dev.erst.gridgrind.core.model.style.ResolvedRichTextRun;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableColumnSpec;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingShapeKind;
import dev.erst.gridgrind.core.vocabulary.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectPayload;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectSnapshot;
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
    var color = new ExcelAutofilterFilterCriterion.Color(false, ColorSpec.indexed(10, -0.25d));
    var icon = new ExcelAutofilterFilterCriterion.Icon("3TrafficLights1", 2);
    var column = new ExcelAutofilterFilterColumn(4L, false, color);
    var sortCondition =
        new ExcelAutofilterSortCondition.CellColor(
            "A2:A5", true, ColorSpec.rgb(new RgbHex("#AABBCC")));
    var sortState =
        new ExcelAutofilterSortState("A1:F5", true, true, Optional.empty(), List.of(sortCondition));

    assertEquals(List.of("Queued", ""), values.values());
    assertEquals(List.of(customCondition), custom.conditions());
    assertEquals("today", dynamic.type());
    assertEquals(10, top10.value());
    assertEquals(ColorSpec.indexed(10, -0.25d), color.color());
    assertEquals("3TrafficLights1", icon.iconSet());
    assertEquals(4L, column.columnId());
    assertInstanceOf(ExcelAutofilterSortCondition.CellColor.class, sortCondition);
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
        NullPointerException.class, () -> new ExcelAutofilterSortCondition.Value(null, false));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelAutofilterSortCondition.Value(" ", false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterSortCondition.Icon("A1", false, -1));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelAutofilterSortState("A1:B2", false, false, null, List.of(sortCondition)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelAutofilterSortState(
                " ", false, false, Optional.empty(), List.of(sortCondition)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterSortState("A1:B2", false, false, Optional.empty(), List.of()));
  }

  @Test
  void styleValueObjectsNormalizeDefaultsAndValidate() {
    var themeColor = ColorSpec.theme(6, -0.35d);
    var indexedColor = ColorSpec.indexed(10);
    var gradient =
        CellGradientFillSpec.path(
            Optional.of(0.1d),
            Optional.of(0.2d),
            Optional.of(0.3d),
            Optional.of(0.4d),
            List.of(
                new CellGradientStopSpec(0.0d, ColorSpec.rgb(new RgbHex("#112233"))),
                new CellGradientStopSpec(1.0d, indexedColor)));
    var gradientFill = CellFillSpec.gradient(gradient);
    var printSetup =
        new PrintSetupSpec(
            new PrintMarginsSpec(0.5d, 0.6d, 0.7d, 0.8d, 0.2d, 0.3d),
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
        new ResolvedPrintLayout(
            new PrintAreaSpec.Range(new RangeRef("A1:C3")),
            ExcelPrintOrientation.LANDSCAPE,
            new PrintScalingSpec.Fit(1, 0),
            new PrintTitleRowsSpec.Band(0, 0),
            new PrintTitleColumnsSpec.Band(0, 1),
            new HeaderFooterTextReport("Left", "Center", "Right"),
            new HeaderFooterTextReport("Footer", "", ""));

    assertEquals("#AA00CC", ColorSpec.rgb(new RgbHex("#aa00cc")).rgb().value());
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
    assertEquals(PrintSetupSpec.defaults(), printLayout.setup());
    assertEquals(new PrintAreaSpec.None(), ResolvedPrintLayout.defaults().printArea());

    assertThrows(IllegalArgumentException.class, () -> ColorSpec.rgb(new RgbHex(" ")));
    assertThrows(IllegalArgumentException.class, () -> ColorSpec.rgb(new RgbHex("#12345")));
    assertThrows(IllegalArgumentException.class, () -> ColorSpec.theme(-1));
    assertThrows(IllegalArgumentException.class, () -> ColorSpec.indexed(-1));
    assertThrows(NullPointerException.class, () -> ColorSpec.rgb(new RgbHex(null)));
    assertThrows(
        IllegalArgumentException.class, () -> ColorSpec.theme(1, Double.POSITIVE_INFINITY));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellGradientFillSpec.path(
                Optional.of(Double.NaN),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                List.of()));
    assertEquals(
        "LINEAR",
        gradientType(
            CellGradientFillSpec.linear(
                Optional.empty(),
                List.of(
                    new CellGradientStopSpec(0.0d, ColorSpec.rgb(new RgbHex("#112233"))),
                    new CellGradientStopSpec(1.0d, ColorSpec.rgb(new RgbHex("#445566")))))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellGradientFillSpec.linear(
                Optional.of(Double.POSITIVE_INFINITY),
                List.of(
                    new CellGradientStopSpec(0.0d, ColorSpec.rgb(new RgbHex("#112233"))),
                    new CellGradientStopSpec(1.0d, ColorSpec.rgb(new RgbHex("#445566"))))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellGradientFillSpec.path(
                Optional.of(Double.POSITIVE_INFINITY),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                List.of(
                    new CellGradientStopSpec(0.0d, ColorSpec.rgb(new RgbHex("#112233"))),
                    new CellGradientStopSpec(1.0d, ColorSpec.rgb(new RgbHex("#445566"))))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellGradientFillSpec.linear(
                Optional.of(Double.NaN),
                List.of(
                    new CellGradientStopSpec(0.0d, ColorSpec.rgb(new RgbHex("#112233"))),
                    new CellGradientStopSpec(1.0d, ColorSpec.rgb(new RgbHex("#445566"))))));
    assertThrows(
        NullPointerException.class, () -> CellGradientFillSpec.linear(Optional.empty(), null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellGradientFillSpec.linear(
                Optional.empty(),
                List.of(new CellGradientStopSpec(0.0d, ColorSpec.rgb(new RgbHex("#112233"))))));
    assertThrows(
        NullPointerException.class,
        () ->
            CellGradientFillSpec.linear(
                Optional.empty(),
                Arrays.asList(
                    new CellGradientStopSpec(0.0d, ColorSpec.rgb(new RgbHex("#112233"))), null)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientStopSpec(-0.1d, ColorSpec.rgb(new RgbHex("#112233"))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientStopSpec(1.5d, ColorSpec.rgb(new RgbHex("#112233"))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientStopSpec(Double.NaN, ColorSpec.rgb(new RgbHex("#112233"))));
    assertThrows(NullPointerException.class, () -> CellFillSpec.pattern(null));
    assertThrows(
        NullPointerException.class,
        () -> CellFillSpec.patternBackground(null, ColorSpec.rgb(new RgbHex("#112233"))));
    assertThrows(NullPointerException.class, () -> CellFillSpec.gradient(null));
    assertThrows(
        NullPointerException.class,
        () -> CellFillSpec.patternForeground(null, ColorSpec.rgb(new RgbHex("#112233"))));
    assertThrows(
        NullPointerException.class,
        () -> CellFillSpec.patternBackground(null, ColorSpec.rgb(new RgbHex("#112233"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellFillSpec.patternForeground(
                ExcelFillPattern.NONE, ColorSpec.rgb(new RgbHex("#112233"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellFillSpec.patternColors(
                ExcelFillPattern.SOLID,
                ColorSpec.rgb(new RgbHex("#112233")),
                ColorSpec.rgb(new RgbHex("#445566"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupSpec(
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
            new PrintSetupSpec(
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
            new PrintSetupSpec(
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
            new PrintSetupSpec(
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
        IllegalArgumentException.class, () -> new PrintMarginsSpec(0, 0, Double.NaN, 0, 0, 0));
  }

  @Test
  void drawingValueObjectsNormalizeDefaultsAndValidate() {
    BinaryContent binary = new BinaryContent(new byte[] {1, 2, 3});
    DrawingMarkerSpec from = new DrawingMarkerSpec(1, 2, 3, 4);
    DrawingMarkerSpec to = new DrawingMarkerSpec(4, 6, 0, 0);
    DrawingAnchorSpec.TwoCell twoCell =
        new DrawingAnchorSpec.TwoCell(from, to, ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
    ResolvedPicture picture =
        new ResolvedPicture(
            "OpsPicture", binary, ExcelPictureFormat.PNG, twoCell, Optional.empty());
    ResolvedShape shape =
        new ResolvedShape.SimpleShape("OpsShape", twoCell, " rect ", Optional.of("Queue"));
    ResolvedEmbeddedObject embeddedObject =
        new ResolvedEmbeddedObject(
            "OpsEmbed",
            "Payload",
            "payload.txt",
            "payload.txt",
            new BinaryContent("payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
            ExcelPictureFormat.PNG,
            binary,
            twoCell);
    ExcelDrawingObjectSnapshot.Shape groupSnapshot =
        new ExcelDrawingObjectSnapshot.Shape(
            "OpsGroup",
            new DrawingAnchorReport.Absolute(1L, 2L, 10L, 20L, null),
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
            new BinaryContent("payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
            "Payload",
            "payload.txt");

    assertNotEquals(binary.bytes(), binary.bytes());
    assertEquals(3, binary.size());
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE, twoCell.behavior());
    assertEquals(ExcelPictureFormat.PNG, picture.format());
    assertEquals("rect", ((ResolvedShape.SimpleShape) shape).presetGeometryToken());
    assertEquals("payload.txt", embeddedObject.fileName());
    assertEquals(2, groupSnapshot.childCount());
    assertEquals("Payload", payload.label());
    assertThrows(IllegalArgumentException.class, () -> new BinaryContent(new byte[0]));
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerReport(-1, 0, 0, 0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorSpec.TwoCell(
                new DrawingMarkerSpec(2, 2, 0, 0),
                new DrawingMarkerSpec(1, 2, 0, 0),
                dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertDoesNotThrow(
        () ->
            new DrawingAnchorSpec.TwoCell(
                new DrawingMarkerSpec(2, 1, 0, 3),
                new DrawingMarkerSpec(2, 3, 0, 4),
                dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertDoesNotThrow(
        () ->
            new DrawingAnchorSpec.TwoCell(
                new DrawingMarkerSpec(1, 2, 3, 0),
                new DrawingMarkerSpec(3, 2, 4, 0),
                dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ResolvedShape.SimpleShape("OpsShape", twoCell, " ", Optional.empty()));
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
        new ResolvedComment(
            "Quarterly Review",
            "GridGrind",
            true,
            Optional.of(
                new ResolvedRichText(
                    List.of(
                        new ResolvedRichTextRun("Quarterly", Optional.empty()),
                        new ResolvedRichTextRun(
                            " Review",
                            Optional.of(
                                new CellFontPatch(
                                    Optional.of(true),
                                    Optional.of(false),
                                    Optional.of("Aptos"),
                                    Optional.empty(),
                                    Optional.of(ColorSpec.rgb(new RgbHex("#112233"))),
                                    Optional.empty(),
                                    Optional.empty())))))),
            Optional.of(new CommentAnchorSpec(1, 2, 4, 6)));
    var commentSnapshot =
        new ExcelCommentSnapshot(
            "Quarterly Review",
            "GridGrind",
            true,
            Optional.of(
                List.copyOf(
                    List.of(
                        new RichTextRunReport(
                            "Quarterly",
                            new CellFontReport(
                                false,
                                false,
                                "Aptos",
                                FontHeightReport.fromPoints(java.math.BigDecimal.valueOf(11)),
                                null,
                                false,
                                false)),
                        new RichTextRunReport(
                            " Review",
                            new CellFontReport(
                                true,
                                false,
                                "Aptos",
                                FontHeightReport.fromPoints(java.math.BigDecimal.valueOf(11)),
                                CellColorReport.rgb(new RgbHex("#112233")),
                                false,
                                false))))),
            Optional.of(new CommentAnchorReport(1, 2, 4, 6)));
    var formulaTarget = ExcelNamedRangeTarget.formula("SUM(Ops!$B$2:$B$5)");
    var table =
        new ResolvedTable(
            new TableName("Queue"),
            new SheetName("Ops"),
            new RangeRef("A1:B4"),
            true,
            true,
            new TableStyleSpec.None(),
            "",
            true,
            true,
            false,
            "",
            "",
            "",
            List.of(new TableColumnSpec(1, "", "", " SUM ", "B2*2")));
    var protection =
        new ExcelWorkbookProtectionSettings(
            true, false, true, Optional.of("book"), Optional.of("review"));
    ResolvedComment copiedComment = commentSnapshot.toAuthoringComment();

    assertEquals("SUM(Ops!$B$2:$B$5)", formulaTarget.refersToFormula());
    assertEquals(authoringComment.text(), copiedComment.text());
    assertEquals(authoringComment.author(), copiedComment.author());
    assertEquals(authoringComment.visible(), copiedComment.visible());
    assertEquals(authoringComment.anchor(), copiedComment.anchor());
    assertEquals(
        authoringComment.runs().orElseThrow().plainText(),
        copiedComment.runs().orElseThrow().plainText());
    assertEquals(
        ColorSpec.rgb(new RgbHex("#112233")),
        copiedComment
            .runs()
            .orElseThrow()
            .runs()
            .get(1)
            .font()
            .orElseThrow()
            .fontColor()
            .orElseThrow());
    assertEquals("", table.comment());
    assertEquals("", table.headerRowCellStyle());
    assertEquals("sum", table.columns().getFirst().totalsRowFunction());
    assertEquals("B2*2", table.columns().getFirst().calculatedColumnFormula());
    assertEquals("book", protection.workbookPassword().orElseThrow());
    assertEquals("review", protection.revisionsPassword().orElseThrow());
    assertEquals(
        "A1", ((ExcelNamedRangeTarget.Range) ExcelNamedRangeTarget.range("Ops", "A1")).range());
    assertEquals("Ops!$A$1", ExcelNamedRangeTarget.range("Ops", "A1").refersToFormula());
    assertEquals(
        "A1:B2",
        ((ExcelNamedRangeTarget.Range) ExcelNamedRangeTarget.range("Ops", "$A$1:$B$2")).range());
    ResolvedComment plainComment =
        new ExcelCommentSnapshot("Done", "GridGrind", false, Optional.empty(), Optional.empty())
            .toPlainComment();
    assertEquals("Done", plainComment.text());
    assertEquals("GridGrind", plainComment.author());
    ResolvedComment plainAuthoringComment =
        new ExcelCommentSnapshot("Done", "GridGrind", false, Optional.empty(), Optional.empty())
            .toAuthoringComment();
    assertEquals("Done", plainAuthoringComment.text());
    assertTrue(plainAuthoringComment.runs().isEmpty());
    assertTrue(plainAuthoringComment.anchor().isEmpty());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedComment(
                "Quarterly Review",
                "GridGrind",
                true,
                Optional.of(
                    new ResolvedRichText(
                        List.of(new ResolvedRichTextRun("Mismatch", Optional.empty())))),
                Optional.empty()));
    assertThrows(IllegalArgumentException.class, () -> ExcelNamedRangeTarget.formula(" "));
    assertThrows(IllegalArgumentException.class, () -> ExcelNamedRangeTarget.range("Ops", " "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelWorkbookProtectionSettings(
                false, false, false, Optional.of(" "), Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelWorkbookProtectionSettings(
                false, false, false, Optional.empty(), Optional.of(" ")));
    assertThrows(IllegalArgumentException.class, () -> new TableColumnSpec(-1, "", "", "", ""));
    TableColumnSpec normalizedColumn = new TableColumnSpec(0, "", "", "", "");
    assertEquals("", normalizedColumn.uniqueName());
    assertEquals("", normalizedColumn.totalsRowLabel());
    assertEquals("", normalizedColumn.totalsRowFunction());
    assertEquals("", normalizedColumn.calculatedColumnFormula());
    ResolvedTable normalizedTable =
        new ResolvedTable(
            new TableName("Queue"),
            new SheetName("Ops"),
            new RangeRef("A1:B2"),
            false,
            true,
            new TableStyleSpec.None(),
            "",
            false,
            false,
            false,
            "",
            "",
            "",
            java.util.List.of());
    assertEquals("", normalizedTable.comment());
    assertEquals(List.of(), normalizedTable.columns());
    assertThrows(
        NullPointerException.class,
        () ->
            new ResolvedTable(
                new TableName("Queue"),
                new SheetName("Ops"),
                new RangeRef("A1:B2"),
                false,
                true,
                new TableStyleSpec.None(),
                "",
                false,
                false,
                false,
                "",
                "",
                "",
                Arrays.asList((TableColumnSpec) null)));
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
                ColorSpec.rgb(new RgbHex("#112233")),
                ColorSpec.rgb(new RgbHex("#445566")),
                ColorSpec.rgb(new RgbHex("#778899"))),
            false);
    var dataBar =
        new ExcelConditionalFormattingRule.DataBarRule(
            ColorSpec.rgb(new RgbHex("#102030")), true, 10, 90, min, max, false);
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
    var top10 =
        new ExcelConditionalFormattingRule.Top10Rule(10, true, false, false, Optional.empty());
    var anchor = new CommentAnchorSpec(1, 2, 3, 4);

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
                List.of(min), List.of(ColorSpec.rgb(new RgbHex("#112233"))), false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRule.ColorScaleRule(
                List.of(min, max), List.of(ColorSpec.rgb(new RgbHex("#112233"))), false));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelConditionalFormattingRule.ColorScaleRule(
                null, List.of(ColorSpec.rgb(new RgbHex("#112233"))), false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRule.DataBarRule(
                ColorSpec.rgb(new RgbHex("#102030")), false, 10, 9, min, max, false));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelConditionalFormattingRule.DataBarRule(null, false, 0, 0, min, max, false));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelConditionalFormattingRule.DataBarRule(
                ColorSpec.rgb(new RgbHex("#102030")), false, 0, 0, null, max, false));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelConditionalFormattingRule.DataBarRule(
                ColorSpec.rgb(new RgbHex("#102030")), false, 0, 0, min, null, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRule.DataBarRule(
                ColorSpec.rgb(new RgbHex("#102030")), false, -1, 0, min, max, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRule.DataBarRule(
                ColorSpec.rgb(new RgbHex("#102030")), false, 0, -1, min, max, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRule.DataBarRule(
                ColorSpec.rgb(new RgbHex("#102030")), false, 101, 101, min, max, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRule.DataBarRule(
                ColorSpec.rgb(new RgbHex("#102030")), false, 0, 101, min, max, false));
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
        NullPointerException.class,
        () -> new ExcelConditionalFormattingRule.Top10Rule(0, false, false, false, null));
    assertThrows(IllegalArgumentException.class, () -> new CommentAnchorSpec(-1, 0, 1, 1));
    assertThrows(IllegalArgumentException.class, () -> new CommentAnchorSpec(1, 0, 0, 1));
    assertThrows(IllegalArgumentException.class, () -> new CommentAnchorSpec(1, 2, 1, 1));
    assertThrows(IllegalArgumentException.class, () -> new PrintMarginsSpec(-0.1d, 0, 0, 0, 0, 0));
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
