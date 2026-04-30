package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.excel.foundation.ExcelPaneRegion;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import java.math.BigDecimal;
import java.util.Arrays;
import java.util.Base64;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Closes remaining constructor and validation branches for standalone DTO families. */
class DtoEdgeCoverageTest {
  @Test
  void drawingAnchorsMarkersAndObjectsValidateAllRemainingBranches() {
    DrawingMarkerReport marker = new DrawingMarkerReport(0, 0, 0, 0);

    assertEquals(
        "rowIndex must not be negative",
        assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerReport(0, -1, 0, 0))
            .getMessage());
    assertEquals(
        "dx must not be negative",
        assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerReport(0, 0, -1, 0))
            .getMessage());
    assertEquals(
        "dy must not be negative",
        assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerReport(0, 0, 0, -1))
            .getMessage());
    assertEquals(
        "widthEmu must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingAnchorReport.OneCell(
                        marker, 0L, 1L, ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE))
            .getMessage());
    assertEquals(
        "heightEmu must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingAnchorReport.OneCell(
                        marker, 1L, 0L, ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE))
            .getMessage());
    assertEquals(
        "xEmu must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingAnchorReport.Absolute(
                        -1L, 0L, 1L, 1L, ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE))
            .getMessage());
    assertEquals(
        "yEmu must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new DrawingAnchorReport.Absolute(0L, -1L, 1L, 1L, null))
            .getMessage());
    assertEquals(
        "widthEmu must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> new DrawingAnchorReport.Absolute(0L, 0L, 0L, 1L, null))
            .getMessage());
    assertEquals(
        "heightEmu must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> new DrawingAnchorReport.Absolute(0L, 0L, 1L, 0L, null))
            .getMessage());

    String base64 = Base64.getEncoder().encodeToString(new byte[] {1, 2, 3});
    DrawingAnchorReport.TwoCell twoCell =
        new DrawingAnchorReport.TwoCell(marker, new DrawingMarkerReport(1, 1, 0, 0), null);

    assertEquals(
        "byteSize must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectReport.Picture(
                        "Pic",
                        twoCell,
                        ExcelPictureFormat.PNG,
                        "image/png",
                        -1L,
                        "sha",
                        10,
                        10,
                        null))
            .getMessage());
    assertEquals(
        "description must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectReport.Picture(
                        "Pic",
                        twoCell,
                        ExcelPictureFormat.PNG,
                        "image/png",
                        1L,
                        "sha",
                        10,
                        10,
                        " "))
            .getMessage());
    assertNull(
        new DrawingObjectReport.Picture(
                "Pic", twoCell, ExcelPictureFormat.PNG, "image/png", 1L, "sha", 10, 10, null)
            .description());
    assertEquals(
        "heightPixels must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectReport.Picture(
                        "Pic",
                        twoCell,
                        ExcelPictureFormat.PNG,
                        "image/png",
                        1L,
                        "sha",
                        10,
                        -1,
                        null))
            .getMessage());
    assertEquals(
        "plotTypeTokens value must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new DrawingObjectReport.Chart("Chart", twoCell, true, List.of(" "), "Title"))
            .getMessage());
    assertEquals(
        "presetGeometryToken must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectReport.Shape(
                        "Shape", twoCell, ExcelDrawingShapeKind.SIMPLE_SHAPE, " ", null, 0))
            .getMessage());
    assertEquals(
        "text must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectReport.Shape(
                        "Shape", twoCell, ExcelDrawingShapeKind.SIMPLE_SHAPE, null, " ", 0))
            .getMessage());
    assertEquals(
        "label must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectReport.EmbeddedObject(
                        "Object",
                        twoCell,
                        ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                        " ",
                        "file.bin",
                        null,
                        "application/octet-stream",
                        1L,
                        "sha",
                        null,
                        null,
                        null))
            .getMessage());
    assertEquals(
        "fileName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectReport.EmbeddedObject(
                        "Object",
                        twoCell,
                        ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                        null,
                        " ",
                        null,
                        "application/octet-stream",
                        1L,
                        "sha",
                        null,
                        null,
                        null))
            .getMessage());
    assertEquals(
        "command must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectReport.EmbeddedObject(
                        "Object",
                        twoCell,
                        ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                        null,
                        null,
                        " ",
                        "application/octet-stream",
                        1L,
                        "sha",
                        null,
                        null,
                        null))
            .getMessage());
    assertEquals(
        "previewSha256 must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectReport.EmbeddedObject(
                        "Object",
                        twoCell,
                        ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                        null,
                        null,
                        null,
                        "application/octet-stream",
                        1L,
                        "sha",
                        ExcelPictureFormat.PNG,
                        1L,
                        " "))
            .getMessage());
    assertEquals(
        "previewSha256 requires previewFormat",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectReport.EmbeddedObject(
                        "Object",
                        twoCell,
                        ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                        null,
                        null,
                        null,
                        "application/octet-stream",
                        1L,
                        "sha",
                        null,
                        null,
                        "preview"))
            .getMessage());
    assertEquals(
        "previewByteSize must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectReport.EmbeddedObject(
                        "Object",
                        twoCell,
                        ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                        null,
                        null,
                        null,
                        "application/octet-stream",
                        1L,
                        "sha",
                        ExcelPictureFormat.PNG,
                        -1L,
                        null))
            .getMessage());
    assertEquals(
        "byteSize must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectReport.EmbeddedObject(
                        "Object",
                        twoCell,
                        ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                        null,
                        null,
                        null,
                        "application/octet-stream",
                        -1L,
                        "sha",
                        null,
                        null,
                        null))
            .getMessage());
    assertEquals(
        "previewContentType requires previewFormat",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectReport.SignatureLine(
                        "Signature",
                        twoCell,
                        null,
                        null,
                        null,
                        null,
                        null,
                        null,
                        null,
                        "image/png",
                        null,
                        null,
                        null,
                        null))
            .getMessage());
    assertEquals(
        "previewHeightPixels must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectReport.SignatureLine(
                        "Signature",
                        twoCell,
                        null,
                        null,
                        null,
                        null,
                        null,
                        null,
                        ExcelPictureFormat.PNG,
                        "image/png",
                        1L,
                        "sha",
                        10,
                        -1))
            .getMessage());
    assertEquals(
        "description must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectPayloadReport.Picture(
                        "Pic", ExcelPictureFormat.PNG, "image/png", "pic.png", "sha", base64, " "))
            .getMessage());
    assertEquals(
        "label must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectPayloadReport.EmbeddedObject(
                        "Object",
                        ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                        "application/octet-stream",
                        "file.bin",
                        "sha",
                        base64,
                        " ",
                        null))
            .getMessage());
    assertEquals(
        "command must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectPayloadReport.EmbeddedObject(
                        "Object",
                        ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                        "application/octet-stream",
                        "file.bin",
                        "sha",
                        base64,
                        null,
                        " "))
            .getMessage());
    assertEquals(
        "name must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectPayloadReport.EmbeddedObject(
                        " ",
                        ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                        "application/octet-stream",
                        "file.bin",
                        "sha",
                        base64,
                        null,
                        null))
            .getMessage());
    assertEquals(
        "fileName must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new DrawingObjectPayloadReport.Picture(
                        "Pic", ExcelPictureFormat.PNG, "image/png", null, "sha", base64, null))
            .getMessage());
    assertEquals(
        "lastRow must not be negative",
        assertThrows(IllegalArgumentException.class, () -> new CommentAnchorReport(0, 0, 0, -1))
            .getMessage());
  }

  @Test
  void chartAutofilterPaneAndEncryptionDtosCoverRemainingValidationBranches() {
    assertEquals(
        "firstSliceAngle must be between 0 and 360",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ChartReport.Pie(
                        false,
                        -1,
                        List.of(
                            chartSeries(
                                new ChartReport.Title.Text("S1"),
                                new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                                new ChartReport.DataSource.NumericLiteral(null, List.of("1"))))))
            .getMessage());
    assertEquals(
        "detail must not be blank",
        assertThrows(
                IllegalArgumentException.class, () -> new ChartReport.Unsupported("RADAR", " "))
            .getMessage());
    assertEquals(
        "formula must not be blank",
        assertThrows(
                IllegalArgumentException.class, () -> new ChartReport.Title.Formula(" ", "cache"))
            .getMessage());
    assertEquals(
        "formatCode must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new ChartReport.DataSource.NumericReference("A1:A2", " ", List.of("1")))
            .getMessage());
    assertEquals(
        "formatCode must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new ChartReport.DataSource.NumericLiteral(" ", List.of("1")))
            .getMessage());

    assertEquals(
        "conditions must not contain null values",
        assertThrows(
                NullPointerException.class,
                () ->
                    new AutofilterFilterCriterionReport.Custom(
                        true,
                        Arrays.asList(
                            new AutofilterFilterCriterionReport.CustomConditionReport(
                                "equals", "1"),
                            null)))
            .getMessage());
    assertEquals(
        "filterValue must be finite when provided",
        assertThrows(
                IllegalArgumentException.class,
                () -> new AutofilterFilterCriterionReport.Top10(true, false, 1.0d, Double.NaN))
            .getMessage());
    assertEquals(
        "iconSet must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new AutofilterFilterCriterionReport.Icon(" ", 0))
            .getMessage());
    assertEquals(
        "operator must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new AutofilterFilterCriterionReport.CustomConditionReport(" ", "1"))
            .getMessage());
    assertEquals(
        "value must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new AutofilterFilterCriterionReport.CustomConditionReport("equals", " "))
            .getMessage());

    assertEquals(
        "leftmostColumn must be greater than or equal to splitColumn",
        assertThrows(IllegalArgumentException.class, () -> new PaneInput.Frozen(2, 0, 1, 0))
            .getMessage());
    assertEquals(
        "topRow must be greater than or equal to splitRow",
        assertThrows(IllegalArgumentException.class, () -> new PaneInput.Frozen(0, 2, 0, 1))
            .getMessage());
    assertEquals(
        "topRow must be 0 when ySplitPosition is 0: 1",
        assertThrows(
                IllegalArgumentException.class,
                () -> new PaneInput.Split(1, 0, 0, 1, ExcelPaneRegion.LOWER_RIGHT))
            .getMessage());

    assertEquals(
        "keyBits must be positive when encrypted",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new OoxmlEncryptionReport(
                        true, ExcelOoxmlEncryptionMode.AGILE, "AES", "SHA-512", "CBC", 0, 16, 100))
            .getMessage());
    assertEquals(
        "blockSize must be positive when encrypted",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new OoxmlEncryptionReport(
                        true, ExcelOoxmlEncryptionMode.AGILE, "AES", "SHA-512", "CBC", 128, 0, 100))
            .getMessage());
    assertEquals(
        "spinCount must be zero or positive when encrypted",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new OoxmlEncryptionReport(
                        true, ExcelOoxmlEncryptionMode.AGILE, "AES", "SHA-512", "CBC", 128, 16, -1))
            .getMessage());
    assertEquals(
        "theme must not be negative",
        assertThrows(IllegalArgumentException.class, () -> CellColorReport.theme(-1)).getMessage());
    assertEquals(
        "indexed must not be negative",
        assertThrows(IllegalArgumentException.class, () -> CellColorReport.indexed(-1))
            .getMessage());
    assertEquals(
        "rgb must not be blank",
        assertThrows(IllegalArgumentException.class, () -> CellColorReport.rgb(" ")).getMessage());
    assertEquals(
        "rgb must match #RRGGBB",
        assertThrows(IllegalArgumentException.class, () -> CellColorReport.rgb("#GGGGGG"))
            .getMessage());
    assertEquals(
        "degree must be finite when provided",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    CellGradientFillReport.linear(
                        Double.POSITIVE_INFINITY,
                        List.of(new CellGradientStopReport(0.0d, CellColorReport.rgb("#AABBCC")))))
            .getMessage());
  }

  @Test
  void pivotPrintAndMiscellaneousReportsCoverRemainingValidationBranches() {
    assertEquals(
        "sourceColumnIndex must not be negative",
        assertThrows(IllegalArgumentException.class, () -> new PivotTableReport.Field(-1, "Amount"))
            .getMessage());
    assertEquals(
        "sourceColumnIndex must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new PivotTableReport.DataField(
                        -1, "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total", null))
            .getMessage());
    assertEquals(
        "detail must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new PivotTableReport.Unsupported(
                        "Sales Pivot 2026",
                        "Report",
                        new PivotTableReport.Anchor("A1", "A1:B3"),
                        " "))
            .getMessage());
    assertEquals(
        "paperSize must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new PrintSetupReport(
                        PrintMarginsReport.defaults(),
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
                        List.of()))
            .getMessage());
    assertEquals(
        "copies must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new PrintSetupReport(
                        PrintMarginsReport.defaults(),
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
                        List.of()))
            .getMessage());
    assertEquals(
        "firstPageNumber must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new PrintSetupReport(
                        PrintMarginsReport.defaults(),
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
                        List.of()))
            .getMessage());
    assertEquals(
        "columnBreaks must not contain null indexes",
        assertThrows(
                NullPointerException.class,
                () ->
                    new PrintSetupReport(
                        PrintMarginsReport.defaults(),
                        false,
                        false,
                        false,
                        1,
                        false,
                        false,
                        1,
                        false,
                        1,
                        List.of(),
                        Arrays.asList((Integer) null)))
            .getMessage());
    assertEquals(
        "lastRow must be greater than or equal to firstRow",
        assertThrows(IllegalArgumentException.class, () -> new CommentAnchorReport(0, 1, 0, 0))
            .getMessage());
    assertEquals(
        "range must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new AutofilterEntryReport.SheetOwned(
                        null,
                        List.of(
                            new AutofilterFilterColumnReport(
                                0L,
                                true,
                                new AutofilterFilterCriterionReport.Values(List.of("A"), false))),
                        null))
            .getMessage());
    assertEquals(
        "filterColumns must not contain null values",
        assertThrows(
                NullPointerException.class,
                () ->
                    new AutofilterEntryReport.SheetOwned(
                        "A1:B2", Arrays.asList((AutofilterFilterColumnReport) null), null))
            .getMessage());
    assertEquals(
        "iconId must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new AutofilterSortConditionReport.Icon("A1:A2", false, -1))
            .getMessage());
    assertEquals(
        "conditions must not contain null values",
        assertThrows(
                NullPointerException.class,
                () ->
                    AutofilterSortStateReport.withoutSortMethod(
                        "A1:B2", false, false, Arrays.asList((AutofilterSortConditionReport) null)))
            .getMessage());

    TableEntryReport table =
        new TableEntryReport(
            "BudgetTable",
            "Budget",
            "A1:B2",
            1,
            0,
            List.of("Owner"),
            List.of(
                TableColumnReport.create(
                    1L,
                    "Owner",
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty())),
            new TableStyleReport.Named("TableStyleMedium2", true, false, false, false),
            true,
            Optional.empty(),
            false,
            false,
            false,
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    assertEquals(Optional.empty(), table.comment());
    assertEquals(Optional.empty(), table.headerRowCellStyle());
    assertEquals(Optional.empty(), table.dataCellStyle());
    assertEquals(Optional.empty(), table.totalsRowCellStyle());
    assertEquals(
        "columnNames must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () ->
                    new TableEntryReport(
                        "BudgetTable",
                        "Budget",
                        "A1:B2",
                        1,
                        0,
                        Arrays.asList((String) null),
                        List.of(),
                        new TableStyleReport.None(),
                        false,
                        Optional.empty(),
                        false,
                        false,
                        false,
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty()))
            .getMessage());
    assertEquals(
        "columns must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () ->
                    new TableEntryReport(
                        "BudgetTable",
                        "Budget",
                        "A1:B2",
                        1,
                        0,
                        List.of("Owner"),
                        Arrays.asList((TableColumnReport) null),
                        new TableStyleReport.None(),
                        false,
                        Optional.empty(),
                        false,
                        false,
                        false,
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty()))
            .getMessage());

    assertEquals(
        "range must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new IgnoredErrorReport(
                        " ", List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT)))
            .getMessage());
    assertEquals(
        "errorTypes must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new IgnoredErrorReport(
                        "A1:A2",
                        List.of(
                            ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT,
                            ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT)))
            .getMessage());
    assertEquals(
        "packagePartName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new OoxmlSignatureReport(" ", null, null, null, ExcelOoxmlSignatureState.VALID))
            .getMessage());
    assertEquals(
        "text must not be empty",
        assertThrows(IllegalArgumentException.class, () -> new RichTextRunReport("", font()))
            .getMessage());
    assertEquals(
        "columnId must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new AutofilterFilterColumnReport(
                        -1L, true, new AutofilterFilterCriterionReport.Values(List.of("A"), false)))
            .getMessage());
    assertEquals(
        "checkedValidationCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DataValidationHealthReport(
                        -1,
                        new GridGrindAnalysisReports.AnalysisSummaryReport(0, 0, 0, 0),
                        List.of()))
            .getMessage());
    assertEquals(
        "checkedPivotTableCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new PivotTableHealthReport(
                        -1,
                        new GridGrindAnalysisReports.AnalysisSummaryReport(0, 0, 0, 0),
                        List.of()))
            .getMessage());
    assertInstanceOf(
        GridGrindWorkbookSurfaceReports.NamedRangeReport.FormulaReport.class,
        new GridGrindWorkbookSurfaceReports.NamedRangeReport.FormulaReport(
            "Expr", new NamedRangeScope.Workbook(), "SUM(A1:A2)"));
    assertEquals(
        "color must not be null",
        assertThrows(NullPointerException.class, () -> new CellGradientStopReport(0.0d, null))
            .getMessage());
    assertTrue(PrintMarginsReport.defaults().left() > 0.0d);
    assertEquals(
        "left must be finite and non-negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new PrintMarginsReport(Double.NaN, 0, 0, 0, 0, 0))
            .getMessage());
    assertEquals(
        "left must be finite and non-negative",
        assertThrows(
                IllegalArgumentException.class, () -> new PrintMarginsReport(-1.0d, 0, 0, 0, 0, 0))
            .getMessage());
    assertEquals(
        "rank must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ConditionalFormattingRuleReport.Top10Rule(1, false, -1, false, false, null))
            .getMessage());
  }

  private static CellFontReport font() {
    return new CellFontReport(
        false,
        false,
        "Aptos",
        new FontHeightReport(220, BigDecimal.valueOf(11)),
        null,
        false,
        false);
  }

  private static ChartReport.Series chartSeries(
      ChartReport.Title title, ChartReport.DataSource categories, ChartReport.DataSource values) {
    return new ChartReport.Series(title, categories, values, null, null, null, null);
  }
}
