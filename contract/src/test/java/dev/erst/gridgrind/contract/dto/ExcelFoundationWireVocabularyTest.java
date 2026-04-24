package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarShape;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartMarkerStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartRadarStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartScatterStyle;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingUnsupportedFeature;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelGradientFillGeometry;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.excel.foundation.ExcelPaneRegion;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;

/** Guards the published foundation tokens re-exported through the contract module. */
class ExcelFoundationWireVocabularyTest {
  @Test
  void contractExposedFoundationEnumsKeepTheirPublishedTokens() {
    assertWireVocabulary(ExcelAuthoredDrawingShapeKind.class, "SIMPLE_SHAPE", "CONNECTOR");
    assertWireVocabulary(
        ExcelBorderStyle.class,
        "NONE",
        "THIN",
        "MEDIUM",
        "DASHED",
        "DOTTED",
        "THICK",
        "DOUBLE",
        "HAIR",
        "MEDIUM_DASHED",
        "DASH_DOT",
        "MEDIUM_DASH_DOT",
        "DASH_DOT_DOT",
        "MEDIUM_DASH_DOT_DOT",
        "SLANTED_DASH_DOT");
    assertWireVocabulary(ExcelChartAxisCrosses.class, "AUTO_ZERO", "MAX", "MIN");
    assertWireVocabulary(ExcelChartAxisKind.class, "CATEGORY", "DATE", "SERIES", "VALUE");
    assertWireVocabulary(ExcelChartAxisPosition.class, "BOTTOM", "LEFT", "RIGHT", "TOP");
    assertWireVocabulary(ExcelChartBarDirection.class, "COLUMN", "BAR");
    assertWireVocabulary(
        ExcelChartBarGrouping.class, "STANDARD", "CLUSTERED", "STACKED", "PERCENT_STACKED");
    assertWireVocabulary(
        ExcelChartBarShape.class,
        "BOX",
        "CONE",
        "CONE_TO_MAX",
        "CYLINDER",
        "PYRAMID",
        "PYRAMID_TO_MAX");
    assertWireVocabulary(ExcelChartDisplayBlanksAs.class, "GAP", "SPAN", "ZERO");
    assertWireVocabulary(ExcelChartGrouping.class, "STANDARD", "STACKED", "PERCENT_STACKED");
    assertWireVocabulary(
        ExcelChartLegendPosition.class, "BOTTOM", "LEFT", "RIGHT", "TOP", "TOP_RIGHT");
    assertWireVocabulary(
        ExcelChartMarkerStyle.class,
        "CIRCLE",
        "DASH",
        "DIAMOND",
        "DOT",
        "NONE",
        "PICTURE",
        "PLUS",
        "SQUARE",
        "STAR",
        "TRIANGLE",
        "X");
    assertWireVocabulary(ExcelChartRadarStyle.class, "FILLED", "MARKER", "STANDARD");
    assertWireVocabulary(
        ExcelChartScatterStyle.class,
        "LINE",
        "LINE_MARKER",
        "MARKER",
        "NONE",
        "SMOOTH",
        "SMOOTH_MARKER");
    assertWireVocabulary(
        ExcelComparisonOperator.class,
        "BETWEEN",
        "NOT_BETWEEN",
        "EQUAL",
        "NOT_EQUAL",
        "GREATER_THAN",
        "LESS_THAN",
        "GREATER_OR_EQUAL",
        "LESS_OR_EQUAL");
    assertWireVocabulary(
        ExcelConditionalFormattingIconSet.class,
        "GYR_3_ARROW",
        "GREY_3_ARROWS",
        "GYR_3_FLAGS",
        "GYR_3_TRAFFIC_LIGHTS",
        "GYR_3_TRAFFIC_LIGHTS_BOX",
        "GYR_3_SHAPES",
        "GYR_3_SYMBOLS_CIRCLE",
        "GYR_3_SYMBOLS",
        "GYR_4_ARROWS",
        "GREY_4_ARROWS",
        "RB_4_TRAFFIC_LIGHTS",
        "RATINGS_4",
        "GYRB_4_TRAFFIC_LIGHTS",
        "GYYYR_5_ARROWS",
        "GREY_5_ARROWS",
        "RATINGS_5",
        "QUARTERS_5");
    assertWireVocabulary(
        ExcelConditionalFormattingThresholdType.class,
        "NUMBER",
        "MIN",
        "MAX",
        "PERCENT",
        "PERCENTILE",
        "UNALLOCATED",
        "FORMULA");
    assertWireVocabulary(
        ExcelConditionalFormattingUnsupportedFeature.class,
        "STYLE_REFERENCE",
        "FONT_ATTRIBUTES",
        "FILL_PATTERN",
        "FILL_BACKGROUND_COLOR",
        "BORDER_COMPLEXITY",
        "ALIGNMENT",
        "PROTECTION");
    assertWireVocabulary(ExcelDataValidationErrorStyle.class, "STOP", "WARNING", "INFORMATION");
    assertWireVocabulary(
        ExcelDrawingAnchorBehavior.class,
        "MOVE_AND_RESIZE",
        "MOVE_DONT_RESIZE",
        "DONT_MOVE_AND_RESIZE");
    assertWireVocabulary(
        ExcelDrawingShapeKind.class, "SIMPLE_SHAPE", "CONNECTOR", "GROUP", "GRAPHIC_FRAME");
    assertWireVocabulary(ExcelEmbeddedObjectPackagingKind.class, "OLE10_NATIVE", "RAW_PACKAGE");
    assertWireVocabulary(
        ExcelFillPattern.class,
        "NONE",
        "SOLID",
        "FINE_DOTS",
        "ALT_BARS",
        "SPARSE_DOTS",
        "THICK_HORIZONTAL_BANDS",
        "THICK_VERTICAL_BANDS",
        "THICK_BACKWARD_DIAGONAL",
        "THICK_FORWARD_DIAGONAL",
        "BIG_SPOTS",
        "BRICKS",
        "THIN_HORIZONTAL_BANDS",
        "THIN_VERTICAL_BANDS",
        "THIN_BACKWARD_DIAGONAL",
        "THIN_FORWARD_DIAGONAL",
        "SQUARES",
        "DIAMONDS",
        "LESS_DOTS",
        "LEAST_DOTS");
    assertWireVocabulary(
        ExcelHorizontalAlignment.class,
        "GENERAL",
        "LEFT",
        "CENTER",
        "RIGHT",
        "FILL",
        "JUSTIFY",
        "CENTER_SELECTION",
        "DISTRIBUTED");
    assertWireVocabulary(
        ExcelIgnoredErrorType.class,
        "CALCULATED_COLUMN",
        "EMPTY_CELL_REFERENCE",
        "EVALUATION_ERROR",
        "FORMULA",
        "FORMULA_RANGE",
        "LIST_DATA_VALIDATION",
        "NUMBER_STORED_AS_TEXT",
        "TWO_DIGIT_TEXT_YEAR",
        "UNLOCKED_FORMULA");
    assertWireVocabulary(ExcelOoxmlEncryptionMode.class, "AGILE", "STANDARD");
    assertWireVocabulary(ExcelOoxmlSignatureDigestAlgorithm.class, "SHA256", "SHA384", "SHA512");
    assertWireVocabulary(
        ExcelOoxmlSignatureState.class, "VALID", "INVALID", "INVALIDATED_BY_MUTATION");
    assertWireVocabulary(
        ExcelPaneRegion.class, "UPPER_LEFT", "UPPER_RIGHT", "LOWER_LEFT", "LOWER_RIGHT");
    assertWireVocabulary(
        ExcelPictureFormat.class,
        "EMF",
        "WMF",
        "PICT",
        "JPEG",
        "PNG",
        "DIB",
        "GIF",
        "TIFF",
        "EPS",
        "BMP",
        "WPG");
    assertWireVocabulary(
        ExcelPivotDataConsolidateFunction.class,
        "SUM",
        "COUNT",
        "COUNT_NUMS",
        "AVERAGE",
        "MAX",
        "MIN",
        "PRODUCT",
        "STD_DEV",
        "STD_DEVP",
        "VAR",
        "VARP");
    assertWireVocabulary(ExcelPrintOrientation.class, "PORTRAIT", "LANDSCAPE");
    assertWireVocabulary(ExcelSheetVisibility.class, "VISIBLE", "HIDDEN", "VERY_HIDDEN");
    assertWireVocabulary(
        ExcelVerticalAlignment.class, "TOP", "CENTER", "BOTTOM", "JUSTIFY", "DISTRIBUTED");
  }

  @Test
  void gradientFillGeometryKeepsItsPublishedTypeTokens() {
    assertEquals("LINEAR", ExcelGradientFillGeometry.normalizeType("linear"));
    assertEquals("PATH", ExcelGradientFillGeometry.normalizeType("PATH"));
    assertEquals("LINEAR", ExcelGradientFillGeometry.effectiveType(null, null, null, null, null));
    assertEquals("PATH", ExcelGradientFillGeometry.effectiveType(null, 0.1d, null, null, null));
  }

  private static void assertWireVocabulary(
      Class<? extends Enum<?>> enumType, String firstToken, String... remainingTokens) {
    List<String> expectedTokens =
        Stream.concat(Stream.of(firstToken), Arrays.stream(remainingTokens)).toList();
    List<String> actualTokens = Arrays.stream(enumType.getEnumConstants()).map(Enum::name).toList();
    assertEquals(expectedTokens, actualTokens, () -> enumType.getSimpleName() + " wire tokens");
  }
}
