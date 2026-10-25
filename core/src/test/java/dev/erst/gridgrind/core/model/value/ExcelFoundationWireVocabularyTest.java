package dev.erst.gridgrind.core.model.value;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.vocabulary.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarShape;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartMarkerStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartRadarStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartScatterStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingUnsupportedFeature;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingShapeKind;
import dev.erst.gridgrind.core.vocabulary.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelGradientFillGeometry;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelIgnoredErrorType;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteHash;
import dev.erst.gridgrind.core.vocabulary.ExcelPaneRegion;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.ExcelReportedCellErrorLiteral;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.ExcelStoredCellErrorLiteral;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
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
    assertEquals(
        List.of("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"),
        ExcelStoredCellErrorLiteral.orderedWireValues());
    assertEquals(
        List.of(
            "#NULL!",
            "#DIV/0!",
            "#VALUE!",
            "#REF!",
            "#NAME?",
            "#NUM!",
            "#N/A",
            "#CIRCULAR_REF!",
            "#FUNCTION_NOT_IMPLEMENTED!"),
        ExcelReportedCellErrorLiteral.orderedWireValues());
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
    assertWireVocabulary(ExcelOoxmlWriteCipher.class, "AES_256", "AES_192");
    assertWireVocabulary(ExcelOoxmlWriteHash.class, "SHA_512", "SHA_384", "SHA_256");
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
