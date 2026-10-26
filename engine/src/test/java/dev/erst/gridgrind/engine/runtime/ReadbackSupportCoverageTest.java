package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.drawing.CommentAnchorReport;
import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellTemporalKind;
import dev.erst.gridgrind.core.model.value.CellTemporalReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.excel.ExcelCellMetadataSnapshot;
import dev.erst.gridgrind.excel.ExcelCellReadFacet;
import dev.erst.gridgrind.excel.ExcelCellReadProjection;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import java.math.BigDecimal;
import java.util.List;
import java.util.Optional;
import java.util.Set;
import org.junit.jupiter.api.Test;

/**
 * Coverage tests for runtime readback helpers introduced by projection-aware cell introspection.
 */
class ReadbackSupportCoverageTest {
  @Test
  void styleReportConversionPreservesEveryOwnedDifferentialColorVariant() {
    assertEquals(
        Optional.empty(),
        InspectionResultCellStyleReportSupport.toCellColorReport((ColorSpec) null));
    assertEquals(
        Optional.of(CellColorReport.rgb(new RgbHex("#112233"))),
        InspectionResultCellStyleReportSupport.toCellColorReport(
            ColorSpec.rgb(new RgbHex("#112233"))));
    assertEquals(
        Optional.of(CellColorReport.rgb(new RgbHex("#112233"), 0.25d)),
        InspectionResultCellStyleReportSupport.toCellColorReport(
            new ColorSpec.Rgb(new RgbHex("#112233"), Optional.of(0.25d))));
    assertEquals(
        Optional.of(CellColorReport.theme(4)),
        InspectionResultCellStyleReportSupport.toCellColorReport(ColorSpec.theme(4)));
    assertEquals(
        Optional.of(CellColorReport.theme(4, -0.25d)),
        InspectionResultCellStyleReportSupport.toCellColorReport(
            new ColorSpec.Theme(4, Optional.of(-0.25d))));
    assertEquals(
        Optional.of(CellColorReport.indexed(9)),
        InspectionResultCellStyleReportSupport.toCellColorReport(ColorSpec.indexed(9)));
    assertEquals(
        Optional.of(CellColorReport.indexed(9, 0.5d)),
        InspectionResultCellStyleReportSupport.toCellColorReport(
            new ColorSpec.Indexed(9, Optional.of(0.5d))));
  }

  @Test
  void cellReportConversionProjectsFacetsAcrossAllSnapshotKinds() {
    ExcelCellReadProjection allFacets =
        new ExcelCellReadProjection(Set.of(ExcelCellReadFacet.values()));
    ExcelCellReadProjection valueOnly = ExcelCellReadProjection.defaults();
    ExcelCellReadProjection formulaOnly =
        new ExcelCellReadProjection(Set.of(ExcelCellReadFacet.FORMULA));
    List<RichTextRunReport> richText =
        List.copyOf(List.of(new RichTextRunReport("Ada", fontSnapshot())));
    ExcelCellSnapshot.TextSnapshot textSnapshot =
        new ExcelCellSnapshot.TextSnapshot(
            "A1", "Ada", style("General"), metadata(), "Ada", richText);
    ExcelCellSnapshot.NumberSnapshot dateSnapshot =
        new ExcelCellSnapshot.NumberSnapshot(
            "A2", "7/1/2025", style("m/d/yyyy"), metadata(), 45839.0d);
    ExcelCellSnapshot.NumberSnapshot dateTimeSnapshot =
        new ExcelCellSnapshot.NumberSnapshot(
            "A8", "7/1/2025 12:00", style("m/d/yyyy h:mm"), metadata(), 45839.5d);
    ExcelCellSnapshot.NumberSnapshot timeSnapshot =
        new ExcelCellSnapshot.NumberSnapshot("A9", "12:00", style("h:mm"), metadata(), 0.5d);
    ExcelCellSnapshot.NumberSnapshot plainNumberSnapshot =
        new ExcelCellSnapshot.NumberSnapshot("A3", "42", style("0.00"), metadata(), 42.0d);
    ExcelCellSnapshot.BooleanSnapshot booleanSnapshot =
        new ExcelCellSnapshot.BooleanSnapshot("A4", "TRUE", style("General"), metadata(), true);
    ExcelCellSnapshot.ErrorSnapshot errorSnapshot =
        new ExcelCellSnapshot.ErrorSnapshot("A5", "#REF!", style("General"), metadata(), "#REF!");
    ExcelCellSnapshot.BlankSnapshot blankSnapshot =
        new ExcelCellSnapshot.BlankSnapshot("A6", "", style("General"), metadata());
    ExcelCellSnapshot.FormulaSnapshot formulaSnapshot =
        new ExcelCellSnapshot.FormulaSnapshot(
            "A7",
            "Ada",
            style("General"),
            metadata(),
            "CONCAT(\"A\",\"da\")",
            java.util.Optional.of(textSnapshot));

    CellReport.TextReport textReport =
        assertInstanceOf(
            CellReport.TextReport.class,
            InspectionResultCellReportSupport.toCellReport(textSnapshot, allFacets, false));
    CellReport.NumberReport dateReport =
        assertInstanceOf(
            CellReport.NumberReport.class,
            InspectionResultCellReportSupport.toCellReport(dateSnapshot, allFacets, false));
    CellReport.NumberReport dateTimeReport =
        assertInstanceOf(
            CellReport.NumberReport.class,
            InspectionResultCellReportSupport.toCellReport(dateTimeSnapshot, allFacets, false));
    CellReport.NumberReport timeReport =
        assertInstanceOf(
            CellReport.NumberReport.class,
            InspectionResultCellReportSupport.toCellReport(timeSnapshot, allFacets, false));
    CellReport.NumberReport plainNumberReport =
        assertInstanceOf(
            CellReport.NumberReport.class,
            InspectionResultCellReportSupport.toCellReport(plainNumberSnapshot, allFacets, false));
    CellReport.TextReport textValueOnly =
        assertInstanceOf(
            CellReport.TextReport.class,
            InspectionResultCellReportSupport.toCellReport(textSnapshot, valueOnly, false));
    CellReport.NumberReport dateValueOnly =
        assertInstanceOf(
            CellReport.NumberReport.class,
            InspectionResultCellReportSupport.toCellReport(dateSnapshot, valueOnly, false));
    CellReport.FormulaReport formulaValueOnly =
        assertInstanceOf(
            CellReport.FormulaReport.class,
            InspectionResultCellReportSupport.toCellReport(formulaSnapshot, valueOnly, false));
    CellReport.FormulaReport formulaFormulaOnly =
        assertInstanceOf(
            CellReport.FormulaReport.class,
            InspectionResultCellReportSupport.toCellReport(formulaSnapshot, formulaOnly, false));

    assertEquals(
        Optional.of("Ada"),
        textReport.value().map(dev.erst.gridgrind.core.model.value.CellValue.Text::value));
    assertEquals("Ada", textReport.runs().orElseThrow().getFirst().text());
    assertEquals(HyperlinkTarget.Url.class, textReport.hyperlink().orElseThrow().getClass());
    assertEquals("Review", textReport.comment().orElseThrow().text());
    assertEquals(Optional.of("Ada"), textReport.displayValue());
    assertEquals(CellTemporalKind.DATE, dateReport.temporal().orElseThrow().kind().orElseThrow());
    assertEquals(
        CellTemporalKind.DATE_TIME, dateTimeReport.temporal().orElseThrow().kind().orElseThrow());
    assertEquals(CellTemporalKind.TIME, timeReport.temporal().orElseThrow().kind().orElseThrow());
    assertEquals("12:00", timeReport.temporal().orElseThrow().isoValue().orElseThrow());
    assertFalse(plainNumberReport.temporal().orElseThrow().isDate());
    assertEquals(Optional.empty(), textValueOnly.displayValue());
    assertEquals(Optional.empty(), textValueOnly.style());
    assertEquals(Optional.empty(), textValueOnly.runs());
    assertEquals(Optional.empty(), dateValueOnly.displayValue());
    assertEquals(Optional.empty(), dateValueOnly.style());
    assertEquals(Optional.empty(), dateValueOnly.temporal());
    assertEquals(Optional.empty(), formulaValueOnly.formula());
    assertEquals(Optional.empty(), formulaValueOnly.style());
    assertEquals(Optional.of("CONCAT(\"A\",\"da\")"), formulaFormulaOnly.formula());
    assertEquals(Optional.empty(), formulaFormulaOnly.value());

    assertInstanceOf(
        CellReport.BooleanReport.class,
        InspectionResultCellReportSupport.toCellReport(booleanSnapshot, valueOnly, false));
    assertInstanceOf(
        CellReport.ErrorReport.class,
        InspectionResultCellReportSupport.toCellReport(errorSnapshot, valueOnly, false));
    assertInstanceOf(
        CellReport.BlankReport.class,
        InspectionResultCellReportSupport.toCellReport(blankSnapshot, valueOnly, false));

    assertInstanceOf(
        CellValue.Blank.class, InspectionResultCellReportSupport.toCellValue(blankSnapshot));
    assertInstanceOf(
        CellValue.Text.class, InspectionResultCellReportSupport.toCellValue(textSnapshot));
    assertInstanceOf(
        CellValue.NumberValue.class, InspectionResultCellReportSupport.toCellValue(dateSnapshot));
    assertInstanceOf(
        CellValue.BooleanValue.class,
        InspectionResultCellReportSupport.toCellValue(booleanSnapshot));
    assertInstanceOf(
        CellValue.ErrorValue.class, InspectionResultCellReportSupport.toCellValue(errorSnapshot));

    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class,
            () -> InspectionResultCellReportSupport.toCellValue(formulaSnapshot));
    assertEquals("Formula evaluations must not recursively remain FORMULA", exception.getMessage());
  }

  @Test
  void semanticSelectorKeyMatchingHandlesSupportedScalarKinds() {
    ExcelCellSnapshot.TextSnapshot textSnapshot =
        new ExcelCellSnapshot.TextSnapshot(
            "A1", "Ada", style("General"), ExcelCellMetadataSnapshot.empty(), "Ada", null);
    ExcelCellSnapshot.NumberSnapshot numberSnapshot =
        new ExcelCellSnapshot.NumberSnapshot(
            "A2", "42", style("0.00"), ExcelCellMetadataSnapshot.empty(), 42.0d);
    ExcelCellSnapshot.BooleanSnapshot booleanSnapshot =
        new ExcelCellSnapshot.BooleanSnapshot(
            "A3", "TRUE", style("General"), ExcelCellMetadataSnapshot.empty(), true);
    ExcelCellSnapshot.FormulaSnapshot formulaSnapshot =
        new ExcelCellSnapshot.FormulaSnapshot(
            "A4",
            "42",
            style("General"),
            ExcelCellMetadataSnapshot.empty(),
            "SUM(A1:A2)",
            java.util.Optional.of(numberSnapshot));
    ExcelCellSnapshot.BlankSnapshot blankSnapshot =
        new ExcelCellSnapshot.BlankSnapshot(
            "A5", "", style("General"), ExcelCellMetadataSnapshot.empty());

    assertTrue(
        SemanticSelectorKeyMatchSupport.matchesKeyCell(
            textSnapshot,
            new dev.erst.gridgrind.core.model.value.CellValueSpec.Text(
                TextSourceSpec.inline("Ada")),
            false));
    assertTrue(
        SemanticSelectorKeyMatchSupport.matchesKeyCell(
            numberSnapshot,
            new dev.erst.gridgrind.core.model.value.CellValueSpec.NumberValue(42.0d),
            false));
    assertTrue(
        SemanticSelectorKeyMatchSupport.matchesKeyCell(
            booleanSnapshot,
            new dev.erst.gridgrind.core.model.value.CellValueSpec.BooleanValue(true),
            false));
    assertTrue(
        SemanticSelectorKeyMatchSupport.matchesKeyCell(
            formulaSnapshot,
            new dev.erst.gridgrind.core.model.value.CellValueSpec.Formula(
                TextSourceSpec.inline("SUM(A1:A2)")),
            false));
    assertTrue(
        SemanticSelectorKeyMatchSupport.matchesKeyCell(
            blankSnapshot, new dev.erst.gridgrind.core.model.value.CellValueSpec.Blank(), false));
    assertFalse(
        SemanticSelectorKeyMatchSupport.matchesKeyCell(
            textSnapshot,
            new dev.erst.gridgrind.core.model.value.CellValueSpec.Text(
                TextSourceSpec.inline("Grace")),
            false));
    assertThrows(
        IllegalStateException.class,
        () ->
            SemanticSelectorKeyMatchSupport.matchesKeyCell(
                textSnapshot,
                new dev.erst.gridgrind.core.model.value.CellValueSpec.Text(
                    TextSourceSpec.standardInput()),
                false));
  }

  @Test
  void formulaSnapshotRejectsFormulaEvaluations() {
    ExcelCellSnapshot.NumberSnapshot evaluatedNumber =
        new ExcelCellSnapshot.NumberSnapshot(
            "A2", "42", style("General"), ExcelCellMetadataSnapshot.empty(), 42.0d);
    ExcelCellSnapshot.FormulaSnapshot nestedFormula =
        new ExcelCellSnapshot.FormulaSnapshot(
            "A2",
            "42",
            style("General"),
            ExcelCellMetadataSnapshot.empty(),
            "SUM(A1:A1)",
            java.util.Optional.of(evaluatedNumber));

    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new ExcelCellSnapshot.FormulaSnapshot(
                    "A1",
                    "42",
                    style("General"),
                    ExcelCellMetadataSnapshot.empty(),
                    "SUM(A1:A2)",
                    java.util.Optional.of(nestedFormula)));
    assertEquals("formula evaluation must not itself be FORMULA", exception.getMessage());
  }

  @Test
  void assertionValueEvaluatorMatchesFormulaEvaluationsAcrossPublishedKinds() {
    CellStyleReport style = styleReport();
    CellReport.FormulaReport blankFormula =
        new CellReport.FormulaReport(
            new CellAddress("A1"),
            Optional.empty(),
            Optional.of(style),
            Optional.empty(),
            Optional.empty(),
            Optional.of("A1"),
            java.util.Optional.of(new dev.erst.gridgrind.core.model.value.CellValue.Blank()),
            java.util.Optional.empty(),
            java.util.Optional.empty());
    CellReport.FormulaReport textFormula =
        new CellReport.FormulaReport(
            new CellAddress("A2"),
            Optional.empty(),
            Optional.of(style),
            Optional.empty(),
            Optional.empty(),
            Optional.of("A2"),
            java.util.Optional.of(new dev.erst.gridgrind.core.model.value.CellValue.Text("Ada")),
            Optional.of(List.of(runReport("Ada"))),
            java.util.Optional.empty());
    CellReport.FormulaReport numberFormula =
        new CellReport.FormulaReport(
            new CellAddress("A3"),
            Optional.empty(),
            Optional.of(style),
            Optional.empty(),
            Optional.empty(),
            Optional.of("A3"),
            java.util.Optional.of(
                new dev.erst.gridgrind.core.model.value.CellValue.NumberValue(42.0d)),
            java.util.Optional.empty(),
            Optional.of(CellTemporalReport.temporal(CellTemporalKind.DATE, "2026-07-01")));
    CellReport.FormulaReport booleanFormula =
        new CellReport.FormulaReport(
            new CellAddress("A4"),
            Optional.empty(),
            Optional.of(style),
            Optional.empty(),
            Optional.empty(),
            Optional.of("A4"),
            java.util.Optional.of(
                new dev.erst.gridgrind.core.model.value.CellValue.BooleanValue(true)),
            java.util.Optional.empty(),
            java.util.Optional.empty());
    CellReport.FormulaReport errorFormula =
        new CellReport.FormulaReport(
            new CellAddress("A5"),
            Optional.empty(),
            Optional.of(style),
            Optional.empty(),
            Optional.empty(),
            Optional.of("A5"),
            java.util.Optional.of(
                new dev.erst.gridgrind.core.model.value.CellValue.ErrorValue("#REF!")),
            java.util.Optional.empty(),
            java.util.Optional.empty());
    CellReport.FormulaReport missingEvaluation =
        new CellReport.FormulaReport(
            new CellAddress("A6"),
            Optional.empty(),
            Optional.of(style),
            Optional.empty(),
            Optional.empty(),
            Optional.of("A6"),
            Optional.empty(),
            java.util.Optional.empty(),
            java.util.Optional.empty());

    assertTrue(AssertionValueEvaluator.matchesCellValue(blankFormula, new CellValue.Blank()));
    assertTrue(AssertionValueEvaluator.matchesCellValue(textFormula, new CellValue.Text("Ada")));
    assertFalse(AssertionValueEvaluator.matchesCellValue(textFormula, new CellValue.Text("Grace")));
    assertFalse(AssertionValueEvaluator.matchesCellValue(numberFormula, new CellValue.Text("Ada")));
    assertTrue(
        AssertionValueEvaluator.matchesCellValue(numberFormula, new CellValue.NumberValue(42.0d)));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(numberFormula, new CellValue.NumberValue(41.0d)));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(textFormula, new CellValue.NumberValue(42.0d)));
    assertTrue(
        AssertionValueEvaluator.matchesCellValue(booleanFormula, new CellValue.BooleanValue(true)));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(
            booleanFormula, new CellValue.BooleanValue(false)));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(numberFormula, new CellValue.BooleanValue(true)));
    assertTrue(
        AssertionValueEvaluator.matchesCellValue(errorFormula, new CellValue.ErrorValue("#REF!")));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(
            errorFormula, new CellValue.ErrorValue("#DIV/0!")));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(
            booleanFormula, new CellValue.ErrorValue("#REF!")));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(missingEvaluation, new CellValue.Text("Ada")));
  }

  @Test
  void scalarMetadataAndBlankValuesFollowTheirProjectionMasks() {
    ExcelCellReadProjection allFacets =
        new ExcelCellReadProjection(Set.of(ExcelCellReadFacet.values()));
    ExcelCellReadProjection valueOnly = ExcelCellReadProjection.defaults();
    ExcelCellReadProjection formulaOnly =
        new ExcelCellReadProjection(Set.of(ExcelCellReadFacet.FORMULA));
    var textSnapshot =
        new ExcelCellSnapshot.TextSnapshot(
            "A1",
            "Ada",
            style("General"),
            metadata(),
            "Ada",
            List.copyOf(List.of(new RichTextRunReport("Ada", fontSnapshot()))));
    var formulaSnapshot =
        new ExcelCellSnapshot.FormulaSnapshot(
            "A7", "Ada", style("General"), metadata(), "A1", Optional.of(textSnapshot));
    var blankSnapshot = new ExcelCellSnapshot.BlankSnapshot("A6", "", style("General"), metadata());
    var dateSnapshot =
        new ExcelCellSnapshot.NumberSnapshot(
            "A2", "7/1/2025", style("m/d/yyyy"), metadata(), 45839.0);
    var formulaValueOnly =
        assertInstanceOf(
            CellReport.FormulaReport.class,
            InspectionResultCellReportSupport.toCellReport(formulaSnapshot, valueOnly, false));
    var formulaFormulaOnly =
        assertInstanceOf(
            CellReport.FormulaReport.class,
            InspectionResultCellReportSupport.toCellReport(formulaSnapshot, formulaOnly, false));
    CellReport.FormulaReport formulaAll =
        assertInstanceOf(
            CellReport.FormulaReport.class,
            InspectionResultCellReportSupport.toCellReport(formulaSnapshot, allFacets, false));
    assertEquals(new CellValue.Text("Ada"), formulaAll.value().orElseThrow());
    assertEquals("Ada", formulaAll.runs().orElseThrow().getFirst().text());
    assertEquals(Optional.empty(), formulaValueOnly.runs());
    assertEquals(Optional.empty(), formulaFormulaOnly.runs());
    CellReport.BlankReport blankValue =
        assertInstanceOf(
            CellReport.BlankReport.class,
            InspectionResultCellReportSupport.toCellReport(blankSnapshot, valueOnly, false));
    assertEquals(Optional.of(new CellValue.Blank()), blankValue.value());
    assertEquals(
        Optional.empty(),
        assertInstanceOf(
                CellReport.BlankReport.class,
                InspectionResultCellReportSupport.toCellReport(blankSnapshot, formulaOnly, false))
            .value());
    ExcelCellReadProjection temporalOnly =
        new ExcelCellReadProjection(Set.of(ExcelCellReadFacet.TEMPORAL));
    CellReport.NumberReport dateWithoutValue =
        assertInstanceOf(
            CellReport.NumberReport.class,
            InspectionResultCellReportSupport.toCellReport(dateSnapshot, temporalOnly, false));
    assertEquals(Optional.empty(), dateWithoutValue.value());
    assertEquals(
        CellTemporalKind.DATE, dateWithoutValue.temporal().orElseThrow().kind().orElseThrow());
    var numericFormula =
        new ExcelCellSnapshot.FormulaSnapshot(
            "A7", "7/1/2025", style("m/d/yyyy"), metadata(), "A2", Optional.of(dateSnapshot));
    CellReport.FormulaReport numericFormulaAll =
        assertInstanceOf(
            CellReport.FormulaReport.class,
            InspectionResultCellReportSupport.toCellReport(numericFormula, allFacets, false));
    assertEquals(new CellValue.NumberValue(45839), numericFormulaAll.value().orElseThrow());
    assertEquals(Optional.of("2025-07-01"), numericFormulaAll.temporal().orElseThrow().isoValue());
    assertEquals(
        Optional.empty(),
        assertInstanceOf(
                CellReport.FormulaReport.class,
                InspectionResultCellReportSupport.toCellReport(numericFormula, temporalOnly, false))
            .temporal());
  }

  private static ExcelCellMetadataSnapshot metadata() {
    return ExcelCellMetadataSnapshot.of(
        Optional.of(new ExcelHyperlink.Url("https://example.com/report")),
        Optional.of(
            new CommentReport(
                "Review",
                "Alice",
                true,
                Optional.of(List.copyOf(List.of(new RichTextRunReport("Review", fontSnapshot())))),
                Optional.of(new CommentAnchorReport(0, 0, 1, 2)))));
  }

  private static CellStyleReport style(String numberFormat) {
    CellBorderSideReport emptySide = new CellBorderSideReport.None();
    return new CellStyleReport(
        numberFormat,
        new CellAlignmentReport(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        fontSnapshot(),
        CellFillReport.pattern(ExcelFillPattern.NONE),
        new CellBorderReport(emptySide, emptySide, emptySide, emptySide),
        new CellProtectionReport(true, false));
  }

  private static CellFontReport fontSnapshot() {
    return new CellFontReport(
        false,
        false,
        "Aptos",
        FontHeightReport.fromPoints(BigDecimal.valueOf(11)),
        null,
        false,
        false);
  }

  private static CellStyleReport styleReport() {
    CellBorderSideReport emptySide = new CellBorderSideReport.None();
    return new CellStyleReport(
        "General",
        new dev.erst.gridgrind.core.model.style.CellAlignmentReport(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new CellFontReport(
            false,
            false,
            "Aptos",
            new FontHeightReport(220, BigDecimal.valueOf(11)),
            null,
            false,
            false),
        CellFillReport.pattern(ExcelFillPattern.NONE),
        new CellBorderReport(emptySide, emptySide, emptySide, emptySide),
        new CellProtectionReport(true, false));
  }

  private static RichTextRunReport runReport(String text) {
    return new RichTextRunReport(
        text,
        new CellFontReport(
            false,
            false,
            "Aptos",
            new FontHeightReport(220, BigDecimal.valueOf(11)),
            null,
            false,
            false));
  }
}
