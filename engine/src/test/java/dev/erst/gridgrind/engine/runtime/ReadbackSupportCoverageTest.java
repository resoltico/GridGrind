package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.CellBorderReport;
import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellColorReport;
import dev.erst.gridgrind.contract.dto.CellFillReport;
import dev.erst.gridgrind.contract.dto.CellFontReport;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.CellReport;
import dev.erst.gridgrind.contract.dto.CellScalarValue;
import dev.erst.gridgrind.contract.dto.CellStyleReport;
import dev.erst.gridgrind.contract.dto.CellTemporalKind;
import dev.erst.gridgrind.contract.dto.CellTemporalReport;
import dev.erst.gridgrind.contract.dto.CellValueReport;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.RichTextRunReport;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelBorderSideSnapshot;
import dev.erst.gridgrind.excel.ExcelBorderSnapshot;
import dev.erst.gridgrind.excel.ExcelCellAlignmentSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFillSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFontSnapshot;
import dev.erst.gridgrind.excel.ExcelCellMetadataSnapshot;
import dev.erst.gridgrind.excel.ExcelCellProtectionSnapshot;
import dev.erst.gridgrind.excel.ExcelCellReadFacet;
import dev.erst.gridgrind.excel.ExcelCellReadProjection;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelCellStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelColor;
import dev.erst.gridgrind.excel.ExcelColorSnapshot;
import dev.erst.gridgrind.excel.ExcelCommentAnchorSnapshot;
import dev.erst.gridgrind.excel.ExcelCommentSnapshot;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelRichTextRunSnapshot;
import dev.erst.gridgrind.excel.ExcelRichTextSnapshot;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
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
        InspectionResultCellStyleReportSupport.toCellColorReport((ExcelColor) null));
    assertEquals(
        Optional.of(CellColorReport.rgb("#112233")),
        InspectionResultCellStyleReportSupport.toCellColorReport(ExcelColor.rgb("#112233")));
    assertEquals(
        Optional.of(CellColorReport.rgb("#112233", 0.25d)),
        InspectionResultCellStyleReportSupport.toCellColorReport(
            ExcelColor.rgb("#112233", Optional.of(0.25d))));
    assertEquals(
        Optional.of(CellColorReport.theme(4)),
        InspectionResultCellStyleReportSupport.toCellColorReport(ExcelColor.theme(4)));
    assertEquals(
        Optional.of(CellColorReport.theme(4, -0.25d)),
        InspectionResultCellStyleReportSupport.toCellColorReport(
            ExcelColor.theme(4, Optional.of(-0.25d))));
    assertEquals(
        Optional.of(CellColorReport.indexed(9)),
        InspectionResultCellStyleReportSupport.toCellColorReport(ExcelColor.indexed(9)));
    assertEquals(
        Optional.of(CellColorReport.indexed(9, 0.5d)),
        InspectionResultCellStyleReportSupport.toCellColorReport(
            ExcelColor.indexed(9, Optional.of(0.5d))));
    assertInstanceOf(
        CellBorderSideReport.DefaultColor.class,
        InspectionResultCellStyleReportSupport.toCellBorderSideReport(
            new ExcelBorderSideSnapshot(ExcelBorderStyle.THIN, null)));
  }

  @Test
  void cellReportConversionProjectsFacetsAcrossAllSnapshotKinds() {
    ExcelCellReadProjection allFacets =
        new ExcelCellReadProjection(Set.of(ExcelCellReadFacet.values()));
    ExcelCellReadProjection valueOnly = ExcelCellReadProjection.defaults();
    ExcelCellReadProjection formulaOnly =
        new ExcelCellReadProjection(Set.of(ExcelCellReadFacet.FORMULA));
    ExcelRichTextSnapshot richText =
        new ExcelRichTextSnapshot(List.of(new ExcelRichTextRunSnapshot("Ada", fontSnapshot())));
    ExcelCellSnapshot.TextSnapshot textSnapshot =
        new ExcelCellSnapshot.TextSnapshot(
            "A1", "Ada", style("General"), metadata(), "Ada", richText);
    ExcelCellSnapshot.NumberSnapshot dateSnapshot =
        new ExcelCellSnapshot.NumberSnapshot(
            "A2", "7/1/2025", style("m/d/yyyy"), metadata(), 45839.0d);
    ExcelCellSnapshot.NumberSnapshot dateTimeSnapshot =
        new ExcelCellSnapshot.NumberSnapshot(
            "A2b", "7/1/2025 12:00", style("m/d/yyyy h:mm"), metadata(), 45839.5d);
    ExcelCellSnapshot.NumberSnapshot timeSnapshot =
        new ExcelCellSnapshot.NumberSnapshot("A2c", "12:00", style("h:mm"), metadata(), 0.5d);
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

    assertEquals(Optional.of("Ada"), textReport.textValue());
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
    assertEquals(Optional.empty(), formulaFormulaOnly.evaluation());

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
        CellValueReport.BlankValue.class,
        InspectionResultCellReportSupport.toCellValueReport(blankSnapshot, valueOnly, false));
    assertInstanceOf(
        CellValueReport.TextValue.class,
        InspectionResultCellReportSupport.toCellValueReport(textSnapshot, allFacets, false));
    assertInstanceOf(
        CellValueReport.NumberValue.class,
        InspectionResultCellReportSupport.toCellValueReport(dateSnapshot, allFacets, false));
    assertInstanceOf(
        CellValueReport.BooleanValue.class,
        InspectionResultCellReportSupport.toCellValueReport(booleanSnapshot, valueOnly, false));
    assertInstanceOf(
        CellValueReport.ErrorValue.class,
        InspectionResultCellReportSupport.toCellValueReport(errorSnapshot, valueOnly, false));

    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                InspectionResultCellReportSupport.toCellValueReport(
                    formulaSnapshot, valueOnly, false));
    assertEquals("Formula evaluations must not recursively remain FORMULA", exception.getMessage());

    assertEquals(
        Optional.empty(), InspectionResultCellReportSupport.toRichTextRunReports(Optional.empty()));
    assertEquals(
        Optional.empty(),
        InspectionResultCellReportSupport.toRichTextRunReports((ExcelRichTextSnapshot) null));
    assertEquals(
        CellColorReport.rgb("#112233"),
        InspectionResultCellReportSupport.toCellColorReport(ExcelColorSnapshot.rgb("#112233"))
            .orElseThrow());
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
            new dev.erst.gridgrind.contract.dto.CellInput.Text(TextSourceInput.inline("Ada")),
            false));
    assertTrue(
        SemanticSelectorKeyMatchSupport.matchesKeyCell(
            numberSnapshot,
            new dev.erst.gridgrind.contract.dto.CellInput.NumberValue(42.0d),
            false));
    assertTrue(
        SemanticSelectorKeyMatchSupport.matchesKeyCell(
            booleanSnapshot,
            new dev.erst.gridgrind.contract.dto.CellInput.BooleanValue(true),
            false));
    assertTrue(
        SemanticSelectorKeyMatchSupport.matchesKeyCell(
            formulaSnapshot,
            new dev.erst.gridgrind.contract.dto.CellInput.Formula(
                TextSourceInput.inline("SUM(A1:A2)")),
            false));
    assertTrue(
        SemanticSelectorKeyMatchSupport.matchesKeyCell(
            blankSnapshot, new dev.erst.gridgrind.contract.dto.CellInput.Blank(), false));
    assertFalse(
        SemanticSelectorKeyMatchSupport.matchesKeyCell(
            textSnapshot,
            new dev.erst.gridgrind.contract.dto.CellInput.Text(TextSourceInput.inline("Grace")),
            false));
    assertThrows(
        IllegalStateException.class,
        () ->
            SemanticSelectorKeyMatchSupport.matchesKeyCell(
                textSnapshot,
                new dev.erst.gridgrind.contract.dto.CellInput.Text(TextSourceInput.standardInput()),
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
            "A1",
            Optional.empty(),
            Optional.of(style),
            Optional.empty(),
            Optional.empty(),
            Optional.of("A1"),
            Optional.of(new CellValueReport.BlankValue()));
    CellReport.FormulaReport textFormula =
        new CellReport.FormulaReport(
            "A2",
            Optional.empty(),
            Optional.of(style),
            Optional.empty(),
            Optional.empty(),
            Optional.of("A2"),
            Optional.of(
                new CellValueReport.TextValue("Ada", Optional.of(List.of(runReport("Ada"))))));
    CellReport.FormulaReport numberFormula =
        new CellReport.FormulaReport(
            "A3",
            Optional.empty(),
            Optional.of(style),
            Optional.empty(),
            Optional.empty(),
            Optional.of("A3"),
            Optional.of(
                new CellValueReport.NumberValue(
                    42.0d,
                    Optional.of(
                        CellTemporalReport.temporal(CellTemporalKind.DATE, "2026-07-01")))));
    CellReport.FormulaReport booleanFormula =
        new CellReport.FormulaReport(
            "A4",
            Optional.empty(),
            Optional.of(style),
            Optional.empty(),
            Optional.empty(),
            Optional.of("A4"),
            Optional.of(new CellValueReport.BooleanValue(true)));
    CellReport.FormulaReport errorFormula =
        new CellReport.FormulaReport(
            "A5",
            Optional.empty(),
            Optional.of(style),
            Optional.empty(),
            Optional.empty(),
            Optional.of("A5"),
            Optional.of(new CellValueReport.ErrorValue("#REF!")));
    CellReport.FormulaReport missingEvaluation =
        new CellReport.FormulaReport(
            "A6",
            Optional.empty(),
            Optional.of(style),
            Optional.empty(),
            Optional.empty(),
            Optional.of("A6"),
            Optional.empty());

    assertTrue(AssertionValueEvaluator.matchesCellValue(blankFormula, new CellScalarValue.Blank()));
    assertTrue(
        AssertionValueEvaluator.matchesCellValue(textFormula, new CellScalarValue.Text("Ada")));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(textFormula, new CellScalarValue.Text("Grace")));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(numberFormula, new CellScalarValue.Text("Ada")));
    assertTrue(
        AssertionValueEvaluator.matchesCellValue(
            numberFormula, new CellScalarValue.NumberValue(42.0d)));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(
            numberFormula, new CellScalarValue.NumberValue(41.0d)));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(
            textFormula, new CellScalarValue.NumberValue(42.0d)));
    assertTrue(
        AssertionValueEvaluator.matchesCellValue(
            booleanFormula, new CellScalarValue.BooleanValue(true)));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(
            booleanFormula, new CellScalarValue.BooleanValue(false)));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(
            numberFormula, new CellScalarValue.BooleanValue(true)));
    assertTrue(
        AssertionValueEvaluator.matchesCellValue(
            errorFormula, new CellScalarValue.ErrorValue("#REF!")));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(
            errorFormula, new CellScalarValue.ErrorValue("#DIV/0!")));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(
            booleanFormula, new CellScalarValue.ErrorValue("#REF!")));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(
            missingEvaluation, new CellScalarValue.Text("Ada")));
  }

  private static ExcelCellMetadataSnapshot metadata() {
    return ExcelCellMetadataSnapshot.of(
        Optional.of(new ExcelHyperlink.Url("https://example.com/report")),
        Optional.of(
            new ExcelCommentSnapshot(
                "Review",
                "Alice",
                true,
                Optional.of(
                    new ExcelRichTextSnapshot(
                        List.of(new ExcelRichTextRunSnapshot("Review", fontSnapshot())))),
                Optional.of(new ExcelCommentAnchorSnapshot(0, 0, 1, 2)))));
  }

  private static ExcelCellStyleSnapshot style(String numberFormat) {
    ExcelBorderSideSnapshot emptySide = new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null);
    return new ExcelCellStyleSnapshot(
        numberFormat,
        new ExcelCellAlignmentSnapshot(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        fontSnapshot(),
        ExcelCellFillSnapshot.pattern(ExcelFillPattern.NONE),
        new ExcelBorderSnapshot(emptySide, emptySide, emptySide, emptySide),
        new ExcelCellProtectionSnapshot(true, false));
  }

  private static ExcelCellFontSnapshot fontSnapshot() {
    return new ExcelCellFontSnapshot(
        false,
        false,
        "Aptos",
        ExcelFontHeight.fromPoints(BigDecimal.valueOf(11)),
        null,
        false,
        false);
  }

  private static CellStyleReport styleReport() {
    CellBorderSideReport emptySide = new CellBorderSideReport.None();
    return new CellStyleReport(
        "General",
        new dev.erst.gridgrind.contract.dto.CellAlignmentReport(
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
