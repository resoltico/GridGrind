package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellTemporalKind;
import dev.erst.gridgrind.core.model.value.CellTemporalReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import java.math.BigDecimal;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/**
 * Coverage tests for runtime readback helpers introduced by projection-aware cell introspection.
 */
class ReadbackSupportCoverageTest {
  @Test
  void semanticSelectorKeyMatchingHandlesSupportedScalarKinds() {
    CellReport.TextReport textSnapshot =
        new CellReport.TextReport(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
            Optional.ofNullable("Ada"),
            Optional.ofNullable(style("General")),
            Optional.empty(),
            Optional.empty(),
            Optional.of(new dev.erst.gridgrind.core.model.value.CellValue.Text("Ada")),
            Optional.ofNullable(null));
    CellReport.NumberReport numberSnapshot =
        new CellReport.NumberReport(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A2"),
            Optional.ofNullable("42"),
            Optional.ofNullable(style("0.00")),
            Optional.empty(),
            Optional.empty(),
            Optional.of(new dev.erst.gridgrind.core.model.value.CellValue.NumberValue(42.0d)),
            Optional.empty());
    CellReport.BooleanReport booleanSnapshot =
        new CellReport.BooleanReport(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A3"),
            Optional.ofNullable("TRUE"),
            Optional.ofNullable(style("General")),
            Optional.empty(),
            Optional.empty(),
            Optional.of(new dev.erst.gridgrind.core.model.value.CellValue.BooleanValue(true)));
    CellReport.FormulaReport formulaSnapshot =
        new CellReport.FormulaReport(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A4"),
            Optional.ofNullable("42"),
            Optional.ofNullable(style("General")),
            Optional.empty(),
            Optional.empty(),
            Optional.ofNullable("SUM(A1:A2)"),
            java.util.Optional.of(numberSnapshot)
                .flatMap(CellReport::value)
                .map(value -> (dev.erst.gridgrind.core.model.value.CellValue) value),
            Optional.empty(),
            Optional.empty());
    CellReport.BlankReport blankSnapshot =
        new CellReport.BlankReport(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A5"),
            Optional.ofNullable(""),
            Optional.ofNullable(style("General")),
            Optional.empty(),
            Optional.empty(),
            Optional.of(new dev.erst.gridgrind.core.model.value.CellValue.Blank()));

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
