package dev.erst.gridgrind.core.model.value;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.layout.WindowDimensionsReport;
import dev.erst.gridgrind.core.model.layout.WindowReport;
import dev.erst.gridgrind.core.model.layout.WindowRowReport;
import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import java.math.BigDecimal;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Coverage tests for cell readback DTOs added by the projection-aware contract. */
class CellReadbackCoverageTest {
  @Test
  void cellReadbackReportsExposePublishedTypesAndValidateConstructors() {
    CellReport.BlankReport blank =
        new CellReport.BlankReport(
            new CellAddress("A1"),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            java.util.Optional.empty());
    CellReport.TextReport text =
        new CellReport.TextReport(
            new CellAddress("A2"),
            Optional.of("Ada"),
            Optional.of(style()),
            Optional.empty(),
            Optional.empty(),
            java.util.Optional.of(new dev.erst.gridgrind.core.model.value.CellValue.Text("Ada")),
            Optional.of(List.of(run("Ada"))));
    CellReport.NumberReport number =
        new CellReport.NumberReport(
            new CellAddress("A3"),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            java.util.Optional.empty(),
            Optional.empty());
    CellReport.BooleanReport bool =
        new CellReport.BooleanReport(
            new CellAddress("A4"),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            java.util.Optional.of(
                new dev.erst.gridgrind.core.model.value.CellValue.BooleanValue(true)));
    CellReport.ErrorReport error =
        new CellReport.ErrorReport(
            new CellAddress("A5"),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            java.util.Optional.of(
                new dev.erst.gridgrind.core.model.value.CellValue.ErrorValue("#REF!")));
    CellReport.FormulaReport formula =
        new CellReport.FormulaReport(
            new CellAddress("A6"),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.of("SUM(A1:A3)"),
            Optional.empty(),
            java.util.Optional.empty(),
            java.util.Optional.empty());

    assertEquals("BLANK", blank.type());
    assertEquals("TEXT", text.type());
    assertEquals("NUMBER", number.type());
    assertEquals("BOOLEAN", bool.type());
    assertEquals("ERROR", error.type());
    assertEquals("FORMULA", formula.type());

    CellBorderSideReport.DefaultColor defaultColor =
        new CellBorderSideReport.DefaultColor(ExcelBorderStyle.THIN);
    assertEquals(ExcelBorderStyle.THIN, defaultColor.style());
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellBorderSideReport.DefaultColor(ExcelBorderStyle.NONE));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.BlankReport(
                new CellAddress(" "),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                java.util.Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.TextReport(
                new CellAddress(" "),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                java.util.Optional.of(
                    new dev.erst.gridgrind.core.model.value.CellValue.Text("Ada")),
                Optional.of(List.of(run("Ada")))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.NumberReport(
                new CellAddress(" "),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                java.util.Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.BooleanReport(
                new CellAddress(" "),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                java.util.Optional.of(
                    new dev.erst.gridgrind.core.model.value.CellValue.BooleanValue(true))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.ErrorReport(
                new CellAddress(" "),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                java.util.Optional.of(
                    new dev.erst.gridgrind.core.model.value.CellValue.ErrorValue("#REF!"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.FormulaReport(
                new CellAddress(" "),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of("SUM(A1:A3)"),
                Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty()));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.TextReport(
                new CellAddress("A1"),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                java.util.Optional.empty(),
                Optional.of(List.of(run("Ada")))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.TextReport(
                new CellAddress("A1"),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                java.util.Optional.of(
                    new dev.erst.gridgrind.core.model.value.CellValue.Text("Ada")),
                Optional.of(List.of())));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.TextReport(
                new CellAddress("A1"),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                java.util.Optional.of(
                    new dev.erst.gridgrind.core.model.value.CellValue.Text("Ada")),
                Optional.of(List.of(run("Mismatch")))));
    assertThrows(
        NullPointerException.class,
        () ->
            new CellReport.TextReport(
                new CellAddress("A1"),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                java.util.Optional.of(
                    new dev.erst.gridgrind.core.model.value.CellValue.Text("Ada")),
                Optional.of(java.util.Arrays.asList(run("Ada"), null))));
  }

  @Test
  void cellValueTemporalAndWindowReportsCoverResidualValidationBranches() {
    CellValue.Blank blank = new CellValue.Blank();
    CellValue.Text text = new CellValue.Text("Ada");
    CellTemporalReport notDate = CellTemporalReport.notDate();
    CellTemporalReport date = CellTemporalReport.temporal(CellTemporalKind.DATE, "2026-07-01");
    CellTemporalReport time = CellTemporalReport.temporal(CellTemporalKind.TIME, "12:00");
    CellValue.NumberValue number = new CellValue.NumberValue(42.0d);
    CellValue.BooleanValue bool = new CellValue.BooleanValue(true);
    CellValue.ErrorValue error = new CellValue.ErrorValue("#CIRCULAR_REF!");
    WindowDimensionsReport dimensions = new WindowDimensionsReport(2, 2);
    WindowReport.Sparse sparse =
        new WindowReport.Sparse(
            new SheetLabel("Budget"),
            new CellAddress("A1"),
            dimensions,
            List.of(
                new CellReport.TextReport(
                    new CellAddress("A1"),
                    Optional.of("Ada"),
                    Optional.of(style()),
                    Optional.empty(),
                    Optional.empty(),
                    java.util.Optional.of(
                        new dev.erst.gridgrind.core.model.value.CellValue.Text("Ada")),
                    Optional.empty())));
    WindowReport.Dense dense =
        new WindowReport.Dense(
            new SheetLabel("Budget"),
            new CellAddress("A1"),
            dimensions,
            List.of(
                new WindowRowReport(
                    0,
                    List.of(
                        new CellReport.BlankReport(
                            new CellAddress("A1"),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            java.util.Optional.empty())))));

    assertEquals("BLANK", blank.type());
    assertEquals("TEXT", text.type());
    assertEquals("NUMBER", number.type());
    assertEquals("BOOLEAN", bool.type());
    assertEquals("ERROR", error.type());
    assertEquals("#CIRCULAR_REF!", error.value());
    assertFalse(notDate.isDate());
    assertTrue(date.isDate());
    assertEquals(Optional.of(CellTemporalKind.DATE), date.kind());
    assertEquals(Optional.of(CellTemporalKind.TIME), time.kind());
    assertEquals("SPARSE", sparse.type());
    assertEquals("DENSE", dense.type());

    assertThrows(IllegalArgumentException.class, () -> new CellValue.ErrorValue(" "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.ErrorReport(
                new CellAddress("A1"),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                java.util.Optional.of(
                    new dev.erst.gridgrind.core.model.value.CellValue.ErrorValue(
                        "~CIRCULAR~REF~"))));
    assertThrows(IllegalArgumentException.class, () -> formulaText("Ada", Optional.of(List.of())));
    assertThrows(
        IllegalArgumentException.class,
        () -> formulaText("Ada", Optional.of(List.of(run("Mismatch")))));
    assertThrows(
        NullPointerException.class,
        () -> formulaText("Ada", Optional.of(java.util.Arrays.asList(run("Ada"), null))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellTemporalReport(true, Optional.empty(), Optional.of("2026-07-01")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellTemporalReport(true, Optional.of(CellTemporalKind.DATE), Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellTemporalReport(true, Optional.of(CellTemporalKind.DATE), Optional.of(" ")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellTemporalReport(false, Optional.of(CellTemporalKind.DATE), Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellTemporalReport(false, Optional.empty(), Optional.of("2026-07-01")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WindowReport.Sparse(
                new SheetLabel(" "),
                new CellAddress("A1"),
                dimensions,
                List.of(
                    new CellReport.TextReport(
                        new CellAddress("A1"),
                        Optional.of("Ada"),
                        Optional.of(style()),
                        Optional.empty(),
                        Optional.empty(),
                        java.util.Optional.of(
                            new dev.erst.gridgrind.core.model.value.CellValue.Text("Ada")),
                        Optional.empty()))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WindowReport.Sparse(
                new SheetLabel("Budget"),
                new CellAddress(" "),
                dimensions,
                List.of(
                    new CellReport.TextReport(
                        new CellAddress("A1"),
                        Optional.of("Ada"),
                        Optional.of(style()),
                        Optional.empty(),
                        Optional.empty(),
                        java.util.Optional.of(
                            new dev.erst.gridgrind.core.model.value.CellValue.Text("Ada")),
                        Optional.empty()))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WindowReport.Sparse(
                new SheetLabel("Budget"),
                new CellAddress("A1"),
                dimensions,
                List.of(
                    new CellReport.BlankReport(
                        new CellAddress("A1"),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        java.util.Optional.empty()))));
  }

  private static RichTextRunReport run(String text) {
    return new RichTextRunReport(text, style().font());
  }

  private static CellStyleReport style() {
    CellBorderSideReport emptySide = new CellBorderSideReport.None();
    return new CellStyleReport(
        "General",
        new CellAlignmentReport(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new CellFontReport(
            false,
            false,
            "Aptos",
            new FontHeightReport(220, BigDecimal.valueOf(11)),
            Optional.empty(),
            false,
            false),
        CellFillReport.pattern(ExcelFillPattern.NONE),
        new CellBorderReport(emptySide, emptySide, emptySide, emptySide),
        new CellProtectionReport(true, false));
  }

  private static CellReport.FormulaReport formulaText(
      String text, Optional<List<RichTextRunReport>> runs) {
    return new CellReport.FormulaReport(
        new CellAddress("A1"),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.of("A1"),
        Optional.of(new CellValue.Text(text)),
        runs,
        Optional.empty());
  }
}
