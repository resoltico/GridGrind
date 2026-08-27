package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
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
            "A1", Optional.empty(), Optional.empty(), Optional.empty(), Optional.empty());
    CellReport.TextReport text =
        new CellReport.TextReport(
            "A2",
            Optional.of("Ada"),
            Optional.of(style()),
            Optional.empty(),
            Optional.empty(),
            Optional.of("Ada"),
            Optional.of(List.of(run("Ada"))));
    CellReport.NumberReport number =
        new CellReport.NumberReport(
            "A3",
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    CellReport.BooleanReport bool =
        new CellReport.BooleanReport(
            "A4",
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.of(true));
    CellReport.ErrorReport error =
        new CellReport.ErrorReport(
            "A5",
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.of("#REF!"));
    CellReport.FormulaReport formula =
        new CellReport.FormulaReport(
            "A6",
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.of("SUM(A1:A3)"),
            Optional.empty());

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
                " ", Optional.empty(), Optional.empty(), Optional.empty(), Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.TextReport(
                " ",
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of("Ada"),
                Optional.of(List.of(run("Ada")))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.NumberReport(
                " ",
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.BooleanReport(
                " ",
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(true)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.ErrorReport(
                " ",
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of("#REF!")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.FormulaReport(
                " ",
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of("SUM(A1:A3)"),
                Optional.empty()));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.TextReport(
                "A1",
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(List.of(run("Ada")))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.TextReport(
                "A1",
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of("Ada"),
                Optional.of(List.of())));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.TextReport(
                "A1",
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of("Ada"),
                Optional.of(List.of(run("Mismatch")))));
    assertThrows(
        NullPointerException.class,
        () ->
            new CellReport.TextReport(
                "A1",
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of("Ada"),
                Optional.of(java.util.Arrays.asList(run("Ada"), null))));
  }

  @Test
  void cellValueTemporalAndWindowReportsCoverResidualValidationBranches() {
    CellValueReport.BlankValue blank = new CellValueReport.BlankValue();
    CellValueReport.TextValue text = new CellValueReport.TextValue("Ada", nullOptional());
    CellTemporalReport notDate = CellTemporalReport.notDate();
    CellTemporalReport date = CellTemporalReport.temporal(CellTemporalKind.DATE, "2026-07-01");
    CellTemporalReport time = CellTemporalReport.temporal(CellTemporalKind.TIME, "12:00");
    CellValueReport.NumberValue number = new CellValueReport.NumberValue(42.0d, Optional.of(date));
    CellValueReport.BooleanValue bool = new CellValueReport.BooleanValue(true);
    CellValueReport.ErrorValue error = new CellValueReport.ErrorValue("#CIRCULAR_REF!");
    WindowDimensionsReport dimensions = new WindowDimensionsReport(2, 2);
    WindowReport.Sparse sparse =
        new WindowReport.Sparse(
            "Budget",
            "A1",
            dimensions,
            List.of(
                new CellReport.TextReport(
                    "A1",
                    Optional.of("Ada"),
                    Optional.of(style()),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.of("Ada"),
                    Optional.empty())));
    WindowReport.Dense dense =
        new WindowReport.Dense(
            "Budget",
            "A1",
            dimensions,
            List.of(
                new WindowRowReport(
                    0,
                    List.of(
                        new CellReport.BlankReport(
                            "A1",
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty())))));

    assertEquals("BLANK", blank.type());
    assertEquals("TEXT", text.type());
    assertEquals("NUMBER", number.type());
    assertEquals("BOOLEAN", bool.type());
    assertEquals("ERROR", error.type());
    assertEquals("#CIRCULAR_REF!", error.errorValue());
    assertFalse(notDate.isDate());
    assertTrue(date.isDate());
    assertEquals(Optional.of(CellTemporalKind.DATE), date.kind());
    assertEquals(Optional.of(CellTemporalKind.TIME), time.kind());
    assertEquals("SPARSE", sparse.shape());
    assertEquals("DENSE", dense.shape());

    assertThrows(
        NullPointerException.class, () -> new CellValueReport.NumberValue(null, Optional.empty()));
    assertThrows(NullPointerException.class, () -> new CellValueReport.BooleanValue(null));
    assertThrows(IllegalArgumentException.class, () -> new CellValueReport.ErrorValue(" "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellReport.ErrorReport(
                "A1",
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of("~CIRCULAR~REF~")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellValueReport.TextValue("Ada", Optional.of(List.of())));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellValueReport.TextValue("Ada", Optional.of(List.of(run("Mismatch")))));
    assertThrows(
        NullPointerException.class,
        () ->
            new CellValueReport.TextValue(
                "Ada", Optional.of(java.util.Arrays.asList(run("Ada"), null))));
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
                " ",
                "A1",
                dimensions,
                List.of(
                    new CellReport.TextReport(
                        "A1",
                        Optional.of("Ada"),
                        Optional.of(style()),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.of("Ada"),
                        Optional.empty()))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WindowReport.Sparse(
                "Budget",
                " ",
                dimensions,
                List.of(
                    new CellReport.TextReport(
                        "A1",
                        Optional.of("Ada"),
                        Optional.of(style()),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.of("Ada"),
                        Optional.empty()))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WindowReport.Sparse(
                "Budget",
                "A1",
                dimensions,
                List.of(
                    new CellReport.BlankReport(
                        "A1",
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty()))));
  }

  @SuppressWarnings("NullOptional")
  private static <T> Optional<T> nullOptional() {
    return null;
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
            null,
            false,
            false),
        CellFillReport.pattern(ExcelFillPattern.NONE),
        new CellBorderReport(emptySide, emptySide, emptySide, emptySide),
        new CellProtectionReport(true, false));
  }
}
