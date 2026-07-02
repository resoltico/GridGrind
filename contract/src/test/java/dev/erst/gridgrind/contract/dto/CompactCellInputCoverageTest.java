package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Covers compact cell input encodings and residual constructor validation paths. */
class CompactCellInputCoverageTest {
  @Test
  void compactGridAndRowEncodingsExpandIntoCanonicalPerCellInputs() {
    CellGridInput.TextRows textRows = new CellGridInput.TextRows(List.of(List.of("Ada", "Ready")));
    CellGridInput.NumberRows numberRows = new CellGridInput.NumberRows(List.of(List.of(1.5d, 2d)));
    CellGridInput.BooleanRows booleanRows =
        new CellGridInput.BooleanRows(List.of(List.of(true, false)));
    CellGridInput.ErrorRows errorRows =
        new CellGridInput.ErrorRows(List.of(List.of("#REF!", "#DIV/0!")));
    CellGridInput.DateRows dateRows =
        new CellGridInput.DateRows(List.of(List.of(LocalDate.of(2026, 6, 12))));
    CellGridInput.DateTimeRows dateTimeRows =
        new CellGridInput.DateTimeRows(List.of(List.of(LocalDateTime.of(2026, 6, 12, 9, 30))));
    CellGridInput.FormulaRows formulaRows =
        new CellGridInput.FormulaRows(List.of(List.of("=SUM(A1:A2)")));

    assertEquals(
        List.of(
            List.of(
                new CellInput.Text(TextSourceInput.inline("Ada")),
                new CellInput.Text(TextSourceInput.inline("Ready")))),
        textRows.toCellInputRows());
    assertEquals(
        List.of(List.of(new CellInput.NumberValue(1.5d), new CellInput.NumberValue(2d))),
        numberRows.toCellInputRows());
    assertEquals(
        List.of(List.of(new CellInput.BooleanValue(true), new CellInput.BooleanValue(false))),
        booleanRows.toCellInputRows());
    assertEquals(
        List.of(List.of(new CellInput.ErrorValue("#REF!"), new CellInput.ErrorValue("#DIV/0!"))),
        errorRows.toCellInputRows());
    assertEquals(
        List.of(List.of(new CellInput.Date(LocalDate.of(2026, 6, 12)))),
        dateRows.toCellInputRows());
    assertEquals(
        List.of(List.of(new CellInput.DateTime(LocalDateTime.of(2026, 6, 12, 9, 30)))),
        dateTimeRows.toCellInputRows());
    assertEquals(
        "SUM(A1:A2)",
        ((TextSourceInput.Inline)
                assertInstanceOf(
                        CellInput.Formula.class,
                        formulaRows.toCellInputRows().getFirst().getFirst())
                    .source())
            .text());

    assertEquals(
        List.of(
            new CellInput.Text(TextSourceInput.inline("Ada")),
            new CellInput.Text(TextSourceInput.inline("Ready"))),
        new CellRowInput.TextValues(List.of("Ada", "Ready")).toCellInputs());
    assertEquals(
        List.of(new CellInput.NumberValue(1.5d), new CellInput.NumberValue(2d)),
        new CellRowInput.NumberValues(List.of(1.5d, 2d)).toCellInputs());
    assertEquals(
        List.of(new CellInput.BooleanValue(true), new CellInput.BooleanValue(false)),
        new CellRowInput.BooleanValues(List.of(true, false)).toCellInputs());
    assertEquals(
        List.of(new CellInput.ErrorValue("#REF!"), new CellInput.ErrorValue("#DIV/0!")),
        new CellRowInput.ErrorValues(List.of("#REF!", "#DIV/0!")).toCellInputs());
    assertEquals(
        List.of(new CellInput.Date(LocalDate.of(2026, 6, 12))),
        new CellRowInput.DateValues(List.of(LocalDate.of(2026, 6, 12))).toCellInputs());
    assertEquals(
        List.of(new CellInput.DateTime(LocalDateTime.of(2026, 6, 12, 9, 30))),
        new CellRowInput.DateTimeValues(List.of(LocalDateTime.of(2026, 6, 12, 9, 30)))
            .toCellInputs());
    assertEquals(
        "SUM(B1:B2)",
        ((TextSourceInput.Inline)
                assertInstanceOf(
                        CellInput.Formula.class,
                        new CellRowInput.FormulaValues(List.of("=SUM(B1:B2)"))
                            .toCellInputs()
                            .getFirst())
                    .source())
            .text());
  }

  @Test
  void compactCellInputsRejectNullsEmptiesAndNonRectangularShapes() {
    assertEquals(
        "cells must not be empty",
        assertThrows(IllegalArgumentException.class, () -> new CellGridInput.Typed(List.of()))
            .getMessage());
    assertEquals(
        "cells must describe a rectangular matrix",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new CellGridInput.Typed(
                        List.of(
                            List.of(new CellInput.Blank()),
                            List.of(new CellInput.Blank(), new CellInput.Blank()))))
            .getMessage());
    assertEquals(
        "cells must not contain empty rows",
        assertThrows(
                IllegalArgumentException.class, () -> new CellGridInput.Typed(List.of(List.of())))
            .getMessage());
    assertEquals(
        "cells must not contain null rows",
        assertThrows(
                NullPointerException.class,
                () -> new CellGridInput.Typed(Arrays.asList((List<CellInput>) null)))
            .getMessage());
    assertEquals(
        "cells must not be null",
        assertThrows(NullPointerException.class, () -> new CellGridInput.TextRows(null))
            .getMessage());
    assertEquals(
        "cells must not be empty",
        assertThrows(IllegalArgumentException.class, () -> new CellGridInput.TextRows(List.of()))
            .getMessage());
    assertEquals(
        "cells must not contain null rows",
        assertThrows(
                NullPointerException.class,
                () -> new CellGridInput.TextRows(Arrays.asList((List<String>) null)))
            .getMessage());
    assertEquals(
        "cells must describe a rectangular matrix",
        assertThrows(
                IllegalArgumentException.class,
                () -> new CellGridInput.TextRows(List.of(List.of("Ada"), List.of("Ada", "Bob"))))
            .getMessage());
    assertEquals(
        List.of(
            List.of(new CellInput.Text(TextSourceInput.inline("Ada"))),
            List.of(new CellInput.Text(TextSourceInput.inline("Bob")))),
        new CellGridInput.TextRows(List.of(List.of("Ada"), List.of("Bob"))).toCellInputRows());
    assertEquals(
        "cells must not contain empty rows",
        assertThrows(
                IllegalArgumentException.class,
                () -> new CellGridInput.TextRows(List.of(List.of())))
            .getMessage());
    assertEquals(
        "cells must not contain null cell values",
        assertThrows(
                NullPointerException.class,
                () -> new CellGridInput.Typed(List.of(Arrays.asList(new CellInput.Blank(), null))))
            .getMessage());
    assertEquals(
        "cells must not contain null values",
        assertThrows(
                NullPointerException.class,
                () -> new CellGridInput.TextRows(List.of(Arrays.asList("Ada", null))))
            .getMessage());

    assertEquals(
        "cells must not be empty",
        assertThrows(IllegalArgumentException.class, () -> new CellRowInput.Typed(List.of()))
            .getMessage());
    assertEquals(
        "cells must not be null",
        assertThrows(NullPointerException.class, () -> new CellRowInput.TextValues(null))
            .getMessage());
    assertEquals(
        "cells must not be empty",
        assertThrows(IllegalArgumentException.class, () -> new CellRowInput.TextValues(List.of()))
            .getMessage());
    assertEquals(
        "cells must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () -> new CellRowInput.Typed(Arrays.asList(new CellInput.Blank(), null)))
            .getMessage());
    assertEquals(
        "cells must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () -> new CellRowInput.TextValues(Arrays.asList("Ada", null)))
            .getMessage());
  }

  @Test
  void explicitErrorLiteralValidationAndPhaseFactoriesCoverResidualBranches() {
    assertEquals("#REF!", new CellInput.ErrorValue("#REF!").error());
    IllegalArgumentException invalidError =
        assertThrows(
            IllegalArgumentException.class, () -> new CellInput.ErrorValue("NOT_AN_ERROR"));
    assertTrue(invalidError.getMessage().contains("#REF!"));
    assertEquals(ExecutionJournal.Status.NOT_STARTED, ExecutionJournal.Phase.notStarted().status());
    assertEquals(
        ExecutionJournal.Status.NOT_REQUESTED, ExecutionJournal.Phase.notRequested().status());
    assertEquals(
        25L,
        assertInstanceOf(
                ExecutionJournal.Phase.Succeeded.class,
                ExecutionJournal.Phase.succeeded(
                    "2026-06-12T09:30:00Z", "2026-06-12T09:30:25Z", 25))
            .timing()
            .orElseThrow()
            .durationMillis());
    assertEquals(
        0L,
        assertInstanceOf(
                ExecutionJournal.Phase.Succeeded.class,
                ExecutionJournal.Phase.succeededWithoutTiming())
            .timing()
            .map(ExecutionJournal.Timing::durationMillis)
            .orElse(0L));
    assertEquals(
        "durationMillis must be >= 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Timing("2026-06-12T09:30:00Z", "2026-06-12T09:30:25Z", -1))
            .getMessage());
  }
}
