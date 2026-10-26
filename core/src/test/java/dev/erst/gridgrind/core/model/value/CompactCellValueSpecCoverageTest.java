package dev.erst.gridgrind.core.model.value;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.result.ExecutionJournal;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Covers compact cell input encodings and residual constructor validation paths. */
class CompactCellValueSpecCoverageTest {
  @Test
  void compactGridAndRowEncodingsExpandIntoCanonicalPerCellInputs() {
    CellGridSpec.TextRows textRows = new CellGridSpec.TextRows(List.of(List.of("Ada", "Ready")));
    CellGridSpec.NumberRows numberRows = new CellGridSpec.NumberRows(List.of(List.of(1.5d, 2d)));
    CellGridSpec.BooleanRows booleanRows =
        new CellGridSpec.BooleanRows(List.of(List.of(true, false)));
    CellGridSpec.ErrorRows errorRows =
        new CellGridSpec.ErrorRows(List.of(List.of("#REF!", "#DIV/0!")));
    CellGridSpec.DateRows dateRows =
        new CellGridSpec.DateRows(List.of(List.of(LocalDate.of(2026, 6, 12))));
    CellGridSpec.DateTimeRows dateTimeRows =
        new CellGridSpec.DateTimeRows(List.of(List.of(LocalDateTime.of(2026, 6, 12, 9, 30))));
    CellGridSpec.FormulaRows formulaRows =
        new CellGridSpec.FormulaRows(List.of(List.of("SUM(A1:A2)")));

    assertEquals(
        List.of(
            List.of(
                new CellValueSpec.Text(TextSourceSpec.inline("Ada")),
                new CellValueSpec.Text(TextSourceSpec.inline("Ready")))),
        textRows.toCellInputRows());
    assertEquals(
        List.of(List.of(new CellValueSpec.NumberValue(1.5d), new CellValueSpec.NumberValue(2d))),
        numberRows.toCellInputRows());
    assertEquals(
        List.of(
            List.of(new CellValueSpec.BooleanValue(true), new CellValueSpec.BooleanValue(false))),
        booleanRows.toCellInputRows());
    assertEquals(
        List.of(
            List.of(
                new CellValueSpec.ErrorValue("#REF!"), new CellValueSpec.ErrorValue("#DIV/0!"))),
        errorRows.toCellInputRows());
    assertEquals(
        List.of(List.of(new CellValueSpec.Date(LocalDate.of(2026, 6, 12)))),
        dateRows.toCellInputRows());
    assertEquals(
        List.of(List.of(new CellValueSpec.DateTime(LocalDateTime.of(2026, 6, 12, 9, 30)))),
        dateTimeRows.toCellInputRows());
    assertEquals(
        "SUM(A1:A2)",
        ((TextSourceSpec.Inline)
                assertInstanceOf(
                        CellValueSpec.Formula.class,
                        formulaRows.toCellInputRows().getFirst().getFirst())
                    .source())
            .text());

    assertEquals(
        List.of(
            new CellValueSpec.Text(TextSourceSpec.inline("Ada")),
            new CellValueSpec.Text(TextSourceSpec.inline("Ready"))),
        new CellRowSpec.TextValues(List.of("Ada", "Ready")).toCellInputs());
    assertEquals(
        List.of(new CellValueSpec.NumberValue(1.5d), new CellValueSpec.NumberValue(2d)),
        new CellRowSpec.NumberValues(List.of(1.5d, 2d)).toCellInputs());
    assertEquals(
        List.of(new CellValueSpec.BooleanValue(true), new CellValueSpec.BooleanValue(false)),
        new CellRowSpec.BooleanValues(List.of(true, false)).toCellInputs());
    assertEquals(
        List.of(new CellValueSpec.ErrorValue("#REF!"), new CellValueSpec.ErrorValue("#DIV/0!")),
        new CellRowSpec.ErrorValues(List.of("#REF!", "#DIV/0!")).toCellInputs());
    assertEquals(
        List.of(new CellValueSpec.Date(LocalDate.of(2026, 6, 12))),
        new CellRowSpec.DateValues(List.of(LocalDate.of(2026, 6, 12))).toCellInputs());
    assertEquals(
        List.of(new CellValueSpec.DateTime(LocalDateTime.of(2026, 6, 12, 9, 30))),
        new CellRowSpec.DateTimeValues(List.of(LocalDateTime.of(2026, 6, 12, 9, 30)))
            .toCellInputs());
    assertEquals(
        "SUM(B1:B2)",
        ((TextSourceSpec.Inline)
                assertInstanceOf(
                        CellValueSpec.Formula.class,
                        new CellRowSpec.FormulaValues(List.of("SUM(B1:B2)"))
                            .toCellInputs()
                            .getFirst())
                    .source())
            .text());
  }

  @Test
  void compactCellInputsRejectNullsEmptiesAndNonRectangularShapes() {
    assertEquals(
        "cells must not be empty",
        assertThrows(IllegalArgumentException.class, () -> new CellGridSpec.Typed(List.of()))
            .getMessage());
    assertEquals(
        "cells must describe a rectangular matrix",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new CellGridSpec.Typed(
                        List.of(
                            List.of(new CellValueSpec.Blank()),
                            List.of(new CellValueSpec.Blank(), new CellValueSpec.Blank()))))
            .getMessage());
    assertEquals(
        "cells must not contain empty rows",
        assertThrows(
                IllegalArgumentException.class, () -> new CellGridSpec.Typed(List.of(List.of())))
            .getMessage());
    assertEquals(
        "cells must not contain null rows",
        assertThrows(
                NullPointerException.class,
                () -> new CellGridSpec.Typed(Arrays.asList((List<CellValueSpec>) null)))
            .getMessage());
    assertEquals(
        "cells must not be null",
        assertThrows(NullPointerException.class, () -> new CellGridSpec.TextRows(null))
            .getMessage());
    assertEquals(
        "cells must not be empty",
        assertThrows(IllegalArgumentException.class, () -> new CellGridSpec.TextRows(List.of()))
            .getMessage());
    assertEquals(
        "cells must not contain null rows",
        assertThrows(
                NullPointerException.class,
                () -> new CellGridSpec.TextRows(Arrays.asList((List<String>) null)))
            .getMessage());
    assertEquals(
        "cells must describe a rectangular matrix",
        assertThrows(
                IllegalArgumentException.class,
                () -> new CellGridSpec.TextRows(List.of(List.of("Ada"), List.of("Ada", "Bob"))))
            .getMessage());
    assertEquals(
        List.of(
            List.of(new CellValueSpec.Text(TextSourceSpec.inline("Ada"))),
            List.of(new CellValueSpec.Text(TextSourceSpec.inline("Bob")))),
        new CellGridSpec.TextRows(List.of(List.of("Ada"), List.of("Bob"))).toCellInputRows());
    assertEquals(
        "cells must not contain empty rows",
        assertThrows(
                IllegalArgumentException.class, () -> new CellGridSpec.TextRows(List.of(List.of())))
            .getMessage());
    assertEquals(
        "cells must not contain null cell values",
        assertThrows(
                NullPointerException.class,
                () ->
                    new CellGridSpec.Typed(List.of(Arrays.asList(new CellValueSpec.Blank(), null))))
            .getMessage());
    assertEquals(
        "cells must not contain null values",
        assertThrows(
                NullPointerException.class,
                () -> new CellGridSpec.TextRows(List.of(Arrays.asList("Ada", null))))
            .getMessage());

    assertEquals(
        "cells must not be empty",
        assertThrows(IllegalArgumentException.class, () -> new CellRowSpec.Typed(List.of()))
            .getMessage());
    assertEquals(
        "cells must not be null",
        assertThrows(NullPointerException.class, () -> new CellRowSpec.TextValues(null))
            .getMessage());
    assertEquals(
        "cells must not be empty",
        assertThrows(IllegalArgumentException.class, () -> new CellRowSpec.TextValues(List.of()))
            .getMessage());
    assertEquals(
        "cells must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () -> new CellRowSpec.Typed(Arrays.asList(new CellValueSpec.Blank(), null)))
            .getMessage());
    assertEquals(
        "cells must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () -> new CellRowSpec.TextValues(Arrays.asList("Ada", null)))
            .getMessage());
  }

  @Test
  void explicitErrorLiteralValidationAndPhaseFactoriesCoverResidualBranches() {
    assertEquals("#REF!", new CellValueSpec.ErrorValue("#REF!").error());
    assertEquals(
        "storedError must not be null",
        assertThrows(
                IllegalArgumentException.class,
                () -> CellErrorLiteralValidation.requireStoredErrorLiteral(null, "storedError"))
            .getMessage());
    assertEquals(
        "reportedError must not be null",
        assertThrows(
                IllegalArgumentException.class,
                () -> CellErrorLiteralValidation.requireReportedErrorLiteral(null, "reportedError"))
            .getMessage());
    IllegalArgumentException computedState =
        assertThrows(
            IllegalArgumentException.class, () -> new CellValueSpec.ErrorValue("#CIRCULAR_REF!"));
    assertTrue(computedState.getMessage().contains("cannot be authored as stored cell values"));
    assertTrue(computedState.getMessage().contains("#FUNCTION_NOT_IMPLEMENTED!"));
    IllegalArgumentException invalidError =
        assertThrows(
            IllegalArgumentException.class, () -> new CellValueSpec.ErrorValue("NOT_AN_ERROR"));
    assertTrue(invalidError.getMessage().contains("#REF!"));
    assertFalse(invalidError.getMessage().contains("#FUNCTION_NOT_IMPLEMENTED!"));
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

  @Test
  void compactErrorRowsRejectEvaluationOnlyReportedStates() {
    IllegalArgumentException rowError =
        assertThrows(
            IllegalArgumentException.class,
            () -> new CellRowSpec.ErrorValues(List.of("#REF!", "#CIRCULAR_REF!")));
    IllegalArgumentException gridError =
        assertThrows(
            IllegalArgumentException.class,
            () -> new CellGridSpec.ErrorRows(List.of(List.of("#FUNCTION_NOT_IMPLEMENTED!"))));

    assertTrue(rowError.getMessage().contains("cannot be authored as stored cell values"));
    assertTrue(gridError.getMessage().contains("cannot be authored as stored cell values"));
  }
}
