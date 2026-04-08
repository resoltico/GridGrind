package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.protocol.dto.CellInput;
import dev.erst.gridgrind.protocol.dto.GridGrindRequest;
import dev.erst.gridgrind.protocol.dto.RequestWarning;
import dev.erst.gridgrind.protocol.dto.SheetCopyPosition;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests request warnings derived from same-request spaced sheet-name formula references. */
class GridGrindRequestWarningsTest {
  @Test
  void warnsWhenSetCellFormulaReferencesSameRequestSpacedSheetWithoutQuotes() {
    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.None(),
            List.of(
                new WorkbookOperation.EnsureSheet("Budget Review"),
                new WorkbookOperation.EnsureSheet("Summary"),
                new WorkbookOperation.SetCell(
                    "Summary", "A1", new CellInput.Formula("SUM(Budget Review!A1)"))),
            List.of());

    List<RequestWarning> warnings = GridGrindRequestWarnings.collect(request);

    assertEquals(1, warnings.size());
    assertEquals(2, warnings.getFirst().operationIndex());
    assertEquals("SET_CELL", warnings.getFirst().operationType());
    assertTrue(warnings.getFirst().message().contains("Budget Review"));
    assertTrue(warnings.getFirst().message().contains("'Sheet Name'!A1"));
  }

  @Test
  void doesNotWarnWhenSpacedSheetReferencesAreQuotedOrOnlyAppearInsideStrings() {
    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.None(),
            List.of(
                new WorkbookOperation.EnsureSheet("Budget Review"),
                new WorkbookOperation.EnsureSheet("Summary"),
                new WorkbookOperation.SetCell(
                    "Summary", "A1", new CellInput.Formula("SUM('Budget Review'!A1)")),
                new WorkbookOperation.SetCell(
                    "Summary", "A2", new CellInput.Formula("INDIRECT(\"Budget Review!A1\")"))),
            List.of());

    assertEquals(List.of(), GridGrindRequestWarnings.collect(request));
  }

  @Test
  void doesNotWarnWhenAnotherSheetNameOnlyContainsTheSpacedNameAsASuffix() {
    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.None(),
            List.of(
                new WorkbookOperation.EnsureSheet("Budget Review"),
                new WorkbookOperation.EnsureSheet("Summary"),
                new WorkbookOperation.SetCell(
                    "Summary", "A1", new CellInput.Formula("SUM(Annual Budget Review!A1)"))),
            List.of());

    assertEquals(List.of(), GridGrindRequestWarnings.collect(request));
  }

  @Test
  void warnsWhenRenameAndCopyOperationsIntroduceSpacedSheetNames() {
    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.None(),
            List.of(
                new WorkbookOperation.EnsureSheet("Budget"),
                new WorkbookOperation.RenameSheet("Budget", "Budget Review"),
                new WorkbookOperation.CopySheet(
                    "Budget Review", "Annual Budget Review", new SheetCopyPosition.AppendAtEnd()),
                new WorkbookOperation.EnsureSheet("Summary"),
                new WorkbookOperation.SetCell(
                    "Summary",
                    "A1",
                    new CellInput.Formula("SUM(Budget Review!A1,Annual Budget Review!$A$1)"))),
            List.of());

    assertEquals(
        List.of(
            new RequestWarning(
                4,
                "SET_CELL",
                "Formula references same-request sheet names with spaces without single quotes: Budget Review, Annual Budget Review. Use 'Sheet Name'!A1 syntax.")),
        GridGrindRequestWarnings.collect(request));
  }

  @Test
  void doesNotWarnWhenUnquotedSheetTokensDoNotStartACellReference() {
    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.None(),
            List.of(
                new WorkbookOperation.EnsureSheet("Budget Review"),
                new WorkbookOperation.EnsureSheet("Summary"),
                new WorkbookOperation.SetCell(
                    "Summary", "A1", new CellInput.Formula("Budget Review!")),
                new WorkbookOperation.SetCell(
                    "Summary", "A2", new CellInput.Formula("T(Budget Review!1)")),
                new WorkbookOperation.SetCell(
                    "Summary", "A3", new CellInput.Formula("T(\"Budget Review!A1\"\" suffix\")")),
                new WorkbookOperation.SetCell(
                    "Summary", "A4", new CellInput.Formula("\"Budget Review!A1\""))),
            List.of());

    assertEquals(List.of(), GridGrindRequestWarnings.collect(request));
  }

  @Test
  void collectsWarningsFromSetRangeAndAppendRowFormulaCells() {
    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.None(),
            List.of(
                new WorkbookOperation.EnsureSheet("Budget Review"),
                new WorkbookOperation.EnsureSheet("Summary"),
                new WorkbookOperation.SetRange(
                    "Summary",
                    "A1:B1",
                    List.of(
                        List.of(
                            new CellInput.Formula("Budget Review!A1"), new CellInput.Text("ok")))),
                new WorkbookOperation.AppendRow(
                    "Summary",
                    List.of(
                        new CellInput.Text("Total"),
                        new CellInput.Formula("SUM(Budget Review!A1)")))),
            List.of());

    List<RequestWarning> warnings = GridGrindRequestWarnings.collect(request);

    assertEquals(
        List.of(
            new RequestWarning(
                2,
                "SET_RANGE",
                "Formula references same-request sheet names with spaces without single quotes: Budget Review. Use 'Sheet Name'!A1 syntax."),
            new RequestWarning(
                3,
                "APPEND_ROW",
                "Formula references same-request sheet names with spaces without single quotes: Budget Review. Use 'Sheet Name'!A1 syntax.")),
        warnings);
  }
}
