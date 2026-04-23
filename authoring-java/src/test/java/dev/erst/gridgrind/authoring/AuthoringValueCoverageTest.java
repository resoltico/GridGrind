package dev.erst.gridgrind.authoring;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Focused coverage for authored value/helper wrappers and their contract conversions. */
class AuthoringValueCoverageTest {
  @Test
  void linksCoverFactoriesConversionsAndNullValidation() {
    assertEquals(
        "target must not be null",
        assertThrows(NullPointerException.class, () -> new Links.Url(null)).getMessage());
    assertEquals(
        "email must not be null",
        assertThrows(NullPointerException.class, () -> new Links.Email(null)).getMessage());
    assertEquals(
        "path must not be null",
        assertThrows(NullPointerException.class, () -> new Links.FileTarget(null)).getMessage());
    assertEquals(
        "target must not be null",
        assertThrows(NullPointerException.class, () -> new Links.DocumentTarget(null))
            .getMessage());

    HyperlinkTarget.Url url =
        assertInstanceOf(
            HyperlinkTarget.Url.class, Links.toHyperlinkTarget(Links.url("https://example.com")));
    HyperlinkTarget.Email email =
        assertInstanceOf(
            HyperlinkTarget.Email.class, Links.toHyperlinkTarget(Links.email("ada@example.com")));
    HyperlinkTarget.File file =
        assertInstanceOf(
            HyperlinkTarget.File.class, Links.toHyperlinkTarget(Links.file("budget.xlsx")));
    HyperlinkTarget.Document document =
        assertInstanceOf(
            HyperlinkTarget.Document.class,
            Links.toHyperlinkTarget(Links.document("Dashboard!A1")));
    assertEquals("https://example.com", url.target());
    assertEquals("ada@example.com", email.email());
    assertEquals("budget.xlsx", file.path());
    assertEquals("Dashboard!A1", document.target());
    assertEquals(
        "target must not be null",
        assertThrows(NullPointerException.class, () -> Links.toHyperlinkTarget(null)).getMessage());
  }

  @Test
  void tablesCoverFactoriesConversionsAndNullValidation() {
    assertEquals(
        "name must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new Tables.NamedStyle(null, true, false, true, false))
            .getMessage());
    assertEquals(
        "name must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new Tables.Definition(null, "Budget", "A1:B2", false, Tables.noStyle()))
            .getMessage());
    assertEquals(
        "sheetName must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new Tables.Definition("BudgetTable", null, "A1:B2", false, Tables.noStyle()))
            .getMessage());
    assertEquals(
        "range must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new Tables.Definition("BudgetTable", "Budget", null, false, Tables.noStyle()))
            .getMessage());
    assertEquals(
        "style must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new Tables.Definition("BudgetTable", "Budget", "A1:B2", false, null))
            .getMessage());

    TableInput noStyleTable =
        Tables.toTableInput(
            Tables.define("BudgetTable", "Budget", "A1:B2", false, Tables.noStyle()));
    Tables.Style namedStyle = Tables.namedStyle("TableStyleMedium2", true, false, true, false);
    TableInput styledTable =
        Tables.toTableInput(Tables.define("BudgetTable", "Budget", "A1:B2", true, namedStyle));
    assertInstanceOf(TableStyleInput.None.class, noStyleTable.style());
    TableStyleInput.Named named =
        assertInstanceOf(TableStyleInput.Named.class, styledTable.style());
    assertEquals("TableStyleMedium2", named.name());
    assertTrue(named.showFirstColumn());
    assertFalse(named.showLastColumn());
    assertTrue(named.showRowStripes());
    assertFalse(named.showColumnStripes());
    assertEquals(
        "definition must not be null",
        assertThrows(NullPointerException.class, () -> Tables.toTableInput(null)).getMessage());
  }

  @Test
  void valuesRejectNullsForFocusedWrappers() {
    assertEquals(
        "text must not be null",
        assertThrows(NullPointerException.class, () -> new Values.InlineText(null)).getMessage());
    assertEquals(
        "path must not be null",
        assertThrows(NullPointerException.class, () -> new Values.Utf8FileText(null)).getMessage());
    assertEquals(
        "source must not be null",
        assertThrows(NullPointerException.class, () -> new Values.Text(null)).getMessage());
    assertEquals(
        "value must not be null",
        assertThrows(NullPointerException.class, () -> new Values.DateValue(null)).getMessage());
    assertEquals(
        "value must not be null",
        assertThrows(NullPointerException.class, () -> new Values.DateTimeValue(null))
            .getMessage());
    assertEquals(
        "source must not be null",
        assertThrows(NullPointerException.class, () -> new Values.Formula(null)).getMessage());
    assertEquals(
        "source must not be null",
        assertThrows(NullPointerException.class, () -> new Values.Comment(null, "Ada", false))
            .getMessage());
    assertEquals(
        "author must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new Values.Comment(Values.inlineText("note"), null, false))
            .getMessage());
    assertEquals(
        "text must not be null",
        assertThrows(NullPointerException.class, () -> new Values.ExpectedText(null)).getMessage());
    assertEquals(
        "error must not be null",
        assertThrows(NullPointerException.class, () -> new Values.ExpectedError(null))
            .getMessage());
  }

  @Test
  void valuesConvertTypedCellInputsAcrossEverySupportedBranch() {
    Path textPath = Path.of("authored-inputs", "item.txt");
    LocalDate date = LocalDate.of(2026, 4, 23);
    LocalDateTime dateTime = LocalDateTime.of(2026, 4, 23, 10, 15);

    assertInstanceOf(CellInput.Blank.class, Values.toCellInput(Values.blank()));
    CellInput.Text text =
        assertInstanceOf(CellInput.Text.class, Values.toCellInput(Values.text("Owner")));
    CellInput.Text sourcedText =
        assertInstanceOf(
            CellInput.Text.class, Values.toCellInput(Values.text(Values.inlineText("Owner 2"))));
    CellInput.Text fileText =
        assertInstanceOf(CellInput.Text.class, Values.toCellInput(Values.textFile(textPath)));
    CellInput.Text stdinText =
        assertInstanceOf(CellInput.Text.class, Values.toCellInput(Values.textFromStandardInput()));
    CellInput.Numeric number =
        assertInstanceOf(CellInput.Numeric.class, Values.toCellInput(Values.number(42.5)));
    CellInput.BooleanValue bool =
        assertInstanceOf(CellInput.BooleanValue.class, Values.toCellInput(Values.bool(true)));
    CellInput.Date dateInput =
        assertInstanceOf(CellInput.Date.class, Values.toCellInput(Values.date(date)));
    CellInput.DateTime dateTimeInput =
        assertInstanceOf(CellInput.DateTime.class, Values.toCellInput(Values.dateTime(dateTime)));
    CellInput.Formula inlineFormula =
        assertInstanceOf(CellInput.Formula.class, Values.toCellInput(Values.formula("SUM(A1:A2)")));
    CellInput.Formula fileFormula =
        assertInstanceOf(CellInput.Formula.class, Values.toCellInput(Values.formulaFile(textPath)));
    CellInput.Formula stdinFormula =
        assertInstanceOf(
            CellInput.Formula.class, Values.toCellInput(Values.formulaFromStandardInput()));
    assertInstanceOf(TextSourceInput.Inline.class, text.source());
    assertInstanceOf(TextSourceInput.Inline.class, sourcedText.source());
    assertInstanceOf(TextSourceInput.Utf8File.class, fileText.source());
    assertInstanceOf(TextSourceInput.StandardInput.class, stdinText.source());
    assertEquals(42.5, number.number());
    assertTrue(bool.bool());
    assertEquals(date, dateInput.date());
    assertEquals(dateTime, dateTimeInput.dateTime());
    assertInstanceOf(TextSourceInput.Inline.class, inlineFormula.source());
    assertInstanceOf(TextSourceInput.Utf8File.class, fileFormula.source());
    assertInstanceOf(TextSourceInput.StandardInput.class, stdinFormula.source());
    assertEquals(
        "value must not be null",
        assertThrows(NullPointerException.class, () -> Values.toCellInput(null)).getMessage());
  }

  @Test
  void valuesConvertCommentsExpectedValuesAndTextSources() {
    Path textPath = Path.of("authored-inputs", "item.txt");
    CommentInput inlineComment = Values.toCommentInput(Values.comment("note", "Ada"));
    CommentInput sourcedComment =
        Values.toCommentInput(Values.comment(Values.textSourceFile(textPath), "Ada", true));
    assertFalse(inlineComment.visible());
    assertTrue(sourcedComment.visible());
    assertInstanceOf(TextSourceInput.Inline.class, inlineComment.text());
    assertInstanceOf(TextSourceInput.Utf8File.class, sourcedComment.text());
    assertEquals(
        "comment must not be null",
        assertThrows(NullPointerException.class, () -> Values.toCommentInput(null)).getMessage());

    ExpectedCellValue.Blank expectedBlank =
        assertInstanceOf(
            ExpectedCellValue.Blank.class, Values.toExpectedCellValue(Values.expectedBlank()));
    ExpectedCellValue.Text expectedText =
        assertInstanceOf(
            ExpectedCellValue.Text.class, Values.toExpectedCellValue(Values.expectedText("Owner")));
    ExpectedCellValue.NumericValue expectedNumber =
        assertInstanceOf(
            ExpectedCellValue.NumericValue.class,
            Values.toExpectedCellValue(Values.expectedNumber(42.5)));
    ExpectedCellValue.BooleanValue expectedBoolean =
        assertInstanceOf(
            ExpectedCellValue.BooleanValue.class,
            Values.toExpectedCellValue(Values.expectedBoolean(true)));
    ExpectedCellValue.ErrorValue expectedError =
        assertInstanceOf(
            ExpectedCellValue.ErrorValue.class,
            Values.toExpectedCellValue(Values.expectedError("#N/A")));
    assertInstanceOf(ExpectedCellValue.Blank.class, expectedBlank);
    assertEquals("Owner", expectedText.text());
    assertEquals(42.5, expectedNumber.number());
    assertTrue(expectedBoolean.value());
    assertEquals("#N/A", expectedError.error());
    assertEquals(
        "expectedValue must not be null",
        assertThrows(NullPointerException.class, () -> Values.toExpectedCellValue(null))
            .getMessage());

    TextSourceInput.Inline inlineSource =
        assertInstanceOf(
            TextSourceInput.Inline.class, Values.toTextSourceInput(Values.inlineText("Owner")));
    TextSourceInput.Utf8File fileSource =
        assertInstanceOf(
            TextSourceInput.Utf8File.class,
            Values.toTextSourceInput(Values.textSourceFile(textPath)));
    TextSourceInput.StandardInput stdinSource =
        assertInstanceOf(
            TextSourceInput.StandardInput.class,
            Values.toTextSourceInput(Values.textSourceFromStandardInput()));
    assertEquals("Owner", inlineSource.text());
    assertEquals(textPath.toString(), fileSource.path());
    assertInstanceOf(TextSourceInput.StandardInput.class, stdinSource);
    assertEquals(
        "source must not be null",
        assertThrows(NullPointerException.class, () -> Values.toTextSourceInput(null))
            .getMessage());

    List<Values.CellValue> row =
        Values.row(Values.blank(), Values.text("Owner"), Values.number(42.5));
    assertEquals(3, row.size());
    assertInstanceOf(Values.Blank.class, row.getFirst());
  }
}
