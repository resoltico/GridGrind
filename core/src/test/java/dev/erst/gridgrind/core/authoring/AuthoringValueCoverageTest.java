package dev.erst.gridgrind.core.authoring;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
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
                () ->
                    new Tables.Definition(
                        null,
                        new SheetName("Budget"),
                        new RangeRef("A1:B2"),
                        false,
                        Tables.noStyle()))
            .getMessage());
    assertEquals(
        "sheetName must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new Tables.Definition(
                        new TableName("BudgetTable"),
                        null,
                        new RangeRef("A1:B2"),
                        false,
                        Tables.noStyle()))
            .getMessage());
    assertEquals(
        "range must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new Tables.Definition(
                        new TableName("BudgetTable"),
                        new SheetName("Budget"),
                        null,
                        false,
                        Tables.noStyle()))
            .getMessage());
    assertEquals(
        "style must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new Tables.Definition(
                        new TableName("BudgetTable"),
                        new SheetName("Budget"),
                        new RangeRef("A1:B2"),
                        false,
                        null))
            .getMessage());

    TableSpec noStyleTable =
        Tables.toTableInput(
            Tables.define(
                new TableName("BudgetTable"),
                new SheetName("Budget"),
                new RangeRef("A1:B2"),
                false,
                Tables.noStyle()));
    Tables.Style namedStyle = Tables.namedStyle("TableStyleMedium2", true, false, true, false);
    TableSpec styledTable =
        Tables.toTableInput(
            Tables.define(
                new TableName("BudgetTable"),
                new SheetName("Budget"),
                new RangeRef("A1:B2"),
                true,
                namedStyle));
    assertInstanceOf(TableStyleSpec.None.class, noStyleTable.style());
    TableStyleSpec.Named named = assertInstanceOf(TableStyleSpec.Named.class, styledTable.style());
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
        "value must not be null",
        assertThrows(NullPointerException.class, () -> ExpectedValues.text(null)).getMessage());
    assertEquals(
        "value must not be null",
        assertThrows(IllegalArgumentException.class, () -> ExpectedValues.error(null))
            .getMessage());
  }

  @Test
  void valuesConvertTypedCellInputsAcrossEverySupportedBranch() {
    Path textPath = Path.of("authored-inputs", "item.txt");
    LocalDate date = LocalDate.of(2026, 4, 23);
    LocalDateTime dateTime = LocalDateTime.of(2026, 4, 23, 10, 15);

    assertInstanceOf(CellValueSpec.Blank.class, Values.toCellInput(Values.blank()));
    CellValueSpec.Text text =
        assertInstanceOf(CellValueSpec.Text.class, Values.toCellInput(Values.text("Owner")));
    CellValueSpec.Text sourcedText =
        assertInstanceOf(
            CellValueSpec.Text.class,
            Values.toCellInput(Values.text(Values.inlineText("Owner 2"))));
    CellValueSpec.Text fileText =
        assertInstanceOf(CellValueSpec.Text.class, Values.toCellInput(Values.textFile(textPath)));
    CellValueSpec.Text stdinText =
        assertInstanceOf(
            CellValueSpec.Text.class, Values.toCellInput(Values.textFromStandardInput()));
    CellValueSpec.NumberValue number =
        assertInstanceOf(CellValueSpec.NumberValue.class, Values.toCellInput(Values.number(42.5)));
    CellValueSpec.BooleanValue bool =
        assertInstanceOf(CellValueSpec.BooleanValue.class, Values.toCellInput(Values.bool(true)));
    CellValueSpec.Date dateInput =
        assertInstanceOf(CellValueSpec.Date.class, Values.toCellInput(Values.date(date)));
    CellValueSpec.DateTime dateTimeInput =
        assertInstanceOf(
            CellValueSpec.DateTime.class, Values.toCellInput(Values.dateTime(dateTime)));
    CellValueSpec.Formula inlineFormula =
        assertInstanceOf(
            CellValueSpec.Formula.class, Values.toCellInput(Values.formula("SUM(A1:A2)")));
    CellValueSpec.Formula fileFormula =
        assertInstanceOf(
            CellValueSpec.Formula.class, Values.toCellInput(Values.formulaFile(textPath)));
    CellValueSpec.Formula stdinFormula =
        assertInstanceOf(
            CellValueSpec.Formula.class, Values.toCellInput(Values.formulaFromStandardInput()));
    CellValueSpec.ErrorValue errorValue =
        assertInstanceOf(CellValueSpec.ErrorValue.class, Values.toCellInput(Values.error("#N/A")));
    assertEquals("Owner", assertInstanceOf(TextSourceSpec.Inline.class, text.value()).text());
    assertEquals(
        "Owner 2", assertInstanceOf(TextSourceSpec.Inline.class, sourcedText.value()).text());
    assertInstanceOf(TextSourceSpec.Utf8File.class, fileText.value());
    assertInstanceOf(TextSourceSpec.Stdin.class, stdinText.value());
    assertEquals(42.5, number.value());
    assertTrue(bool.value());
    assertEquals(date, dateInput.value());
    assertEquals(dateTime, dateTimeInput.value());
    assertEquals(
        "SUM(A1:A2)", assertInstanceOf(TextSourceSpec.Inline.class, inlineFormula.value()).text());
    assertInstanceOf(TextSourceSpec.Utf8File.class, fileFormula.value());
    assertInstanceOf(TextSourceSpec.Stdin.class, stdinFormula.value());
    assertEquals("#N/A", errorValue.value());
    assertEquals(
        "value must not be null",
        assertThrows(NullPointerException.class, () -> Values.toCellInput(null)).getMessage());
  }

  @Test
  void valuesConvertCommentsExpectedValuesAndTextSources() {
    Path textPath = Path.of("authored-inputs", "item.txt");
    CommentSpec inlineComment = Values.toCommentInput(Values.comment("note", "Ada"));
    CommentSpec sourcedComment =
        Values.toCommentInput(Values.comment(Values.textSourceFile(textPath), "Ada", true));
    assertFalse(inlineComment.visible());
    assertTrue(sourcedComment.visible());
    assertInstanceOf(TextSourceSpec.Inline.class, inlineComment.text());
    assertInstanceOf(TextSourceSpec.Utf8File.class, sourcedComment.text());
    assertEquals(
        "comment must not be null",
        assertThrows(NullPointerException.class, () -> Values.toCommentInput(null)).getMessage());

    CellValue.Blank expectedBlank = assertInstanceOf(CellValue.Blank.class, ExpectedValues.blank());
    CellValue.Text expectedText =
        assertInstanceOf(CellValue.Text.class, ExpectedValues.text("Owner"));
    CellValue.NumberValue expectedNumber =
        assertInstanceOf(CellValue.NumberValue.class, ExpectedValues.number(42.5));
    CellValue.BooleanValue expectedBoolean =
        assertInstanceOf(CellValue.BooleanValue.class, ExpectedValues.bool(true));
    CellValue.ErrorValue expectedError =
        assertInstanceOf(CellValue.ErrorValue.class, ExpectedValues.error("#N/A"));
    assertInstanceOf(CellValue.Blank.class, expectedBlank);
    assertEquals("Owner", expectedText.value());
    assertEquals(42.5, expectedNumber.value());
    assertTrue(expectedBoolean.value());
    assertEquals("#N/A", expectedError.value());

    TextSourceSpec.Inline inlineSource =
        assertInstanceOf(
            TextSourceSpec.Inline.class, Values.toTextSourceInput(Values.inlineText("Owner")));
    TextSourceSpec.Utf8File fileSource =
        assertInstanceOf(
            TextSourceSpec.Utf8File.class,
            Values.toTextSourceInput(Values.textSourceFile(textPath)));
    TextSourceSpec.Stdin stdinSource =
        assertInstanceOf(
            TextSourceSpec.Stdin.class,
            Values.toTextSourceInput(Values.textSourceFromStandardInput()));
    assertEquals("Owner", inlineSource.text());
    assertEquals(textPath.toString(), fileSource.path());
    assertInstanceOf(TextSourceSpec.Stdin.class, stdinSource);
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
