package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

/** Typed authored value helpers for common cell content, comments, and expected values. */
public final class Values {
  /** Authored UTF-8 text sources for cell text, formulas, and comments. */
  public sealed interface TextSource permits InlineText, Utf8FileText, StandardInputText {}

  /** Authored cell values for the focused fluent Java surface. */
  public sealed interface CellValue
      permits Blank, Text, NumericValue, BooleanValue, DateValue, DateTimeValue, Formula {}

  /** Authored expected effective cell values for assertion helpers. */
  public sealed interface ExpectedValue
      permits ExpectedBlank, ExpectedText, ExpectedNumber, ExpectedBoolean, ExpectedError {}

  /** Inline UTF-8 text source. */
  public record InlineText(String text) implements TextSource {
    public InlineText {
      Objects.requireNonNull(text, "text must not be null");
    }
  }

  /** File-backed UTF-8 text source. */
  public record Utf8FileText(Path path) implements TextSource {
    public Utf8FileText {
      Objects.requireNonNull(path, "path must not be null");
    }
  }

  /** Standard-input UTF-8 text source. */
  public record StandardInputText() implements TextSource {}

  /** Blank cell payload. */
  public record Blank() implements CellValue {}

  /** Text cell payload. */
  public record Text(TextSource source) implements CellValue {
    public Text {
      Objects.requireNonNull(source, "source must not be null");
    }
  }

  /** Numeric cell payload. */
  public record NumericValue(double value) implements CellValue {}

  /** Boolean cell payload. */
  public record BooleanValue(boolean value) implements CellValue {}

  /** Date cell payload. */
  public record DateValue(LocalDate value) implements CellValue {
    public DateValue {
      Objects.requireNonNull(value, "value must not be null");
    }
  }

  /** Date-time cell payload. */
  public record DateTimeValue(LocalDateTime value) implements CellValue {
    public DateTimeValue {
      Objects.requireNonNull(value, "value must not be null");
    }
  }

  /** Formula cell payload. */
  public record Formula(TextSource source) implements CellValue {
    public Formula {
      Objects.requireNonNull(source, "source must not be null");
    }
  }

  /** Plain-text comment payload. */
  public record Comment(TextSource source, String author, boolean visible) {
    public Comment {
      Objects.requireNonNull(source, "source must not be null");
      Objects.requireNonNull(author, "author must not be null");
    }
  }

  /** Expected blank effective cell value. */
  public record ExpectedBlank() implements ExpectedValue {}

  /** Expected text effective cell value. */
  public record ExpectedText(String text) implements ExpectedValue {
    public ExpectedText {
      Objects.requireNonNull(text, "text must not be null");
    }
  }

  /** Expected numeric effective cell value. */
  public record ExpectedNumber(double value) implements ExpectedValue {}

  /** Expected boolean effective cell value. */
  public record ExpectedBoolean(boolean value) implements ExpectedValue {}

  /** Expected Excel error effective cell value. */
  public record ExpectedError(String error) implements ExpectedValue {
    public ExpectedError {
      Objects.requireNonNull(error, "error must not be null");
    }
  }

  private Values() {}

  /** Returns a blank cell payload. */
  public static CellValue blank() {
    return new Blank();
  }

  /** Returns an inline UTF-8 text cell payload. */
  public static CellValue text(String text) {
    return new Text(inlineText(text));
  }

  /** Returns a UTF-8 text cell payload from an authored text source. */
  public static CellValue text(TextSource source) {
    return new Text(source);
  }

  /** Returns a UTF-8 text cell payload backed by a file path. */
  public static CellValue textFile(Path path) {
    return text(textSourceFile(path));
  }

  /** Returns a UTF-8 text cell payload sourced from standard input. */
  public static CellValue textFromStandardInput() {
    return text(textSourceFromStandardInput());
  }

  /** Returns a numeric cell payload. */
  public static CellValue number(double number) {
    return new NumericValue(number);
  }

  /** Returns a boolean cell payload. */
  public static CellValue bool(boolean value) {
    return new BooleanValue(value);
  }

  /** Returns a date cell payload. */
  public static CellValue date(LocalDate value) {
    return new DateValue(value);
  }

  /** Returns a date-time cell payload. */
  public static CellValue dateTime(LocalDateTime value) {
    return new DateTimeValue(value);
  }

  /** Returns an inline formula cell payload. */
  public static CellValue formula(String formula) {
    return new Formula(inlineText(formula));
  }

  /** Returns a formula cell payload backed by a UTF-8 file. */
  public static CellValue formulaFile(Path path) {
    return new Formula(textSourceFile(path));
  }

  /** Returns a formula cell payload sourced from standard input. */
  public static CellValue formulaFromStandardInput() {
    return new Formula(textSourceFromStandardInput());
  }

  /** Returns a plain-text comment payload with the given author. */
  public static Comment comment(String text, String author) {
    return new Comment(inlineText(text), author, false);
  }

  /** Returns a plain-text comment payload from an authored text source. */
  public static Comment comment(TextSource textSource, String author, boolean visible) {
    return new Comment(textSource, author, visible);
  }

  /** Returns an inline UTF-8 text source. */
  public static TextSource inlineText(String text) {
    return new InlineText(text);
  }

  /** Returns a file-backed UTF-8 text source. */
  public static TextSource textSourceFile(Path path) {
    return new Utf8FileText(path);
  }

  /** Returns a standard-input UTF-8 text source. */
  public static TextSource textSourceFromStandardInput() {
    return new StandardInputText();
  }

  /** Returns one expected blank effective cell value. */
  public static ExpectedValue expectedBlank() {
    return new ExpectedBlank();
  }

  /** Returns one expected text effective cell value. */
  public static ExpectedValue expectedText(String text) {
    return new ExpectedText(text);
  }

  /** Returns one expected numeric effective cell value. */
  public static ExpectedValue expectedNumber(double number) {
    return new ExpectedNumber(number);
  }

  /** Returns one expected boolean effective cell value. */
  public static ExpectedValue expectedBoolean(boolean value) {
    return new ExpectedBoolean(value);
  }

  /** Returns one expected error effective cell value. */
  public static ExpectedValue expectedError(String error) {
    return new ExpectedError(error);
  }

  /** Returns one authored row payload for range helpers. */
  public static List<CellValue> row(CellValue... cells) {
    return List.copyOf(Arrays.asList(cells));
  }

  static CellInput toCellInput(CellValue value) {
    return switch (Objects.requireNonNull(value, "value must not be null")) {
      case Blank _ -> new CellInput.Blank();
      case Text text -> new CellInput.Text(toTextSourceInput(text.source()));
      case NumericValue number -> new CellInput.Numeric(number.value());
      case BooleanValue booleanValue -> new CellInput.BooleanValue(booleanValue.value());
      case DateValue date -> new CellInput.Date(date.value());
      case DateTimeValue dateTime -> new CellInput.DateTime(dateTime.value());
      case Formula formula -> new CellInput.Formula(toTextSourceInput(formula.source()));
    };
  }

  static CommentInput toCommentInput(Comment comment) {
    Comment nonNullComment = Objects.requireNonNull(comment, "comment must not be null");
    return CommentInput.plain(
        toTextSourceInput(nonNullComment.source()),
        nonNullComment.author(),
        nonNullComment.visible());
  }

  static ExpectedCellValue toExpectedCellValue(ExpectedValue expectedValue) {
    return switch (Objects.requireNonNull(expectedValue, "expectedValue must not be null")) {
      case ExpectedBlank _ -> new ExpectedCellValue.Blank();
      case ExpectedText text -> new ExpectedCellValue.Text(text.text());
      case ExpectedNumber number -> new ExpectedCellValue.NumericValue(number.value());
      case ExpectedBoolean booleanValue -> new ExpectedCellValue.BooleanValue(booleanValue.value());
      case ExpectedError error -> new ExpectedCellValue.ErrorValue(error.error());
    };
  }

  static TextSourceInput toTextSourceInput(TextSource source) {
    return switch (Objects.requireNonNull(source, "source must not be null")) {
      case InlineText inline -> TextSourceInput.inline(inline.text());
      case Utf8FileText file -> TextSourceInput.utf8File(file.path().toString());
      case StandardInputText _ -> TextSourceInput.standardInput();
    };
  }
}
