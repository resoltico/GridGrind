package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.List;

/** Typed authored value helpers for cells, comments, and source-backed payloads. */
public final class Values {
  private Values() {}

  /** Returns a blank cell payload. */
  public static CellInput blank() {
    return new CellInput.Blank();
  }

  /** Returns a UTF-8 text cell payload embedded inline. */
  public static CellInput text(String text) {
    return new CellInput.Text(TextSourceInput.inline(text));
  }

  /** Returns a UTF-8 text cell payload from an explicit source model. */
  public static CellInput text(TextSourceInput source) {
    return new CellInput.Text(source);
  }

  /** Returns a UTF-8 text cell payload backed by a file path. */
  public static CellInput textFile(Path path) {
    return text(textSourceFile(path));
  }

  /** Returns a UTF-8 text cell payload sourced from standard input. */
  public static CellInput textFromStandardInput() {
    return text(TextSourceInput.standardInput());
  }

  /** Returns a rich-text cell payload. */
  public static CellInput richText(List<RichTextRunInput> runs) {
    return new CellInput.RichText(runs);
  }

  /** Returns a rich-text cell payload from authored runs. */
  public static CellInput richText(RichTextRunInput... runs) {
    return richText(List.of(runs));
  }

  /** Returns a numeric cell payload. */
  public static CellInput number(double number) {
    return new CellInput.Numeric(number);
  }

  /** Returns a boolean cell payload. */
  public static CellInput bool(boolean value) {
    return new CellInput.BooleanValue(value);
  }

  /** Returns a date cell payload. */
  public static CellInput date(LocalDate value) {
    return new CellInput.Date(value);
  }

  /** Returns a date-time cell payload. */
  public static CellInput dateTime(LocalDateTime value) {
    return new CellInput.DateTime(value);
  }

  /** Returns an inline formula cell payload, stripping a leading '=' when present. */
  public static CellInput formula(String formula) {
    return new CellInput.Formula(TextSourceInput.inline(formula));
  }

  /** Returns a formula cell payload backed by a UTF-8 file. */
  public static CellInput formulaFile(Path path) {
    return new CellInput.Formula(textSourceFile(path));
  }

  /** Returns a formula cell payload sourced from standard input. */
  public static CellInput formulaFromStandardInput() {
    return new CellInput.Formula(TextSourceInput.standardInput());
  }

  /** Returns a plain-text comment payload with the given author. */
  public static CommentInput comment(String text, String author) {
    return new CommentInput(TextSourceInput.inline(text), author, false);
  }

  /** Returns a plain-text comment payload from an explicit source model with the given author. */
  public static CommentInput comment(TextSourceInput textSource, String author, boolean visible) {
    return new CommentInput(textSource, author, visible);
  }

  /** Returns an inline UTF-8 text source model. */
  public static TextSourceInput inlineText(String text) {
    return TextSourceInput.inline(text);
  }

  /** Returns a file-backed UTF-8 text source model. */
  public static TextSourceInput textSourceFile(Path path) {
    return TextSourceInput.utf8File(path.toString());
  }

  /** Returns a standard-input UTF-8 text source model. */
  public static TextSourceInput textSourceFromStandardInput() {
    return TextSourceInput.standardInput();
  }

  /** Returns an inline base64-authored binary source model. */
  public static BinarySourceInput inlineBase64(String base64Data) {
    return BinarySourceInput.inlineBase64(base64Data);
  }

  /** Returns a file-backed binary source model. */
  public static BinarySourceInput binaryFile(Path path) {
    return BinarySourceInput.file(path.toString());
  }

  /** Returns a standard-input binary source model. */
  public static BinarySourceInput binaryFromStandardInput() {
    return BinarySourceInput.standardInput();
  }

  /** Returns one expected blank effective cell value. */
  public static ExpectedCellValue expectedBlank() {
    return new ExpectedCellValue.Blank();
  }

  /** Returns one expected text effective cell value. */
  public static ExpectedCellValue expectedText(String text) {
    return new ExpectedCellValue.Text(text);
  }

  /** Returns one expected numeric effective cell value. */
  public static ExpectedCellValue expectedNumber(double number) {
    return new ExpectedCellValue.NumericValue(number);
  }

  /** Returns one expected boolean effective cell value. */
  public static ExpectedCellValue expectedBoolean(boolean value) {
    return new ExpectedCellValue.BooleanValue(value);
  }

  /** Returns one expected error effective cell value. */
  public static ExpectedCellValue expectedError(String error) {
    return new ExpectedCellValue.ErrorValue(error);
  }

  /** Returns one authored row payload for range and append-row helpers. */
  public static List<CellInput> row(CellInput... cells) {
    return List.copyOf(Arrays.asList(cells));
  }
}
