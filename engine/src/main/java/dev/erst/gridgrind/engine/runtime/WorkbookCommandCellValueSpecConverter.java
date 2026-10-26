package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.formula.ArrayFormulaSpec;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.style.FontHeightSpec;
import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.excel.ExcelArrayFormulaDefinition;
import dev.erst.gridgrind.excel.ExcelCellFont;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelCommentAnchor;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelRichText;
import dev.erst.gridgrind.excel.ExcelRichTextRun;
import java.util.Optional;

/** Converts authored cell values, fonts and annotations for workbook operations. */
final class WorkbookCommandCellValueSpecConverter {
  private WorkbookCommandCellValueSpecConverter() {}

  static ExcelCellValue toExcelCellValue(CellValueSpec value) {
    return switch (value) {
      case CellValueSpec.Blank _ -> ExcelCellValue.blank();
      case CellValueSpec.Text text ->
          ExcelCellValue.text(WorkbookCommandSourceSupport.inlineText(text.value(), "cell text"));
      case CellValueSpec.RichText richText -> ExcelCellValue.richText(toExcelRichText(richText));
      case CellValueSpec.NumberValue numberValue -> ExcelCellValue.number(numberValue.value());
      case CellValueSpec.BooleanValue booleanValue -> ExcelCellValue.bool(booleanValue.value());
      case CellValueSpec.ErrorValue errorValue -> ExcelCellValue.error(errorValue.value());
      case CellValueSpec.Date date -> ExcelCellValue.date(date.value());
      case CellValueSpec.DateTime dateTime -> ExcelCellValue.dateTime(dateTime.value());
      case CellValueSpec.Formula formula ->
          ExcelCellValue.formula(
              WorkbookCommandSourceSupport.inlineText(formula.value(), "formula"));
      case CellValueSpec.RawFormula rawFormula ->
          ExcelCellValue.rawFormula(
              WorkbookCommandSourceSupport.inlineText(rawFormula.value(), "raw formula"));
    };
  }

  static ExcelArrayFormulaDefinition toExcelArrayFormulaDefinition(ArrayFormulaSpec input) {
    return new ExcelArrayFormulaDefinition(
        WorkbookCommandSourceSupport.inlineText(input.source(), "array formula"));
  }

  static ExcelRichText toExcelRichText(CellValueSpec.RichText richText) {
    return new ExcelRichText(
        richText.runs().stream()
            .map(WorkbookCommandCellValueSpecConverter::toExcelRichTextRun)
            .toList());
  }

  static ExcelRichTextRun toExcelRichTextRun(RichTextRunSpec run) {
    return new ExcelRichTextRun(
        WorkbookCommandSourceSupport.inlineText(run.text(), "rich-text run"),
        run.font().flatMap(WorkbookCommandCellValueSpecConverter::toExcelCellFont));
  }

  static ExcelHyperlink toExcelHyperlink(HyperlinkTarget target) {
    return switch (target) {
      case HyperlinkTarget.Url url -> new ExcelHyperlink.Url(url.target());
      case HyperlinkTarget.Email email -> new ExcelHyperlink.Email(email.email());
      case HyperlinkTarget.File file -> new ExcelHyperlink.File(file.path());
      case HyperlinkTarget.Document document -> new ExcelHyperlink.Document(document.target());
    };
  }

  static ExcelComment toExcelComment(CommentSpec comment) {
    return new ExcelComment(
        WorkbookCommandSourceSupport.inlineText(comment.text(), "comment text"),
        comment.author(),
        comment.visible(),
        commentRuns(comment),
        commentAnchor(comment));
  }

  static ExcelCellStyle toExcelCellStyle(CellStylePatch style) {
    return new ExcelCellStyle(
        style.numberFormat(),
        style.alignment(),
        style.font().flatMap(WorkbookCommandCellValueSpecConverter::toExcelCellFont),
        style.fill(),
        style.border(),
        style.protection());
  }

  static Optional<ExcelCellFont> toExcelCellFont(CellFontPatch font) {
    if (font == null) {
      return Optional.empty();
    }
    return Optional.of(
        new ExcelCellFont(
            font.bold(),
            font.italic(),
            font.fontName(),
            font.fontHeight().flatMap(WorkbookCommandCellValueSpecConverter::toFontHeightReport),
            font.fontColor(),
            font.underline(),
            font.strikeout()));
  }

  static Optional<FontHeightReport> toFontHeightReport(FontHeightSpec fontHeight) {
    if (fontHeight == null) {
      return Optional.empty();
    }
    return Optional.of(
        switch (fontHeight) {
          case FontHeightSpec.Points points -> FontHeightReport.fromPoints(points.points());
          case FontHeightSpec.Twips twips -> FontHeightReport.fromTwips(twips.twips());
        });
  }

  private static Optional<ExcelRichText> commentRuns(CommentSpec comment) {
    return comment.runs().isEmpty()
        ? Optional.empty()
        : Optional.of(
            new ExcelRichText(
                comment.runs().orElseThrow().stream()
                    .map(WorkbookCommandCellValueSpecConverter::toExcelRichTextRun)
                    .toList()));
  }

  private static Optional<ExcelCommentAnchor> commentAnchor(CommentSpec comment) {
    return comment.anchor().isEmpty()
        ? Optional.empty()
        : Optional.of(
            new ExcelCommentAnchor(
                comment.anchor().orElseThrow().firstColumn(),
                comment.anchor().orElseThrow().firstRow(),
                comment.anchor().orElseThrow().lastColumn(),
                comment.anchor().orElseThrow().lastRow()));
  }
}
