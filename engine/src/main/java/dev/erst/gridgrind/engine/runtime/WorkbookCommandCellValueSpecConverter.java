package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.formula.ArrayFormulaSpec;
import dev.erst.gridgrind.core.model.formula.ResolvedArrayFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedRawFormula;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.style.FontHeightSpec;
import dev.erst.gridgrind.core.model.style.ResolvedRichText;
import dev.erst.gridgrind.core.model.style.ResolvedRichTextRun;
import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import java.util.Optional;

/** Converts authored cell values, fonts and annotations for workbook operations. */
final class WorkbookCommandCellValueSpecConverter {
  private WorkbookCommandCellValueSpecConverter() {}

  static ResolvedCellWrite toResolvedCellWrite(CellValueSpec value) {
    return switch (value) {
      case CellValueSpec.Blank _ -> new CellValue.Blank();
      case CellValueSpec.Text text ->
          new CellValue.Text(WorkbookCommandSourceSupport.inlineText(text.value(), "cell text"));
      case CellValueSpec.RichText richText -> toResolvedRichText(richText);
      case CellValueSpec.NumberValue numberValue -> new CellValue.NumberValue(numberValue.value());
      case CellValueSpec.BooleanValue booleanValue ->
          new CellValue.BooleanValue(booleanValue.value());
      case CellValueSpec.ErrorValue errorValue -> new CellValue.ErrorValue(errorValue.value());
      case CellValueSpec.Date date -> date;
      case CellValueSpec.DateTime dateTime -> dateTime;
      case CellValueSpec.Formula formula ->
          new ResolvedFormula(WorkbookCommandSourceSupport.inlineText(formula.value(), "formula"));
      case CellValueSpec.RawFormula rawFormula ->
          new ResolvedRawFormula(
              WorkbookCommandSourceSupport.inlineText(rawFormula.value(), "raw formula"));
    };
  }

  static ResolvedArrayFormula toResolvedArrayFormula(ArrayFormulaSpec input) {
    return new ResolvedArrayFormula(
        WorkbookCommandSourceSupport.inlineText(input.source(), "array formula"));
  }

  static ResolvedRichText toResolvedRichText(CellValueSpec.RichText richText) {
    return new ResolvedRichText(
        richText.runs().stream()
            .map(WorkbookCommandCellValueSpecConverter::toResolvedRichTextRun)
            .toList());
  }

  static ResolvedRichTextRun toResolvedRichTextRun(RichTextRunSpec run) {
    return new ResolvedRichTextRun(
        WorkbookCommandSourceSupport.inlineText(run.text(), "rich-text run"), run.font());
  }

  static ExcelHyperlink toExcelHyperlink(HyperlinkTarget target) {
    return switch (target) {
      case HyperlinkTarget.Url url -> new ExcelHyperlink.Url(url.target());
      case HyperlinkTarget.Email email -> new ExcelHyperlink.Email(email.email());
      case HyperlinkTarget.File file -> new ExcelHyperlink.File(file.path());
      case HyperlinkTarget.Document document -> new ExcelHyperlink.Document(document.target());
    };
  }

  static ResolvedComment toResolvedComment(CommentSpec comment) {
    return new ResolvedComment(
        WorkbookCommandSourceSupport.inlineText(comment.text(), "comment text"),
        comment.author(),
        comment.visible(),
        commentRuns(comment),
        comment.anchor());
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

  private static Optional<ResolvedRichText> commentRuns(CommentSpec comment) {
    return comment.runs().isEmpty()
        ? Optional.empty()
        : Optional.of(
            new ResolvedRichText(
                comment.runs().orElseThrow().stream()
                    .map(WorkbookCommandCellValueSpecConverter::toResolvedRichTextRun)
                    .toList()));
  }
}
