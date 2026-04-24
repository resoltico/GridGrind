package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.CellAlignmentReport;
import dev.erst.gridgrind.contract.dto.CellBorderReport;
import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellColorReport;
import dev.erst.gridgrind.contract.dto.CellFillReport;
import dev.erst.gridgrind.contract.dto.CellFontReport;
import dev.erst.gridgrind.contract.dto.CellGradientFillReport;
import dev.erst.gridgrind.contract.dto.CellGradientStopReport;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.CommentAnchorReport;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.RichTextRunReport;
import dev.erst.gridgrind.excel.ExcelBorderSideSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFontSnapshot;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelCellStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelColorSnapshot;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelCommentSnapshot;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelGradientFillSnapshot;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelRichTextSnapshot;
import java.util.List;
import java.util.Optional;

/** Converts cell-local workbook snapshots into protocol report records. */
final class InspectionResultCellReportSupport {
  private InspectionResultCellReportSupport() {}

  static Optional<HyperlinkTarget> toHyperlinkTarget(ExcelHyperlink hyperlink) {
    if (hyperlink == null) {
      return Optional.empty();
    }
    return Optional.of(
        switch (hyperlink) {
          case ExcelHyperlink.Url url -> new HyperlinkTarget.Url(url.target());
          case ExcelHyperlink.Email email -> new HyperlinkTarget.Email(email.target());
          case ExcelHyperlink.File file -> new HyperlinkTarget.File(file.path());
          case ExcelHyperlink.Document document -> new HyperlinkTarget.Document(document.target());
        });
  }

  static Optional<GridGrindResponse.CommentReport> toCommentReport(ExcelComment comment) {
    if (comment == null) {
      return Optional.empty();
    }
    return Optional.of(
        new GridGrindResponse.CommentReport(comment.text(), comment.author(), comment.visible()));
  }

  static Optional<GridGrindResponse.CommentReport> toCommentReport(ExcelCommentSnapshot comment) {
    if (comment == null) {
      return Optional.empty();
    }
    return Optional.of(
        new GridGrindResponse.CommentReport(
            comment.text(),
            comment.author(),
            comment.visible(),
            toRichTextRunReports(comment.runs()).orElse(null),
            commentAnchorReport(comment).orElse(null)));
  }

  static FontHeightReport toFontHeightReport(ExcelFontHeight fontHeight) {
    return fontHeight == null
        ? null
        : new FontHeightReport(fontHeight.twips(), fontHeight.points());
  }

  static GridGrindResponse.CellStyleReport toCellStyleReport(ExcelCellStyleSnapshot style) {
    return new GridGrindResponse.CellStyleReport(
        style.numberFormat(),
        new CellAlignmentReport(
            style.alignment().wrapText(),
            style.alignment().horizontalAlignment(),
            style.alignment().verticalAlignment(),
            style.alignment().textRotation(),
            style.alignment().indentation()),
        toCellFontReport(style.font()),
        new CellFillReport(
            style.fill().pattern(),
            toCellColorReport(style.fill().foregroundColor()).orElse(null),
            toCellColorReport(style.fill().backgroundColor()).orElse(null),
            toCellGradientFillReport(style.fill().gradient()).orElse(null)),
        new CellBorderReport(
            toCellBorderSideReport(style.border().top()),
            toCellBorderSideReport(style.border().right()),
            toCellBorderSideReport(style.border().bottom()),
            toCellBorderSideReport(style.border().left())),
        new CellProtectionReport(style.protection().locked(), style.protection().hiddenFormula()));
  }

  static CellFontReport toCellFontReport(ExcelCellFontSnapshot font) {
    return new CellFontReport(
        font.bold(),
        font.italic(),
        font.fontName(),
        toFontHeightReport(font.fontHeight()),
        toCellColorReport(font.fontColor()).orElse(null),
        font.underline(),
        font.strikeout());
  }

  static CellBorderSideReport toCellBorderSideReport(ExcelBorderSideSnapshot side) {
    return new CellBorderSideReport(side.style(), toCellColorReport(side.color()).orElse(null));
  }

  static GridGrindResponse.CellReport toCellReport(ExcelCellSnapshot snapshot) {
    GridGrindResponse.CellStyleReport style = toCellStyleReport(snapshot.style());
    HyperlinkTarget hyperlink =
        toHyperlinkTarget(snapshot.metadata().hyperlink().orElse(null)).orElse(null);
    GridGrindResponse.CommentReport comment =
        toCommentReport(snapshot.metadata().comment().orElse(null)).orElse(null);

    return switch (snapshot) {
      case ExcelCellSnapshot.BlankSnapshot s ->
          new GridGrindResponse.CellReport.BlankReport(
              s.address(), s.declaredType(), s.displayValue(), style, hyperlink, comment);
      case ExcelCellSnapshot.TextSnapshot s ->
          new GridGrindResponse.CellReport.TextReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.stringValue(),
              toRichTextRunReports(s.richText()).orElse(null));
      case ExcelCellSnapshot.NumberSnapshot s ->
          new GridGrindResponse.CellReport.NumberReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.numberValue());
      case ExcelCellSnapshot.BooleanSnapshot s ->
          new GridGrindResponse.CellReport.BooleanReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.booleanValue());
      case ExcelCellSnapshot.ErrorSnapshot s ->
          new GridGrindResponse.CellReport.ErrorReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.errorValue());
      case ExcelCellSnapshot.FormulaSnapshot s ->
          new GridGrindResponse.CellReport.FormulaReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.formula(),
              toCellReport(s.evaluation()));
    };
  }

  static Optional<List<RichTextRunReport>> toRichTextRunReports(ExcelRichTextSnapshot richText) {
    if (richText == null) {
      return Optional.empty();
    }
    return Optional.of(
        richText.runs().stream()
            .map(run -> new RichTextRunReport(run.text(), toCellFontReport(run.font())))
            .toList());
  }

  static Optional<CellColorReport> toCellColorReport(ExcelColorSnapshot color) {
    return color == null
        ? Optional.empty()
        : Optional.of(
            new CellColorReport(color.rgb(), color.theme(), color.indexed(), color.tint()));
  }

  private static Optional<CellGradientFillReport> toCellGradientFillReport(
      ExcelGradientFillSnapshot gradient) {
    return gradient == null
        ? Optional.empty()
        : Optional.of(
            new CellGradientFillReport(
                gradient.type(),
                gradient.degree(),
                gradient.left(),
                gradient.right(),
                gradient.top(),
                gradient.bottom(),
                gradient.stops().stream()
                    .map(
                        stop ->
                            new CellGradientStopReport(
                                stop.position(), toCellColorReport(stop.color()).orElse(null)))
                    .toList()));
  }

  private static Optional<CommentAnchorReport> commentAnchorReport(ExcelCommentSnapshot comment) {
    return comment.anchor() == null
        ? Optional.empty()
        : Optional.of(
            new CommentAnchorReport(
                comment.anchor().firstColumn(),
                comment.anchor().firstRow(),
                comment.anchor().lastColumn(),
                comment.anchor().lastRow()));
  }
}
