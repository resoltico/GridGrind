package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.CellAlignmentReport;
import dev.erst.gridgrind.contract.dto.CellBorderReport;
import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellColorReport;
import dev.erst.gridgrind.contract.dto.CellFillReport;
import dev.erst.gridgrind.contract.dto.CellFontReport;
import dev.erst.gridgrind.contract.dto.CellGradientFillReport;
import dev.erst.gridgrind.contract.dto.CellGradientStopReport;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.CellStyleReport;
import dev.erst.gridgrind.contract.dto.CommentAnchorReport;
import dev.erst.gridgrind.contract.dto.CommentReport;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.RichTextRunReport;
import dev.erst.gridgrind.excel.ExcelBorderSideSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFillSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFontSnapshot;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelCellStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelColorSnapshot;
import dev.erst.gridgrind.excel.ExcelCommentSnapshot;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelGradientFillSnapshot;
import dev.erst.gridgrind.excel.ExcelRichTextSnapshot;
import java.util.List;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Converts cell-local workbook snapshots into protocol report records. */
final class InspectionResultCellReportSupport {
  private InspectionResultCellReportSupport() {}

  static Optional<HyperlinkTarget> toHyperlinkTarget(
      dev.erst.gridgrind.excel.@Nullable ExcelHyperlink hyperlink) {
    if (hyperlink == null) {
      return Optional.empty();
    }
    return Optional.of(
        switch (hyperlink) {
          case dev.erst.gridgrind.excel.ExcelHyperlink.Url url ->
              new HyperlinkTarget.Url(url.target());
          case dev.erst.gridgrind.excel.ExcelHyperlink.Email email ->
              new HyperlinkTarget.Email(email.target());
          case dev.erst.gridgrind.excel.ExcelHyperlink.File file ->
              new HyperlinkTarget.File(file.path());
          case dev.erst.gridgrind.excel.ExcelHyperlink.Document document ->
              new HyperlinkTarget.Document(document.target());
        });
  }

  static Optional<CommentReport> toCommentReport(
      dev.erst.gridgrind.excel.@Nullable ExcelComment comment) {
    if (comment == null) {
      return Optional.empty();
    }
    return Optional.of(new CommentReport(comment.text(), comment.author(), comment.visible()));
  }

  static Optional<CommentReport> toCommentReport(@Nullable ExcelCommentSnapshot comment) {
    if (comment == null) {
      return Optional.empty();
    }
    return Optional.of(
        new CommentReport(
            comment.text(),
            comment.author(),
            comment.visible(),
            toRichTextRunReports(comment.runs()),
            commentAnchorReport(comment)));
  }

  static FontHeightReport toFontHeightReport(ExcelFontHeight fontHeight) {
    return new FontHeightReport(fontHeight.twips(), fontHeight.points());
  }

  static CellStyleReport toCellStyleReport(ExcelCellStyleSnapshot style) {
    return new CellStyleReport(
        style.numberFormat(),
        new CellAlignmentReport(
            style.alignment().wrapText(),
            style.alignment().horizontalAlignment(),
            style.alignment().verticalAlignment(),
            style.alignment().textRotation(),
            style.alignment().indentation()),
        toCellFontReport(style.font()),
        toCellFillReport(style.fill()),
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

  static dev.erst.gridgrind.contract.dto.CellReport toCellReport(ExcelCellSnapshot snapshot) {
    CellStyleReport style = toCellStyleReport(snapshot.style());
    Optional<HyperlinkTarget> hyperlink =
        toHyperlinkTarget(snapshot.metadata().hyperlink().orElse(null));
    Optional<CommentReport> comment = toCommentReport(snapshot.metadata().comment().orElse(null));

    return switch (snapshot) {
      case ExcelCellSnapshot.BlankSnapshot s ->
          new dev.erst.gridgrind.contract.dto.CellReport.BlankReport(
              s.address(), s.declaredType(), s.displayValue(), style, hyperlink, comment);
      case ExcelCellSnapshot.TextSnapshot s ->
          new dev.erst.gridgrind.contract.dto.CellReport.TextReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.stringValue(),
              toRichTextRunReports(s.richText()));
      case ExcelCellSnapshot.NumberSnapshot s ->
          new dev.erst.gridgrind.contract.dto.CellReport.NumberReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.numberValue());
      case ExcelCellSnapshot.BooleanSnapshot s ->
          new dev.erst.gridgrind.contract.dto.CellReport.BooleanReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.booleanValue());
      case ExcelCellSnapshot.ErrorSnapshot s ->
          new dev.erst.gridgrind.contract.dto.CellReport.ErrorReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.errorValue());
      case ExcelCellSnapshot.FormulaSnapshot s ->
          new dev.erst.gridgrind.contract.dto.CellReport.FormulaReport(
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

  static Optional<List<RichTextRunReport>> toRichTextRunReports(
      Optional<ExcelRichTextSnapshot> richText) {
    if (richText.isEmpty()) {
      return Optional.empty();
    }
    return Optional.of(
        richText.orElseThrow().runs().stream()
            .map(run -> new RichTextRunReport(run.text(), toCellFontReport(run.font())))
            .toList());
  }

  static Optional<List<RichTextRunReport>> toRichTextRunReports(
      @Nullable ExcelRichTextSnapshot richText) {
    return toRichTextRunReports(Optional.ofNullable(richText));
  }

  static Optional<CellColorReport> toCellColorReport(@Nullable ExcelColorSnapshot color) {
    return color == null
        ? Optional.empty()
        : Optional.of(
            switch (color) {
              case ExcelColorSnapshot.Rgb rgb ->
                  rgb.tint().isPresent()
                      ? CellColorReport.rgb(rgb.rgb(), rgb.tint().orElseThrow())
                      : CellColorReport.rgb(rgb.rgb());
              case ExcelColorSnapshot.Theme theme ->
                  theme.tint().isPresent()
                      ? CellColorReport.theme(theme.theme(), theme.tint().orElseThrow())
                      : CellColorReport.theme(theme.theme());
              case ExcelColorSnapshot.Indexed indexed ->
                  indexed.tint().isPresent()
                      ? CellColorReport.indexed(indexed.indexed(), indexed.tint().orElseThrow())
                      : CellColorReport.indexed(indexed.indexed());
            });
  }

  private static CellFillReport toCellFillReport(ExcelCellFillSnapshot fill) {
    return switch (fill) {
      case ExcelCellFillSnapshot.PatternOnly pattern -> CellFillReport.pattern(pattern.pattern());
      case ExcelCellFillSnapshot.PatternForeground pattern ->
          CellFillReport.patternForeground(
              pattern.pattern(), toCellColorReport(pattern.foregroundColor()).orElseThrow());
      case ExcelCellFillSnapshot.PatternBackground pattern ->
          CellFillReport.patternBackground(
              pattern.pattern(), toCellColorReport(pattern.backgroundColor()).orElseThrow());
      case ExcelCellFillSnapshot.PatternForegroundBackground pattern ->
          CellFillReport.patternColors(
              pattern.pattern(),
              toCellColorReport(pattern.foregroundColor()).orElseThrow(),
              toCellColorReport(pattern.backgroundColor()).orElseThrow());
      case ExcelCellFillSnapshot.Gradient gradient ->
          CellFillReport.gradient(toCellGradientFillReport(gradient.gradient()));
    };
  }

  private static CellGradientFillReport toCellGradientFillReport(
      ExcelGradientFillSnapshot gradient) {
    return switch (gradient) {
      case ExcelGradientFillSnapshot.Linear linear ->
          CellGradientFillReport.linear(
              linear.degree(),
              linear.stops().stream()
                  .map(
                      stop ->
                          new CellGradientStopReport(
                              stop.position(), toCellColorReport(stop.color()).orElseThrow()))
                  .toList());
      case ExcelGradientFillSnapshot.Path path ->
          CellGradientFillReport.path(
              path.left(),
              path.right(),
              path.top(),
              path.bottom(),
              path.stops().stream()
                  .map(
                      stop ->
                          new CellGradientStopReport(
                              stop.position(), toCellColorReport(stop.color()).orElseThrow()))
                  .toList());
    };
  }

  private static Optional<CommentAnchorReport> commentAnchorReport(ExcelCommentSnapshot comment) {
    return comment
        .anchor()
        .map(
            anchor ->
                new CommentAnchorReport(
                    anchor.firstColumn(),
                    anchor.firstRow(),
                    anchor.lastColumn(),
                    anchor.lastRow()));
  }
}
