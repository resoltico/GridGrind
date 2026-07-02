package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.CellColorReport;
import dev.erst.gridgrind.contract.dto.CellReport;
import dev.erst.gridgrind.contract.dto.CellStyleReport;
import dev.erst.gridgrind.contract.dto.CellTemporalKind;
import dev.erst.gridgrind.contract.dto.CellTemporalReport;
import dev.erst.gridgrind.contract.dto.CellValueReport;
import dev.erst.gridgrind.contract.dto.CommentAnchorReport;
import dev.erst.gridgrind.contract.dto.CommentReport;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.RichTextRunReport;
import dev.erst.gridgrind.excel.ExcelCellReadFacet;
import dev.erst.gridgrind.excel.ExcelCellReadProjection;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelColorSnapshot;
import dev.erst.gridgrind.excel.ExcelCommentSnapshot;
import dev.erst.gridgrind.excel.ExcelRichTextSnapshot;
import dev.erst.gridgrind.excel.ExcelTemporalFormatSupport;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Optional;
import org.apache.poi.ss.usermodel.DateUtil;
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

  static CellReport toCellReport(
      ExcelCellSnapshot snapshot, ExcelCellReadProjection projection, boolean date1904) {
    Optional<String> displayValue = projectedDisplayValue(snapshot, projection);
    Optional<CellStyleReport> style = projectedStyle(snapshot, projection);
    Optional<HyperlinkTarget> hyperlink = projectedHyperlink(snapshot, projection);
    Optional<CommentReport> comment = projectedComment(snapshot, projection);

    return switch (snapshot) {
      case ExcelCellSnapshot.BlankSnapshot s ->
          new CellReport.BlankReport(s.address(), displayValue, style, hyperlink, comment);
      case ExcelCellSnapshot.TextSnapshot s ->
          new CellReport.TextReport(
              s.address(),
              displayValue,
              style,
              hyperlink,
              comment,
              projectedValue(projection, s.textValue()),
              projectedRuns(projection, s.richText()));
      case ExcelCellSnapshot.NumberSnapshot s ->
          new CellReport.NumberReport(
              s.address(),
              displayValue,
              style,
              hyperlink,
              comment,
              projectedValue(projection, s.numberValue()),
              projectedTemporal(s, projection, date1904));
      case ExcelCellSnapshot.BooleanSnapshot s ->
          new CellReport.BooleanReport(
              s.address(),
              displayValue,
              style,
              hyperlink,
              comment,
              projectedValue(projection, s.booleanValue()));
      case ExcelCellSnapshot.ErrorSnapshot s ->
          new CellReport.ErrorReport(
              s.address(),
              displayValue,
              style,
              hyperlink,
              comment,
              projectedValue(projection, s.errorValue()));
      case ExcelCellSnapshot.FormulaSnapshot s ->
          new CellReport.FormulaReport(
              s.address(),
              displayValue,
              style,
              hyperlink,
              comment,
              projectedFormula(projection, s.formula()),
              projectedEvaluation(s.evaluation(), projection, date1904));
    };
  }

  static CellValueReport toCellValueReport(
      ExcelCellSnapshot snapshot, ExcelCellReadProjection projection, boolean date1904) {
    return switch (snapshot) {
      case ExcelCellSnapshot.BlankSnapshot _ -> new CellValueReport.BlankValue();
      case ExcelCellSnapshot.TextSnapshot s ->
          new CellValueReport.TextValue(
              s.textValue(),
              projection.includes(ExcelCellReadFacet.RICH_TEXT_RUNS)
                  ? toRichTextRunReports(s.richText())
                  : Optional.empty());
      case ExcelCellSnapshot.NumberSnapshot s ->
          new CellValueReport.NumberValue(
              s.numberValue(), projectedTemporal(s, projection, date1904));
      case ExcelCellSnapshot.BooleanSnapshot s ->
          new CellValueReport.BooleanValue(s.booleanValue());
      case ExcelCellSnapshot.ErrorSnapshot s -> new CellValueReport.ErrorValue(s.errorValue());
      case ExcelCellSnapshot.FormulaSnapshot _ ->
          throw new IllegalArgumentException(
              "Formula evaluations must not recursively remain FORMULA");
    };
  }

  static Optional<List<RichTextRunReport>> toRichTextRunReports(
      Optional<ExcelRichTextSnapshot> richText) {
    if (richText.isEmpty()) {
      return Optional.empty();
    }
    return Optional.of(
        richText.orElseThrow().runs().stream()
            .map(
                run ->
                    new RichTextRunReport(
                        run.text(),
                        InspectionResultCellStyleReportSupport.toCellFontReport(run.font())))
            .toList());
  }

  static Optional<List<RichTextRunReport>> toRichTextRunReports(
      @Nullable ExcelRichTextSnapshot richText) {
    return toRichTextRunReports(Optional.ofNullable(richText));
  }

  static Optional<CellColorReport> toCellColorReport(@Nullable ExcelColorSnapshot color) {
    return InspectionResultCellStyleReportSupport.toCellColorReport(color);
  }

  private static Optional<String> projectedDisplayValue(
      ExcelCellSnapshot snapshot, ExcelCellReadProjection projection) {
    return projection.includes(ExcelCellReadFacet.FORMAT)
        ? Optional.of(snapshot.displayValue())
        : Optional.empty();
  }

  private static Optional<CellStyleReport> projectedStyle(
      ExcelCellSnapshot snapshot, ExcelCellReadProjection projection) {
    return projection.includes(ExcelCellReadFacet.STYLE)
        ? Optional.of(InspectionResultCellStyleReportSupport.toCellStyleReport(snapshot.style()))
        : Optional.empty();
  }

  private static Optional<HyperlinkTarget> projectedHyperlink(
      ExcelCellSnapshot snapshot, ExcelCellReadProjection projection) {
    return projection.includes(ExcelCellReadFacet.HYPERLINK)
        ? toHyperlinkTarget(snapshot.metadata().hyperlink().orElse(null))
        : Optional.empty();
  }

  private static Optional<CommentReport> projectedComment(
      ExcelCellSnapshot snapshot, ExcelCellReadProjection projection) {
    return projection.includes(ExcelCellReadFacet.COMMENT)
        ? toCommentReport(snapshot.metadata().comment().orElse(null))
        : Optional.empty();
  }

  private static <T> Optional<T> projectedValue(ExcelCellReadProjection projection, T value) {
    return projection.includes(ExcelCellReadFacet.VALUE) ? Optional.of(value) : Optional.empty();
  }

  private static Optional<String> projectedFormula(
      ExcelCellReadProjection projection, String formula) {
    return projection.includes(ExcelCellReadFacet.FORMULA)
        ? Optional.of(formula)
        : Optional.empty();
  }

  private static Optional<CellValueReport> projectedEvaluation(
      ExcelCellSnapshot evaluation, ExcelCellReadProjection projection, boolean date1904) {
    return projection.includes(ExcelCellReadFacet.VALUE)
        ? Optional.of(toCellValueReport(evaluation, projection, date1904))
        : Optional.empty();
  }

  private static Optional<List<RichTextRunReport>> projectedRuns(
      ExcelCellReadProjection projection, @Nullable ExcelRichTextSnapshot richText) {
    return projection.includes(ExcelCellReadFacet.RICH_TEXT_RUNS)
        ? toRichTextRunReports(richText)
        : Optional.empty();
  }

  private static Optional<CellTemporalReport> projectedTemporal(
      ExcelCellSnapshot.NumberSnapshot snapshot,
      ExcelCellReadProjection projection,
      boolean date1904) {
    if (!projection.includes(ExcelCellReadFacet.TEMPORAL)) {
      return Optional.empty();
    }
    return Optional.of(
        temporalReport(snapshot.numberValue(), snapshot.style().numberFormat(), date1904));
  }

  private static CellTemporalReport temporalReport(
      double numberValue, String numberFormat, boolean date1904) {
    Optional<ExcelTemporalFormatSupport.ObservedKind> observedKind =
        ExcelTemporalFormatSupport.observedKind(numberFormat);
    if (observedKind.isEmpty()) {
      return CellTemporalReport.notDate();
    }
    LocalDateTime localDateTime = DateUtil.getLocalDateTime(numberValue, date1904);
    return switch (observedKind.orElseThrow()) {
      case DATE ->
          CellTemporalReport.temporal(
              CellTemporalKind.DATE, localDateTime.toLocalDate().toString());
      case TIME ->
          CellTemporalReport.temporal(
              CellTemporalKind.TIME, localDateTime.toLocalTime().toString());
      case DATE_TIME ->
          CellTemporalReport.temporal(CellTemporalKind.DATE_TIME, localDateTime.toString());
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
