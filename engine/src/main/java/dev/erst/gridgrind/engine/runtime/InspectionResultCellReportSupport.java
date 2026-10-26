package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellTemporalKind;
import dev.erst.gridgrind.core.model.value.CellTemporalReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.excel.ExcelCellReadFacet;
import dev.erst.gridgrind.excel.ExcelCellReadProjection;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
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

  static Optional<CommentReport> toCommentReport(@Nullable ResolvedComment comment) {
    if (comment == null) {
      return Optional.empty();
    }
    return Optional.of(new CommentReport(comment.text(), comment.author(), comment.visible()));
  }

  static CellReport toCellReport(
      ExcelCellSnapshot snapshot, ExcelCellReadProjection projection, boolean date1904) {
    Optional<String> displayValue = projectedDisplayValue(snapshot, projection);
    Optional<CellStyleReport> style = projectedStyle(snapshot, projection);
    Optional<HyperlinkTarget> hyperlink = projectedHyperlink(snapshot, projection);
    Optional<CommentReport> comment = projectedComment(snapshot, projection);

    return switch (snapshot) {
      case ExcelCellSnapshot.BlankSnapshot cell ->
          new CellReport.BlankReport(
              new CellAddress(cell.address()),
              displayValue,
              style,
              hyperlink,
              comment,
              projectedValue(projection, new CellValue.Blank()));
      case ExcelCellSnapshot.TextSnapshot cell ->
          new CellReport.TextReport(
              new CellAddress(cell.address()),
              displayValue,
              style,
              hyperlink,
              comment,
              projectedValue(projection, new CellValue.Text(cell.textValue())),
              projectedRuns(projection, cell.richText()));
      case ExcelCellSnapshot.NumberSnapshot cell ->
          new CellReport.NumberReport(
              new CellAddress(cell.address()),
              displayValue,
              style,
              hyperlink,
              comment,
              projectedValue(projection, new CellValue.NumberValue(cell.numberValue())),
              projectedTemporal(cell, projection, date1904));
      case ExcelCellSnapshot.BooleanSnapshot cell ->
          new CellReport.BooleanReport(
              new CellAddress(cell.address()),
              displayValue,
              style,
              hyperlink,
              comment,
              projectedValue(projection, new CellValue.BooleanValue(cell.booleanValue())));
      case ExcelCellSnapshot.ErrorSnapshot cell ->
          new CellReport.ErrorReport(
              new CellAddress(cell.address()),
              displayValue,
              style,
              hyperlink,
              comment,
              projectedValue(projection, new CellValue.ErrorValue(cell.errorValue())));
      case ExcelCellSnapshot.FormulaSnapshot cell -> {
        Optional<ExcelCellSnapshot> evaluation =
            projection.includes(ExcelCellReadFacet.VALUE)
                ? Optional.of(cell.evaluation().orElseThrow())
                : Optional.empty();
        yield new CellReport.FormulaReport(
            new CellAddress(cell.address()),
            displayValue,
            style,
            hyperlink,
            comment,
            projectedFormula(projection, cell.formula()),
            evaluation.map(InspectionResultCellReportSupport::toCellValue),
            evaluation
                .filter(ExcelCellSnapshot.TextSnapshot.class::isInstance)
                .map(ExcelCellSnapshot.TextSnapshot.class::cast)
                .flatMap(text -> projectedRuns(projection, text.richText())),
            evaluation
                .filter(ExcelCellSnapshot.NumberSnapshot.class::isInstance)
                .map(ExcelCellSnapshot.NumberSnapshot.class::cast)
                .flatMap(number -> projectedTemporal(number, projection, date1904)));
      }
    };
  }

  static CellValue toCellValue(ExcelCellSnapshot snapshot) {
    return switch (snapshot) {
      case ExcelCellSnapshot.BlankSnapshot _ -> new CellValue.Blank();
      case ExcelCellSnapshot.TextSnapshot cell -> new CellValue.Text(cell.textValue());
      case ExcelCellSnapshot.NumberSnapshot cell -> new CellValue.NumberValue(cell.numberValue());
      case ExcelCellSnapshot.BooleanSnapshot cell ->
          new CellValue.BooleanValue(cell.booleanValue());
      case ExcelCellSnapshot.ErrorSnapshot cell -> new CellValue.ErrorValue(cell.errorValue());
      case ExcelCellSnapshot.FormulaSnapshot _ ->
          throw new IllegalArgumentException(
              "Formula evaluations must not recursively remain FORMULA");
    };
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
        ? Optional.of(snapshot.style())
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
        ? snapshot.metadata().comment()
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

  private static Optional<List<RichTextRunReport>> projectedRuns(
      ExcelCellReadProjection projection, @Nullable List<RichTextRunReport> richText) {
    return projection.includes(ExcelCellReadFacet.RICH_TEXT_RUNS)
        ? Optional.ofNullable(richText)
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
}
