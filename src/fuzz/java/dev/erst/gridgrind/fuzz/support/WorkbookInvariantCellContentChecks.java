package dev.erst.gridgrind.fuzz.support;

import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.model.value.CellTemporalReport;
import dev.erst.gridgrind.core.model.value.CellValue;

/** Owns invariant checks for cell content, comments, hyperlinks, and named ranges. */
final class WorkbookInvariantCellContentChecks {
  private WorkbookInvariantCellContentChecks() {}

  static void requireCellReportShape(dev.erst.gridgrind.core.model.value.CellReport cellReport) {
    WorkbookInvariantChecks.require(cellReport.address() != null, "cell address must not be null");
    WorkbookInvariantChecks.require(
        !cellReport.address().value().isBlank(), "cell address must not be blank");
    WorkbookInvariantChecks.require(cellReport.type() != null, "cell type must not be null");
    WorkbookInvariantChecks.require(!cellReport.type().isBlank(), "cell type must not be blank");
    WorkbookInvariantChecks.require(
        cellReport.displayValue() != null, "displayValue must not be null");
    WorkbookInvariantChecks.require(cellReport.style() != null, "style must not be null");
    cellReport.style().ifPresent(WorkbookInvariantCellStyleChecks::requireCellStyleShape);
    WorkbookInvariantChecks.require(cellReport.hyperlink() != null, "hyperlink must not be null");
    WorkbookInvariantChecks.require(cellReport.comment() != null, "comment must not be null");

    WorkbookInvariantChecks.require(cellReport.value() != null, "value must not be null");
    cellReport.value().ifPresent(WorkbookInvariantCellContentChecks::requireCellValueShape);
    if (!(cellReport instanceof dev.erst.gridgrind.core.model.value.CellReport.FormulaReport)) {
      cellReport
          .value()
          .ifPresent(
              value ->
                  WorkbookInvariantChecks.require(
                      cellReport.type().equals(value.type()),
                      "stored and scalar value types must agree"));
    }
    switch (cellReport) {
      case dev.erst.gridgrind.core.model.value.CellReport.TextReport text ->
          requireTextRuns(text.runs(), text.value().map(CellValue.Text::value));
      case dev.erst.gridgrind.core.model.value.CellReport.NumberReport number -> {
        WorkbookInvariantChecks.require(number.temporal() != null, "temporal must not be null");
        number.temporal().ifPresent(WorkbookInvariantCellContentChecks::requireTemporalShape);
      }
      case dev.erst.gridgrind.core.model.value.CellReport.FormulaReport formula -> {
        WorkbookInvariantChecks.require(formula.formula() != null, "formula must not be null");
        requireTextRuns(
            formula.runs(),
            formula
                .value()
                .filter(CellValue.Text.class::isInstance)
                .map(CellValue.Text.class::cast)
                .map(CellValue.Text::value));
        WorkbookInvariantChecks.require(formula.temporal() != null, "temporal must not be null");
        if (formula.temporal().isPresent()) {
          WorkbookInvariantChecks.require(
              formula.value().orElse(null) instanceof CellValue.NumberValue,
              "temporal requires a numeric value");
          requireTemporalShape(formula.temporal().orElseThrow());
        }
      }
      default -> {}
    }
    if (cellReport.hyperlink().isPresent()) {
      requireHyperlinkShape(cellReport.hyperlink().orElseThrow());
    }
    if (cellReport.comment().isPresent()) {
      requireCommentReportShape(cellReport.comment().orElseThrow());
    }
  }

  private static void requireCellValueShape(CellValue value) {
    WorkbookInvariantChecks.require(value != null, "value must not be null");
    WorkbookInvariantChecks.require(value.type() != null, "value type must not be null");
    WorkbookInvariantChecks.require(!value.type().isBlank(), "value type must not be blank");
    switch (value) {
      case CellValue.Text text ->
          WorkbookInvariantChecks.require(text.value() != null, "text value must not be null");
      case CellValue.ErrorValue error -> {
        WorkbookInvariantChecks.require(error.value() != null, "error value must not be null");
        WorkbookInvariantChecks.require(
            dev.erst.gridgrind.core.vocabulary.ExcelReportedCellErrorLiteral.orderedWireValues()
                .contains(error.value()),
            "error value must be a reported Excel literal");
      }
      case CellValue.Blank _, CellValue.NumberValue _, CellValue.BooleanValue _ -> {}
    }
  }

  private static void requireTextRuns(
      java.util.Optional<java.util.List<RichTextRunReport>> runs, java.util.Optional<String> text) {
    WorkbookInvariantChecks.require(runs != null, "runs must not be null");
    if (runs.isEmpty()) {
      return;
    }
    WorkbookInvariantChecks.require(text.isPresent(), "runs require a text value");
    WorkbookInvariantChecks.require(!runs.orElseThrow().isEmpty(), "runs must not be empty");
    StringBuilder builder = new StringBuilder();
    for (RichTextRunReport run : runs.orElseThrow()) {
      WorkbookInvariantChecks.require(run != null, "runs must not contain null values");
      WorkbookInvariantChecks.require(run.text() != null, "run text must not be null");
      WorkbookInvariantChecks.require(!run.text().isEmpty(), "run text must not be empty");
      WorkbookInvariantCellStyleChecks.requireCellFontShape(run.font());
      builder.append(run.text());
    }
    WorkbookInvariantChecks.require(
        text.orElseThrow().equals(builder.toString()), "runs must concatenate to value");
  }

  private static void requireTemporalShape(CellTemporalReport temporal) {
    WorkbookInvariantChecks.require(temporal != null, "temporal must not be null");
    WorkbookInvariantChecks.require(temporal.kind() != null, "temporal kind must not be null");
    WorkbookInvariantChecks.require(
        temporal.isoValue() != null, "temporal isoValue must not be null");
    if (temporal.isDate()) {
      WorkbookInvariantChecks.require(
          temporal.kind().isPresent(), "temporal kind must be present when isDate is true");
      WorkbookInvariantChecks.require(
          temporal.isoValue().isPresent(), "temporal isoValue must be present when isDate is true");
      WorkbookInvariantChecks.require(
          !temporal.isoValue().orElseThrow().isBlank(), "temporal isoValue must not be blank");
    } else {
      WorkbookInvariantChecks.require(
          temporal.kind().isEmpty(), "temporal kind must be absent when isDate is false");
      WorkbookInvariantChecks.require(
          temporal.isoValue().isEmpty(), "temporal isoValue must be absent when isDate is false");
    }
  }

  static void requireCommentReportShape(CommentReport comment) {
    WorkbookInvariantChecks.require(comment.text() != null, "comment text must not be null");
    WorkbookInvariantChecks.require(comment.author() != null, "comment author must not be null");
    WorkbookInvariantChecks.require(!comment.text().isBlank(), "comment text must not be blank");
    WorkbookInvariantChecks.require(
        !comment.author().isBlank(), "comment author must not be blank");
    if (comment.runs().isPresent()) {
      WorkbookInvariantChecks.require(
          !comment.runs().orElseThrow().isEmpty(), "comment runs must not be empty");
      StringBuilder builder = new StringBuilder();
      for (RichTextRunReport run : comment.runs().orElseThrow()) {
        WorkbookInvariantChecks.require(run != null, "comment runs must not contain null values");
        WorkbookInvariantChecks.require(run.text() != null, "comment run text must not be null");
        WorkbookInvariantChecks.require(
            !run.text().isEmpty(), "comment run text must not be empty");
        WorkbookInvariantCellStyleChecks.requireCellFontShape(run.font());
        builder.append(run.text());
      }
      WorkbookInvariantChecks.require(
          builder.toString().equals(comment.text()), "comment runs must concatenate to text");
    }
    if (comment.anchor().isPresent()) {
      WorkbookInvariantCellMetadataChecks.requireCommentAnchorShape(comment.anchor().orElseThrow());
    }
  }

  static void requireNamedRangeShape(NamedRangeReport namedRange) {
    WorkbookInvariantChecks.require(namedRange.name() != null, "namedRange name must not be null");
    WorkbookInvariantChecks.require(
        !namedRange.name().value().isBlank(), "namedRange name must not be blank");
    WorkbookInvariantChecks.require(
        namedRange.scope() != null, "namedRange scope must not be null");
    WorkbookInvariantChecks.require(
        namedRange.refersToFormula() != null, "namedRange formula must not be null");

    switch (namedRange) {
      case NamedRangeReport.RangeReport range -> {
        WorkbookInvariantChecks.require(
            range.target() != null, "namedRange target must not be null");
        WorkbookInvariantChecks.require(
            range.target() instanceof NamedRangeTarget.Range,
            "namedRange range report must use a range target");
        NamedRangeTarget.Range target = (NamedRangeTarget.Range) range.target();
        WorkbookInvariantChecks.require(
            target.sheetName() != null, "namedRange target sheet must not be null");
        WorkbookInvariantChecks.require(
            target.range() != null, "namedRange target range must not be null");
        WorkbookInvariantChecks.require(
            !target.sheetName().value().isBlank(), "namedRange target sheet must not be blank");
        WorkbookInvariantChecks.require(
            !target.range().value().isBlank(), "namedRange target range must not be blank");
      }
      case NamedRangeReport.FormulaReport _ -> {}
    }
  }

  static void requireHyperlinkShape(dev.erst.gridgrind.core.model.value.HyperlinkTarget hyperlink) {
    WorkbookInvariantChecks.require(hyperlink != null, "hyperlink must not be null");
    switch (hyperlink) {
      case dev.erst.gridgrind.core.model.value.HyperlinkTarget.Url url -> {
        WorkbookInvariantChecks.require(url.target() != null, "hyperlink target must not be null");
        WorkbookInvariantChecks.require(
            !url.target().isBlank(), "hyperlink target must not be blank");
        WorkbookInvariantChecks.require(
            !url.target().regionMatches(true, 0, "file:", 0, 5),
            "URL hyperlink targets must not use file: schemes");
        WorkbookInvariantChecks.require(
            !url.target().regionMatches(true, 0, "mailto:", 0, 7),
            "URL hyperlink targets must not use mailto: schemes");
      }
      case dev.erst.gridgrind.core.model.value.HyperlinkTarget.Email email -> {
        WorkbookInvariantChecks.require(email.email() != null, "hyperlink email must not be null");
        WorkbookInvariantChecks.require(
            !email.email().isBlank(), "hyperlink email must not be blank");
        WorkbookInvariantChecks.require(
            !email.email().regionMatches(true, 0, "mailto:", 0, 7),
            "EMAIL hyperlink targets must omit the mailto: prefix");
      }
      case dev.erst.gridgrind.core.model.value.HyperlinkTarget.File file -> {
        WorkbookInvariantChecks.require(file.path() != null, "hyperlink path must not be null");
        WorkbookInvariantChecks.require(!file.path().isBlank(), "hyperlink path must not be blank");
        WorkbookInvariantChecks.require(
            !file.path().regionMatches(true, 0, "file:", 0, 5),
            "FILE hyperlink targets must be normalized path strings");
      }
      case dev.erst.gridgrind.core.model.value.HyperlinkTarget.Document document -> {
        WorkbookInvariantChecks.require(
            document.target() != null, "hyperlink target must not be null");
        WorkbookInvariantChecks.require(
            !document.target().isBlank(), "hyperlink target must not be blank");
      }
    }
  }
}
