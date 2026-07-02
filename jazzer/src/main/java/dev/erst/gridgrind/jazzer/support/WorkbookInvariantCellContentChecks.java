package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.CellTemporalReport;
import dev.erst.gridgrind.contract.dto.CellValueReport;
import dev.erst.gridgrind.contract.dto.CommentReport;
import dev.erst.gridgrind.contract.dto.NamedRangeReport;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.RichTextRunReport;

/** Owns invariant checks for cell content, comments, hyperlinks, and named ranges. */
final class WorkbookInvariantCellContentChecks {
  private WorkbookInvariantCellContentChecks() {}

  static void requireCellReportShape(dev.erst.gridgrind.contract.dto.CellReport cellReport) {
    WorkbookInvariantChecks.require(cellReport.address() != null, "cell address must not be null");
    WorkbookInvariantChecks.require(
        !cellReport.address().isBlank(), "cell address must not be blank");
    WorkbookInvariantChecks.require(cellReport.type() != null, "cell type must not be null");
    WorkbookInvariantChecks.require(!cellReport.type().isBlank(), "cell type must not be blank");
    WorkbookInvariantChecks.require(
        cellReport.displayValue() != null, "displayValue must not be null");
    WorkbookInvariantChecks.require(cellReport.style() != null, "style must not be null");
    cellReport.style().ifPresent(WorkbookInvariantCellStyleChecks::requireCellStyleShape);
    WorkbookInvariantChecks.require(cellReport.hyperlink() != null, "hyperlink must not be null");
    WorkbookInvariantChecks.require(cellReport.comment() != null, "comment must not be null");

    switch (cellReport) {
      case dev.erst.gridgrind.contract.dto.CellReport.BlankReport _ -> {}
      case dev.erst.gridgrind.contract.dto.CellReport.TextReport text -> {
        WorkbookInvariantChecks.require(text.textValue() != null, "textValue must not be null");
        WorkbookInvariantChecks.require(text.runs() != null, "runs must not be null");
        if (text.runs().isPresent()) {
          WorkbookInvariantChecks.require(
              text.textValue().isPresent(), "runs require textValue to be present");
          WorkbookInvariantChecks.require(
              !text.runs().orElseThrow().isEmpty(), "runs must not be empty");
          StringBuilder builder = new StringBuilder();
          for (var run : text.runs().orElseThrow()) {
            WorkbookInvariantChecks.require(
                run.text() != null, "richText run text must not be null");
            WorkbookInvariantChecks.require(
                !run.text().isEmpty(), "richText run text must not be empty");
            WorkbookInvariantCellStyleChecks.requireCellFontShape(run.font());
            builder.append(run.text());
          }
          WorkbookInvariantChecks.require(
              text.textValue().orElseThrow().equals(builder.toString()),
              "richText run text must concatenate to textValue");
        }
      }
      case dev.erst.gridgrind.contract.dto.CellReport.NumberReport number -> {
        WorkbookInvariantChecks.require(
            number.numberValue() != null, "numberValue must not be null");
        WorkbookInvariantChecks.require(number.temporal() != null, "temporal must not be null");
        number.temporal().ifPresent(WorkbookInvariantCellContentChecks::requireTemporalShape);
      }
      case dev.erst.gridgrind.contract.dto.CellReport.BooleanReport bool ->
          WorkbookInvariantChecks.require(
              bool.booleanValue() != null, "booleanValue must not be null");
      case dev.erst.gridgrind.contract.dto.CellReport.ErrorReport error ->
          WorkbookInvariantChecks.require(
              error.errorValue() != null, "errorValue must not be null");
      case dev.erst.gridgrind.contract.dto.CellReport.FormulaReport formula -> {
        WorkbookInvariantChecks.require(formula.formula() != null, "formula must not be null");
        WorkbookInvariantChecks.require(
            formula.evaluation() != null, "formula evaluation must not be null");
        formula
            .evaluation()
            .ifPresent(WorkbookInvariantCellContentChecks::requireCellValueReportShape);
      }
    }
    if (cellReport.hyperlink().isPresent()) {
      requireHyperlinkShape(cellReport.hyperlink().orElseThrow());
    }
    if (cellReport.comment().isPresent()) {
      requireCommentReportShape(cellReport.comment().orElseThrow());
    }
  }

  private static void requireCellValueReportShape(CellValueReport valueReport) {
    WorkbookInvariantChecks.require(valueReport != null, "valueReport must not be null");
    WorkbookInvariantChecks.require(valueReport.type() != null, "value type must not be null");
    WorkbookInvariantChecks.require(!valueReport.type().isBlank(), "value type must not be blank");
    switch (valueReport) {
      case CellValueReport.BlankValue _ -> {}
      case CellValueReport.TextValue text -> {
        WorkbookInvariantChecks.require(text.textValue() != null, "textValue must not be null");
        WorkbookInvariantChecks.require(text.runs() != null, "runs must not be null");
        if (text.runs().isPresent()) {
          WorkbookInvariantChecks.require(
              !text.runs().orElseThrow().isEmpty(), "runs must not be empty");
          StringBuilder builder = new StringBuilder();
          for (RichTextRunReport run : text.runs().orElseThrow()) {
            WorkbookInvariantChecks.require(run != null, "runs must not contain null values");
            WorkbookInvariantChecks.require(run.text() != null, "run text must not be null");
            WorkbookInvariantChecks.require(!run.text().isEmpty(), "run text must not be empty");
            WorkbookInvariantCellStyleChecks.requireCellFontShape(run.font());
            builder.append(run.text());
          }
          WorkbookInvariantChecks.require(
              text.textValue().equals(builder.toString()), "runs must concatenate to textValue");
        }
      }
      case CellValueReport.NumberValue number -> {
        WorkbookInvariantChecks.require(
            number.numberValue() != null, "numberValue must not be null");
        WorkbookInvariantChecks.require(number.temporal() != null, "temporal must not be null");
        number.temporal().ifPresent(WorkbookInvariantCellContentChecks::requireTemporalShape);
      }
      case CellValueReport.BooleanValue bool ->
          WorkbookInvariantChecks.require(
              bool.booleanValue() != null, "booleanValue must not be null");
      case CellValueReport.ErrorValue error -> {
        WorkbookInvariantChecks.require(error.errorValue() != null, "errorValue must not be null");
        WorkbookInvariantChecks.require(
            !error.errorValue().isBlank(), "errorValue must not be blank");
      }
    }
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
        !namedRange.name().isBlank(), "namedRange name must not be blank");
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
            !target.sheetName().isBlank(), "namedRange target sheet must not be blank");
        WorkbookInvariantChecks.require(
            !target.range().isBlank(), "namedRange target range must not be blank");
      }
      case NamedRangeReport.FormulaReport _ -> {}
    }
  }

  static void requireHyperlinkShape(dev.erst.gridgrind.contract.dto.HyperlinkTarget hyperlink) {
    WorkbookInvariantChecks.require(hyperlink != null, "hyperlink must not be null");
    switch (hyperlink) {
      case dev.erst.gridgrind.contract.dto.HyperlinkTarget.Url url -> {
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
      case dev.erst.gridgrind.contract.dto.HyperlinkTarget.Email email -> {
        WorkbookInvariantChecks.require(email.email() != null, "hyperlink email must not be null");
        WorkbookInvariantChecks.require(
            !email.email().isBlank(), "hyperlink email must not be blank");
        WorkbookInvariantChecks.require(
            !email.email().regionMatches(true, 0, "mailto:", 0, 7),
            "EMAIL hyperlink targets must omit the mailto: prefix");
      }
      case dev.erst.gridgrind.contract.dto.HyperlinkTarget.File file -> {
        WorkbookInvariantChecks.require(file.path() != null, "hyperlink path must not be null");
        WorkbookInvariantChecks.require(!file.path().isBlank(), "hyperlink path must not be blank");
        WorkbookInvariantChecks.require(
            !file.path().regionMatches(true, 0, "file:", 0, 5),
            "FILE hyperlink targets must be normalized path strings");
      }
      case dev.erst.gridgrind.contract.dto.HyperlinkTarget.Document document -> {
        WorkbookInvariantChecks.require(
            document.target() != null, "hyperlink target must not be null");
        WorkbookInvariantChecks.require(
            !document.target().isBlank(), "hyperlink target must not be blank");
      }
    }
  }
}
