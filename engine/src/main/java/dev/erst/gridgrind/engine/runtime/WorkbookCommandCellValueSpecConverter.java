package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.formula.ArrayFormulaSpec;
import dev.erst.gridgrind.core.model.protection.CellProtectionPatch;
import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.CellAlignmentPatch;
import dev.erst.gridgrind.core.model.style.CellBorderPatch;
import dev.erst.gridgrind.core.model.style.CellFillSpec;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellGradientFillSpec;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.style.FontHeightSpec;
import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.excel.ExcelArrayFormulaDefinition;
import dev.erst.gridgrind.excel.ExcelBorder;
import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelCellAlignment;
import dev.erst.gridgrind.excel.ExcelCellFill;
import dev.erst.gridgrind.excel.ExcelCellFont;
import dev.erst.gridgrind.excel.ExcelCellProtection;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelColor;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelCommentAnchor;
import dev.erst.gridgrind.excel.ExcelGradientFill;
import dev.erst.gridgrind.excel.ExcelGradientStop;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelRichText;
import dev.erst.gridgrind.excel.ExcelRichTextRun;
import java.util.Objects;
import java.util.Optional;

/**
 * Converts cell-local contract inputs into workbook-core value, style, and annotation records.
 *
 * <p>This helper intentionally spans the cell input surface, so PMD's import-count heuristic is not
 * a useful coupling signal here.
 */
@SuppressWarnings("PMD.ExcessiveImports")
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
        style.alignment().flatMap(WorkbookCommandCellValueSpecConverter::toExcelCellAlignment),
        style.font().flatMap(WorkbookCommandCellValueSpecConverter::toExcelCellFont),
        style.fill().flatMap(WorkbookCommandCellValueSpecConverter::toExcelCellFill),
        style.border().flatMap(WorkbookCommandCellValueSpecConverter::toExcelBorder),
        style.protection().flatMap(WorkbookCommandCellValueSpecConverter::toExcelCellProtection));
  }

  static Optional<ExcelCellAlignment> toExcelCellAlignment(CellAlignmentPatch alignment) {
    if (alignment == null) {
      return Optional.empty();
    }
    return Optional.of(
        new ExcelCellAlignment(
            alignment.wrapText(),
            alignment.horizontalAlignment(),
            alignment.verticalAlignment(),
            alignment.textRotation(),
            alignment.indentation()));
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
            font.fontColor().flatMap(WorkbookCommandCellValueSpecConverter::toExcelColor),
            font.underline(),
            font.strikeout()));
  }

  static Optional<ExcelCellFill> toExcelCellFill(CellFillSpec fill) {
    if (fill == null) {
      return Optional.empty();
    }
    return Optional.of(
        switch (fill) {
          case CellFillSpec.PatternOnly pattern -> ExcelCellFill.pattern(pattern.pattern());
          case CellFillSpec.PatternForeground pattern ->
              ExcelCellFill.patternForeground(
                  pattern.pattern(),
                  toRequiredExcelColor(pattern.foregroundColor(), "foregroundColor"));
          case CellFillSpec.PatternBackground pattern ->
              ExcelCellFill.patternBackground(
                  pattern.pattern(),
                  toRequiredExcelColor(pattern.backgroundColor(), "backgroundColor"));
          case CellFillSpec.PatternForegroundBackground pattern ->
              ExcelCellFill.patternColors(
                  pattern.pattern(),
                  toRequiredExcelColor(pattern.foregroundColor(), "foregroundColor"),
                  toRequiredExcelColor(pattern.backgroundColor(), "backgroundColor"));
          case CellFillSpec.Gradient gradient ->
              ExcelCellFill.gradient(toExcelGradientFill(gradient.gradient()));
        });
  }

  static Optional<ExcelCellProtection> toExcelCellProtection(CellProtectionPatch protection) {
    if (protection == null) {
      return Optional.empty();
    }
    return Optional.of(new ExcelCellProtection(protection.locked(), protection.hiddenFormula()));
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

  static Optional<ExcelBorder> toExcelBorder(CellBorderPatch border) {
    if (border == null) {
      return Optional.empty();
    }
    return Optional.of(
        new ExcelBorder(
            border.all().flatMap(WorkbookCommandCellValueSpecConverter::toExcelBorderSide),
            border.top().flatMap(WorkbookCommandCellValueSpecConverter::toExcelBorderSide),
            border.right().flatMap(WorkbookCommandCellValueSpecConverter::toExcelBorderSide),
            border.bottom().flatMap(WorkbookCommandCellValueSpecConverter::toExcelBorderSide),
            border.left().flatMap(WorkbookCommandCellValueSpecConverter::toExcelBorderSide)));
  }

  static Optional<ExcelBorderSide> toExcelBorderSide(BorderSidePatch side) {
    return side == null
        ? Optional.empty()
        : Optional.of(
            new ExcelBorderSide(
                side.style(),
                side.color().flatMap(WorkbookCommandCellValueSpecConverter::toExcelColor)));
  }

  private static ExcelGradientFill toExcelGradientFill(CellGradientFillSpec gradient) {
    return switch (gradient) {
      case CellGradientFillSpec.Linear linear ->
          ExcelGradientFill.linear(
              linear.degree(),
              linear.stops().stream()
                  .map(
                      stop ->
                          new ExcelGradientStop(
                              stop.position(), toRequiredExcelColor(stop.color(), "color")))
                  .toList());
      case CellGradientFillSpec.Path path ->
          ExcelGradientFill.path(
              path.left(),
              path.right(),
              path.top(),
              path.bottom(),
              path.stops().stream()
                  .map(
                      stop ->
                          new ExcelGradientStop(
                              stop.position(), toRequiredExcelColor(stop.color(), "color")))
                  .toList());
    };
  }

  static Optional<ExcelColor> toExcelColor(ColorSpec color) {
    if (color == null) {
      return Optional.empty();
    }
    return Optional.of(toRequiredExcelColor(color, "color"));
  }

  static ExcelColor toRequiredExcelColor(ColorSpec color, String fieldName) {
    Objects.requireNonNull(color, fieldName + " must not be null");
    return switch (color) {
      case ColorSpec.Rgb rgb -> ExcelColor.rgb(rgb.rgb().value(), rgb.tint());
      case ColorSpec.Theme theme -> ExcelColor.theme(theme.theme(), theme.tint());
      case ColorSpec.Indexed indexed -> ExcelColor.indexed(indexed.indexed(), indexed.tint());
    };
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
