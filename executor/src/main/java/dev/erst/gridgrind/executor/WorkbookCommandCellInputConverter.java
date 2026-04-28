package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.ArrayFormulaInput;
import dev.erst.gridgrind.contract.dto.CellAlignmentInput;
import dev.erst.gridgrind.contract.dto.CellBorderInput;
import dev.erst.gridgrind.contract.dto.CellBorderSideInput;
import dev.erst.gridgrind.contract.dto.CellFillInput;
import dev.erst.gridgrind.contract.dto.CellFontInput;
import dev.erst.gridgrind.contract.dto.CellGradientFillInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellProtectionInput;
import dev.erst.gridgrind.contract.dto.CellStyleInput;
import dev.erst.gridgrind.contract.dto.ColorInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.FontHeightInput;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
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
import dev.erst.gridgrind.excel.ExcelFontHeight;
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
final class WorkbookCommandCellInputConverter {
  private WorkbookCommandCellInputConverter() {}

  static ExcelCellValue toExcelCellValue(CellInput value) {
    return switch (value) {
      case CellInput.Blank _ -> ExcelCellValue.blank();
      case CellInput.Text text ->
          ExcelCellValue.text(WorkbookCommandSourceSupport.inlineText(text.source(), "cell text"));
      case CellInput.RichText richText -> ExcelCellValue.richText(toExcelRichText(richText));
      case CellInput.Numeric numeric -> ExcelCellValue.number(numeric.number());
      case CellInput.BooleanValue booleanValue -> ExcelCellValue.bool(booleanValue.bool());
      case CellInput.Date date -> ExcelCellValue.date(date.date());
      case CellInput.DateTime dateTime -> ExcelCellValue.dateTime(dateTime.dateTime());
      case CellInput.Formula formula ->
          ExcelCellValue.formula(
              WorkbookCommandSourceSupport.inlineText(formula.source(), "formula"));
    };
  }

  static ExcelArrayFormulaDefinition toExcelArrayFormulaDefinition(ArrayFormulaInput input) {
    return new ExcelArrayFormulaDefinition(
        WorkbookCommandSourceSupport.inlineText(input.source(), "array formula"));
  }

  static ExcelRichText toExcelRichText(CellInput.RichText richText) {
    return new ExcelRichText(
        richText.runs().stream()
            .map(WorkbookCommandCellInputConverter::toExcelRichTextRun)
            .toList());
  }

  static ExcelRichTextRun toExcelRichTextRun(RichTextRunInput run) {
    return new ExcelRichTextRun(
        WorkbookCommandSourceSupport.inlineText(run.source(), "rich-text run"),
        toExcelCellFont(run.font()).orElse(null));
  }

  static ExcelHyperlink toExcelHyperlink(HyperlinkTarget target) {
    return switch (target) {
      case HyperlinkTarget.Url url -> new ExcelHyperlink.Url(url.target());
      case HyperlinkTarget.Email email -> new ExcelHyperlink.Email(email.email());
      case HyperlinkTarget.File file -> new ExcelHyperlink.File(file.path());
      case HyperlinkTarget.Document document -> new ExcelHyperlink.Document(document.target());
    };
  }

  static ExcelComment toExcelComment(CommentInput comment) {
    return new ExcelComment(
        WorkbookCommandSourceSupport.inlineText(comment.text(), "comment text"),
        comment.author(),
        comment.visible(),
        commentRuns(comment).orElse(null),
        commentAnchor(comment).orElse(null));
  }

  static ExcelCellStyle toExcelCellStyle(CellStyleInput style) {
    return new ExcelCellStyle(
        style.numberFormat(),
        toExcelCellAlignment(style.alignment()).orElse(null),
        toExcelCellFont(style.font()).orElse(null),
        toExcelCellFill(style.fill()).orElse(null),
        toExcelBorder(style.border()).orElse(null),
        toExcelCellProtection(style.protection()).orElse(null));
  }

  static Optional<ExcelCellAlignment> toExcelCellAlignment(CellAlignmentInput alignment) {
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

  static Optional<ExcelCellFont> toExcelCellFont(CellFontInput font) {
    if (font == null) {
      return Optional.empty();
    }
    return Optional.of(
        new ExcelCellFont(
            font.bold(),
            font.italic(),
            font.fontName(),
            toExcelFontHeight(font.fontHeight()).orElse(null),
            toExcelColor(font.fontColor()).orElse(null),
            font.underline(),
            font.strikeout()));
  }

  static Optional<ExcelCellFill> toExcelCellFill(CellFillInput fill) {
    if (fill == null) {
      return Optional.empty();
    }
    return Optional.of(
        switch (fill) {
          case CellFillInput.PatternOnly pattern -> ExcelCellFill.pattern(pattern.pattern());
          case CellFillInput.PatternForeground pattern ->
              ExcelCellFill.patternForeground(
                  pattern.pattern(),
                  toRequiredExcelColor(pattern.foregroundColor(), "foregroundColor"));
          case CellFillInput.PatternBackground pattern ->
              ExcelCellFill.patternBackground(
                  pattern.pattern(),
                  toRequiredExcelColor(pattern.backgroundColor(), "backgroundColor"));
          case CellFillInput.PatternForegroundBackground pattern ->
              ExcelCellFill.patternColors(
                  pattern.pattern(),
                  toRequiredExcelColor(pattern.foregroundColor(), "foregroundColor"),
                  toRequiredExcelColor(pattern.backgroundColor(), "backgroundColor"));
          case CellFillInput.Gradient gradient ->
              ExcelCellFill.gradient(toExcelGradientFill(gradient.gradient()));
        });
  }

  static Optional<ExcelCellProtection> toExcelCellProtection(CellProtectionInput protection) {
    if (protection == null) {
      return Optional.empty();
    }
    return Optional.of(new ExcelCellProtection(protection.locked(), protection.hiddenFormula()));
  }

  static Optional<ExcelFontHeight> toExcelFontHeight(FontHeightInput fontHeight) {
    if (fontHeight == null) {
      return Optional.empty();
    }
    return Optional.of(
        switch (fontHeight) {
          case FontHeightInput.Points points -> ExcelFontHeight.fromPoints(points.points());
          case FontHeightInput.Twips twips -> new ExcelFontHeight(twips.twips());
        });
  }

  static Optional<ExcelBorder> toExcelBorder(CellBorderInput border) {
    if (border == null) {
      return Optional.empty();
    }
    return Optional.of(
        new ExcelBorder(
            toExcelBorderSide(border.all()).orElse(null),
            toExcelBorderSide(border.top()).orElse(null),
            toExcelBorderSide(border.right()).orElse(null),
            toExcelBorderSide(border.bottom()).orElse(null),
            toExcelBorderSide(border.left()).orElse(null)));
  }

  static Optional<ExcelBorderSide> toExcelBorderSide(CellBorderSideInput side) {
    return side == null
        ? Optional.empty()
        : Optional.of(new ExcelBorderSide(side.style(), toExcelColor(side.color()).orElse(null)));
  }

  private static ExcelGradientFill toExcelGradientFill(CellGradientFillInput gradient) {
    return switch (gradient) {
      case CellGradientFillInput.Linear linear ->
          ExcelGradientFill.linear(
              linear.degree(),
              linear.stops().stream()
                  .map(
                      stop ->
                          new ExcelGradientStop(
                              stop.position(), toRequiredExcelColor(stop.color(), "color")))
                  .toList());
      case CellGradientFillInput.Path path ->
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

  static Optional<ExcelColor> toExcelColor(ColorInput color) {
    if (color == null) {
      return Optional.empty();
    }
    return Optional.of(toRequiredExcelColor(color, "color"));
  }

  static ExcelColor toRequiredExcelColor(ColorInput color, String fieldName) {
    Objects.requireNonNull(color, fieldName + " must not be null");
    return switch (color) {
      case ColorInput.Rgb rgb -> ExcelColor.rgb(rgb.rgb(), rgb.tint());
      case ColorInput.Theme theme -> ExcelColor.theme(theme.theme(), theme.tint());
      case ColorInput.Indexed indexed -> ExcelColor.indexed(indexed.indexed(), indexed.tint());
    };
  }

  private static Optional<ExcelRichText> commentRuns(CommentInput comment) {
    return comment.runs().isEmpty()
        ? Optional.empty()
        : Optional.of(
            new ExcelRichText(
                comment.runs().orElseThrow().stream()
                    .map(WorkbookCommandCellInputConverter::toExcelRichTextRun)
                    .toList()));
  }

  private static Optional<ExcelCommentAnchor> commentAnchor(CommentInput comment) {
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
