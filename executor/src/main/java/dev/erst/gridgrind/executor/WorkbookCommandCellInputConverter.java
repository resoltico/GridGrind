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
        toExcelCellFont(run.font()));
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
        comment.runs() == null
            ? null
            : new ExcelRichText(
                comment.runs().stream()
                    .map(WorkbookCommandCellInputConverter::toExcelRichTextRun)
                    .toList()),
        comment.anchor() == null
            ? null
            : new ExcelCommentAnchor(
                comment.anchor().firstColumn(),
                comment.anchor().firstRow(),
                comment.anchor().lastColumn(),
                comment.anchor().lastRow()));
  }

  static ExcelCellStyle toExcelCellStyle(CellStyleInput style) {
    return new ExcelCellStyle(
        style.numberFormat(),
        toExcelCellAlignment(style.alignment()),
        toExcelCellFont(style.font()),
        toExcelCellFill(style.fill()),
        toExcelBorder(style.border()),
        toExcelCellProtection(style.protection()));
  }

  static ExcelCellAlignment toExcelCellAlignment(CellAlignmentInput alignment) {
    if (alignment == null) {
      return null;
    }
    return new ExcelCellAlignment(
        alignment.wrapText(),
        alignment.horizontalAlignment(),
        alignment.verticalAlignment(),
        alignment.textRotation(),
        alignment.indentation());
  }

  static ExcelCellFont toExcelCellFont(CellFontInput font) {
    if (font == null) {
      return null;
    }
    return new ExcelCellFont(
        font.bold(),
        font.italic(),
        font.fontName(),
        toExcelFontHeight(font.fontHeight()),
        toExcelColor(
            font.fontColor(), font.fontColorTheme(), font.fontColorIndexed(), font.fontColorTint()),
        font.underline(),
        font.strikeout());
  }

  static ExcelCellFill toExcelCellFill(CellFillInput fill) {
    if (fill == null) {
      return null;
    }
    return new ExcelCellFill(
        fill.pattern(),
        toExcelColor(
            fill.foregroundColor(),
            fill.foregroundColorTheme(),
            fill.foregroundColorIndexed(),
            fill.foregroundColorTint()),
        toExcelColor(
            fill.backgroundColor(),
            fill.backgroundColorTheme(),
            fill.backgroundColorIndexed(),
            fill.backgroundColorTint()),
        toExcelGradientFill(fill.gradient()));
  }

  static ExcelCellProtection toExcelCellProtection(CellProtectionInput protection) {
    if (protection == null) {
      return null;
    }
    return new ExcelCellProtection(protection.locked(), protection.hiddenFormula());
  }

  static ExcelFontHeight toExcelFontHeight(FontHeightInput fontHeight) {
    if (fontHeight == null) {
      return null;
    }
    return switch (fontHeight) {
      case FontHeightInput.Points points -> ExcelFontHeight.fromPoints(points.points());
      case FontHeightInput.Twips twips -> new ExcelFontHeight(twips.twips());
    };
  }

  static ExcelBorder toExcelBorder(CellBorderInput border) {
    if (border == null) {
      return null;
    }
    return new ExcelBorder(
        toExcelBorderSide(border.all()),
        toExcelBorderSide(border.top()),
        toExcelBorderSide(border.right()),
        toExcelBorderSide(border.bottom()),
        toExcelBorderSide(border.left()));
  }

  static ExcelBorderSide toExcelBorderSide(CellBorderSideInput side) {
    return side == null
        ? null
        : new ExcelBorderSide(
            side.style(),
            toExcelColor(side.color(), side.colorTheme(), side.colorIndexed(), side.colorTint()));
  }

  static ExcelColor toExcelColor(String rgb, Integer theme, Integer indexed, Double tint) {
    if (rgb == null && theme == null && indexed == null) {
      return null;
    }
    return new ExcelColor(rgb, theme, indexed, tint);
  }

  static ExcelGradientFill toExcelGradientFill(CellGradientFillInput gradient) {
    if (gradient == null) {
      return null;
    }
    return new ExcelGradientFill(
        gradient.type(),
        gradient.degree(),
        gradient.left(),
        gradient.right(),
        gradient.top(),
        gradient.bottom(),
        gradient.stops().stream()
            .map(stop -> new ExcelGradientStop(stop.position(), toExcelColor(stop.color())))
            .toList());
  }

  static ExcelColor toExcelColor(ColorInput color) {
    if (color == null) {
      return null;
    }
    return new ExcelColor(color.rgb(), color.theme(), color.indexed(), color.tint());
  }
}
