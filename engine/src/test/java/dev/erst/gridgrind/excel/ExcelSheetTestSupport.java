package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.ExcelStyleTestAccess.*;
import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Shared helpers for ExcelSheet coverage slices. */
class ExcelSheetTestSupport {
  org.apache.poi.ss.usermodel.Hyperlink hyperlink(
      XSSFWorkbook workbook, HyperlinkType hyperlinkType, String address) {
    org.apache.poi.ss.usermodel.Hyperlink hyperlink =
        workbook.getCreationHelper().createHyperlink(hyperlinkType);
    if (hyperlinkType != HyperlinkType.NONE) {
      hyperlink.setAddress(address);
    }
    return hyperlink;
  }

  static org.apache.poi.ss.usermodel.Hyperlink hyperlinkWithNullType() {
    return new org.apache.poi.ss.usermodel.Hyperlink() {
      @Override
      public HyperlinkType getType() {
        return null;
      }

      @Override
      public String getAddress() {
        return "ignored";
      }

      @Override
      public void setAddress(String address) {}

      @Override
      public String getLabel() {
        return null;
      }

      @Override
      public void setLabel(String label) {}

      @Override
      public int getFirstRow() {
        return 0;
      }

      @Override
      public void setFirstRow(int row) {}

      @Override
      public int getLastRow() {
        return 0;
      }

      @Override
      public void setLastRow(int row) {}

      @Override
      public int getFirstColumn() {
        return 0;
      }

      @Override
      public void setFirstColumn(int column) {}

      @Override
      public int getLastColumn() {
        return 0;
      }

      @Override
      public void setLastColumn(int column) {}
    };
  }

  Comment comment(XSSFWorkbook workbook, Sheet sheet, String text, String author, boolean visible) {
    Comment comment = emptyComment(workbook, sheet);
    comment.setString(workbook.getCreationHelper().createRichTextString(text));
    comment.setAuthor(author);
    comment.setVisible(visible);
    return comment;
  }

  Comment emptyComment(XSSFWorkbook workbook, Sheet sheet) {
    ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();
    anchor.setRow1(0);
    anchor.setRow2(3);
    anchor.setCol1(0);
    anchor.setCol2(3);
    return sheet.createDrawingPatriarch().createCellComment(anchor);
  }

  static Comment commentWithoutAnchor(Sheet sheet, Comment comment) {
    var commentsTable = new org.apache.poi.xssf.model.CommentsTable();
    commentsTable.setSheet(sheet);
    var anchorless =
        new org.apache.poi.xssf.usermodel.XSSFComment(
            commentsTable,
            commentsTable.newComment(new org.apache.poi.ss.util.CellAddress("A1")),
            null);
    anchorless.setString(comment.getString());
    anchorless.setAuthor(comment.getAuthor());
    anchorless.setVisible(comment.isVisible());
    return anchorless;
  }

  static ExcelColorSnapshot rgb(String rgb) {
    return ExcelColorSnapshot.rgb(rgb);
  }

  /** Local test-only type name that trips missing-workbook classification. */
  static final class LocalWorkbookNotFoundException extends RuntimeException {
    static final long serialVersionUID = 1L;

    LocalWorkbookNotFoundException(String message) {
      super(message);
    }
  }

  /** Test-only formula runtime with explicit context and optional evaluation/display failures. */
  record StaticContextFormulaRuntime(
      ExcelFormulaRuntimeContext context,
      RuntimeException evaluateFailure,
      RuntimeException displayFailure)
      implements ExcelFormulaRuntime {
    @Override
    public CellValue evaluate(Cell cell) {
      if (evaluateFailure != null) {
        throw evaluateFailure;
      }
      return null;
    }

    @Override
    public CellType evaluateFormulaCell(Cell cell) {
      if (evaluateFailure != null) {
        throw evaluateFailure;
      }
      return CellType._NONE;
    }

    @Override
    public void clearCachedResults() {}

    @Override
    public String displayValue(DataFormatter formatter, Cell cell) {
      if (displayFailure != null) {
        throw displayFailure;
      }
      return "";
    }
  }
}
