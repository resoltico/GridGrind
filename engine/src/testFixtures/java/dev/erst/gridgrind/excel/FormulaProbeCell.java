package dev.erst.gridgrind.excel;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.CellBase;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellRangeAddress;

/** Minimal formula-cell stub for probing scalar-decoding branches without reflection. */
final class FormulaProbeCell extends CellBase {
  private final CellType cachedFormulaResultType;
  private final String stringValue;
  private final double numericValue;
  private final boolean booleanValue;

  FormulaProbeCell(
      CellType cachedFormulaResultType,
      String stringValue,
      double numericValue,
      boolean booleanValue) {
    this.cachedFormulaResultType = cachedFormulaResultType;
    this.stringValue = stringValue;
    this.numericValue = numericValue;
    this.booleanValue = booleanValue;
  }

  @Override
  public CellType getCellType() {
    return CellType.FORMULA;
  }

  @Override
  public CellType getCachedFormulaResultType() {
    return cachedFormulaResultType;
  }

  @Override
  public int getColumnIndex() {
    return 0;
  }

  @Override
  public int getRowIndex() {
    return 0;
  }

  @Override
  public org.apache.poi.ss.usermodel.Sheet getSheet() {
    throw new UnsupportedOperationException();
  }

  @Override
  public org.apache.poi.ss.usermodel.Row getRow() {
    throw new UnsupportedOperationException();
  }

  @Override
  public String getCellFormula() {
    throw new UnsupportedOperationException();
  }

  @Override
  public String getStringCellValue() {
    return stringValue;
  }

  @Override
  public double getNumericCellValue() {
    return numericValue;
  }

  @Override
  public boolean getBooleanCellValue() {
    return booleanValue;
  }

  @Override
  public java.util.Date getDateCellValue() {
    throw new UnsupportedOperationException();
  }

  @Override
  public java.time.LocalDateTime getLocalDateTimeCellValue() {
    throw new UnsupportedOperationException();
  }

  @Override
  public RichTextString getRichStringCellValue() {
    throw new UnsupportedOperationException();
  }

  @Override
  public byte getErrorCellValue() {
    throw new UnsupportedOperationException();
  }

  @Override
  public void setCellValue(boolean value) {
    throw new UnsupportedOperationException();
  }

  @Override
  public void setCellErrorValue(byte value) {
    throw new UnsupportedOperationException();
  }

  @Override
  public void setCellStyle(org.apache.poi.ss.usermodel.CellStyle style) {
    throw new UnsupportedOperationException();
  }

  @Override
  public org.apache.poi.ss.usermodel.CellStyle getCellStyle() {
    throw new UnsupportedOperationException();
  }

  @Override
  public void setAsActiveCell() {
    throw new UnsupportedOperationException();
  }

  @Override
  public void setCellComment(org.apache.poi.ss.usermodel.Comment comment) {
    throw new UnsupportedOperationException();
  }

  @Override
  public org.apache.poi.ss.usermodel.Comment getCellComment() {
    throw new UnsupportedOperationException();
  }

  @Override
  public void removeCellComment() {
    throw new UnsupportedOperationException();
  }

  @Override
  public org.apache.poi.ss.usermodel.Hyperlink getHyperlink() {
    throw new UnsupportedOperationException();
  }

  @Override
  public void setHyperlink(org.apache.poi.ss.usermodel.Hyperlink link) {
    throw new UnsupportedOperationException();
  }

  @Override
  public void removeHyperlink() {
    throw new UnsupportedOperationException();
  }

  @Override
  protected void setCellTypeImpl(CellType cellType) {
    throw new UnsupportedOperationException();
  }

  @Override
  protected void setCellFormulaImpl(String formula) {
    throw new UnsupportedOperationException();
  }

  @Override
  protected void removeFormulaImpl() {
    throw new UnsupportedOperationException();
  }

  @Override
  protected void setCellValueImpl(double value) {
    throw new UnsupportedOperationException();
  }

  @Override
  @SuppressWarnings("PMD.ReplaceJavaUtilDate")
  protected void setCellValueImpl(java.util.Date value) {
    throw new UnsupportedOperationException();
  }

  @Override
  protected void setCellValueImpl(java.time.LocalDateTime value) {
    throw new UnsupportedOperationException();
  }

  @Override
  @SuppressWarnings("PMD.ReplaceJavaUtilCalendar")
  protected void setCellValueImpl(java.util.Calendar value) {
    throw new UnsupportedOperationException();
  }

  @Override
  protected void setCellValueImpl(String value) {
    throw new UnsupportedOperationException();
  }

  @Override
  protected void setCellValueImpl(RichTextString value) {
    throw new UnsupportedOperationException();
  }

  @Override
  protected SpreadsheetVersion getSpreadsheetVersion() {
    return SpreadsheetVersion.EXCEL2007;
  }

  @Override
  public CellRangeAddress getArrayFormulaRange() {
    throw new UnsupportedOperationException();
  }

  @Override
  public boolean isPartOfArrayFormulaGroup() {
    return false;
  }
}
