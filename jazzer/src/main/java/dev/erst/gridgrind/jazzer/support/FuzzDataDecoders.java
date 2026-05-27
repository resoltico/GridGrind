package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import java.util.List;

/** Decodes bounded structured values from a Jazzer data provider. */
public final class FuzzDataDecoders {
  private FuzzDataDecoders() {}

  /** Returns a bounded sheet name, optionally valid for Excel sheet creation. */
  public static String nextSheetName(GridGrindFuzzData data, boolean valid) {
    return FuzzAddressDecoders.nextSheetName(data, valid);
  }

  /** Returns a bounded A1-style cell address. */
  public static String nextCellAddress(GridGrindFuzzData data, boolean valid) {
    return FuzzAddressDecoders.nextCellAddress(data, valid);
  }

  /** Returns a bounded non-blank A1-style cell address that may still be semantically invalid. */
  public static String nextNonBlankCellAddress(GridGrindFuzzData data, boolean valid) {
    return FuzzAddressDecoders.nextNonBlankCellAddress(data, valid);
  }

  /** Returns a bounded rectangular A1-style range. */
  public static String nextRange(GridGrindFuzzData data, boolean valid) {
    return FuzzAddressDecoders.nextRange(data, valid);
  }

  /**
   * Returns a bounded non-blank rectangular A1-style range that may still be semantically invalid.
   */
  public static String nextNonBlankRange(GridGrindFuzzData data, boolean valid) {
    return FuzzAddressDecoders.nextNonBlankRange(data, valid);
  }

  /** Returns a protocol-layer cell input value. */
  public static CellInput nextCellInput(GridGrindFuzzData data) {
    return FuzzValueDecoders.nextCellInput(data);
  }

  /** Returns a workbook-core cell value. */
  public static ExcelCellValue nextExcelCellValue(GridGrindFuzzData data) {
    return FuzzValueDecoders.nextExcelCellValue(data);
  }

  /** Returns a bounded style patch. */
  public static ExcelCellStyle nextStyle(GridGrindFuzzData data) {
    return FuzzStyleDecoders.nextStyle(data);
  }

  /** Returns a non-empty rectangular matrix of protocol cell inputs. */
  public static List<List<CellInput>> nextProtocolMatrix(
      GridGrindFuzzData data, int rowCount, int columnCount) {
    return FuzzValueDecoders.nextProtocolMatrix(data, rowCount, columnCount);
  }

  /** Returns a non-empty rectangular matrix of workbook-core cell values. */
  public static List<List<ExcelCellValue>> nextExcelMatrix(
      GridGrindFuzzData data, int rowCount, int columnCount) {
    return FuzzValueDecoders.nextExcelMatrix(data, rowCount, columnCount);
  }
}
