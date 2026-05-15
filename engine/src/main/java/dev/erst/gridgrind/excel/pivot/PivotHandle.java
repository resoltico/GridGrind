package dev.erst.gridgrind.excel.pivot;

import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Internal workbook-global identity for one discovered pivot table. */
public record PivotHandle(
    int sheetIndex, int ordinalOnSheet, String sheetName, XSSFSheet sheet, XSSFPivotTable table) {}
