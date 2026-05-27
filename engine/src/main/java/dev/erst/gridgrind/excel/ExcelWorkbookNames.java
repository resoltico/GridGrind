package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Name;

/** Workbook defined-name authoring and inspection operations. */
public final class ExcelWorkbookNames {
  private final ExcelWorkbook workbook;

  ExcelWorkbookNames(ExcelWorkbook workbook) {
    this.workbook = Objects.requireNonNull(workbook, "workbook must not be null");
  }

  /** Creates or replaces one named range in workbook or sheet scope. */
  public ExcelWorkbookNames setNamedRange(ExcelNamedRangeDefinition definition) {
    ExcelWorkbookNamedRangeSupport.setNamedRange(workbook, definition);
    return this;
  }

  /** Deletes one named range from workbook or sheet scope. */
  public ExcelWorkbookNames deleteNamedRange(String name, ExcelNamedRangeScope scope) {
    ExcelWorkbookNamedRangeSupport.deleteNamedRange(workbook, name, scope);
    return this;
  }

  /** Returns the number of analyzable named ranges currently present in the workbook. */
  public int namedRangeCount() {
    return namedRanges().size();
  }

  /** Returns every analyzable named range currently present in the workbook. */
  public List<ExcelNamedRangeSnapshot> namedRanges() {
    return ExcelWorkbookNamedRangeSupport.namedRanges(workbook);
  }

  /** Returns whether the POI defined name belongs to the GridGrind user-facing surface. */
  public static boolean shouldExpose(Name name) {
    return ExcelWorkbookNamedRangeSupport.shouldExpose(name);
  }

  /** Returns whether a defined-name triple is user-facing and analyzable by GridGrind. */
  public static boolean shouldExpose(String nameName, boolean functionName, boolean hidden) {
    return ExcelWorkbookNamedRangeSupport.shouldExpose(nameName, functionName, hidden);
  }
}
