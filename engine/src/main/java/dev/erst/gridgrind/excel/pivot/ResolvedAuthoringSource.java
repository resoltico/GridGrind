package dev.erst.gridgrind.excel.pivot;

import java.util.Objects;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;

/** Resolved workbook source for pivot-table authoring. */
@SuppressWarnings("PMD.CommentRequired")
public sealed interface ResolvedAuthoringSource
    permits ResolvedAuthoringSource.Range,
        ResolvedAuthoringSource.NamedRange,
        ResolvedAuthoringSource.Table {
  XSSFSheet sheet();

  AreaReference area();

  String description();

  static Range range(XSSFSheet sheet, AreaReference area) {
    return new Range(sheet, area, sheet.getSheetName() + "!" + area.formatAsString());
  }

  static NamedRange namedRange(XSSFSheet sheet, AreaReference area, Name namedRange) {
    return new NamedRange(sheet, area, "named range " + namedRange.getNameName(), namedRange);
  }

  static Table table(XSSFSheet sheet, AreaReference area, XSSFTable table) {
    return new Table(sheet, area, "table " + table.getName(), table);
  }

  record Range(XSSFSheet sheet, AreaReference area, String description)
      implements ResolvedAuthoringSource {
    public Range {
      Objects.requireNonNull(sheet, "sheet must not be null");
      Objects.requireNonNull(area, "area must not be null");
      Objects.requireNonNull(description, "description must not be null");
    }
  }

  record NamedRange(XSSFSheet sheet, AreaReference area, String description, Name namedRange)
      implements ResolvedAuthoringSource {
    public NamedRange {
      Objects.requireNonNull(sheet, "sheet must not be null");
      Objects.requireNonNull(area, "area must not be null");
      Objects.requireNonNull(description, "description must not be null");
      Objects.requireNonNull(namedRange, "namedRange must not be null");
    }
  }

  record Table(XSSFSheet sheet, AreaReference area, String description, XSSFTable table)
      implements ResolvedAuthoringSource {
    public Table {
      Objects.requireNonNull(sheet, "sheet must not be null");
      Objects.requireNonNull(area, "area must not be null");
      Objects.requireNonNull(description, "description must not be null");
      Objects.requireNonNull(table, "table must not be null");
    }
  }
}
