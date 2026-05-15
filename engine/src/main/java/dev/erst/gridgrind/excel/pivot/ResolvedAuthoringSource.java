package dev.erst.gridgrind.excel.pivot;

import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;

/** Resolved workbook source for pivot-table authoring. */
@SuppressWarnings("PMD.CommentRequired")
public record ResolvedAuthoringSource(
    ResolvedAuthoringSourceKind kind,
    XSSFSheet sheet,
    AreaReference area,
    String description,
    Optional<Name> namedRange,
    Optional<XSSFTable> table) {
  public ResolvedAuthoringSource {
    Objects.requireNonNull(kind, "kind must not be null");
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(area, "area must not be null");
    Objects.requireNonNull(description, "description must not be null");
    Objects.requireNonNull(namedRange, "namedRange must not be null");
    Objects.requireNonNull(table, "table must not be null");
  }

  public static ResolvedAuthoringSource range(XSSFSheet sheet, AreaReference area) {
    return new ResolvedAuthoringSource(
        ResolvedAuthoringSourceKind.RANGE,
        sheet,
        area,
        sheet.getSheetName() + "!" + area.formatAsString(),
        Optional.empty(),
        Optional.empty());
  }

  public static ResolvedAuthoringSource namedRange(
      XSSFSheet sheet, AreaReference area, Name namedRange) {
    return new ResolvedAuthoringSource(
        ResolvedAuthoringSourceKind.NAMED_RANGE,
        sheet,
        area,
        "named range " + namedRange.getNameName(),
        Optional.of(namedRange),
        Optional.empty());
  }

  public static ResolvedAuthoringSource table(
      XSSFSheet sheet, AreaReference area, XSSFTable table) {
    return new ResolvedAuthoringSource(
        ResolvedAuthoringSourceKind.TABLE,
        sheet,
        area,
        "table " + table.getName(),
        Optional.empty(),
        Optional.of(table));
  }
}
