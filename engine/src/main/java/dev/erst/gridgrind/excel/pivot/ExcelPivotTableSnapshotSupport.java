package dev.erst.gridgrind.excel.pivot;

import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheDefinition;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheRecords;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCache;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheetSource;

/** Owns pivot-table snapshot normalization and cache/readback helpers. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelPivotTableSnapshotSupport {
  private ExcelPivotTableSnapshotSupport() {}

  public static PivotTableReport snapshot(XSSFWorkbook workbook, PivotHandle handle) {
    String name = ExcelPivotTableIdentitySupport.resolvedName(handle);
    PivotLocation location = ExcelPivotTableIdentitySupport.safeLocation(handle).orElse(null);
    PivotTableReport.Anchor anchor =
        location == null
            ? new PivotTableReport.Anchor(
                new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1"))
            : new PivotTableReport.Anchor(
                new dev.erst.gridgrind.core.vocabulary.CellAddress(location.topLeftAddress()),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel(location.locationRange()));
    if (location == null) {
      return unsupportedSnapshot(
          handle, name, anchor, "Pivot table location range is missing or malformed.");
    }
    try {
      CTPivotTableDefinition definition = handle.table().getCTPivotTableDefinition();
      List<String> sourceColumnNames = cacheFieldNames(handle.table());
      boolean valuesAxisOnColumns = false;
      List<PivotTableReport.Field> columnLabels = new ArrayList<>();
      if (definition.getColFields() != null) {
        for (CTField field : definition.getColFields().getFieldArray()) {
          if (field.getX() == -2) {
            valuesAxisOnColumns = true;
            continue;
          }
          columnLabels.add(sourceField(sourceColumnNames, field.getX()));
        }
      }
      List<PivotTableReport.DataField> dataFields =
          snapshotDataFields(workbook, definition, sourceColumnNames);
      if (dataFields.isEmpty()) {
        return unsupportedSnapshot(
            handle, name, anchor, "Pivot table does not contain any data fields.");
      }
      List<PivotTableReport.Field> rowFields =
          definition.getRowFields() == null
              ? List.of()
              : snapshotFields(
                  Objects.requireNonNull(
                      definition.getRowFields().getFieldArray(),
                      "pivot row field array must not be null"),
                  sourceColumnNames);
      List<PivotTableReport.Field> pageFields =
          definition.getPageFields() == null
              ? List.of()
              : snapshotPageFields(
                  Objects.requireNonNull(
                      definition.getPageFields().getPageFieldArray(),
                      "pivot page field array must not be null"),
                  sourceColumnNames);
      return new PivotTableReport.Supported(
          name,
          new dev.erst.gridgrind.core.vocabulary.SheetName(handle.sheetName()),
          anchor,
          snapshotSource(workbook, handle.table()),
          rowFields,
          columnLabels,
          pageFields,
          dataFields,
          valuesAxisOnColumns);
    } catch (IllegalArgumentException exception) {
      return unsupportedSnapshot(
          handle,
          name,
          anchor,
          Objects.requireNonNullElse(
              exception.getMessage(),
              "The pivot table could not be normalized into the modeled contract."));
    }
  }

  public static List<PivotTableReport.Field> snapshotFields(
      @Nullable CTField[] fields, List<String> sourceColumnNames) {
    if (fields == null || fields.length == 0) {
      return List.of();
    }
    List<PivotTableReport.Field> snapshots = new ArrayList<>(fields.length);
    for (CTField field : fields) {
      Objects.requireNonNull(field, "pivot field entries must not be null");
      snapshots.add(sourceField(sourceColumnNames, field.getX()));
    }
    return List.copyOf(snapshots);
  }

  public static List<PivotTableReport.Field> snapshotPageFields(
      @Nullable CTPageField[] pageFields, List<String> sourceColumnNames) {
    if (pageFields == null || pageFields.length == 0) {
      return List.of();
    }
    List<PivotTableReport.Field> snapshots = new ArrayList<>(pageFields.length);
    for (CTPageField pageField : pageFields) {
      Objects.requireNonNull(pageField, "pivot page-field entries must not be null");
      snapshots.add(sourceField(sourceColumnNames, pageField.getFld()));
    }
    return List.copyOf(snapshots);
  }

  public static List<PivotTableReport.DataField> snapshotDataFields(
      XSSFWorkbook workbook, CTPivotTableDefinition definition, List<String> sourceColumnNames) {
    if (definition.getDataFields() == null
        || definition.getDataFields().sizeOfDataFieldArray() == 0) {
      return List.of();
    }
    List<PivotTableReport.DataField> dataFields = new ArrayList<>();
    for (CTDataField dataField : definition.getDataFields().getDataFieldArray()) {
      int sourceColumnIndex = Math.toIntExact(dataField.getFld());
      String sourceColumnName =
          ExcelPivotTableSourceSupport.sourceColumnName(sourceColumnNames, sourceColumnIndex);
      dataFields.add(
          new PivotTableReport.DataField(
              sourceColumnIndex,
              sourceColumnName,
              fromSubtotal(
                  dataField.getSubtotal() == null
                      ? DataConsolidateFunction.SUM.getValue()
                      : dataField.getSubtotal().intValue()),
              ExcelPivotTableIdentitySupport.nonBlankOrDefault(
                  dataField.getName(), sourceColumnName),
              numberFormat(workbook, dataField.isSetNumFmtId() ? dataField.getNumFmtId() : null)));
    }
    return List.copyOf(dataFields);
  }

  public static PivotTableReport.Unsupported unsupportedSnapshot(
      PivotHandle handle, String name, PivotTableReport.Anchor anchor, String detail) {
    return new PivotTableReport.Unsupported(
        name, new dev.erst.gridgrind.core.vocabulary.SheetName(handle.sheetName()), anchor, detail);
  }

  public static PivotTableReport.PivotSourceReport snapshotSource(
      XSSFWorkbook workbook, XSSFPivotTable pivotTable) {
    XSSFPivotCacheDefinition cacheDefinition = requiredCacheDefinition(pivotTable);
    if (cacheDefinition.getCTPivotCacheDefinition().getCacheSource() == null
        || !cacheDefinition.getCTPivotCacheDefinition().getCacheSource().isSetWorksheetSource()) {
      throw new IllegalArgumentException("Pivot cache source is missing its worksheetSource.");
    }
    CTWorksheetSource worksheetSource =
        cacheDefinition.getCTPivotCacheDefinition().getCacheSource().getWorksheetSource();
    String sheetName =
        ExcelPivotTableIdentitySupport.requireNonBlank(
            worksheetSource.getSheet(), "Pivot source sheet is missing.");
    if (worksheetSource.getRef() != null && !worksheetSource.getRef().isBlank()) {
      AreaReference area =
          new AreaReference(worksheetSource.getRef(), SpreadsheetVersion.EXCEL2007);
      return new PivotTableReport.PivotSourceReport.Range(
          new dev.erst.gridgrind.core.vocabulary.SheetName(sheetName),
          new dev.erst.gridgrind.core.vocabulary.RangeLabel(
              ExcelPivotTableIdentitySupport.normalizeArea(area)));
    }

    String sourceName =
        ExcelPivotTableIdentitySupport.requireNonBlank(
            worksheetSource.getName(), "Pivot source name is missing.");
    List<Name> namedRanges =
        ExcelPivotTableSourceSupport.matchingNamedRanges(workbook, sourceName, sheetName);
    if (namedRanges.size() == 1) {
      AreaReference area = ExcelPivotTableSourceSupport.namedRangeArea(namedRanges.getFirst());
      return new PivotTableReport.PivotSourceReport.NamedRange(
          new dev.erst.gridgrind.core.vocabulary.DefinedName(sourceName),
          new dev.erst.gridgrind.core.vocabulary.SheetName(
              ExcelPivotTableSourceSupport.sourceSheetName(
                  area, namedRanges.getFirst(), sheetName)),
          new dev.erst.gridgrind.core.vocabulary.RangeLabel(
              ExcelPivotTableIdentitySupport.normalizeArea(area)));
    }
    if (namedRanges.size() > 1) {
      throw new IllegalArgumentException(
          "Pivot source name '"
              + sourceName
              + "' is ambiguous because multiple matching named ranges exist.");
    }

    Optional<XSSFTable> table =
        ExcelPivotTableSourceSupport.tableByName(workbook, sourceName, sheetName);
    if (table.isPresent()) {
      return new PivotTableReport.PivotSourceReport.Table(
          new dev.erst.gridgrind.core.vocabulary.TableName(sourceName),
          new dev.erst.gridgrind.core.vocabulary.SheetName(table.orElseThrow().getSheetName()),
          new dev.erst.gridgrind.core.vocabulary.RangeLabel(
              ExcelPivotTableIdentitySupport.normalizeArea(
                  new AreaReference(
                      table.orElseThrow().getStartCellReference(),
                      table.orElseThrow().getEndCellReference(),
                      SpreadsheetVersion.EXCEL2007))));
    }
    throw new IllegalArgumentException(
        "Pivot source name '"
            + sourceName
            + "' does not resolve to an existing named range or table.");
  }

  static List<String> cacheFieldNames(XSSFPivotTable pivotTable) {
    XSSFPivotCacheDefinition cacheDefinition = requiredCacheDefinition(pivotTable);
    if (cacheDefinition.getCTPivotCacheDefinition().getCacheFields() == null) {
      throw new IllegalArgumentException("Pivot cache definition is missing cacheFields.");
    }
    List<String> fieldNames = new ArrayList<>();
    for (var cacheField :
        cacheDefinition.getCTPivotCacheDefinition().getCacheFields().getCacheFieldArray()) {
      fieldNames.add(
          ExcelPivotTableIdentitySupport.requireNonBlank(
              cacheField.getName(), "Pivot cache field name is missing."));
    }
    return List.copyOf(fieldNames);
  }

  public static PivotTableReport.Field sourceField(
      List<String> sourceColumnNames, int sourceColumnIndex) {
    return new PivotTableReport.Field(
        sourceColumnIndex,
        ExcelPivotTableSourceSupport.sourceColumnName(sourceColumnNames, sourceColumnIndex));
  }

  public static ExcelPivotDataConsolidateFunction fromSubtotal(int subtotalValue) {
    for (DataConsolidateFunction function : DataConsolidateFunction.values()) {
      if (function.getValue() == subtotalValue) {
        return ExcelPivotDataPoiBridge.fromPoi(function);
      }
    }
    throw new IllegalArgumentException(
        "Unsupported pivot data consolidate function value: " + subtotalValue);
  }

  public static Optional<String> numberFormat(XSSFWorkbook workbook, @Nullable Long numFmtId) {
    if (numFmtId == null) {
      return Optional.empty();
    }
    short formatIndex = numFmtId > Short.MAX_VALUE ? Short.MAX_VALUE : (short) numFmtId.longValue();
    String format = workbook.createDataFormat().getFormat(formatIndex);
    return format == null || format.isBlank() ? Optional.empty() : Optional.of(format);
  }

  public static XSSFPivotCacheDefinition requiredCacheDefinition(XSSFPivotTable pivotTable) {
    return cacheDefinition(pivotTable)
        .orElseThrow(
            () ->
                new IllegalArgumentException(
                    "Pivot table is missing its cache definition relation."));
  }

  public static Optional<XSSFPivotCacheDefinition> cacheDefinition(XSSFPivotTable pivotTable) {
    return Optional.ofNullable(pivotTable.getPivotCacheDefinition());
  }

  public static Optional<XSSFPivotCacheRecords> cacheRecords(
      XSSFPivotCacheDefinition cacheDefinition) {
    return firstRelation(cacheDefinition, XSSFPivotCacheRecords.class);
  }

  public static <T extends POIXMLDocumentPart> Optional<T> firstRelation(
      POIXMLDocumentPart parent, Class<T> relationType) {
    for (POIXMLDocumentPart relation : parent.getRelations()) {
      if (relationType.isInstance(relation)) {
        return Optional.of(relationType.cast(relation));
      }
    }
    return Optional.empty();
  }

  public static Optional<CTPivotCache> workbookPivotCache(XSSFWorkbook workbook, long cacheId) {
    for (CTPivotCache cache : workbookPivotCaches(workbook)) {
      if (cache.getCacheId() == cacheId) {
        return Optional.of(cache);
      }
    }
    return Optional.empty();
  }

  public static List<CTPivotCache> workbookPivotCaches(XSSFWorkbook workbook) {
    if (!workbook.getCTWorkbook().isSetPivotCaches()) {
      return List.of();
    }
    return List.of(workbook.getCTWorkbook().getPivotCaches().getPivotCacheArray());
  }
}
