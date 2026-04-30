package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
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
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCache;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheetSource;

/** Owns pivot-table snapshot normalization and cache/readback helpers. */
final class ExcelPivotTableSnapshotSupport {
  private ExcelPivotTableSnapshotSupport() {}

  static ExcelPivotTableSnapshot snapshot(XSSFWorkbook workbook, PivotHandle handle) {
    String name = ExcelPivotTableIdentitySupport.resolvedName(handle);
    PivotLocation location = ExcelPivotTableIdentitySupport.safeLocation(handle).orElse(null);
    ExcelPivotTableSnapshot.Anchor anchor =
        location == null
            ? new ExcelPivotTableSnapshot.Anchor("A1", "A1")
            : new ExcelPivotTableSnapshot.Anchor(
                location.topLeftAddress(), location.locationRange());
    if (location == null) {
      return unsupportedSnapshot(
          handle, name, anchor, "Pivot table location range is missing or malformed.");
    }
    try {
      CTPivotTableDefinition definition = handle.table().getCTPivotTableDefinition();
      List<String> sourceColumnNames = cacheFieldNames(handle.table());
      ColumnAxisSnapshot columns = snapshotColumnLabels(definition, sourceColumnNames);
      List<ExcelPivotTableSnapshot.DataField> dataFields =
          snapshotDataFields(workbook, definition, sourceColumnNames);
      if (dataFields.isEmpty()) {
        return unsupportedSnapshot(
            handle, name, anchor, "Pivot table does not contain any data fields.");
      }
      return new ExcelPivotTableSnapshot.Supported(
          name,
          handle.sheetName(),
          anchor,
          snapshotSource(workbook, handle.table()),
          snapshotFields(
              definition.getRowFields() == null ? null : definition.getRowFields().getFieldArray(),
              sourceColumnNames),
          columns.columnLabels(),
          snapshotPageFields(
              definition.getPageFields() == null
                  ? null
                  : definition.getPageFields().getPageFieldArray(),
              sourceColumnNames),
          dataFields,
          columns.valuesAxisOnColumns());
    } catch (RuntimeException exception) {
      return unsupportedSnapshot(
          handle,
          name,
          anchor,
          Objects.requireNonNullElse(
              exception.getMessage(),
              "The pivot table could not be normalized into the modeled contract."));
    }
  }

  static ColumnAxisSnapshot snapshotColumnLabels(
      CTPivotTableDefinition definition, List<String> sourceColumnNames) {
    boolean valuesAxisOnColumns = false;
    List<ExcelPivotTableSnapshot.Field> columnLabels = new ArrayList<>();
    if (definition.getColFields() != null) {
      for (CTField field : definition.getColFields().getFieldArray()) {
        if (field.getX() == -2) {
          valuesAxisOnColumns = true;
          continue;
        }
        columnLabels.add(sourceField(sourceColumnNames, field.getX()));
      }
    }
    return new ColumnAxisSnapshot(List.copyOf(columnLabels), valuesAxisOnColumns);
  }

  static List<ExcelPivotTableSnapshot.Field> snapshotFields(
      CTField[] fields, List<String> sourceColumnNames) {
    if (fields == null || fields.length == 0) {
      return List.of();
    }
    List<ExcelPivotTableSnapshot.Field> snapshots = new ArrayList<>(fields.length);
    for (CTField field : fields) {
      snapshots.add(sourceField(sourceColumnNames, field.getX()));
    }
    return List.copyOf(snapshots);
  }

  static List<ExcelPivotTableSnapshot.Field> snapshotPageFields(
      CTPageField[] pageFields, List<String> sourceColumnNames) {
    if (pageFields == null || pageFields.length == 0) {
      return List.of();
    }
    List<ExcelPivotTableSnapshot.Field> snapshots = new ArrayList<>(pageFields.length);
    for (CTPageField pageField : pageFields) {
      snapshots.add(sourceField(sourceColumnNames, pageField.getFld()));
    }
    return List.copyOf(snapshots);
  }

  static List<ExcelPivotTableSnapshot.DataField> snapshotDataFields(
      XSSFWorkbook workbook, CTPivotTableDefinition definition, List<String> sourceColumnNames) {
    if (definition.getDataFields() == null
        || definition.getDataFields().sizeOfDataFieldArray() == 0) {
      return List.of();
    }
    List<ExcelPivotTableSnapshot.DataField> dataFields = new ArrayList<>();
    for (CTDataField dataField : definition.getDataFields().getDataFieldArray()) {
      int sourceColumnIndex = Math.toIntExact(dataField.getFld());
      String sourceColumnName =
          ExcelPivotTableSourceSupport.sourceColumnName(sourceColumnNames, sourceColumnIndex);
      dataFields.add(
          new ExcelPivotTableSnapshot.DataField(
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

  static ExcelPivotTableSnapshot.Unsupported unsupportedSnapshot(
      PivotHandle handle, String name, ExcelPivotTableSnapshot.Anchor anchor, String detail) {
    return new ExcelPivotTableSnapshot.Unsupported(name, handle.sheetName(), anchor, detail);
  }

  static ExcelPivotTableSnapshot.Source snapshotSource(
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
      return new ExcelPivotTableSnapshot.Source.Range(
          sheetName, ExcelPivotTableIdentitySupport.normalizeArea(area));
    }

    String sourceName =
        ExcelPivotTableIdentitySupport.requireNonBlank(
            worksheetSource.getName(), "Pivot source name is missing.");
    List<Name> namedRanges =
        ExcelPivotTableSourceSupport.matchingNamedRanges(workbook, sourceName, sheetName);
    if (namedRanges.size() == 1) {
      AreaReference area = ExcelPivotTableSourceSupport.namedRangeArea(namedRanges.getFirst());
      return new ExcelPivotTableSnapshot.Source.NamedRange(
          sourceName,
          ExcelPivotTableSourceSupport.sourceSheetName(area, namedRanges.getFirst(), sheetName),
          ExcelPivotTableIdentitySupport.normalizeArea(area));
    }
    if (namedRanges.size() > 1) {
      throw new IllegalArgumentException(
          "Pivot source name '"
              + sourceName
              + "' is ambiguous because multiple matching named ranges exist.");
    }

    XSSFTable table = ExcelPivotTableSourceSupport.tableByName(workbook, sourceName, sheetName);
    if (table != null) {
      return new ExcelPivotTableSnapshot.Source.Table(
          sourceName,
          table.getSheetName(),
          ExcelPivotTableIdentitySupport.normalizeArea(
              new AreaReference(
                  table.getStartCellReference(),
                  table.getEndCellReference(),
                  SpreadsheetVersion.EXCEL2007)));
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

  static ExcelPivotTableSnapshot.Field sourceField(
      List<String> sourceColumnNames, int sourceColumnIndex) {
    return new ExcelPivotTableSnapshot.Field(
        sourceColumnIndex,
        ExcelPivotTableSourceSupport.sourceColumnName(sourceColumnNames, sourceColumnIndex));
  }

  static ExcelPivotDataConsolidateFunction fromSubtotal(int subtotalValue) {
    for (DataConsolidateFunction function : DataConsolidateFunction.values()) {
      if (function.getValue() == subtotalValue) {
        return ExcelPivotDataPoiBridge.fromPoi(function);
      }
    }
    throw new IllegalArgumentException(
        "Unsupported pivot data consolidate function value: " + subtotalValue);
  }

  static String numberFormat(XSSFWorkbook workbook, Long numFmtId) {
    if (numFmtId == null) {
      return null;
    }
    short formatIndex = numFmtId > Short.MAX_VALUE ? Short.MAX_VALUE : (short) numFmtId.longValue();
    String format = workbook.createDataFormat().getFormat(formatIndex);
    return format == null || format.isBlank() ? null : format;
  }

  static XSSFPivotCacheDefinition requiredCacheDefinition(XSSFPivotTable pivotTable) {
    XSSFPivotCacheDefinition cacheDefinition = cacheDefinition(pivotTable);
    if (cacheDefinition == null) {
      throw new IllegalArgumentException("Pivot table is missing its cache definition relation.");
    }
    return cacheDefinition;
  }

  static XSSFPivotCacheDefinition cacheDefinition(XSSFPivotTable pivotTable) {
    try {
      return pivotTable.getPivotCacheDefinition();
    } catch (RuntimeException exception) {
      return null;
    }
  }

  static XSSFPivotCacheRecords cacheRecords(XSSFPivotCacheDefinition cacheDefinition) {
    return firstRelation(cacheDefinition, XSSFPivotCacheRecords.class);
  }

  static <T extends POIXMLDocumentPart> T firstRelation(
      POIXMLDocumentPart parent, Class<T> relationType) {
    for (POIXMLDocumentPart relation : parent.getRelations()) {
      if (relationType.isInstance(relation)) {
        return relationType.cast(relation);
      }
    }
    return null;
  }

  static CTPivotCache workbookPivotCache(XSSFWorkbook workbook, long cacheId) {
    for (CTPivotCache cache : workbookPivotCaches(workbook)) {
      if (cache.getCacheId() == cacheId) {
        return cache;
      }
    }
    return null;
  }

  static List<CTPivotCache> workbookPivotCaches(XSSFWorkbook workbook) {
    if (!workbook.getCTWorkbook().isSetPivotCaches()) {
      return List.of();
    }
    return List.of(workbook.getCTWorkbook().getPivotCaches().getPivotCacheArray());
  }
}
