package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import java.util.List;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFPivotCache;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheDefinition;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;

/**
 * Shared helpers for pivot-table coverage slices.
 *
 * <p>This support surface intentionally centralizes terse nested fixtures and checked-exception
 * helpers so the individual coverage slices can stay focused on behavioral theory.
 */
@SuppressWarnings({"PMD.CommentRequired", "PMD.SignatureDeclareThrowsException"})
class ExcelPivotTableCoverageTestSupport {
  final ExcelPivotTableController controller = new ExcelPivotTableController();

  XSSFPivotTable pivotWorkbookWithSecondPivot(ExcelWorkbook workbook) {
    workbook.getOrCreateSheet("SecondReport");
    controller.setPivotTable(
        workbook,
        definition(
            "Second Pivot",
            "SecondReport",
            new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
            "C5",
            List.of(),
            List.of("Stage"),
            List.of()));
    return workbook.xssfWorkbook().getPivotTables().get(1);
  }

  ExcelWorkbook pivotWorkbook() throws Exception {
    ExcelWorkbook workbook = ExcelWorkbook.create();
    populatePivotSource(workbook, "Data");
    workbook.getOrCreateSheet("Report");
    controller.setPivotTable(
        workbook,
        definition(
            "Sales Pivot",
            "Report",
            new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
            "C5",
            List.of(),
            List.of("Region"),
            List.of()));
    return workbook;
  }

  ExcelPivotTableDefinition definition(
      String name,
      String sheetName,
      ExcelPivotTableDefinition.Source source,
      String anchor,
      List<String> reportFilters,
      List<String> rowLabels,
      List<String> columnLabels) {
    return new ExcelPivotTableDefinition(
        name,
        sheetName,
        source,
        new ExcelPivotTableDefinition.Anchor(anchor),
        rowLabels,
        columnLabels,
        reportFilters,
        List.of(
            new ExcelPivotTableDefinition.DataField(
                "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total Amount", "#,##0.00")));
  }

  void populatePivotSource(ExcelWorkbook workbook, String sheetName) {
    workbook
        .getOrCreateSheet(sheetName)
        .setRange(
            "A1:D5",
            List.of(
                List.of(
                    ExcelCellValue.text("Region"),
                    ExcelCellValue.text("Stage"),
                    ExcelCellValue.text("Owner"),
                    ExcelCellValue.text("Amount")),
                List.of(
                    ExcelCellValue.text("North"),
                    ExcelCellValue.text("Plan"),
                    ExcelCellValue.text("Ada"),
                    ExcelCellValue.number(10)),
                List.of(
                    ExcelCellValue.text("North"),
                    ExcelCellValue.text("Do"),
                    ExcelCellValue.text("Ada"),
                    ExcelCellValue.number(15)),
                List.of(
                    ExcelCellValue.text("South"),
                    ExcelCellValue.text("Plan"),
                    ExcelCellValue.text("Lin"),
                    ExcelCellValue.number(7)),
                List.of(
                    ExcelCellValue.text("South"),
                    ExcelCellValue.text("Do"),
                    ExcelCellValue.text("Lin"),
                    ExcelCellValue.number(12))));
  }

  List<String> fieldNames(List<ExcelPivotTableSnapshot.Field> fields) {
    return fields.stream().map(ExcelPivotTableSnapshot.Field::sourceColumnName).toList();
  }

  @SuppressWarnings("unchecked")
  List<Object> allPivotHandles(ExcelWorkbook workbook) throws Exception {
    return (List<Object>) invoke(controller, "allPivotTables", List.class, workbook);
  }

  Object newPivotHandle(
      int sheetIndex,
      int ordinalOnSheet,
      String sheetName,
      org.apache.poi.xssf.usermodel.XSSFSheet sheet,
      XSSFPivotTable table) {
    return new PivotHandle(sheetIndex, ordinalOnSheet, sheetName, sheet, table);
  }

  CTPivotTableDefinition pivotTableDefinition(String name, String locationRange, long cacheId) {
    CTPivotTableDefinition definition = CTPivotTableDefinition.Factory.newInstance();
    definition.setName(name);
    definition.addNewLocation().setRef(locationRange);
    definition.setCacheId(cacheId);
    return definition;
  }

  static <T> T invoke(Object target, String name, Class<T> returnType, Object... args)
      throws Exception {
    return returnType.cast(dispatch(target, name, args));
  }

  static void invokeVoid(Object target, String name, Object... args) throws Exception {
    dispatch(target, name, args);
  }

  static Object dispatch(Object target, String name, Object... args) throws Exception {
    if (target instanceof ExcelPivotTableController controller) {
      return switch (name) {
        case "actualName" -> ExcelPivotTableIdentitySupport.actualName((PivotHandle) args[0]);
        case "allPivotTables" -> controller.allPivotTables((ExcelWorkbook) args[0]);
        case "cacheDefinition" -> controller.cacheDefinition((XSSFPivotTable) args[0]);
        case "cacheDefinitionShared" ->
            controller.cacheDefinitionShared(
                (ExcelWorkbook) args[0], (PivotHandle) args[1], (XSSFPivotCacheDefinition) args[2]);
        case "cacheFieldNames" -> controller.cacheFieldNames((XSSFPivotTable) args[0]);
        case "cleanupPackagePartIfUnused" -> {
          controller.cleanupPackagePartIfUnused(
              (org.apache.poi.openxml4j.opc.OPCPackage) args[0],
              (org.apache.poi.openxml4j.opc.PackagePartName) args[1]);
          yield null;
        }
        case "contiguousArea" ->
            ExcelPivotTableIdentitySupport.contiguousArea((String) args[0], (String) args[1]);
        case "deletePivotHandle" -> {
          controller.deletePivotHandle((ExcelWorkbook) args[0], (PivotHandle) args[1]);
          yield null;
        }
        case "finding" ->
            controller.finding(
                (AnalysisFindingCode) args[0],
                (AnalysisSeverity) args[1],
                (PivotHandle) args[2],
                (String) args[3],
                (String) args[4],
                (List<String>) args[5]);
        case "firstRelation" ->
            controller.firstRelation(
                (POIXMLDocumentPart) args[0], (Class<? extends POIXMLDocumentPart>) args[1]);
        case "fromSubtotal" -> controller.fromSubtotal((Integer) args[0]);
        case "matchingNamedRanges" ->
            ExcelPivotTableSourceSupport.matchingNamedRanges(
                (XSSFWorkbook) args[0], (String) args[1], (String) args[2]);
        case "namedRangeArea" -> ExcelPivotTableSourceSupport.namedRangeArea((Name) args[0]);
        case "nonBlankOrDefault" ->
            ExcelPivotTableIdentitySupport.nonBlankOrDefault((String) args[0], (String) args[1]);
        case "normalizeCacheId" -> {
          controller.normalizeCacheId((XSSFWorkbook) args[0], (XSSFPivotTable) args[1]);
          yield null;
        }
        case "normalizeArea" ->
            ExcelPivotTableIdentitySupport.normalizeArea((AreaReference) args[0]);
        case "numberFormat" -> controller.numberFormat((XSSFWorkbook) args[0], (Long) args[1]);
        case "packagePartIndex" -> {
          if (args[0] instanceof PackagePart part) {
            yield controller.packagePartIndex(part, (String) args[1]);
          }
          yield controller.packagePartIndex((POIXMLDocumentPart) args[0], (String) args[1]);
        }
        case "pivotTableHealthFindings" ->
            controller.pivotTableHealthFindings((XSSFWorkbook) args[0], (PivotHandle) args[1]);
        case "pivotTableIdHighWaterMark" ->
            controller.pivotTableIdHighWaterMark((XSSFWorkbook) args[0]);
        case "primePivotTableAllocator" -> {
          controller.primePivotTableAllocator((XSSFWorkbook) args[0], (XSSFPivotTable) args[1]);
          yield null;
        }
        case "rawLocationRange" ->
            ExcelPivotTableIdentitySupport.rawLocationRange((PivotHandle) args[0]);
        case "removePoiRelation" ->
            controller.removePoiRelation(
                (POIXMLDocumentPart) args[0], (POIXMLDocumentPart) args[1]);
        case "removeWorkbookPivotCacheRegistration" -> {
          controller.removeWorkbookPivotCacheRegistration(
              (XSSFWorkbook) args[0], (Long) args[1], (String) args[2]);
          yield null;
        }
        case "requireNonBlank" ->
            ExcelPivotTableIdentitySupport.requireNonBlank((String) args[0], (String) args[1]);
        case "requiredCacheDefinition" ->
            controller.requiredCacheDefinition((XSSFPivotTable) args[0]);
        case "requiredTableByName" ->
            ExcelPivotTableSourceSupport.requiredTableByName(
                (ExcelWorkbook) args[0], (String) args[1]);
        case "resolvedName" -> ExcelPivotTableIdentitySupport.resolvedName((PivotHandle) args[0]);
        case "safeLocation" -> ExcelPivotTableIdentitySupport.safeLocation((PivotHandle) args[0]);
        case "sanitize" -> ExcelPivotTableIdentitySupport.sanitize((String) args[0]);
        case "selectHandlesByName" ->
            controller.selectHandlesByName((List<PivotHandle>) args[0], (List<String>) args[1]);
        case "snapshot" -> controller.snapshot((XSSFWorkbook) args[0], (PivotHandle) args[1]);
        case "snapshotColumnLabels" ->
            controller.snapshotColumnLabels(
                (CTPivotTableDefinition) args[0], (List<String>) args[1]);
        case "snapshotDataFields" ->
            controller.snapshotDataFields(
                (XSSFWorkbook) args[0], (CTPivotTableDefinition) args[1], (List<String>) args[2]);
        case "snapshotFields" ->
            controller.snapshotFields(
                args[0] == null ? null : (CTField[]) args[0], (List<String>) args[1]);
        case "snapshotPageFields" ->
            controller.snapshotPageFields(
                (org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageField[]) args[0],
                (List<String>) args[1]);
        case "snapshotSource" ->
            controller.snapshotSource((XSSFWorkbook) args[0], (XSSFPivotTable) args[1]);
        case "sourceColumnName" ->
            ExcelPivotTableSourceSupport.sourceColumnName(
                (List<String>) args[0], (Integer) args[1]);
        case "sourceColumns" ->
            ExcelPivotTableSourceSupport.sourceColumns(
                (org.apache.poi.xssf.usermodel.XSSFSheet) args[0],
                (AreaReference) args[1],
                (String) args[2]);
        case "sourceField" -> controller.sourceField((List<String>) args[0], (Integer) args[1]);
        case "sourceSheetName" ->
            ExcelPivotTableSourceSupport.sourceSheetName(
                (AreaReference) args[0], (Name) args[1], (String) args[2]);
        case "tableByName" ->
            ExcelPivotTableSourceSupport.tableByName(
                (XSSFWorkbook) args[0], (String) args[1], (String) args[2]);
        case "unsupportedSnapshot" ->
            controller.unsupportedSnapshot(
                (PivotHandle) args[0],
                (String) args[1],
                (ExcelPivotTableSnapshot.Anchor) args[2],
                (String) args[3]);
        case "cacheRecords" -> controller.cacheRecords((XSSFPivotCacheDefinition) args[0]);
        case "workbookPivotCache" ->
            controller.workbookPivotCache((XSSFWorkbook) args[0], (Long) args[1]);
        default -> throw new IllegalArgumentException("Unsupported helper invocation: " + name);
      };
    }
    if (target instanceof SourceColumns sourceColumns) {
      if ("relativeIndex".equals(name)) {
        return sourceColumns.relativeIndex((String) args[0]);
      }
      throw new IllegalArgumentException("Unsupported helper invocation: " + name);
    }
    if (target instanceof ColumnAxisSnapshot columnAxisSnapshot) {
      if ("columnLabels".equals(name)) {
        return columnAxisSnapshot.columnLabels();
      }
      if ("valuesAxisOnColumns".equals(name)) {
        return columnAxisSnapshot.valuesAxisOnColumns();
      }
      throw new IllegalArgumentException("Unsupported helper invocation: " + name);
    }
    throw new IllegalArgumentException(
        "Unsupported helper target: " + target.getClass().getName() + "#" + name);
  }

  static <T extends Throwable> T assertInvocationFailure(Class<T> type, ThrowingRunnable runnable) {
    return assertThrows(type, runnable::run);
  }

  @FunctionalInterface
  interface ThrowingRunnable {
    void run() throws Exception;
  }

  static final class NullDocumentPart extends POIXMLDocumentPart {}

  static final class SyntheticDocumentPart extends POIXMLDocumentPart {
    SyntheticDocumentPart(PackagePart packagePart) {
      super(packagePart);
    }
  }

  static final class ThrowingPivotTable extends XSSFPivotTable {
    final CTPivotTableDefinition definition;

    ThrowingPivotTable(CTPivotTableDefinition definition) {
      this.definition = definition;
    }

    @Override
    public CTPivotTableDefinition getCTPivotTableDefinition() {
      return definition;
    }

    @Override
    public XSSFPivotCacheDefinition getPivotCacheDefinition() {
      throw new IllegalStateException("broken cache relation");
    }
  }

  static final class NoCachePivotTable extends XSSFPivotTable {
    final CTPivotTableDefinition definition;

    NoCachePivotTable(CTPivotTableDefinition definition) {
      this.definition = definition;
    }

    @Override
    public CTPivotTableDefinition getCTPivotTableDefinition() {
      return definition;
    }

    @Override
    public org.apache.poi.xssf.usermodel.XSSFPivotCache getPivotCache() {
      return null;
    }
  }

  static final class CacheOnlyPivotTable extends XSSFPivotTable {
    final CTPivotTableDefinition definition;
    final XSSFPivotCache pivotCache;

    CacheOnlyPivotTable(CTPivotTableDefinition definition, XSSFPivotCache pivotCache) {
      this.definition = definition;
      this.pivotCache = pivotCache;
    }

    @Override
    public CTPivotTableDefinition getCTPivotTableDefinition() {
      return definition;
    }

    @Override
    public XSSFPivotCache getPivotCache() {
      return pivotCache;
    }

    @Override
    public XSSFPivotCacheDefinition getPivotCacheDefinition() {
      return null;
    }
  }

  static final class SyntheticPivotTable extends XSSFPivotTable {
    final CTPivotTableDefinition definition;
    final XSSFPivotCacheDefinition cacheDefinition;

    SyntheticPivotTable(
        CTPivotTableDefinition definition, XSSFPivotCacheDefinition cacheDefinition) {
      this.definition = definition;
      this.cacheDefinition = cacheDefinition;
    }

    @Override
    public CTPivotTableDefinition getCTPivotTableDefinition() {
      return definition;
    }

    @Override
    public XSSFPivotCacheDefinition getPivotCacheDefinition() {
      return cacheDefinition;
    }
  }

  static void removeChild(org.apache.xmlbeans.XmlObject xmlObject, String localName) {
    try (var cursor = xmlObject.newCursor()) {
      if (cursor.toChild("http://schemas.openxmlformats.org/spreadsheetml/2006/main", localName)) {
        cursor.removeXml();
      } else {
        fail("Expected child element not found: " + localName);
      }
    }
  }
}
