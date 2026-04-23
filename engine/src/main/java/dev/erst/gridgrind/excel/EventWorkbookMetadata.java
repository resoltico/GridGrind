package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBookView;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedName;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSheetState;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorkbookDocument;

/** Workbook-scoped metadata collected once and reused by EVENT_READ commands. */
record EventWorkbookMetadata(
    List<EventSheetReference> sheets,
    Map<String, EventSheetReference> sheetByName,
    int activeSheetIndex,
    int namedRangeCount,
    boolean forceFormulaRecalculationOnOpen) {

  List<String> sheetNames() {
    return sheets.stream().map(EventSheetReference::name).toList();
  }

  static EventWorkbookMetadata workbookMetadata(XSSFReader reader)
      throws IOException, OpenXML4JException, XmlException {
    try (InputStream inputStream = reader.getWorkbookData()) {
      CTWorkbook workbook = WorkbookDocument.Factory.parse(inputStream).getWorkbook();
      List<EventSheetReference> sheets = new ArrayList<>();
      Map<String, EventSheetReference> sheetByName = new ConcurrentHashMap<>();
      if (workbook.getSheets() != null) {
        for (CTSheet sheet : workbook.getSheets().getSheetList()) {
          EventSheetReference reference = sheetReference(sheet);
          sheets.add(reference);
          sheetByName.put(reference.name(), reference);
        }
      }
      return new EventWorkbookMetadata(
          List.copyOf(sheets),
          Map.copyOf(sheetByName),
          activeSheetIndex(workbook),
          namedRangeCount(workbook),
          workbook.isSetCalcPr()
              && workbook.getCalcPr().isSetFullCalcOnLoad()
              && workbook.getCalcPr().getFullCalcOnLoad());
    }
  }

  static int activeSheetIndex(CTWorkbook workbook) {
    if (!workbook.isSetBookViews() || workbook.getBookViews().sizeOfWorkbookViewArray() == 0) {
      return 0;
    }
    CTBookView workbookView = workbook.getBookViews().getWorkbookViewArray(0);
    return workbookView.isSetActiveTab() ? (int) workbookView.getActiveTab() : 0;
  }

  static int namedRangeCount(CTWorkbook workbook) {
    if (!workbook.isSetDefinedNames()) {
      return 0;
    }
    int count = 0;
    for (CTDefinedName definedName : workbook.getDefinedNames().getDefinedNameList()) {
      boolean function = definedName.isSetFunction() && definedName.getFunction();
      boolean hidden = definedName.isSetHidden() && definedName.getHidden();
      if (ExcelWorkbook.shouldExpose(definedName.getName(), function, hidden)) {
        count++;
      }
    }
    return count;
  }

  static EventSheetReference sheetReference(CTSheet sheet) {
    return new EventSheetReference(
        sheet.getName(), sheet.getId(), visibility(sheet.isSetState() ? sheet.getState() : null));
  }

  static ExcelSheetVisibility visibility(STSheetState.Enum state) {
    if (state == null || state == STSheetState.VISIBLE) {
      return ExcelSheetVisibility.VISIBLE;
    }
    if (state == STSheetState.HIDDEN) {
      return ExcelSheetVisibility.HIDDEN;
    }
    return ExcelSheetVisibility.VERY_HIDDEN;
  }

  static String activeSheetName(EventWorkbookMetadata metadata) {
    if (metadata.sheets().isEmpty()) {
      throw new IllegalStateException("workbook metadata must contain at least one sheet");
    }
    int activeSheetIndex = metadata.activeSheetIndex();
    if (activeSheetIndex < 0 || activeSheetIndex >= metadata.sheets().size()) {
      return metadata.sheets().getFirst().name();
    }
    return metadata.sheets().get(activeSheetIndex).name();
  }
}
