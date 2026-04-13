package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ConcurrentHashMap;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBookView;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedName;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSheetState;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorkbookDocument;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

/** Low-memory factual workbook reader backed by POI's XSSF event-model package access. */
public final class ExcelEventWorkbookReader {
  /** Executes supported introspection commands against one workbook path using the event model. */
  public List<WorkbookReadResult.Introspection> apply(
      Path workbookPath, Iterable<WorkbookReadCommand.Introspection> commands) throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    Objects.requireNonNull(commands, "commands must not be null");

    try (OPCPackage pkg = OPCPackage.open(workbookPath.toFile(), PackageAccess.READ)) {
      XSSFReader reader = new XSSFReader(pkg);
      EventWorkbookMetadata metadata = workbookMetadata(reader);
      Map<String, EventSheetSummary> sheetSummaries = new ConcurrentHashMap<>();
      List<WorkbookReadResult.Introspection> results = new ArrayList<>();
      for (WorkbookReadCommand.Introspection command : commands) {
        Objects.requireNonNull(command, "command must not contain nulls");
        results.add(applyOne(reader, metadata, sheetSummaries, command));
      }
      return List.copyOf(results);
    } catch (NotOfficeXmlFileException exception) {
      throw new IllegalArgumentException("Only .xlsx workbooks are supported", exception);
    } catch (OpenXML4JException | XmlException exception) {
      throw new IOException(
          "Failed to read workbook through the XSSF event model: " + workbookPath, exception);
    }
  }

  private WorkbookReadResult.Introspection applyOne(
      XSSFReader reader,
      EventWorkbookMetadata metadata,
      Map<String, EventSheetSummary> sheetSummaries,
      WorkbookReadCommand.Introspection command)
      throws IOException {
    return switch (command) {
      case WorkbookReadCommand.GetWorkbookSummary getWorkbookSummary ->
          new WorkbookReadResult.WorkbookSummaryResult(
              getWorkbookSummary.requestId(), workbookSummary(reader, metadata, sheetSummaries));
      case WorkbookReadCommand.GetSheetSummary getSheetSummary ->
          new WorkbookReadResult.SheetSummaryResult(
              getSheetSummary.requestId(),
              sheetSummary(reader, metadata, sheetSummaries, getSheetSummary.sheetName()));
      default ->
          throw new IllegalArgumentException(
              "executionMode.readMode=EVENT_READ supports GET_WORKBOOK_SUMMARY and"
                  + " GET_SHEET_SUMMARY only; unsupported read type: "
                  + commandType(command));
    };
  }

  private WorkbookReadResult.WorkbookSummary workbookSummary(
      XSSFReader reader,
      EventWorkbookMetadata metadata,
      Map<String, EventSheetSummary> sheetSummaries)
      throws IOException {
    if (metadata.sheets().isEmpty()) {
      return new WorkbookReadResult.WorkbookSummary.Empty(
          0, List.of(), metadata.namedRangeCount(), metadata.forceFormulaRecalculationOnOpen());
    }
    List<String> selectedSheetNames = new ArrayList<>();
    for (EventSheetReference sheet : metadata.sheets()) {
      if (sheetSnapshot(reader, metadata, sheetSummaries, sheet.name()).selected()) {
        selectedSheetNames.add(sheet.name());
      }
    }
    if (selectedSheetNames.isEmpty()) {
      selectedSheetNames = List.of(activeSheetName(metadata));
    }
    return new WorkbookReadResult.WorkbookSummary.WithSheets(
        metadata.sheets().size(),
        metadata.sheetNames(),
        activeSheetName(metadata),
        List.copyOf(selectedSheetNames),
        metadata.namedRangeCount(),
        metadata.forceFormulaRecalculationOnOpen());
  }

  private WorkbookReadResult.SheetSummary sheetSummary(
      XSSFReader reader,
      EventWorkbookMetadata metadata,
      Map<String, EventSheetSummary> sheetSummaries,
      String sheetName)
      throws IOException {
    EventSheetReference sheet = metadata.sheetByName().get(sheetName);
    EventSheetSummary summary = sheetSnapshot(reader, metadata, sheetSummaries, sheetName);
    return new WorkbookReadResult.SheetSummary(
        sheetName,
        sheet.visibility(),
        summary.protection(),
        summary.physicalRowCount(),
        summary.lastRowIndex(),
        summary.lastColumnIndex());
  }

  private EventSheetSummary sheetSnapshot(
      XSSFReader reader,
      EventWorkbookMetadata metadata,
      Map<String, EventSheetSummary> sheetSummaries,
      String sheetName)
      throws IOException {
    EventSheetReference sheet = metadata.sheetByName().get(sheetName);
    if (sheet == null) {
      throw new SheetNotFoundException(sheetName);
    }
    EventSheetSummary summary = sheetSummaries.get(sheetName);
    if (summary != null) {
      return summary;
    }
    EventSheetSummary scannedSummary = scanSheet(reader, sheet);
    sheetSummaries.put(sheetName, scannedSummary);
    return scannedSummary;
  }

  private static EventWorkbookMetadata workbookMetadata(XSSFReader reader)
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

  private static int activeSheetIndex(CTWorkbook workbook) {
    if (!workbook.isSetBookViews() || workbook.getBookViews().sizeOfWorkbookViewArray() == 0) {
      return 0;
    }
    CTBookView workbookView = workbook.getBookViews().getWorkbookViewArray(0);
    return workbookView.isSetActiveTab() ? (int) workbookView.getActiveTab() : 0;
  }

  private static int namedRangeCount(CTWorkbook workbook) {
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

  private static EventSheetReference sheetReference(CTSheet sheet) {
    return new EventSheetReference(
        sheet.getName(), sheet.getId(), visibility(sheet.isSetState() ? sheet.getState() : null));
  }

  private static ExcelSheetVisibility visibility(STSheetState.Enum state) {
    if (state == null || state == STSheetState.VISIBLE) {
      return ExcelSheetVisibility.VISIBLE;
    }
    if (state == STSheetState.HIDDEN) {
      return ExcelSheetVisibility.HIDDEN;
    }
    return ExcelSheetVisibility.VERY_HIDDEN;
  }

  private static String activeSheetName(EventWorkbookMetadata metadata) {
    if (metadata.sheets().isEmpty()) {
      throw new IllegalStateException("workbook metadata must contain at least one sheet");
    }
    int activeSheetIndex = metadata.activeSheetIndex();
    if (activeSheetIndex < 0 || activeSheetIndex >= metadata.sheets().size()) {
      return metadata.sheets().getFirst().name();
    }
    return metadata.sheets().get(activeSheetIndex).name();
  }

  private static EventSheetSummary scanSheet(XSSFReader reader, EventSheetReference sheet)
      throws IOException {
    SheetSummaryHandler handler = new SheetSummaryHandler();
    try (InputStream inputStream = reader.getSheet(sheet.relationshipId())) {
      XMLReader xmlReader = XMLHelper.newXMLReader();
      xmlReader.setContentHandler(handler);
      xmlReader.parse(new InputSource(inputStream));
      return handler.summary();
    } catch (IOException | SAXException exception) {
      throw new IOException("Failed to parse sheet " + sheet.name(), exception);
    } catch (Exception exception) {
      throw new IOException("Failed to parse sheet " + sheet.name(), exception);
    }
  }

  private static String commandType(WorkbookReadCommand.Introspection command) {
    return switch (command) {
      case WorkbookReadCommand.GetWorkbookSummary _ -> "GET_WORKBOOK_SUMMARY";
      case WorkbookReadCommand.GetWorkbookProtection _ -> "GET_WORKBOOK_PROTECTION";
      case WorkbookReadCommand.GetNamedRanges _ -> "GET_NAMED_RANGES";
      case WorkbookReadCommand.GetSheetSummary _ -> "GET_SHEET_SUMMARY";
      case WorkbookReadCommand.GetCells _ -> "GET_CELLS";
      case WorkbookReadCommand.GetWindow _ -> "GET_WINDOW";
      case WorkbookReadCommand.GetMergedRegions _ -> "GET_MERGED_REGIONS";
      case WorkbookReadCommand.GetHyperlinks _ -> "GET_HYPERLINKS";
      case WorkbookReadCommand.GetComments _ -> "GET_COMMENTS";
      case WorkbookReadCommand.GetDrawingObjects _ -> "GET_DRAWING_OBJECTS";
      case WorkbookReadCommand.GetCharts _ -> "GET_CHARTS";
      case WorkbookReadCommand.GetPivotTables _ -> "GET_PIVOT_TABLES";
      case WorkbookReadCommand.GetDrawingObjectPayload _ -> "GET_DRAWING_OBJECT_PAYLOAD";
      case WorkbookReadCommand.GetSheetLayout _ -> "GET_SHEET_LAYOUT";
      case WorkbookReadCommand.GetPrintLayout _ -> "GET_PRINT_LAYOUT";
      case WorkbookReadCommand.GetDataValidations _ -> "GET_DATA_VALIDATIONS";
      case WorkbookReadCommand.GetConditionalFormatting _ -> "GET_CONDITIONAL_FORMATTING";
      case WorkbookReadCommand.GetAutofilters _ -> "GET_AUTOFILTERS";
      case WorkbookReadCommand.GetTables _ -> "GET_TABLES";
      case WorkbookReadCommand.GetFormulaSurface _ -> "GET_FORMULA_SURFACE";
      case WorkbookReadCommand.GetSheetSchema _ -> "GET_SHEET_SCHEMA";
      case WorkbookReadCommand.GetNamedRangeSurface _ -> "GET_NAMED_RANGE_SURFACE";
    };
  }

  private record EventWorkbookMetadata(
      List<EventSheetReference> sheets,
      Map<String, EventSheetReference> sheetByName,
      int activeSheetIndex,
      int namedRangeCount,
      boolean forceFormulaRecalculationOnOpen) {
    private List<String> sheetNames() {
      return sheets.stream().map(EventSheetReference::name).toList();
    }
  }

  private record EventSheetReference(
      String name, String relationshipId, ExcelSheetVisibility visibility) {}

  private record EventSheetSummary(
      boolean selected,
      WorkbookReadResult.SheetProtection protection,
      int physicalRowCount,
      int lastRowIndex,
      int lastColumnIndex) {}

  /** SAX handler that extracts the factual sheet-summary surface supported by EVENT_READ. */
  private static final class SheetSummaryHandler extends DefaultHandler {
    private int physicalRowCount;
    private int lastRowIndex = -1;
    private int lastColumnIndex = -1;
    private int nextRowIndex;
    private boolean selected;
    private WorkbookReadResult.SheetProtection protection =
        new WorkbookReadResult.SheetProtection.Unprotected();

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) {
      String element = localName == null || localName.isEmpty() ? qName : localName;
      switch (element) {
        case "sheetView" -> selected = selected || booleanAttribute(attributes, "tabSelected");
        case "sheetProtection" -> protection = sheetProtection(attributes);
        case "row" -> handleRow(attributes);
        case "c" -> handleCell(attributes);
        case "col" -> handleColumn(attributes);
        default -> {}
      }
    }

    private void handleRow(Attributes attributes) {
      physicalRowCount++;
      String rowRef = attributes.getValue("r");
      int rowIndex = rowRef == null ? nextRowIndex : Integer.parseInt(rowRef) - 1;
      lastRowIndex = Math.max(lastRowIndex, rowIndex);
      nextRowIndex = rowIndex + 1;
    }

    private void handleCell(Attributes attributes) {
      String cellRef = attributes.getValue("r");
      if (cellRef == null || cellRef.isBlank()) {
        return;
      }
      lastColumnIndex = Math.max(lastColumnIndex, new CellReference(cellRef).getCol());
    }

    private void handleColumn(Attributes attributes) {
      String max = attributes.getValue("max");
      if (max == null || max.isBlank()) {
        return;
      }
      lastColumnIndex = Math.max(lastColumnIndex, Integer.parseInt(max) - 1);
    }

    private WorkbookReadResult.SheetProtection sheetProtection(Attributes attributes) {
      if (!booleanAttribute(attributes, "sheet")) {
        return new WorkbookReadResult.SheetProtection.Unprotected();
      }
      return new WorkbookReadResult.SheetProtection.Protected(
          new ExcelSheetProtectionSettings(
              booleanAttribute(attributes, "autoFilter"),
              booleanAttribute(attributes, "deleteColumns"),
              booleanAttribute(attributes, "deleteRows"),
              booleanAttribute(attributes, "formatCells"),
              booleanAttribute(attributes, "formatColumns"),
              booleanAttribute(attributes, "formatRows"),
              booleanAttribute(attributes, "insertColumns"),
              booleanAttribute(attributes, "insertHyperlinks"),
              booleanAttribute(attributes, "insertRows"),
              booleanAttribute(attributes, "objects"),
              booleanAttribute(attributes, "pivotTables"),
              booleanAttribute(attributes, "scenarios"),
              booleanAttribute(attributes, "selectLockedCells"),
              booleanAttribute(attributes, "selectUnlockedCells"),
              booleanAttribute(attributes, "sort")));
    }

    private static boolean booleanAttribute(Attributes attributes, String name) {
      String value = attributes.getValue(name);
      return "1".equals(value) || "true".equalsIgnoreCase(value);
    }

    private EventSheetSummary summary() {
      return new EventSheetSummary(
          selected, protection, physicalRowCount, lastRowIndex, lastColumnIndex);
    }
  }
}
