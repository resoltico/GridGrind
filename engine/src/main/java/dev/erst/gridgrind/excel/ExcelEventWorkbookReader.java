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
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.xmlbeans.XmlException;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

/** Low-memory factual workbook reader backed by POI's XSSF event-model package access. */
public final class ExcelEventWorkbookReader {
  /** Executes supported introspection commands against one workbook path using the event model. */
  public List<WorkbookReadIntrospectionResult> apply(
      Path workbookPath, Iterable<WorkbookReadCommand.Introspection> commands) throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    Objects.requireNonNull(commands, "commands must not be null");

    try (OPCPackage pkg = OPCPackage.open(workbookPath.toFile(), PackageAccess.READ)) {
      XSSFReader reader = new XSSFReader(pkg);
      EventWorkbookMetadata metadata = EventWorkbookMetadata.workbookMetadata(reader);
      Map<String, EventSheetSummary> sheetSummaries = new ConcurrentHashMap<>();
      List<WorkbookReadIntrospectionResult> results = new ArrayList<>();
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

  private WorkbookReadIntrospectionResult applyOne(
      XSSFReader reader,
      EventWorkbookMetadata metadata,
      Map<String, EventSheetSummary> sheetSummaries,
      WorkbookReadCommand.Introspection command)
      throws IOException {
    return switch (command) {
      case WorkbookReadCommand.GetWorkbookSummary getWorkbookSummary ->
          new WorkbookCoreResult.WorkbookSummaryResult(
              getWorkbookSummary.stepId(), workbookSummary(reader, metadata, sheetSummaries));
      case WorkbookReadCommand.GetSheetSummary getSheetSummary ->
          new WorkbookSheetResult.SheetSummaryResult(
              getSheetSummary.stepId(),
              sheetSummary(reader, metadata, sheetSummaries, getSheetSummary.sheetName()));
      default ->
          throw new IllegalArgumentException(
              "executionMode.readMode=EVENT_READ supports GET_WORKBOOK_SUMMARY and"
                  + " GET_SHEET_SUMMARY only; unsupported read type: "
                  + EventReadCommandTypes.commandType(command));
    };
  }

  private WorkbookCoreResult.WorkbookSummary workbookSummary(
      XSSFReader reader,
      EventWorkbookMetadata metadata,
      Map<String, EventSheetSummary> sheetSummaries)
      throws IOException {
    if (metadata.sheets().isEmpty()) {
      return new WorkbookCoreResult.WorkbookSummary.Empty(
          0, List.of(), metadata.namedRangeCount(), metadata.forceFormulaRecalculationOnOpen());
    }
    List<String> selectedSheetNames = new ArrayList<>();
    for (EventSheetReference sheet : metadata.sheets()) {
      if (sheetSnapshot(reader, metadata, sheetSummaries, sheet.name()).selected()) {
        selectedSheetNames.add(sheet.name());
      }
    }
    if (selectedSheetNames.isEmpty()) {
      selectedSheetNames = List.of(EventWorkbookMetadata.activeSheetName(metadata));
    }
    return new WorkbookCoreResult.WorkbookSummary.WithSheets(
        metadata.sheets().size(),
        metadata.sheetNames(),
        EventWorkbookMetadata.activeSheetName(metadata),
        List.copyOf(selectedSheetNames),
        metadata.namedRangeCount(),
        metadata.forceFormulaRecalculationOnOpen());
  }

  private WorkbookSheetResult.SheetSummary sheetSummary(
      XSSFReader reader,
      EventWorkbookMetadata metadata,
      Map<String, EventSheetSummary> sheetSummaries,
      String sheetName)
      throws IOException {
    EventSheetReference sheet = metadata.sheetByName().get(sheetName);
    EventSheetSummary summary = sheetSnapshot(reader, metadata, sheetSummaries, sheetName);
    return new WorkbookSheetResult.SheetSummary(
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

  private static EventSheetSummary scanSheet(XSSFReader reader, EventSheetReference sheet)
      throws IOException {
    EventSheetSummaryHandler handler = new EventSheetSummaryHandler();
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
}
