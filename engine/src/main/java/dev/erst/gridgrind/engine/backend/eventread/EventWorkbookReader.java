package dev.erst.gridgrind.engine.backend.eventread;

import dev.erst.gridgrind.core.plan.InspectionIntrospectionResult;
import dev.erst.gridgrind.core.result.SheetSummaryReport;
import dev.erst.gridgrind.core.result.WorkbookSummary;
import dev.erst.gridgrind.engine.backend.xssf.SheetNotFoundException;
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
public final class EventWorkbookReader {
  /** Executes supported introspection commands against one workbook path using the event model. */
  public List<InspectionIntrospectionResult> apply(
      Path workbookPath, Iterable<dev.erst.gridgrind.core.plan.ResolvedInspectionStep> commands)
      throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    Objects.requireNonNull(commands, "commands must not be null");

    try (OPCPackage pkg = openSourcePackage(workbookPath)) {
      XSSFReader reader = sourceReader(pkg, workbookPath);
      EventWorkbookMetadata metadata = sourceMetadata(reader, workbookPath);
      Map<String, EventSheetSummary> sheetSummaries = new ConcurrentHashMap<>();
      List<InspectionIntrospectionResult> results = new ArrayList<>();
      for (dev.erst.gridgrind.core.plan.ResolvedInspectionStep command : commands) {
        Objects.requireNonNull(command, "command must not contain nulls");
        results.add(applyOne(reader, metadata, sheetSummaries, command));
      }
      return List.copyOf(results);
    }
  }

  private static OPCPackage openSourcePackage(Path path) throws IOException {
    try {
      return OPCPackage.open(path.toFile(), PackageAccess.READ);
    } catch (NotOfficeXmlFileException
        | org.apache.poi.openxml4j.exceptions.InvalidFormatException failure) {
      throw new dev.erst.gridgrind.engine.backend.xssf.WorkbookNotOpenableException(path, failure);
    } catch (RuntimeException failure) {
      throw dev.erst.gridgrind.engine.backend.xssf.WorkbookSourceXmlRefusals.translate(failure);
    }
  }

  private static XSSFReader sourceReader(OPCPackage pkg, Path path) throws IOException {
    try {
      return new XSSFReader(pkg);
    } catch (org.apache.poi.openxml4j.exceptions.InvalidFormatException failure) {
      throw new dev.erst.gridgrind.engine.backend.xssf.WorkbookNotOpenableException(path, failure);
    } catch (OpenXML4JException | RuntimeException failure) {
      throw dev.erst.gridgrind.engine.backend.xssf.WorkbookSourceXmlRefusals.translate(failure);
    }
  }

  private static EventWorkbookMetadata sourceMetadata(XSSFReader reader, Path path)
      throws IOException {
    try {
      return EventWorkbookMetadata.workbookMetadata(reader);
    } catch (org.apache.poi.openxml4j.exceptions.InvalidFormatException failure) {
      throw new dev.erst.gridgrind.engine.backend.xssf.WorkbookNotOpenableException(path, failure);
    } catch (OpenXML4JException | XmlException | RuntimeException failure) {
      throw dev.erst.gridgrind.engine.backend.xssf.WorkbookSourceXmlRefusals.translate(failure);
    }
  }

  private InspectionIntrospectionResult applyOne(
      XSSFReader reader,
      EventWorkbookMetadata metadata,
      Map<String, EventSheetSummary> sheetSummaries,
      dev.erst.gridgrind.core.plan.ResolvedInspectionStep command)
      throws IOException {
    return switch (command.query()) {
      case dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery.GetWorkbookSummary _ ->
          new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookSummaryResult(
              command.stepId(), workbookSummary(reader, metadata, sheetSummaries));
      case dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetSheetSummary _ ->
          new dev.erst.gridgrind.core.plan.SheetInspectionResult.SheetSummaryResult(
              command.stepId(),
              sheetSummary(
                  reader,
                  metadata,
                  sheetSummaries,
                  ((dev.erst.gridgrind.core.plan.SheetSelector.ByName) command.target())
                      .name()
                      .value()));
      default ->
          throw new IllegalArgumentException(
              "execution.mode.type=EVENT_READ supports GET_WORKBOOK_SUMMARY and"
                  + " GET_SHEET_SUMMARY only; unsupported read type: "
                  + command.query().queryType());
    };
  }

  private WorkbookSummary workbookSummary(
      XSSFReader reader,
      EventWorkbookMetadata metadata,
      Map<String, EventSheetSummary> sheetSummaries)
      throws IOException {
    if (metadata.sheets().isEmpty()) {
      return new WorkbookSummary.Empty(
          0, List.of(), metadata.namedRangeCount(), metadata.forceFormulaRecalculationOnOpen());
    }
    List<String> selectedSheetNames = new ArrayList<>();
    for (EventSheetReference sheet : metadata.sheets()) {
      if (sheetSnapshot(reader, sheetSummaries, sheet).selected()) {
        selectedSheetNames.add(sheet.name());
      }
    }
    if (selectedSheetNames.isEmpty()) {
      selectedSheetNames = List.of(EventWorkbookMetadata.activeSheetName(metadata));
    }
    return new WorkbookSummary.WithSheets(
        metadata.sheets().size(),
        metadata.sheetNames().stream()
            .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
            .toList(),
        new dev.erst.gridgrind.core.vocabulary.SheetLabel(
            EventWorkbookMetadata.activeSheetName(metadata)),
        List.copyOf(selectedSheetNames).stream()
            .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
            .toList(),
        metadata.namedRangeCount(),
        metadata.forceFormulaRecalculationOnOpen());
  }

  private SheetSummaryReport sheetSummary(
      XSSFReader reader,
      EventWorkbookMetadata metadata,
      Map<String, EventSheetSummary> sheetSummaries,
      String sheetName)
      throws IOException {
    EventSheetReference sheet = metadata.sheetByName().get(sheetName);
    if (sheet == null) {
      throw new SheetNotFoundException(sheetName);
    }
    EventSheetSummary summary = sheetSnapshot(reader, sheetSummaries, sheet);
    return new SheetSummaryReport(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetName),
        sheet.visibility(),
        summary.protection(),
        summary.physicalRowCount(),
        summary.lastRowIndex(),
        summary.lastColumnIndex());
  }

  private EventSheetSummary sheetSnapshot(
      XSSFReader reader, Map<String, EventSheetSummary> sheetSummaries, EventSheetReference sheet)
      throws IOException {
    EventSheetSummary summary = sheetSummaries.get(sheet.name());
    if (summary != null) {
      return summary;
    }
    EventSheetSummary scannedSummary = scanSheet(reader, sheet);
    sheetSummaries.put(sheet.name(), scannedSummary);
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
    } catch (org.apache.poi.openxml4j.exceptions.InvalidFormatException failure) {
      throw new dev.erst.gridgrind.engine.backend.xssf.WorkbookContentRefusalException(
          dev.erst.gridgrind.core.problem.WorkbookContentIssue.INVALID_NATIVE_VALUE,
          java.util.Optional.of(
              new dev.erst.gridgrind.core.problem.WorkbookContentLocation(
                  java.util.Optional.of(
                      new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheet.name())),
                  java.util.Optional.empty(),
                  java.util.Optional.of("sheet.relationship"),
                  java.util.Optional.empty(),
                  java.util.Optional.empty())),
          failure);
    } catch (SAXException
        | javax.xml.parsers.ParserConfigurationException
        | RuntimeException failure) {
      throw dev.erst.gridgrind.engine.backend.xssf.WorkbookSourceXmlRefusals.translate(failure);
    }
  }
}
