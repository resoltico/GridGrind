package dev.erst.gridgrind.jazzer.support;

import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.nio.file.Path;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

/** Verifies canonical persisted OOXML picture relations inside saved `.xlsx` drawing parts. */
final class XlsxPicturePackageInvariantSupport {
  private static final String DRAWINGML_NAMESPACE =
      "http://schemas.openxmlformats.org/drawingml/2006/main";
  private static final String SPREADSHEETML_NAMESPACE =
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  private static final String RELATIONSHIP_NAMESPACE =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  private static final String PACKAGE_RELATIONSHIP_NAMESPACE =
      "http://schemas.openxmlformats.org/package/2006/relationships";

  private XlsxPicturePackageInvariantSupport() {}

  static void requireCanonicalPicturePackageState(Path workbookPath) throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    try (ZipFile zipFile = new ZipFile(workbookPath.toFile())) {
      Enumeration<? extends ZipEntry> entries = zipFile.entries();
      while (entries.hasMoreElements()) {
        ZipEntry entry = entries.nextElement();
        if (isDrawingEntry(entry)) {
          requireCanonicalDrawingEntry(zipFile, entry);
        }
        if (isWorksheetEntry(entry)) {
          requireCanonicalWorksheetEntry(zipFile, entry);
        }
      }
    }
  }

  private static boolean isDrawingEntry(ZipEntry entry) {
    return entry.getName().startsWith("xl/drawings/drawing") && entry.getName().endsWith(".xml");
  }

  private static boolean isWorksheetEntry(ZipEntry entry) {
    return entry.getName().startsWith("xl/worksheets/sheet") && entry.getName().endsWith(".xml");
  }

  private static void requireCanonicalDrawingEntry(ZipFile zipFile, ZipEntry drawingEntry)
      throws IOException {
    Document drawingDocument = parseXml(zipFile, drawingEntry, "drawing OOXML");
    Map<String, String> relationships = partRelationships(zipFile, drawingEntry);
    NodeList blips = drawingDocument.getElementsByTagNameNS(DRAWINGML_NAMESPACE, "blip");
    for (int index = 0; index < blips.getLength(); index++) {
      Element blip = (Element) blips.item(index);
      String embedId = blip.getAttributeNS(RELATIONSHIP_NAMESPACE, "embed");
      if (embedId == null || embedId.isBlank()) {
        continue;
      }
      String target = relationships.get(embedId);
      if (target == null) {
        throw new IllegalStateException(
            "picture refs must resolve in " + drawingEntry.getName() + ": " + embedId);
      }
      requireMediaTarget(
          zipFile,
          drawingEntry,
          embedId,
          resolvedTargetEntry(drawingEntry, target),
          "picture refs");
    }
  }

  private static void requireCanonicalWorksheetEntry(ZipFile zipFile, ZipEntry worksheetEntry)
      throws IOException {
    Document worksheetDocument = parseXml(zipFile, worksheetEntry, "worksheet OOXML");
    Map<String, String> relationships = partRelationships(zipFile, worksheetEntry);
    NodeList objectProperties =
        worksheetDocument.getElementsByTagNameNS(SPREADSHEETML_NAMESPACE, "objectPr");
    for (int index = 0; index < objectProperties.getLength(); index++) {
      Element objectPropertiesElement = (Element) objectProperties.item(index);
      String relationId = objectPropertiesElement.getAttributeNS(RELATIONSHIP_NAMESPACE, "id");
      if (relationId == null || relationId.isBlank()) {
        continue;
      }
      String target = relationships.get(relationId);
      if (target == null) {
        throw new IllegalStateException(
            "embedded object preview refs must resolve in "
                + worksheetEntry.getName()
                + ": "
                + relationId);
      }
      requireMediaTarget(
          zipFile,
          worksheetEntry,
          relationId,
          resolvedTargetEntry(worksheetEntry, target),
          "embedded object preview refs");
    }
  }

  private static Map<String, String> partRelationships(ZipFile zipFile, ZipEntry sourceEntry)
      throws IOException {
    String relsEntryName = relationshipEntryName(sourceEntry);
    ZipEntry relsEntry = zipFile.getEntry(relsEntryName);
    if (relsEntry == null) {
      return Map.of();
    }
    Document relsDocument = parseXml(zipFile, relsEntry, "drawing relationships OOXML");
    NodeList relationships =
        relsDocument.getElementsByTagNameNS(PACKAGE_RELATIONSHIP_NAMESPACE, "Relationship");
    Map<String, String> targets = new HashMap<>();
    for (int index = 0; index < relationships.getLength(); index++) {
      Element relationship = (Element) relationships.item(index);
      String targetMode = relationship.getAttribute("TargetMode");
      if ("External".equalsIgnoreCase(targetMode)) {
        continue;
      }
      String id = relationship.getAttribute("Id");
      String target = relationship.getAttribute("Target");
      if (id != null && !id.isBlank() && target != null && !target.isBlank()) {
        targets.put(id, target);
      }
    }
    return Map.copyOf(targets);
  }

  private static String relationshipEntryName(ZipEntry sourceEntry) {
    Path sourcePath = Path.of(sourceEntry.getName());
    Path parent = Objects.requireNonNull(sourcePath.getParent(), "sourceEntry parent must exist");
    return parent + "/_rels/" + sourcePath.getFileName() + ".rels";
  }

  private static void requireMediaTarget(
      ZipFile zipFile,
      ZipEntry sourceEntry,
      String relationId,
      String resolvedEntry,
      String relationDescription) {
    if (!resolvedEntry.startsWith("xl/media/")) {
      throw new IllegalStateException(
          relationDescription
              + " must target /xl/media parts in "
              + sourceEntry.getName()
              + ": "
              + relationId
              + " -> "
              + resolvedEntry);
    }
    if (zipFile.getEntry(resolvedEntry) == null) {
      throw new IllegalStateException(
          relationDescription
              + " must target existing media parts in "
              + sourceEntry.getName()
              + ": "
              + relationId
              + " -> "
              + resolvedEntry);
    }
  }

  private static String resolvedTargetEntry(ZipEntry drawingEntry, String target) {
    try {
      URI drawingUri = PackagingURIHelper.createPartName("/" + drawingEntry.getName()).getURI();
      URI targetUri = URI.create(target);
      URI resolved = PackagingURIHelper.resolvePartUri(drawingUri, targetUri);
      String entryName = resolved.getPath();
      return entryName.startsWith("/") ? entryName.substring(1) : entryName;
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException(
          "failed to resolve picture target '" + target + "' for " + drawingEntry.getName(),
          exception);
    }
  }

  private static Document parseXml(ZipFile zipFile, ZipEntry entry, String description)
      throws IOException {
    try (InputStream inputStream = zipFile.getInputStream(entry)) {
      DocumentBuilderFactory factory = DocumentBuilderFactory.newDefaultInstance();
      factory.setNamespaceAware(true);
      return factory.newDocumentBuilder().parse(inputStream);
    } catch (ParserConfigurationException | SAXException exception) {
      throw new IllegalStateException(
          description + " must parse successfully in " + entry.getName(), exception);
    }
  }
}
